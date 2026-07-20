// HTTP/CORS and cache configuration for the Cloudflare Worker.
const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, HEAD, PUT, POST, DELETE, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, X-User-Name, X-Pod-Token, X-State-Revision',
};

const SCRYFALL_AUTOCOMPLETE_CACHE_TTL_MS = 10 * 60 * 1000;
const SCRYFALL_CARD_CACHE_TTL_MS = 24 * 60 * 60 * 1000;
const AUTH_AUDIT_LOG_KEY = 'pod:default:auth-audit-log';
const AUTH_AUDIT_LOG_LIMIT = 500;
const SESSION_COOKIE_NAME = 'commanderSession';
const SESSION_TTL_SECONDS = 60 * 60 * 24 * 30;
const scryfallAutocompleteCache = new Map();
const scryfallCardCache = new Map();

// Response and normalization helpers.

function jsonResponse(body, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      'Content-Type': 'application/json',
      ...CORS_HEADERS,
      ...extraHeaders,
    },
  });
}

function getCookieValue(request, name) {
  const cookieHeader = getTextValue(request.headers.get('Cookie'));
  if (!cookieHeader || !name) {
    return '';
  }

  const pairs = cookieHeader.split(';');
  for (const pair of pairs) {
    const separatorIndex = pair.indexOf('=');
    if (separatorIndex < 0) {
      continue;
    }

    const key = pair.slice(0, separatorIndex).trim();
    if (key !== name) {
      continue;
    }

    return decodeURIComponent(pair.slice(separatorIndex + 1).trim());
  }

  return '';
}

function buildSessionCookie(value, { maxAge = SESSION_TTL_SECONDS, expires = '' } = {}) {
  const parts = [
    `${SESSION_COOKIE_NAME}=${encodeURIComponent(value || '')}`,
    'Path=/',
    'HttpOnly',
    'SameSite=Lax',
    'Secure',
  ];

  if (Number.isFinite(Number(maxAge))) {
    parts.push(`Max-Age=${Math.max(0, Number(maxAge))}`);
  }

  if (expires) {
    parts.push(`Expires=${expires}`);
  }

  return parts.join('; ');
}

function buildExpiredSessionCookie() {
  return buildSessionCookie('', {
    maxAge: 0,
    expires: 'Thu, 01 Jan 1970 00:00:00 GMT',
  });
}

function getSessionKey(request) {
  return getCookieValue(request, SESSION_COOKIE_NAME);
}

function getSessionStoreKey(sessionKey) {
  return sessionKey ? `pod:default:session:${sessionKey}` : '';
}

async function loadSessionAuth(request, env) {
  if (!env.POD_STATE) {
    return null;
  }

  const sessionKey = getSessionKey(request);
  if (!sessionKey) {
    return null;
  }

  const raw = await env.POD_STATE.get(getSessionStoreKey(sessionKey), 'json');
  if (!raw || typeof raw !== 'object') {
    return null;
  }

  const expiresAt = getTextValue(raw.expiresAt);
  if (!expiresAt || Number.isNaN(new Date(expiresAt).getTime()) || new Date(expiresAt).getTime() <= Date.now()) {
    await env.POD_STATE.delete(getSessionStoreKey(sessionKey));
    return null;
  }

  const auth = raw.auth && typeof raw.auth === 'object' ? raw.auth : null;
  if (!auth) {
    return null;
  }

  return {
    ok: true,
    user: getTextValue(auth.user),
    userId: getTextValue(auth.userId).toLowerCase(),
    displayName: getTextValue(auth.displayName || auth.user),
    role: getTextValue(auth.role || 'member').toLowerCase(),
    authMode: getTextValue(auth.authMode || 'session').toLowerCase(),
    sessionKey,
  };
}

async function persistSessionAuth(env, auth) {
  if (!env.POD_STATE || !auth) {
    return '';
  }

  const sessionKey = crypto.randomUUID?.() || `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  const expiresAt = new Date(Date.now() + (SESSION_TTL_SECONDS * 1000)).toISOString();
  await env.POD_STATE.put(getSessionStoreKey(sessionKey), JSON.stringify({
    auth: {
      user: getTextValue(auth.user),
      userId: getTextValue(auth.userId).toLowerCase(),
      displayName: getTextValue(auth.displayName || auth.user),
      role: getTextValue(auth.role || 'member').toLowerCase(),
      authMode: getTextValue(auth.authMode || 'session').toLowerCase(),
    },
    expiresAt,
  }), {
    expirationTtl: SESSION_TTL_SECONDS,
  });
  return sessionKey;
}

async function clearSessionAuth(request, env) {
  if (!env.POD_STATE) {
    return;
  }

  const sessionKey = getSessionKey(request);
  if (!sessionKey) {
    return;
  }

  await env.POD_STATE.delete(getSessionStoreKey(sessionKey));
}

function normalizeCommanderIdentity(identity) {
  const value = String(identity || '').trim().toLowerCase();
  if (!value) {
    return '';
  }

  if (value === 'c') {
    return 'c';
  }

  const order = ['w', 'u', 'b', 'r', 'g'];
  const normalized = order.filter((symbol) => value.includes(symbol));
  return normalized.length === value.length ? normalized.join('') : '';
}

function normalizeCommanderBuilderMode(mode) {
  const value = String(mode || '').trim().toLowerCase();
  return value === 'keywords' ? 'keywords' : 'identity';
}

function normalizeCommanderKeywords(keywords) {
  const rawValues = Array.isArray(keywords)
    ? keywords
    : String(keywords || '').split(',');

  return [...new Set(rawValues
    .map((value) => getTextValue(value))
    .filter(Boolean)
    .map((value) => value.replace(/\s+/g, ' ').trim()))];
}

function getCardImageUri(card) {
  if (card?.image_uris?.normal) {
    return card.image_uris.normal;
  }

  if (Array.isArray(card?.card_faces)) {
    const faceWithImage = card.card_faces.find((face) => face?.image_uris?.normal);
    if (faceWithImage?.image_uris?.normal) {
      return faceWithImage.image_uris.normal;
    }
  }

  return '';
}

function getCardImageVariant(card, size) {
  if (card?.image_uris?.[size]) {
    return getTextValue(card.image_uris[size]);
  }

  if (Array.isArray(card?.card_faces)) {
    const faceWithImage = card.card_faces.find((face) => face?.image_uris?.[size]);
    if (faceWithImage?.image_uris?.[size]) {
      return getTextValue(faceWithImage.image_uris[size]);
    }
  }

  return '';
}

function buildCommanderSearchQuery(identity) {
  return `game:paper is:commander id=${identity}`;
}

function buildCommanderKeywordSearchQuery(keywords, identity = '') {
  const normalizedKeywords = normalizeCommanderKeywords(keywords);
  if (!normalizedKeywords.length) {
    return '';
  }

  const clauses = normalizedKeywords.map((keyword) => `fo:"${keyword.replace(/"/g, '\\"')}"`);
  const normalizedIdentity = normalizeCommanderIdentity(identity);
  const identityClause = normalizedIdentity ? ` id=${normalizedIdentity}` : '';
  return `game:paper is:commander${identityClause} ${clauses.join(' ')}`;
}

// Card mapping helpers used by deck search, card detail, and commander endpoints.

function getDeckLookupKey(value) {
  return getTextValue(value)
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[^a-z0-9]+/g, '');
}

function mapDeckCard(card, requestOrigin) {
  const imageUri = getCardImageUri(card);
  const imageLargeUri = getCardImageVariant(card, 'large');
  const layout = getTextValue(card?.layout).toLowerCase();
  const typeLine = getTextValue(card?.type_line);
  const isToken = layout.includes('token') || /\btoken\b/i.test(typeLine);

  return {
    id: getTextValue(card?.id),
    oracleId: getTextValue(card?.oracle_id),
    name: getTextValue(card?.name),
    set: getTextValue(card?.set),
    setName: getTextValue(card?.set_name),
    collectorNumber: getTextValue(card?.collector_number),
    lang: getTextValue(card?.lang),
    artist: getTextValue(card?.artist),
    releasedAt: getTextValue(card?.released_at),
    manaCost: getTextValue(card?.mana_cost),
    typeLine: getTextValue(card?.type_line),
    oracleText: getTextValue(card?.oracle_text),
    scryfallUri: getTextValue(card?.scryfall_uri),
    imageUri: buildCommanderImageProxyUrl(imageUri, requestOrigin),
    imageLargeUri: buildCommanderImageProxyUrl(imageLargeUri, requestOrigin),
    cardFaces: getCardFaces(card, requestOrigin),
    colorIdentity: Array.isArray(card?.color_identity) ? card.color_identity : [],
    power: getTextValue(card?.power),
    toughness: getTextValue(card?.toughness),
    loyalty: getTextValue(card?.loyalty),
    defense: getTextValue(card?.defense),
    layout: getTextValue(card?.layout),
    isBanned: getTextValue(card?.legalities?.commander) === 'banned',
    isGameChanger: Boolean(card?.game_changer),
    isCommanderLegal: getTextValue(card?.legalities?.commander) === 'legal',
    isToken,
  };
}

function mapCommanderCard(card, requestOrigin) {
  const imageUri = getCardImageUri(card);
  const imageLargeUri = getCardImageVariant(card, 'large');
  const imagePngUri = getCardImageVariant(card, 'png');
  const buildImageProxyUrl = (source) => {
    const value = getTextValue(source);
    if (!value || !requestOrigin) {
      return '';
    }

    return `${requestOrigin}/api/commander-image?src=${encodeURIComponent(value)}`;
  };

  return {
    name: getTextValue(card?.name),
    manaCost: getTextValue(card?.mana_cost),
    typeLine: getTextValue(card?.type_line),
    colorIdentity: Array.isArray(card?.color_identity) ? card.color_identity : [],
    oracleText: getTextValue(card?.oracle_text),
    power: getTextValue(card?.power),
    toughness: getTextValue(card?.toughness),
    loyalty: getTextValue(card?.loyalty),
    defense: getTextValue(card?.defense),
    cardFaces: getCardFaces(card, requestOrigin),
    scryfallUri: getTextValue(card?.scryfall_uri),
    imageUri: buildImageProxyUrl(imageUri),
    imageLargeUri: buildImageProxyUrl(imageLargeUri),
    imagePngUri: buildImageProxyUrl(imagePngUri),
  };
}

function getTextValue(value) {
  return String(value || '').trim();
}

function normalizeMemberKey(value) {
  return getTextValue(value)
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[^a-z0-9]+/g, '');
}

const BUILTIN_ADMIN_USER_KEY = normalizeMemberKey('Baily');

function isBuiltInAdminUser(value) {
  return normalizeMemberKey(value) === BUILTIN_ADMIN_USER_KEY;
}

function getConfiguredMembers(env) {
  const raw = getTextValue(env.POD_MEMBERS_JSON);
  if (!raw) {
    return [];
  }

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }

    return parsed
      .map((member) => {
        const explicitUserId = getTextValue(member?.userId || member?.id || member?.username);
        const displayName = getTextValue(member?.displayName || member?.user || member?.name || explicitUserId);
        const userId = getTextValue(explicitUserId || displayName);
        const token = getTextValue(
          member?.token
          || member?.accessCode
          || member?.podAccessCode
          || member?.passcode
          || member?.password
          || member?.code
        );
        const role = (
          getTextValue(member?.role).toLowerCase() === 'admin'
          || isBuiltInAdminUser(userId)
          || isBuiltInAdminUser(displayName)
        ) ? 'admin' : 'member';
        if (!userId || !displayName || !token) {
          return null;
        }

        return {
          userId: normalizeMemberKey(userId),
          displayName,
          token,
          role,
          matchKeys: new Set([normalizeMemberKey(userId), normalizeMemberKey(displayName)]),
        };
      })
      .filter(Boolean);
  } catch (error) {
    return [];
  }
}

function buildCommanderImageProxyUrl(source, requestOrigin) {
  const value = getTextValue(source);
  if (!value || !requestOrigin) {
    return '';
  }

  return `${requestOrigin}/api/commander-image?src=${encodeURIComponent(value)}`;
}

function getCardFaces(card, requestOrigin) {
  if (!Array.isArray(card?.card_faces)) {
    return [];
  }

  return card.card_faces
    .map((face) => ({
      name: getTextValue(face?.name),
      manaCost: getTextValue(face?.mana_cost),
      typeLine: getTextValue(face?.type_line),
      oracleText: getTextValue(face?.oracle_text),
      imageUri: buildCommanderImageProxyUrl(face?.image_uris?.normal, requestOrigin),
      imageLargeUri: buildCommanderImageProxyUrl(face?.image_uris?.large, requestOrigin),
      imagePngUri: buildCommanderImageProxyUrl(face?.image_uris?.png, requestOrigin),
      power: getTextValue(face?.power),
      toughness: getTextValue(face?.toughness),
      loyalty: getTextValue(face?.loyalty),
      defense: getTextValue(face?.defense),
    }))
    .filter((face) => face.name || face.oracleText || face.typeLine || face.manaCost || face.power || face.toughness || face.loyalty || face.defense);
}

function getScryfallHeaders() {
  return {
    Accept: 'application/json;q=0.9,*/*;q=0.8',
    'User-Agent': 'CommanderTracker/1.0 (+https://github.com/bailyschmeltzer/Commander)',
  };
}

function getCachedValue(cache, key, ttlMs) {
  const entry = cache.get(key);
  if (!entry) {
    return null;
  }

  if (Date.now() - entry.storedAt > ttlMs) {
    return null;
  }

  return entry.value;
}

function getStaleCachedValue(cache, key) {
  const entry = cache.get(key);
  return entry ? entry.value : null;
}

function setCachedValue(cache, key, value) {
  cache.set(key, {
    value,
    storedAt: Date.now(),
  });
}

function getRetryAfterSeconds(response) {
  const retryAfter = Number.parseInt(String(response.headers.get('Retry-After') || '').trim(), 10);
  return Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter : 60;
}

// Scryfall fetch and cache orchestration.

async function fetchDeckSearchResults(query) {
  const cacheKey = getTextValue(query).toLowerCase();
  const cachedResults = getCachedValue(scryfallAutocompleteCache, cacheKey, SCRYFALL_AUTOCOMPLETE_CACHE_TTL_MS);
  if (cachedResults) {
    return cachedResults;
  }

  const autocompleteUrl = new URL('https://api.scryfall.com/cards/autocomplete');
  autocompleteUrl.searchParams.set('q', query);

  const response = await fetch(autocompleteUrl.toString(), {
    headers: getScryfallHeaders(),
  });

  if (!response.ok) {
    if (response.status === 429) {
      const fallbackResults = getStaleCachedValue(scryfallAutocompleteCache, cacheKey);
      if (fallbackResults) {
        return fallbackResults;
      }
      throw new Error(`Scryfall autocomplete is temporarily rate-limited. Try again in about ${getRetryAfterSeconds(response)} seconds.`);
    }

    const detail = await response.text();
    throw new Error(`Scryfall autocomplete request failed (${response.status}): ${detail}`);
  }

  const payload = await response.json();
  const results = Array.isArray(payload?.data)
    ? payload.data.map((value) => getTextValue(value)).filter(Boolean)
    : [];

  setCachedValue(scryfallAutocompleteCache, cacheKey, results);
  return results;
}

async function fetchTokenSearchResults(query) {
  const normalizedQuery = getTextValue(query);
  const tokenSearchUrl = new URL('https://api.scryfall.com/cards/search');
  tokenSearchUrl.searchParams.set('q', `game:paper t:token ${normalizedQuery}`);
  tokenSearchUrl.searchParams.set('order', 'name');
  tokenSearchUrl.searchParams.set('unique', 'cards');

  const response = await fetch(tokenSearchUrl.toString(), {
    headers: getScryfallHeaders(),
  });

  if (!response.ok) {
    if (response.status === 404) {
      return [];
    }
    if (response.status === 429) {
      throw new Error(`Scryfall token search is temporarily rate-limited. Try again in about ${getRetryAfterSeconds(response)} seconds.`);
    }

    const detail = await response.text();
    throw new Error(`Scryfall token search failed (${response.status}): ${detail}`);
  }

  const payload = await response.json();
  const cards = Array.isArray(payload?.data) ? payload.data : [];
  return cards
    .map((card) => getTextValue(card?.name))
    .filter(Boolean)
    .slice(0, 20);
}

async function fetchDeckCardPrints({ oracleId = '', name = '' }, requestOrigin) {
  const normalizedOracleId = getTextValue(oracleId);
  const normalizedName = getTextValue(name);
  if (!normalizedOracleId && !normalizedName) {
    return [];
  }

  const searchUrl = new URL('https://api.scryfall.com/cards/search');
  if (normalizedOracleId) {
    searchUrl.searchParams.set('q', `oracleid:${normalizedOracleId}`);
  } else {
    const escapedName = normalizedName.replace(/"/g, '\\"');
    searchUrl.searchParams.set('q', `!"${escapedName}"`);
  }
  searchUrl.searchParams.set('unique', 'prints');
  searchUrl.searchParams.set('order', 'released');
  searchUrl.searchParams.set('dir', 'desc');

  const prints = [];
  let nextPageUrl = searchUrl;
  while (nextPageUrl) {
    const response = await fetch(nextPageUrl.toString(), {
      headers: getScryfallHeaders(),
    });

    if (!response.ok) {
      if (response.status === 404) {
        return [];
      }
      if (response.status === 429) {
        throw new Error(`Scryfall print search is temporarily rate-limited. Try again in about ${getRetryAfterSeconds(response)} seconds.`);
      }

      const detail = await response.text();
      throw new Error(`Scryfall print search failed (${response.status}): ${detail}`);
    }

    const payload = await response.json();
    const pageCards = Array.isArray(payload?.data) ? payload.data : [];
    prints.push(...pageCards.map((card) => ({
      id: getTextValue(card?.id),
      oracleId: getTextValue(card?.oracle_id),
      name: getTextValue(card?.name),
      set: getTextValue(card?.set),
      setName: getTextValue(card?.set_name),
      collectorNumber: getTextValue(card?.collector_number),
      lang: getTextValue(card?.lang),
      artist: getTextValue(card?.artist),
      releasedAt: getTextValue(card?.released_at),
      scryfallUri: getTextValue(card?.scryfall_uri),
      imageUri: buildCommanderImageProxyUrl(getCardImageUri(card), requestOrigin),
      imageLargeUri: buildCommanderImageProxyUrl(getCardImageVariant(card, 'large'), requestOrigin),
      imagePngUri: buildCommanderImageProxyUrl(getCardImageVariant(card, 'png'), requestOrigin),
      cardFaces: getCardFaces(card, requestOrigin),
    })).filter((entry) => entry.id && (entry.imageUri || entry.imageLargeUri || entry.cardFaces.length)));

    nextPageUrl = payload?.has_more && payload?.next_page ? new URL(payload.next_page) : null;
  }

  return prints;
}

function normalizeSecretLairSearchName(value) {
  return getTextValue(value)
    .replace(/^secret lair commander deck\s*/i, '')
    .replace(/^secret lair\s*/i, '')
    .replace(/\s+foil edition\s*$/i, '')
    .replace(/[,:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function getSecretLairSearchAliases(value) {
  const normalized = normalizeSecretLairSearchName(value);
  if (!normalized) {
    return [];
  }

  const aliases = new Set();
  const pushAlias = (candidate) => {
    const key = getDeckLookupKey(candidate);
    if (key) {
      aliases.add(key);
    }
  };

  pushAlias(normalized);
  pushAlias(normalized.replace(/^secret\s+lair\s+commander\s+deck\s*/i, ''));
  pushAlias(normalized.replace(/^secret\s+lair\s*/i, ''));

  return [...aliases];
}

function buildDeckImportTextFromMtgjsonDeck(deckData) {
  const data = deckData && typeof deckData === 'object' ? deckData : {};
  const lines = [];
  const appendEntries = (entries) => {
    (Array.isArray(entries) ? entries : []).forEach((entry) => {
      const name = getTextValue(entry?.name);
      if (!name) {
        return;
      }

      const count = Math.max(1, Number.isFinite(Number(entry?.count)) ? Number(entry.count) : 1);
      const setCode = getTextValue(entry?.setCode || entry?.set).toLowerCase();
      lines.push(setCode ? `${count} ${name} (${setCode})` : `${count} ${name}`);
    });
  };

  const commanderLines = [];
  const deckLines = [];
  const pushCommander = (entries) => {
    const bufferStart = lines.length;
    appendEntries(entries);
    const added = lines.slice(bufferStart);
    commanderLines.push(...added);
    lines.length = bufferStart;
  };
  const pushDeck = (entries) => {
    const bufferStart = lines.length;
    appendEntries(entries);
    const added = lines.slice(bufferStart);
    deckLines.push(...added);
    lines.length = bufferStart;
  };

  pushCommander(data?.commander);
  pushCommander(data?.displayCommander);
  pushDeck(data?.mainBoard);
  pushDeck(data?.sideBoard);

  const output = [];
  if (commanderLines.length) {
    output.push('Commander:');
    output.push(...commanderLines);
    output.push('');
  }
  output.push('Deck:');
  output.push(...deckLines);
  return output.join('\n');
}

function decodeBingRedirectUrl(value) {
  const match = String(value || '').match(/u=a1(?<payload>[A-Za-z0-9_-]+)/i);
  const payload = match?.groups?.payload;
  if (!payload) {
    return '';
  }

  let normalized = payload.replace(/-/g, '+').replace(/_/g, '/');
  while (normalized.length % 4) {
    normalized += '=';
  }

  try {
    return atob(normalized);
  } catch (_error) {
    return '';
  }
}

async function fetchSecretLairBundleSourceText(bundleName) {
  const searchName = normalizeSecretLairSearchName(bundleName);
  if (!searchName) {
    throw new Error('A Secret Lair bundle name is required.');
  }

  const aliases = getSecretLairSearchAliases(bundleName);
  if (aliases.length) {
    try {
      const deckListResponse = await fetch('https://mtgjson.com/api/v5/DeckList.json', {
        headers: {
          'User-Agent': 'CommanderTracker/1.0 (+https://github.com/bailyschmeltzer/Commander)',
        },
      });

      if (deckListResponse.ok) {
        const deckListPayload = await deckListResponse.json();
        const entries = Array.isArray(deckListPayload?.data) ? deckListPayload.data : [];
        const commanderEntries = entries.filter((entry) => getTextValue(entry?.type).toLowerCase() === 'commander deck');
        const matchedEntry = commanderEntries.find((entry) => aliases.includes(getDeckLookupKey(entry?.name)));
        const matchedFileName = getTextValue(matchedEntry?.fileName);
        if (matchedFileName) {
          const deckResponse = await fetch(`https://mtgjson.com/api/v5/decks/${encodeURIComponent(matchedFileName)}.json`, {
            headers: {
              'User-Agent': 'CommanderTracker/1.0 (+https://github.com/bailyschmeltzer/Commander)',
            },
          });

          if (deckResponse.ok) {
            const deckPayload = await deckResponse.json();
            const text = buildDeckImportTextFromMtgjsonDeck(deckPayload?.data);
            if (text.trim()) {
              return text;
            }
          }
        }
      }
    } catch (_error) {
      // Fall through to search-based fallback.
    }
  }

  const searchQueries = [
    `site:mtggoldfish.com/deck/ "${searchName}" decklist`,
    `"${searchName}" decklist mtggoldfish`,
    `"${searchName}" mtggoldfish`,
  ];

  let deckUrl = '';
  for (const query of searchQueries) {
    const searchUrl = new URL('https://www.bing.com/search');
    searchUrl.searchParams.set('q', query);

    const response = await fetch(searchUrl.toString(), {
      headers: {
        'User-Agent': 'CommanderTracker/1.0 (+https://github.com/bailyschmeltzer/Commander)',
      },
    });

    if (!response.ok) {
      continue;
    }

    const html = await response.text();
    const candidateUrls = [];
    for (const match of html.matchAll(/u=a1(?<payload>[A-Za-z0-9_-]+)/g)) {
      const decodedUrl = decodeBingRedirectUrl(`u=a1${match.groups.payload}`);
      if (decodedUrl && /mtggoldfish\.com\/deck\/\d+/i.test(decodedUrl) && !candidateUrls.includes(decodedUrl)) {
        candidateUrls.push(decodedUrl);
      }
    }

    deckUrl = candidateUrls[0] || '';
    if (deckUrl) {
      break;
    }
  }

  const deckIdMatch = deckUrl.match(/\/deck\/(\d+)/i);
  const deckId = deckIdMatch?.[1] || '';
  if (!deckId) {
    throw new Error(`Could not find a public decklist for ${bundleName}.`);
  }

  const downloadUrl = `https://www.mtggoldfish.com/deck/download/${deckId}?output=mtggoldfish&type=tabletop`;
  const downloadResponse = await fetch(downloadUrl, {
    headers: {
      'User-Agent': 'CommanderTracker/1.0 (+https://github.com/bailyschmeltzer/Commander)',
    },
  });

  if (!downloadResponse.ok) {
    throw new Error(`Failed to load the decklist source (${downloadResponse.status}).`);
  }

  return await downloadResponse.text();
}

async function fetchDeckCardByPrint(setCode, collectorNumber, requestOrigin) {
  const normalizedSetCode = getTextValue(setCode).toLowerCase();
  const normalizedCollectorNumber = getTextValue(collectorNumber);
  const cacheKey = `print:${normalizedSetCode}:${normalizedCollectorNumber}`;
  const cachedCard = getCachedValue(scryfallCardCache, cacheKey, SCRYFALL_CARD_CACHE_TTL_MS);
  if (cachedCard) {
    return mapDeckCard(cachedCard, requestOrigin);
  }

  const printUrl = new URL(`https://api.scryfall.com/cards/${encodeURIComponent(normalizedSetCode)}/${encodeURIComponent(normalizedCollectorNumber)}`);
  const response = await fetch(printUrl.toString(), {
    headers: getScryfallHeaders(),
  });

  if (!response.ok) {
    if (response.status === 429) {
      const fallbackCard = getStaleCachedValue(scryfallCardCache, cacheKey);
      if (fallbackCard) {
        return mapDeckCard(fallbackCard, requestOrigin);
      }
      throw new Error(`Scryfall card lookup is temporarily rate-limited. Try again in about ${getRetryAfterSeconds(response)} seconds.`);
    }

    const detail = await response.text();
    throw new Error(`Scryfall card lookup failed (${response.status}): ${detail}`);
  }

  const payload = await response.json();
  setCachedValue(scryfallCardCache, cacheKey, payload);
  return mapDeckCard(payload, requestOrigin);
}

async function fetchDeckCardByName(name, requestOrigin) {
  const cacheKey = getTextValue(name).toLowerCase();
  const cachedCard = getCachedValue(scryfallCardCache, cacheKey, SCRYFALL_CARD_CACHE_TTL_MS);
  if (cachedCard) {
    return mapDeckCard(cachedCard, requestOrigin);
  }

  const namedUrl = new URL('https://api.scryfall.com/cards/named');
  namedUrl.searchParams.set('exact', name);

  let response = await fetch(namedUrl.toString(), {
    headers: getScryfallHeaders(),
  });

  if (response.status === 404) {
    namedUrl.searchParams.delete('exact');
    namedUrl.searchParams.set('fuzzy', name);
    response = await fetch(namedUrl.toString(), {
      headers: getScryfallHeaders(),
    });
  }

  if (!response.ok) {
    if (response.status === 429) {
      const fallbackCard = getStaleCachedValue(scryfallCardCache, cacheKey);
      if (fallbackCard) {
        return mapDeckCard(fallbackCard, requestOrigin);
      }
      throw new Error(`Scryfall card lookup is temporarily rate-limited. Try again in about ${getRetryAfterSeconds(response)} seconds.`);
    }

    const detail = await response.text();
    throw new Error(`Scryfall card lookup failed (${response.status}): ${detail}`);
  }

  const payload = await response.json();
  setCachedValue(scryfallCardCache, cacheKey, payload);
  return mapDeckCard(payload, requestOrigin);
}

async function fetchDeckCardsByNames(names, requestOrigin) {
  const requestedNames = Array.isArray(names)
    ? names.map((value) => getTextValue(value)).filter(Boolean)
    : [];

  if (!requestedNames.length) {
    return [];
  }

  const uniqueNames = [];
  const seenKeys = new Set();
  requestedNames.forEach((name) => {
    const key = getDeckLookupKey(name);
    if (!seenKeys.has(key)) {
      seenKeys.add(key);
      uniqueNames.push(name);
    }
  });

  const foundByNameKey = new Map();
  const unresolvedNames = [];

  uniqueNames.forEach((name) => {
    const key = getDeckLookupKey(name);
    const cachedCard = getCachedValue(scryfallCardCache, key, SCRYFALL_CARD_CACHE_TTL_MS);
    if (cachedCard) {
      foundByNameKey.set(key, mapDeckCard(cachedCard, requestOrigin));
      return;
    }
    unresolvedNames.push(name);
  });

  const CHUNK_SIZE = 75;
  for (let offset = 0; offset < unresolvedNames.length; offset += CHUNK_SIZE) {
    const batch = unresolvedNames.slice(offset, offset + CHUNK_SIZE);
    const response = await fetch('https://api.scryfall.com/cards/collection', {
      method: 'POST',
      headers: {
        ...getScryfallHeaders(),
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        identifiers: batch.map((name) => ({ name })),
      }),
    });

    if (!response.ok) {
      if (response.status === 429) {
        throw new Error(`Scryfall card bulk lookup is temporarily rate-limited. Try again in about ${getRetryAfterSeconds(response)} seconds.`);
      }

      const detail = await response.text();
      throw new Error(`Scryfall card bulk lookup failed (${response.status}): ${detail}`);
    }

    const payload = await response.json();
    const cards = Array.isArray(payload?.data) ? payload.data : [];
    cards.forEach((card) => {
      const mapped = mapDeckCard(card, requestOrigin);
      const nameKey = getDeckLookupKey(card?.name);
      if (!nameKey) {
        return;
      }

      foundByNameKey.set(nameKey, mapped);
      setCachedValue(scryfallCardCache, nameKey, card);
    });
  }

  return requestedNames.map((name) => ({
    name,
    card: foundByNameKey.get(getDeckLookupKey(name)) || null,
  }));
}

async function fetchCommanderCandidates(identity) {
  const cards = [];
  let nextPage = new URL('https://api.scryfall.com/cards/search');
  nextPage.searchParams.set('q', buildCommanderSearchQuery(identity));
  nextPage.searchParams.set('order', 'edhrec');
  nextPage.searchParams.set('unique', 'cards');

  while (nextPage) {
    const response = await fetch(nextPage.toString(), {
      headers: getScryfallHeaders(),
    });

    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`Scryfall request failed (${response.status}): ${detail}`);
    }

    const payload = await response.json();
    const pageCards = Array.isArray(payload?.data) ? payload.data : [];
    cards.push(...pageCards.map((card) => ({
      name: getTextValue(card?.name),
      manaCost: getTextValue(card?.mana_cost),
      typeLine: getTextValue(card?.type_line),
      colorIdentity: Array.isArray(card?.color_identity) ? card.color_identity : [],
      oracleText: getTextValue(card?.oracle_text),
      power: getTextValue(card?.power),
      toughness: getTextValue(card?.toughness),
      loyalty: getTextValue(card?.loyalty),
      defense: getTextValue(card?.defense),
      cardFaces: getCardFaces(card),
      scryfallUri: getTextValue(card?.scryfall_uri),
      imageUri: getCardImageUri(card),
      imageLargeUri: getCardImageVariant(card, 'large'),
      imagePngUri: getCardImageVariant(card, 'png'),
    })).filter((card) => card.name && card.scryfallUri));

    nextPage = payload?.has_more && payload?.next_page ? new URL(payload.next_page) : null;
  }

  return cards;
}

async function fetchCommanderSelectionFromSearch(identity, requestOrigin) {
  // Fetch page 1 to get the total count and the first page of cards
  const PAGE_SIZE = 175; // Scryfall default page size
  const searchUrl = new URL('https://api.scryfall.com/cards/search');
  searchUrl.searchParams.set('q', buildCommanderSearchQuery(identity));
  searchUrl.searchParams.set('order', 'edhrec');
  searchUrl.searchParams.set('unique', 'cards');
  searchUrl.searchParams.set('page', '1');

  const response = await fetch(searchUrl.toString(), {
    headers: getScryfallHeaders(),
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Scryfall request failed (${response.status}): ${detail}`);
  }

  const payload = await response.json();
  const totalCards = Number.isFinite(Number(payload?.total_cards)) ? Number(payload.total_cards) : 0;
  const pageCards = Array.isArray(payload?.data) ? payload.data : [];

  if (!pageCards.length) {
    return { totalCards, card: null };
  }

  // If there are more pages, pick a random page and fetch it
  const totalPages = Math.ceil(totalCards / PAGE_SIZE);
  let candidates = pageCards;

  if (totalPages > 1 && Math.random() > 0.5) {
    const randomPage = 2 + Math.floor(Math.random() * (totalPages - 1));
    const pageUrl = new URL('https://api.scryfall.com/cards/search');
    pageUrl.searchParams.set('q', buildCommanderSearchQuery(identity));
    pageUrl.searchParams.set('order', 'edhrec');
    pageUrl.searchParams.set('unique', 'cards');
    pageUrl.searchParams.set('page', String(randomPage));

    try {
      const pageResponse = await fetch(pageUrl.toString(), { headers: getScryfallHeaders() });
      if (pageResponse.ok) {
        const pagePaylod = await pageResponse.json();
        const pageData = Array.isArray(pagePaylod?.data) ? pagePaylod.data : [];
        if (pageData.length) {
          candidates = pageData;
        }
      }
    } catch (_) {
      // Fall back to page 1 candidates
    }
  }

  const raw = candidates[Math.floor(Math.random() * candidates.length)];
  return { totalCards, card: mapCommanderCard(raw, requestOrigin) };
}

async function fetchCommanderSelectionByKeywords(keywords, identity, requestOrigin) {
  const normalizedKeywords = normalizeCommanderKeywords(keywords);
  const normalizedIdentity = normalizeCommanderIdentity(identity);
  const searchQuery = buildCommanderKeywordSearchQuery(normalizedKeywords, normalizedIdentity);
  if (!searchQuery) {
    throw new Error('At least one keyword is required.');
  }

  const PAGE_SIZE = 175;
  const searchUrl = new URL('https://api.scryfall.com/cards/search');
  searchUrl.searchParams.set('q', searchQuery);
  searchUrl.searchParams.set('order', 'edhrec');
  searchUrl.searchParams.set('unique', 'cards');
  searchUrl.searchParams.set('page', '1');

  const response = await fetch(searchUrl.toString(), {
    headers: getScryfallHeaders(),
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Scryfall request failed (${response.status}): ${detail}`);
  }

  const payload = await response.json();
  const totalCards = Number.isFinite(Number(payload?.total_cards)) ? Number(payload.total_cards) : 0;
  const pageCards = Array.isArray(payload?.data) ? payload.data : [];

  if (!pageCards.length) {
    return { totalCards, card: null };
  }

  const totalPages = Math.ceil(totalCards / PAGE_SIZE);
  let candidates = pageCards;

  if (totalPages > 1 && Math.random() > 0.5) {
    const randomPage = 2 + Math.floor(Math.random() * (totalPages - 1));
    const pageUrl = new URL('https://api.scryfall.com/cards/search');
    pageUrl.searchParams.set('q', searchQuery);
    pageUrl.searchParams.set('order', 'edhrec');
    pageUrl.searchParams.set('unique', 'cards');
    pageUrl.searchParams.set('page', String(randomPage));

    try {
      const pageResponse = await fetch(pageUrl.toString(), { headers: getScryfallHeaders() });
      if (pageResponse.ok) {
        const pagePayload = await pageResponse.json();
        const pageData = Array.isArray(pagePayload?.data) ? pagePayload.data : [];
        if (pageData.length) {
          candidates = pageData;
        }
      }
    } catch (_) {
      // Fall back to page 1 candidates.
    }
  }

  const raw = candidates[Math.floor(Math.random() * candidates.length)];
  return {
    totalCards,
    card: mapCommanderCard(raw, requestOrigin),
    keywords: normalizedKeywords,
    identity: normalizedIdentity,
  };
}

function isAllowedCommanderImageSource(value) {
  if (!value) {
    return false;
  }

  try {
    const url = new URL(value);
    return url.protocol === 'https:' && url.hostname === 'cards.scryfall.io';
  } catch (error) {
    return false;
  }
}

function getRequestUser(request) {
  return (request.headers.get('X-User-Name') || '').trim();
}

function getRequestToken(request) {
  return (request.headers.get('X-Pod-Token') || '').trim();
}

function maskToken(token) {
  const value = getTextValue(token);
  if (!value) {
    return '';
  }

  if (value.length <= 4) {
    return '*'.repeat(value.length);
  }

  return `${'*'.repeat(value.length - 4)}${value.slice(-4)}`;
}

function getRequestIp(request) {
  const cfConnectingIp = getTextValue(request.headers.get('CF-Connecting-IP'));
  if (cfConnectingIp) {
    return cfConnectingIp;
  }

  const forwardedFor = getTextValue(request.headers.get('X-Forwarded-For'));
  return forwardedFor.split(',')[0]?.trim() || '';
}

function buildAuthAuditEntry({ request, url, auth = null, success = false, reason = '', status = 0 }) {
  const requestUser = getRequestUser(request);
  const normalizedUser = normalizeMemberKey(requestUser);

  return {
    id: crypto.randomUUID?.() || `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`,
    timestamp: new Date().toISOString(),
    success: Boolean(success),
    status: Number.isFinite(Number(status)) ? Number(status) : 0,
    reason: getTextValue(reason || (success ? 'Authenticated.' : 'Authentication failed.')),
    method: getTextValue(request.method).toUpperCase(),
    path: `${url.pathname}${url.search}`,
    user: getTextValue(requestUser),
    normalizedUser,
    tokenHint: maskToken(getRequestToken(request)),
    ip: getRequestIp(request),
    userAgent: getTextValue(request.headers.get('User-Agent')).slice(0, 180),
    auth: success ? buildAuthPayload(auth) : null,
  };
}

async function appendAuthAuditEntry(env, entry) {
  if (!env.POD_STATE || !entry || typeof entry !== 'object') {
    return;
  }

  try {
    const raw = await env.POD_STATE.get(AUTH_AUDIT_LOG_KEY, 'json');
    const existingLogs = Array.isArray(raw?.logs)
      ? raw.logs
      : (Array.isArray(raw) ? raw : []);
    const nextLogs = [entry, ...existingLogs].slice(0, AUTH_AUDIT_LOG_LIMIT);
    await env.POD_STATE.put(AUTH_AUDIT_LOG_KEY, JSON.stringify({
      logs: nextLogs,
      updatedAt: new Date().toISOString(),
    }));
  } catch (error) {
    // Do not block requests if audit logging fails.
  }
}

function buildAutoProvisionedAuth(user) {
  const normalizedUser = normalizeMemberKey(user);
  if (!normalizedUser) {
    return null;
  }

  const role = isBuiltInAdminUser(normalizedUser) ? 'admin' : 'member';

  return {
    ok: true,
    user: getTextValue(user),
    userId: normalizedUser,
    displayName: getTextValue(user),
    role,
    authMode: 'auto-provisioned',
  };
}

function getRegisteredPlayerKeysFromState(state) {
  const registeredKeys = new Set();
  const games = Array.isArray(state?.games) ? state.games : [];

  const addPlayerValue = (value) => {
    const normalized = normalizeMemberKey(value);
    if (normalized) {
      registeredKeys.add(normalized);
    }
  };

  games.forEach((game) => {
    (Array.isArray(game?.players) ? game.players : []).forEach((player) => {
      addPlayerValue(player);
    });

    (Array.isArray(game?.finishOrder) ? game.finishOrder : []).forEach((player) => {
      addPlayerValue(player);
    });

    (Array.isArray(game?.playerRows) ? game.playerRows : []).forEach((row) => {
      addPlayerValue(row?.player);
    });

    (Array.isArray(game?.playerCommanders) ? game.playerCommanders : []).forEach((entry) => {
      addPlayerValue(entry?.player);
    });
  });

  return registeredKeys;
}

function resolveRegisteredPlayerKey(normalizedUser, registeredPlayerKeys) {
  const candidate = normalizeMemberKey(normalizedUser);
  if (!candidate || !(registeredPlayerKeys instanceof Set)) {
    return '';
  }

  if (registeredPlayerKeys.has(candidate)) {
    return candidate;
  }

  const closeMatches = [...registeredPlayerKeys].filter((key) => key.startsWith(candidate) || candidate.startsWith(key));
  if (closeMatches.length === 1) {
    return closeMatches[0];
  }

  return '';
}

function getRegisteredPlayersFromState(state) {
  const registeredPlayersById = new Map();
  const games = Array.isArray(state?.games) ? state.games : [];

  const addPlayerValue = (value) => {
    const displayName = getTextValue(value);
    const normalized = normalizeMemberKey(displayName);
    if (!normalized) {
      return;
    }

    if (!registeredPlayersById.has(normalized)) {
      registeredPlayersById.set(normalized, displayName || normalized);
    }
  };

  games.forEach((game) => {
    (Array.isArray(game?.players) ? game.players : []).forEach((player) => {
      addPlayerValue(player);
    });

    (Array.isArray(game?.finishOrder) ? game.finishOrder : []).forEach((player) => {
      addPlayerValue(player);
    });

    (Array.isArray(game?.playerRows) ? game.playerRows : []).forEach((row) => {
      addPlayerValue(row?.player);
    });

    (Array.isArray(game?.playerCommanders) ? game.playerCommanders : []).forEach((entry) => {
      addPlayerValue(entry?.player);
    });
  });

  return registeredPlayersById;
}

function buildRegisteredAccounts(configuredMembers, state) {
  const accountsByUserId = new Map();
  const aliasToConfiguredUserId = new Map();

  (Array.isArray(configuredMembers) ? configuredMembers : []).forEach((member) => {
    const userId = getTextValue(member?.userId).toLowerCase();
    if (!userId) {
      return;
    }

    accountsByUserId.set(userId, {
      userId,
      displayName: getTextValue(member?.displayName || userId),
      role: getTextValue(member?.role || (isBuiltInAdminUser(userId) ? 'admin' : 'member')).toLowerCase(),
      fromConfiguredMembers: true,
      fromGameHistory: false,
    });

    if (member?.matchKeys && typeof member.matchKeys.forEach === 'function') {
      member.matchKeys.forEach((key) => {
        const normalizedKey = getTextValue(key).toLowerCase();
        if (normalizedKey) {
          aliasToConfiguredUserId.set(normalizedKey, userId);
        }
      });
    }
  });

  const registeredPlayers = getRegisteredPlayersFromState(state);
  registeredPlayers.forEach((displayName, userId) => {
    const canonicalUserId = aliasToConfiguredUserId.get(userId) || userId;
    const existing = accountsByUserId.get(canonicalUserId);
    if (existing) {
      existing.fromGameHistory = true;
      if (!existing.displayName || normalizeMemberKey(existing.displayName) === existing.userId) {
        existing.displayName = getTextValue(displayName || canonicalUserId);
      }
      return;
    }

    accountsByUserId.set(canonicalUserId, {
      userId: canonicalUserId,
      displayName: getTextValue(displayName || canonicalUserId),
      role: isBuiltInAdminUser(canonicalUserId) ? 'admin' : 'member',
      fromConfiguredMembers: false,
      fromGameHistory: true,
    });
  });

  return Array.from(accountsByUserId.values())
    .map((entry) => ({
      userId: entry.userId,
      displayName: entry.displayName,
      role: entry.role,
      source: entry.fromConfiguredMembers && entry.fromGameHistory
        ? 'Configured + history'
        : (entry.fromConfiguredMembers ? 'Configured member' : 'Game history'),
    }))
    .sort((a, b) => {
      const displayCompare = String(a.displayName || '').localeCompare(String(b.displayName || ''), undefined, {
        numeric: true,
        sensitivity: 'base',
      });
      if (displayCompare !== 0) {
        return displayCompare;
      }
      return String(a.userId || '').localeCompare(String(b.userId || ''), undefined, {
        numeric: true,
        sensitivity: 'base',
      });
    });
}

async function hasValidAuth(request, env) {
  const sessionAuth = await loadSessionAuth(request, env);
  if (sessionAuth?.ok) {
    return sessionAuth;
  }

  const user = getRequestUser(request);
  const token = getRequestToken(request);
  const configuredMembers = getConfiguredMembers(env);
  const configuredMemberTokens = new Set(
    configuredMembers
      .map((entry) => getTextValue(entry?.token))
      .filter(Boolean),
  );
  const stateKey = 'pod:default:state';

  if (!env.POD_STATE) {
    return { ok: false, reason: 'Server missing POD_STATE KV binding.' };
  }

  const rawState = await env.POD_STATE.get(stateKey, 'json');
  const registeredPlayerKeys = getRegisteredPlayerKeysFromState(rawState && typeof rawState === 'object' ? rawState : null);
  const normalizedUser = normalizeMemberKey(user);
  const resolvedRegisteredUserId = resolveRegisteredPlayerKey(normalizedUser, registeredPlayerKeys);
  const autoProvisioned = buildAutoProvisionedAuth(user);
  const configuredToken = (env.POD_ACCESS_TOKEN || '').trim();

  if (!user) {
    return { ok: false, reason: 'Display name is required.' };
  }
  if (!token) {
    return { ok: false, reason: 'Pod access code is required.' };
  }

  if (configuredMembers.length) {
    // Configured member credentials path.
    const member = configuredMembers.find((entry) => entry.matchKeys.has(normalizedUser));
    if (member) {
      if (token === member.token || (configuredToken && token === configuredToken)) {
        const role = (
          getTextValue(member?.role).toLowerCase() === 'admin'
          || isBuiltInAdminUser(user)
          || isBuiltInAdminUser(normalizedUser)
        ) ? 'admin' : 'member';

        return {
          ok: true,
          user: member.displayName,
          userId: member.userId,
          displayName: member.displayName,
          role,
          authMode: token === member.token ? 'member' : 'legacy',
        };
      }
    }

    // Auto-provisioned path for registered game-history players.
    if (autoProvisioned && resolvedRegisteredUserId && token === `commander-${resolvedRegisteredUserId}`) {
      return {
        ok: true,
        user: getTextValue(user),
        userId: resolvedRegisteredUserId,
        displayName: getTextValue(user),
        role: isBuiltInAdminUser(resolvedRegisteredUserId) ? 'admin' : 'member',
        authMode: 'auto-provisioned',
      };
    }

    // Legacy pod token path for registered game-history players.
    if (configuredToken && token === configuredToken) {
      if (!resolvedRegisteredUserId) {
        return { ok: false, reason: `Player "${user}" is not registered in game history.` };
      }

      return {
        ok: true,
        user,
        userId: resolvedRegisteredUserId,
        displayName: user,
        role: isBuiltInAdminUser(resolvedRegisteredUserId) ? 'admin' : 'member',
        authMode: 'legacy',
      };
    }

    // Shared configured-member token path for registered game-history players.
    if (configuredMemberTokens.has(token)) {
      if (!resolvedRegisteredUserId) {
        return { ok: false, reason: `Player "${user}" is not registered in game history.` };
      }

      return {
        ok: true,
        user,
        userId: resolvedRegisteredUserId,
        displayName: user,
        role: isBuiltInAdminUser(resolvedRegisteredUserId) ? 'admin' : 'member',
        authMode: 'legacy',
      };
    }

    // If user is a configured member, keep a specific invalid-code message.
    if (member) {
      return { ok: false, reason: `Incorrect pod access code for "${user}".` };
    }

    if (resolvedRegisteredUserId) {
      return { ok: false, reason: `Incorrect pod access code for "${user}".` };
    }

    return { ok: false, reason: `Username "${user}" not found in pod members and is not registered in game history.` };
  }

  // Auto-provisioned path with no configured member list.
  if (autoProvisioned && resolvedRegisteredUserId && token === `commander-${resolvedRegisteredUserId}`) {
    return {
      ok: true,
      user: getTextValue(user),
      userId: resolvedRegisteredUserId,
      displayName: getTextValue(user),
      role: isBuiltInAdminUser(resolvedRegisteredUserId) ? 'admin' : 'member',
      authMode: 'auto-provisioned',
    };
  }

  if (!configuredToken) {
    return { ok: false, reason: 'Server missing POD_ACCESS_TOKEN.' };
  }

  if (token !== configuredToken) {
    return { ok: false, reason: `Incorrect pod access code. (Tried: "${user}")` };
  }

  if (!resolvedRegisteredUserId) {
    return { ok: false, reason: `Player "${user}" is not registered in game history.` };
  }

  return {
    ok: true,
    user,
    userId: resolvedRegisteredUserId,
    displayName: user,
    role: isBuiltInAdminUser(resolvedRegisteredUserId) ? 'admin' : 'member',
    authMode: 'legacy',
  };
}

function buildAuthPayload(auth) {
  return {
    userId: getTextValue(auth?.userId).toLowerCase(),
    displayName: getTextValue(auth?.displayName || auth?.user),
    role: getTextValue(auth?.role || 'member').toLowerCase(),
    mode: getTextValue(auth?.authMode || 'legacy').toLowerCase(),
  };
}

function resolveOwnerUserIdFromMembers(ownerDisplayName, members) {
  if (!ownerDisplayName) {
    return '';
  }
  const key = normalizeMemberKey(ownerDisplayName);
  if (!key) return '';
  // Exact match against a configured member's userId or displayName.
  if (Array.isArray(members)) {
    const match = members.find((m) => m.matchKeys.has(key));
    if (match) return match.userId;
  }
  // Fall back to the normalized name itself (covers auto-provisioned users).
  return key;
}

function enforceDeckOwnership(currentDecks, nextDecks, auth, members) {
  const currentById = new Map(
    (Array.isArray(currentDecks) ? currentDecks : [])
      .map((deck) => [getTextValue(deck?.id), deck])
      .filter(([deckId]) => deckId),
  );
  const nextById = new Set();
  const normalizedDecks = [];
  const authUserId = getTextValue(auth?.userId).toLowerCase();
  const isAdmin = getTextValue(auth?.role).toLowerCase() === 'admin';

  for (const rawDeck of Array.isArray(nextDecks) ? nextDecks : []) {
    const deckId = getTextValue(rawDeck?.id);
    const currentDeck = currentById.get(deckId) || null;
    const currentOwnerUserId = normalizeMemberKey(currentDeck?.ownerUserId);
    const requestedOwnerUserId = normalizeMemberKey(rawDeck?.ownerUserId);
    const hasChanged = !currentDeck || JSON.stringify(currentDeck) !== JSON.stringify(rawDeck);

    if (deckId) {
      nextById.add(deckId);
    }

    if (hasChanged) {
      if (currentDeck && currentOwnerUserId && !isAdmin && currentOwnerUserId !== authUserId) {
        throw new Error(`Deck "${getTextValue(currentDeck?.name) || 'Untitled Deck'}" is locked to ${getTextValue(currentDeck?.owner) || 'its owner'}.`);
      }

      if (!currentDeck && requestedOwnerUserId && !isAdmin && requestedOwnerUserId !== authUserId) {
        throw new Error('New decks can only be assigned to the authenticated user.');
      }
    }

    // Always re-resolve ownerUserId from owner display name using the member list.
    // The lock check above already prevents unauthorized modifications.
    const resolvedOwnerUserId = resolveOwnerUserIdFromMembers(getTextValue(rawDeck?.owner), members);
    const nextOwnerUserId = resolvedOwnerUserId || requestedOwnerUserId || currentOwnerUserId || authUserId;
    normalizedDecks.push({
      ...rawDeck,
      ownerUserId: nextOwnerUserId,
      owner: getTextValue(rawDeck?.owner) || (nextOwnerUserId === authUserId ? getTextValue(auth?.displayName || auth?.user) : getTextValue(currentDeck?.owner)),
    });
  }

  currentById.forEach((currentDeck, deckId) => {
    if (nextById.has(deckId)) {
      return;
    }

    const currentOwnerUserId = normalizeMemberKey(currentDeck?.ownerUserId);
    if (currentOwnerUserId && !isAdmin && currentOwnerUserId !== authUserId) {
      throw new Error(`Deck "${getTextValue(currentDeck?.name) || 'Untitled Deck'}" is locked to ${getTextValue(currentDeck?.owner) || 'its owner'}.`);
    }
  });

  return normalizedDecks;
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const requestOrigin = url.origin;

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    if (url.pathname === '/api/commanders') {
      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      const mode = normalizeCommanderBuilderMode(url.searchParams.get('mode'));

      try {
        if (mode === 'keywords') {
          const keywords = normalizeCommanderKeywords(url.searchParams.getAll('keyword'));
          if (!keywords.length) {
            const fromCsv = normalizeCommanderKeywords(url.searchParams.get('keywords'));
            keywords.push(...fromCsv);
          }
          const identity = normalizeCommanderIdentity(url.searchParams.get('identity'));

          const normalizedKeywords = normalizeCommanderKeywords(keywords);
          if (!normalizedKeywords.length) {
            return jsonResponse({ error: 'At least one keyword is required.' }, 400);
          }

          const { totalCards, card } = await fetchCommanderSelectionByKeywords(normalizedKeywords, identity, requestOrigin);

          return jsonResponse({
            mode,
            keywords: normalizedKeywords,
            identity,
            totalCards,
            card,
          }, 200, {
            'Cache-Control': 'no-store, no-cache, must-revalidate',
          });
        }

        const identity = normalizeCommanderIdentity(url.searchParams.get('identity'));
        if (!identity) {
          return jsonResponse({ error: 'A valid exact color identity is required.' }, 400);
        }

        const { totalCards, card } = await fetchCommanderSelectionFromSearch(identity, requestOrigin);

        return jsonResponse({
          mode,
          identity,
          totalCards,
          card,
        }, 200, {
          'Cache-Control': 'no-store, no-cache, must-revalidate',
        });
      } catch (error) {
        return jsonResponse({
          error: 'Unable to load commanders from Scryfall right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    if (url.pathname === '/api/commander-image') {
      if (request.method !== 'GET') {
        return new Response('Method not allowed.', { status: 405, headers: CORS_HEADERS });
      }

      const src = getTextValue(url.searchParams.get('src'));
      if (!isAllowedCommanderImageSource(src)) {
        return new Response('Invalid image source.', { status: 400, headers: CORS_HEADERS });
      }

      const imageResponse = await fetch(src, {
        headers: {
          Accept: 'image/avif,image/webp,image/png,image/jpeg,image/*;q=0.8,*/*;q=0.5',
          'User-Agent': 'CommanderTracker/1.0 (+https://github.com/bailyschmeltzer/Commander)',
        },
      });

      if (!imageResponse.ok) {
        return new Response('Unable to load image.', { status: 502, headers: CORS_HEADERS });
      }

      return new Response(imageResponse.body, {
        status: 200,
        headers: {
          ...CORS_HEADERS,
          'Content-Type': imageResponse.headers.get('Content-Type') || 'image/jpeg',
          'Cache-Control': 'public, max-age=604800',
        },
      });
    }

    if (url.pathname === '/api/deck-search') {
      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      const query = getTextValue(url.searchParams.get('q'));
      if (query.length < 3) {
        return jsonResponse({ results: [] }, 200, {
          'Cache-Control': 'no-store, no-cache, must-revalidate',
        });
      }

      try {
        const cfCache = caches.default;
        const cacheKey = new Request(request.url, { method: 'GET' });
        const cachedResponse = await cfCache.match(cacheKey);
        if (cachedResponse) {
          return cachedResponse;
        }
        const results = await fetchDeckSearchResults(query);
        const response = jsonResponse({ results: results.slice(0, 12) }, 200, {
          'Cache-Control': 'public, max-age=300, s-maxage=300',
        });
        await cfCache.put(cacheKey, response.clone());
        return response;
      } catch (error) {
        return jsonResponse({
          error: 'Unable to search cards right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    if (url.pathname === '/api/token-search') {
      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      const query = getTextValue(url.searchParams.get('q'));
      if (query.length < 3) {
        return jsonResponse({ results: [] }, 200, {
          'Cache-Control': 'no-store, no-cache, must-revalidate',
        });
      }

      try {
        const cfCache = caches.default;
        const cacheKey = new Request(request.url, { method: 'GET' });
        const cachedResponse = await cfCache.match(cacheKey);
        if (cachedResponse) {
          return cachedResponse;
        }
        const results = await fetchTokenSearchResults(query);
        const response = jsonResponse({ results }, 200, {
          'Cache-Control': 'public, max-age=300, s-maxage=300',
        });
        await cfCache.put(cacheKey, response.clone());
        return response;
      } catch (error) {
        return jsonResponse({
          error: 'Unable to search token cards right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    if (url.pathname === '/api/deck-card-arts') {
      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      const oracleId = getTextValue(url.searchParams.get('oracleId'));
      const name = getTextValue(url.searchParams.get('name'));
      if (!oracleId && !name) {
        return jsonResponse({ error: 'oracleId or name is required.' }, 400);
      }

      try {
        const cfCache = caches.default;
        const cacheKey = new Request(request.url, { method: 'GET' });
        const cachedResponse = await cfCache.match(cacheKey);
        if (cachedResponse) {
          return cachedResponse;
        }
        const prints = await fetchDeckCardPrints({ oracleId, name }, requestOrigin);
        const response = jsonResponse({ prints }, 200, {
          'Cache-Control': 'public, max-age=86400, s-maxage=86400',
        });
        await cfCache.put(cacheKey, response.clone());
        return response;
      } catch (error) {
        return jsonResponse({
          error: 'Unable to load card arts right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    if (url.pathname === '/api/secret-lair-list') {
      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      const name = getTextValue(url.searchParams.get('name'));
      if (!name) {
        return jsonResponse({ error: 'A Secret Lair bundle name is required.' }, 400);
      }

      try {
        const cfCache = caches.default;
        const cacheKey = new Request(request.url, { method: 'GET' });
        const cachedResponse = await cfCache.match(cacheKey);
        if (cachedResponse) {
          return cachedResponse;
        }

        const text = await fetchSecretLairBundleSourceText(name);
        const response = jsonResponse({ text }, 200, {
          'Cache-Control': 'public, max-age=86400, s-maxage=86400',
        });
        await cfCache.put(cacheKey, response.clone());
        return response;
      } catch (error) {
        return jsonResponse({
          error: 'Unable to load that Secret Lair bundle right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    if (url.pathname === '/api/deck-card') {
      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      const name = getTextValue(url.searchParams.get('name'));
      const setCode = getTextValue(url.searchParams.get('set')).toLowerCase();
      const collectorNumber = getTextValue(url.searchParams.get('collector'));
      const hasPrintSelector = Boolean(setCode && collectorNumber);
      if (!name && !hasPrintSelector) {
        return jsonResponse({ error: 'A card name or set+collector is required.' }, 400);
      }

      const cfCache = caches.default;
      const cacheKey = new Request(request.url, { method: 'GET' });
      const cachedResponse = await cfCache.match(cacheKey);
      if (cachedResponse) {
        return cachedResponse;
      }

      try {
        let card;
        if (hasPrintSelector) {
          card = await fetchDeckCardByPrint(setCode, collectorNumber, requestOrigin);
        } else if (setCode) {
          const prints = await fetchDeckCardPrints({ name }, requestOrigin);
          const exactPrint = prints.find((print) => getTextValue(print?.set).toLowerCase() === setCode) || prints[0];
          if (!exactPrint?.collectorNumber) {
            throw new Error(`Could not find a print of "${name}" in set ${setCode.toUpperCase()}.`);
          }
          card = await fetchDeckCardByPrint(getTextValue(exactPrint.set).toLowerCase(), exactPrint.collectorNumber, requestOrigin);
        } else {
          card = await fetchDeckCardByName(name, requestOrigin);
        }
        const response = jsonResponse({ card }, 200, {
          'Cache-Control': 'public, max-age=86400, s-maxage=86400',
        });
        await cfCache.put(cacheKey, response.clone());
        return response;
      } catch (error) {
        return jsonResponse({
          error: 'Unable to load that card right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    if (url.pathname === '/api/deck-cards-bulk') {
      if (request.method !== 'POST') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      let body;
      try {
        body = await request.json();
      } catch (error) {
        return jsonResponse({ error: 'Invalid JSON payload.' }, 400);
      }

      const names = Array.isArray(body?.names) ? body.names : [];
      if (!names.length) {
        return jsonResponse({ cards: [] }, 200, {
          'Cache-Control': 'no-store, no-cache, must-revalidate',
        });
      }

      try {
        const cards = await fetchDeckCardsByNames(names, requestOrigin);
        return jsonResponse({ cards }, 200, {
          'Cache-Control': 'public, max-age=86400, s-maxage=86400',
        });
      } catch (error) {
        return jsonResponse({
          error: 'Unable to load card batch right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    if (url.pathname === '/api/state') {
      const shouldAuditStateAuth = request.method === 'GET';
      const auth = await hasValidAuth(request, env);
      if (!auth.ok) {
        if (shouldAuditStateAuth) {
          await appendAuthAuditEntry(env, buildAuthAuditEntry({
            request,
            url,
            auth,
            success: false,
            reason: auth.reason,
            status: 401,
          }));
        }
        return jsonResponse({ error: auth.reason }, 401);
      }

      if (shouldAuditStateAuth) {
        await appendAuthAuditEntry(env, buildAuthAuditEntry({
          request,
          url,
          auth,
          success: true,
          reason: 'Authenticated.',
          status: 200,
        }));
      }

      const stateKey = 'pod:default:state';

      if (request.method === 'GET') {
        const raw = await env.POD_STATE.get(stateKey, 'json');
        const state = raw && typeof raw === 'object' ? raw : { games: [], powerLevels: {}, deckLists: [], decks: [], records: [], activeGame: null, activeGameUndo: [] };
        state.deckLists = Array.isArray(state.deckLists) ? state.deckLists : [];
        state.decks = Array.isArray(state.decks) ? state.decks : [];
        state.records = Array.isArray(state.records) ? state.records : [];
        state.activeGame = state.activeGame && typeof state.activeGame === 'object' ? state.activeGame : null;
        state.activeGameUndo = Array.isArray(state.activeGameUndo) ? state.activeGameUndo : [];
        state.revision = Number.isFinite(Number(state.revision)) ? Number(state.revision) : 0;
        state.updatedAt = String(state.updatedAt || '').trim();
        state.updatedBy = String(state.updatedBy || '').trim();

        if (url.searchParams.get('meta') === '1') {
          return jsonResponse({
            revision: state.revision,
            updatedAt: state.updatedAt,
            updatedBy: state.updatedBy,
            auth: buildAuthPayload(auth),
          }, 200);
        }

        return jsonResponse({
          ...state,
          auth: buildAuthPayload(auth),
        }, 200);
      }

      if (request.method === 'PUT') {
        let body;
        try {
          body = await request.json();
        } catch (error) {
          return jsonResponse({ error: 'Invalid JSON payload.' }, 400);
        }

        const hasRequiredPayloadShape = body && typeof body === 'object'
          && Object.prototype.hasOwnProperty.call(body, 'games')
          && Object.prototype.hasOwnProperty.call(body, 'powerLevels')
          && Object.prototype.hasOwnProperty.call(body, 'deckLists')
          && Object.prototype.hasOwnProperty.call(body, 'decks')
          && Object.prototype.hasOwnProperty.call(body, 'records')
          && Object.prototype.hasOwnProperty.call(body, 'activeGame')
          && Object.prototype.hasOwnProperty.call(body, 'activeGameUndo');

        if (!hasRequiredPayloadShape) {
          return jsonResponse({
            error: 'Sync payload is incomplete. Refresh the page and reconnect before syncing again.',
          }, 409);
        }

        const games = Array.isArray(body.games) ? body.games : null;
        const powerLevels = body.powerLevels && typeof body.powerLevels === 'object' && !Array.isArray(body.powerLevels)
          ? body.powerLevels
          : null;
        const deckLists = Array.isArray(body.deckLists) ? body.deckLists : null;
        const decks = Array.isArray(body.decks) ? body.decks : null;
        const records = Array.isArray(body.records) ? body.records : null;
        const activeGame = body.activeGame && typeof body.activeGame === 'object' ? body.activeGame : (body.activeGame === null ? null : undefined);
        const activeGameUndo = Array.isArray(body.activeGameUndo) ? body.activeGameUndo : null;

        if (!games || !powerLevels || !deckLists || !decks || !records || typeof activeGame === 'undefined' || !activeGameUndo) {
          return jsonResponse({
            error: 'Sync payload has invalid field types. Refresh the page and reconnect before syncing again.',
          }, 409);
        }

        const expectedRevisionHeader = request.headers.get('X-State-Revision');
        const expectedRevision = Number.parseInt(String(expectedRevisionHeader || '').trim(), 10);

        if (!Number.isFinite(expectedRevision)) {
          return jsonResponse({
            error: 'Sync client is missing cloud revision metadata. Refresh the page and reconnect before syncing again.',
          }, 409);
        }

        const currentRaw = await env.POD_STATE.get(stateKey, 'json');
        const currentState = currentRaw && typeof currentRaw === 'object' ? currentRaw : null;
        const currentRevision = Number.isFinite(Number(currentState?.revision)) ? Number(currentState.revision) : 0;
        const currentGames = Array.isArray(currentState?.games) ? currentState.games : [];
        const currentPowerLevels = currentState?.powerLevels && typeof currentState.powerLevels === 'object' && !Array.isArray(currentState.powerLevels)
          ? currentState.powerLevels
          : {};
        const currentDeckLists = Array.isArray(currentState?.deckLists) ? currentState.deckLists : [];
        const currentDecks = Array.isArray(currentState?.decks) ? currentState.decks : [];
        const currentRecords = Array.isArray(currentState?.records) ? currentState.records : [];
        const currentActiveGame = currentState?.activeGame && typeof currentState.activeGame === 'object' ? currentState.activeGame : null;
        const currentActiveGameUndo = Array.isArray(currentState?.activeGameUndo) ? currentState.activeGameUndo : [];

        const incomingHasData = games.length > 0
          || Object.keys(powerLevels).length > 0
          || deckLists.length > 0
          || decks.length > 0
          || records.length > 0
          || Boolean(activeGame)
          || activeGameUndo.length > 0;
        const currentHasData = currentGames.length > 0
          || Object.keys(currentPowerLevels).length > 0
          || currentDeckLists.length > 0
          || currentDecks.length > 0
          || currentRecords.length > 0
          || Boolean(currentActiveGame)
          || currentActiveGameUndo.length > 0;
        const allowDestructiveOverwrite = request.headers.get('X-Allow-Destructive-State-Overwrite') === '1';

        if (currentHasData && !incomingHasData && !allowDestructiveOverwrite) {
          return jsonResponse({
            error: 'Blocked an empty sync payload to prevent accidental cloud data loss. Pull the latest cloud state and retry.',
            conflict: {
              revision: currentRevision,
              updatedAt: String(currentState?.updatedAt || '').trim(),
              updatedBy: String(currentState?.updatedBy || '').trim(),
            },
          }, 409);
        }

        if (expectedRevision !== currentRevision) {
          return jsonResponse({
            error: 'Cloud state changed on another device before this sync completed.',
            conflict: {
              revision: currentRevision,
              updatedAt: String(currentState?.updatedAt || '').trim(),
              updatedBy: String(currentState?.updatedBy || '').trim(),
            },
          }, 409);
        }

        const updatedAt = new Date().toISOString();
        const nextRevision = currentRevision + 1;
        const podMembers = getConfiguredMembers(env);
        let normalizedDecks;

        try {
          normalizedDecks = enforceDeckOwnership(currentState?.decks || [], decks, auth, podMembers);
        } catch (error) {
          return jsonResponse({
            error: error instanceof Error ? error.message : 'Deck ownership validation failed.',
          }, 403);
        }

        await env.POD_STATE.put(stateKey, JSON.stringify({
          games,
          powerLevels,
          deckLists,
          decks: normalizedDecks,
          records,
          activeGame,
          activeGameUndo,
          revision: nextRevision,
          updatedAt,
          updatedBy: auth.displayName || auth.user,
        }));

        return jsonResponse({
          ok: true,
          revision: nextRevision,
          updatedAt,
          updatedBy: auth.displayName || auth.user,
          auth: buildAuthPayload(auth),
          decks: normalizedDecks,
        }, 200);
      }

      return jsonResponse({ error: 'Method not allowed.' }, 405);
    }

    if (url.pathname === '/api/session') {
      if (request.method === 'POST') {
        let body;
        try {
          body = await request.json();
        } catch (error) {
          return jsonResponse({ error: 'Invalid JSON payload.' }, 400);
        }

        const headers = new Headers({
          'Content-Type': 'application/json',
          'X-User-Name': getTextValue(body?.user),
          'X-Pod-Token': getTextValue(body?.token),
        });
        const authRequest = new Request(request.url, {
          method: request.method,
          headers,
        });
        const auth = await hasValidAuth(authRequest, env);
        if (!auth.ok) {
          return jsonResponse({ error: auth.reason }, 401);
        }

        const sessionKey = await persistSessionAuth(env, {
          ...auth,
          authMode: auth.authMode || 'session',
        });

        return jsonResponse({
          ok: true,
          auth: buildAuthPayload(auth),
        }, 200, {
          'Set-Cookie': buildSessionCookie(sessionKey),
        });
      }

      if (request.method === 'DELETE') {
        await clearSessionAuth(request, env);
        return jsonResponse({ ok: true }, 200, {
          'Set-Cookie': buildExpiredSessionCookie(),
        });
      }

      return jsonResponse({ error: 'Method not allowed.' }, 405);
    }

    if (url.pathname === '/api/auth-logs') {
      const auth = await hasValidAuth(request, env);
      if (!auth.ok) {
        await appendAuthAuditEntry(env, buildAuthAuditEntry({
          request,
          url,
          auth,
          success: false,
          reason: auth.reason,
          status: 401,
        }));
        return jsonResponse({ error: auth.reason }, 401);
      }

      if (getTextValue(auth?.role).toLowerCase() !== 'admin') {
        await appendAuthAuditEntry(env, buildAuthAuditEntry({
          request,
          url,
          auth,
          success: false,
          reason: 'Only admins can view authentication logs.',
          status: 403,
        }));
        return jsonResponse({ error: 'Only admins can view authentication logs.' }, 403);
      }

      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      const requestedLimit = Number.parseInt(getTextValue(url.searchParams.get('limit')), 10);
      const limit = Number.isFinite(requestedLimit)
        ? Math.min(200, Math.max(1, requestedLimit))
        : 120;
      const raw = env.POD_STATE ? await env.POD_STATE.get(AUTH_AUDIT_LOG_KEY, 'json') : null;
      const logs = Array.isArray(raw?.logs)
        ? raw.logs
        : (Array.isArray(raw) ? raw : []);

      return jsonResponse({
        logs: logs.slice(0, limit),
        total: logs.length,
      }, 200, {
        'Cache-Control': 'no-store, no-cache, must-revalidate',
      });
    }

    if (url.pathname === '/api/accounts') {
      const auth = await hasValidAuth(request, env);
      if (!auth.ok) {
        await appendAuthAuditEntry(env, buildAuthAuditEntry({
          request,
          url,
          auth,
          success: false,
          reason: auth.reason,
          status: 401,
        }));
        return jsonResponse({ error: auth.reason }, 401);
      }

      if (getTextValue(auth?.role).toLowerCase() !== 'admin') {
        await appendAuthAuditEntry(env, buildAuthAuditEntry({
          request,
          url,
          auth,
          success: false,
          reason: 'Only admins can view registered accounts.',
          status: 403,
        }));
        return jsonResponse({ error: 'Only admins can view registered accounts.' }, 403);
      }

      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      const stateKey = 'pod:default:state';
      const rawState = await env.POD_STATE.get(stateKey, 'json');
      const state = rawState && typeof rawState === 'object' ? rawState : null;
      const configuredMembers = getConfiguredMembers(env);
      const accounts = buildRegisteredAccounts(configuredMembers, state);

      return jsonResponse({
        accounts,
        total: accounts.length,
      }, 200, {
        'Cache-Control': 'no-store, no-cache, must-revalidate',
      });
    }

    if (url.pathname === '/api/card-rulings') {
      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      const name = getTextValue(url.searchParams.get('name'));
      if (!name) {
        return jsonResponse({ error: 'A card name is required.' }, 400);
      }

      try {
        const namedUrl = new URL('https://api.scryfall.com/cards/named');
        namedUrl.searchParams.set('exact', name);

        let cardRes = await fetch(namedUrl.toString(), { headers: getScryfallHeaders() });
        if (cardRes.status === 404) {
          namedUrl.searchParams.delete('exact');
          namedUrl.searchParams.set('fuzzy', name);
          cardRes = await fetch(namedUrl.toString(), { headers: getScryfallHeaders() });
        }

        if (!cardRes.ok) {
          if (cardRes.status === 404) {
            return jsonResponse({ error: `No card found matching "${name}".` }, 404);
          }
          if (cardRes.status === 429) {
            throw new Error(`Scryfall is temporarily rate-limited. Try again in about ${getRetryAfterSeconds(cardRes)} seconds.`);
          }
          throw new Error(`Scryfall card lookup failed (${cardRes.status})`);
        }

        const card = await cardRes.json();

        const rulingsRes = await fetch(`https://api.scryfall.com/cards/${card.id}/rulings`, {
          headers: getScryfallHeaders(),
        });
        const rulingsData = rulingsRes.ok ? await rulingsRes.json() : { data: [] };
        const rulings = Array.isArray(rulingsData?.data) ? rulingsData.data : [];

        const cardFaces = Array.isArray(card?.card_faces) ? card.card_faces.map((face) => ({
          name: getTextValue(face?.name),
          manaCost: getTextValue(face?.mana_cost),
          typeLine: getTextValue(face?.type_line),
          oracleText: getTextValue(face?.oracle_text),
          power: getTextValue(face?.power),
          toughness: getTextValue(face?.toughness),
          loyalty: getTextValue(face?.loyalty),
          imageUri: buildCommanderImageProxyUrl(getTextValue(face?.image_uris?.normal), requestOrigin),
          imageLargeUri: buildCommanderImageProxyUrl(getTextValue(face?.image_uris?.large), requestOrigin),
        })) : [];

        return jsonResponse({
          card: {
            id: getTextValue(card?.id),
            name: getTextValue(card?.name),
            manaCost: getTextValue(card?.mana_cost),
            typeLine: getTextValue(card?.type_line),
            oracleText: getTextValue(card?.oracle_text),
            power: getTextValue(card?.power),
            toughness: getTextValue(card?.toughness),
            loyalty: getTextValue(card?.loyalty),
            defense: getTextValue(card?.defense),
            scryfallUri: getTextValue(card?.scryfall_uri),
            setName: getTextValue(card?.set_name),
            releasedAt: getTextValue(card?.released_at),
            imageUri: buildCommanderImageProxyUrl(getCardImageUri(card), requestOrigin),
            imageLargeUri: buildCommanderImageProxyUrl(getCardImageVariant(card, 'large'), requestOrigin),
            layout: getTextValue(card?.layout),
            cardFaces,
          },
          rulings,
        }, 200, {
          'Cache-Control': 'public, max-age=86400, s-maxage=86400',
        });
      } catch (error) {
        return jsonResponse({
          error: 'Unable to look up that card right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    if (url.pathname === '/api/keywords') {
      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      try {
        const [abilitiesRes, actionsRes, abilityWordsRes] = await Promise.all([
          fetch('https://api.scryfall.com/catalog/keyword-abilities', { headers: getScryfallHeaders() }),
          fetch('https://api.scryfall.com/catalog/keyword-actions', { headers: getScryfallHeaders() }),
          fetch('https://api.scryfall.com/catalog/ability-words', { headers: getScryfallHeaders() }),
        ]);

        if (!abilitiesRes.ok || !actionsRes.ok || !abilityWordsRes.ok) {
          throw new Error('Scryfall catalog request failed');
        }

        const [abilitiesData, actionsData, abilityWordsData] = await Promise.all([
          abilitiesRes.json(),
          actionsRes.json(),
          abilityWordsRes.json(),
        ]);

        return jsonResponse({
          keywordAbilities: Array.isArray(abilitiesData?.data) ? abilitiesData.data : [],
          keywordActions: Array.isArray(actionsData?.data) ? actionsData.data : [],
          abilityWords: Array.isArray(abilityWordsData?.data) ? abilityWordsData.data : [],
        }, 200, {
          'Cache-Control': 'public, max-age=86400, s-maxage=86400',
        });
      } catch (error) {
        return jsonResponse({
          error: 'Unable to load keyword catalog right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    return env.ASSETS.fetch(request);
  },
};
