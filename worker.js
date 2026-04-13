const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, HEAD, PUT, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, X-User-Name, X-Pod-Token, X-State-Revision',
};

const SCRYFALL_AUTOCOMPLETE_CACHE_TTL_MS = 10 * 60 * 1000;
const SCRYFALL_CARD_CACHE_TTL_MS = 24 * 60 * 60 * 1000;
const scryfallAutocompleteCache = new Map();
const scryfallCardCache = new Map();

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

function mapDeckCard(card, requestOrigin) {
  const imageUri = getCardImageUri(card);
  const imageLargeUri = getCardImageVariant(card, 'large');

  return {
    id: getTextValue(card?.id),
    oracleId: getTextValue(card?.oracle_id),
    name: getTextValue(card?.name),
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
    isBanned: getTextValue(card?.legalities?.commander) === 'banned',
    isGameChanger: Boolean(card?.game_changer),
    isCommanderLegal: getTextValue(card?.legalities?.commander) === 'legal',
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
    const key = name.toLowerCase();
    if (!seenKeys.has(key)) {
      seenKeys.add(key);
      uniqueNames.push(name);
    }
  });

  const foundByNameKey = new Map();
  const unresolvedNames = [];

  uniqueNames.forEach((name) => {
    const key = name.toLowerCase();
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
      const nameKey = getTextValue(card?.name).toLowerCase();
      if (!nameKey) {
        return;
      }

      foundByNameKey.set(nameKey, mapped);
      setCachedValue(scryfallCardCache, nameKey, card);
    });
  }

  return requestedNames.map((name) => ({
    name,
    card: foundByNameKey.get(name.toLowerCase()) || null,
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

async function fetchCommanderCount(identity) {
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
    throw new Error(`Scryfall count request failed (${response.status}): ${detail}`);
  }

  const payload = await response.json();
  return Number.isFinite(Number(payload?.total_cards)) ? Number(payload.total_cards) : 0;
}

async function fetchRandomCommander(identity, requestOrigin) {
  const randomUrl = new URL('https://api.scryfall.com/cards/random');
  randomUrl.searchParams.set('q', buildCommanderSearchQuery(identity));

  const response = await fetch(randomUrl.toString(), {
    headers: getScryfallHeaders(),
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Scryfall random request failed (${response.status}): ${detail}`);
  }

  const payload = await response.json();
  return mapCommanderCard(payload, requestOrigin);
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

function hasValidAuth(request, env) {
  const user = getRequestUser(request);
  const token = getRequestToken(request);
  const configuredToken = (env.POD_ACCESS_TOKEN || '').trim();

  if (!configuredToken) {
    return { ok: false, reason: 'Server missing POD_ACCESS_TOKEN.' };
  }

  if (!env.POD_STATE) {
    return { ok: false, reason: 'Server missing POD_STATE KV binding.' };
  }

  if (!user || !token) {
    return { ok: false, reason: 'Missing credentials.' };
  }

  if (token !== configuredToken) {
    return { ok: false, reason: 'Invalid pod access code.' };
  }

  return { ok: true, user };
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

      const identity = normalizeCommanderIdentity(url.searchParams.get('identity'));
      if (!identity) {
        return jsonResponse({ error: 'A valid exact color identity is required.' }, 400);
      }

      try {
        const [totalCards, card] = await Promise.all([
          fetchCommanderCount(identity),
          fetchRandomCommander(identity, requestOrigin),
        ]);

        return jsonResponse({
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
      if (query.length < 2) {
        return jsonResponse({ results: [] }, 200, {
          'Cache-Control': 'no-store, no-cache, must-revalidate',
        });
      }

      try {
        const results = await fetchDeckSearchResults(query);
        return jsonResponse({ results: results.slice(0, 12) }, 200, {
          'Cache-Control': 'public, max-age=300, s-maxage=300',
        });
      } catch (error) {
        return jsonResponse({
          error: 'Unable to search cards right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    if (url.pathname === '/api/deck-card') {
      if (request.method !== 'GET') {
        return jsonResponse({ error: 'Method not allowed.' }, 405);
      }

      const name = getTextValue(url.searchParams.get('name'));
      if (!name) {
        return jsonResponse({ error: 'A card name is required.' }, 400);
      }

      try {
        const card = await fetchDeckCardByName(name, requestOrigin);
        return jsonResponse({ card }, 200, {
          'Cache-Control': 'public, max-age=86400, s-maxage=86400',
        });
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
          'Cache-Control': 'public, max-age=600, s-maxage=600',
        });
      } catch (error) {
        return jsonResponse({
          error: 'Unable to load card batch right now.',
          detail: error instanceof Error ? error.message : String(error),
        }, 502);
      }
    }

    if (url.pathname === '/api/state') {
      const auth = hasValidAuth(request, env);
      if (!auth.ok) {
        return jsonResponse({ error: auth.reason }, 401);
      }

      const stateKey = 'pod:default:state';

      if (request.method === 'GET') {
        const raw = await env.POD_STATE.get(stateKey, 'json');
        const state = raw && typeof raw === 'object' ? raw : { games: [], powerLevels: {}, deckLists: [], decks: [], records: [] };
        state.deckLists = Array.isArray(state.deckLists) ? state.deckLists : [];
        state.decks = Array.isArray(state.decks) ? state.decks : [];
        state.records = Array.isArray(state.records) ? state.records : [];
        state.revision = Number.isFinite(Number(state.revision)) ? Number(state.revision) : 0;
        state.updatedAt = String(state.updatedAt || '').trim();
        state.updatedBy = String(state.updatedBy || '').trim();

        if (url.searchParams.get('meta') === '1') {
          return jsonResponse({
            revision: state.revision,
            updatedAt: state.updatedAt,
            updatedBy: state.updatedBy,
          }, 200);
        }

        return jsonResponse(state, 200);
      }

      if (request.method === 'PUT') {
        let body;
        try {
          body = await request.json();
        } catch (error) {
          return jsonResponse({ error: 'Invalid JSON payload.' }, 400);
        }

        const games = Array.isArray(body.games) ? body.games : [];
        const powerLevels = body.powerLevels && typeof body.powerLevels === 'object' ? body.powerLevels : {};
        const deckLists = Array.isArray(body.deckLists) ? body.deckLists : [];
        const decks = Array.isArray(body.decks) ? body.decks : [];
        const records = Array.isArray(body.records) ? body.records : [];
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

        await env.POD_STATE.put(stateKey, JSON.stringify({
          games,
          powerLevels,
          deckLists,
          decks,
          records,
          revision: nextRevision,
          updatedAt,
          updatedBy: auth.user,
        }));

        return jsonResponse({ ok: true, revision: nextRevision, updatedAt, updatedBy: auth.user }, 200);
      }

      return jsonResponse({ error: 'Method not allowed.' }, 405);
    }

    return env.ASSETS.fetch(request);
  },
};
