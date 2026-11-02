// Storage keys, API endpoints, and app-level timing constants.
const STORAGE_KEY = 'commanderTrackerGames';
const EXPECTED_POWER_STORAGE_KEY = 'commanderExpectedPowerLevels';
const DECK_LIST_STORAGE_KEY = 'commanderDeckLists';
const DECKS_STORAGE_KEY = 'commanderDeckRecords';
const RECORDS_STORAGE_KEY = 'commanderTrackerRecords';
const ACTIVE_GAME_STORAGE_KEY = 'commanderTrackerActiveGame';
const ACTIVE_GAME_UNDO_STORAGE_KEY = 'commanderTrackerActiveGameUndo';
const SYNC_USER_STORAGE_KEY = 'commanderTrackerSyncUser';
const SYNC_TOKEN_STORAGE_KEY = 'commanderTrackerSyncToken';
const SYNC_CREDENTIAL_SET_AT_STORAGE_KEY = 'commanderTrackerSyncCredentialSetAt';
const SYNC_PENDING_CHANGES_STORAGE_KEY = 'commanderTrackerSyncPendingChanges';
const SYNC_REVISION_STORAGE_KEY = 'commanderTrackerSyncRevision';
const SYNC_CONFLICT_BACKUP_STORAGE_KEY = 'commanderTrackerSyncConflictBackup';
const CLOUD_SYNC_ENDPOINT = '/api/state';
const CLOUD_SYNC_METADATA_ENDPOINT = '/api/state?meta=1';
const COMMANDER_BUILDER_CACHE_STORAGE_KEY = 'commanderBuilderCacheV4';
const KEYWORDS_ENDPOINT = '/api/keywords';
const COMMANDER_BUILDER_ENDPOINT = '/api/commanders';
const DECK_SEARCH_ENDPOINT = '/api/deck-search';
const TOKEN_SEARCH_ENDPOINT = '/api/token-search';
const DECK_CARD_ENDPOINT = '/api/deck-card?v=2';
const DECK_CARDS_BULK_ENDPOINT = '/api/deck-cards-bulk';
const DECK_CARDS_BULK_REQUEST_CHUNK_SIZE = 75;
const DECK_CARDS_BULK_REQUEST_GAP_MS = 10;
const AUTH_AUDIT_ENDPOINT = '/api/auth-logs';
const REGISTERED_ACCOUNTS_ENDPOINT = '/api/accounts';
const BUILTIN_ADMIN_USER_KEY = 'baily';
const COMMANDER_BUILDER_CACHE_TTL_MS = 1000 * 60 * 60 * 24 * 7;
const ACTIVE_GAME_PERSIST_DEBOUNCE_MS = 180;
const DECKS_PERSIST_DEBOUNCE_MS = 1000;
const DECKS_RAPID_ACTION_PERSIST_DEBOUNCE_MS = 0;
const AUTH_AUDIT_REFRESH_INTERVAL_MS = 60 * 1000;

// Cached DOM references used across all pages.
const form = document.getElementById('game-form');
const dateInput = document.getElementById('game-date');
const playerTableBody = document.getElementById('player-table-body');
const addPlayerRowButton = document.getElementById('add-player-row');
const notesInput = document.getElementById('game-notes');
const firstBloodPlayerInput = document.getElementById('game-first-blood-player');
const firstBloodPlayerMenu = document.getElementById('game-first-blood-menu');
const firstBloodTurnInput = document.getElementById('game-first-blood-turn');
const winningTurnInput = document.getElementById('game-winning-turn');
const summaryEl = document.getElementById('summary');
const playerDatalist = document.getElementById('player-list');
const commanderDatalist = document.getElementById('commander-list');
const playerRenameForm = document.getElementById('player-rename-form');
const playerRenameCurrentInput = document.getElementById('player-rename-current');
const playerRenameNextInput = document.getElementById('player-rename-next');
const playerRenameStatus = document.getElementById('player-rename-status');
const playerRenameDatalist = document.getElementById('player-rename-list');
const playerRenameMenu = document.getElementById('player-rename-menu');
const authAuditSection = document.getElementById('auth-audit-section')
  || document.querySelector('.page-admin-logs main .card');
const authAuditStatus = document.getElementById('auth-audit-status');
const authAuditList = document.getElementById('auth-audit-list');
const authAuditRefreshButton = document.getElementById('auth-audit-refresh');
const authAuditExportJsonButton = document.getElementById('auth-audit-export-json');
const authAuditExportCsvButton = document.getElementById('auth-audit-export-csv');
const authAuditFilterResultSelect = document.getElementById('auth-audit-filter-result');
const authAuditFilterActionSelect = document.getElementById('auth-audit-filter-action');
const authAuditFilterUserInput = document.getElementById('auth-audit-filter-user');
const authAuditFilterFromInput = document.getElementById('auth-audit-filter-from');
const authAuditFilterToInput = document.getElementById('auth-audit-filter-to');
const authAuditFilterClearButton = document.getElementById('auth-audit-filter-clear');
const authAuditPageSizeSelect = document.getElementById('auth-audit-page-size');
const authAuditPagePrevButton = document.getElementById('auth-audit-page-prev');
const authAuditPageNextButton = document.getElementById('auth-audit-page-next');
const authAuditPageSummary = document.getElementById('auth-audit-page-summary');
const syncDebugSection = document.getElementById('sync-debug-section');
const syncDebugStatus = document.getElementById('sync-debug-status');
const syncDebugList = document.getElementById('sync-debug-list');
const syncDebugClearButton = document.getElementById('sync-debug-clear');
const registeredAccountsSection = document.getElementById('registered-accounts-section');
const registeredAccountsStatus = document.getElementById('registered-accounts-status');
const registeredAccountsBody = document.getElementById('registered-accounts-body');
const commanderRenameForm = document.getElementById('commander-rename-form');
const commanderRenameCurrentInput = document.getElementById('commander-rename-current');
const commanderRenameNextInput = document.getElementById('commander-rename-next');
const commanderRenameStatus = document.getElementById('commander-rename-status');
const commanderRenameDatalist = document.getElementById('commander-rename-list');
const commanderRenameMenu = document.getElementById('commander-rename-menu');
const commanderBuilderForm = document.getElementById('commander-builder-form');
const commanderBuilderResult = document.getElementById('commander-builder-result');
const commanderBuilderStatus = document.getElementById('commander-builder-status');
const commanderBuilderCount = document.getElementById('commander-builder-count');
const commanderBuilderRerollButton = document.getElementById('commander-builder-reroll');
const commanderBuilderModeInputs = Array.from(document.querySelectorAll('input[name="commander-builder-mode"]'));
const commanderBuilderIdentityPanel = document.getElementById('commander-builder-identity-panel');
const commanderBuilderIdentityTitle = document.getElementById('commander-builder-identity-title');
const commanderBuilderIdentityCopy = document.getElementById('commander-builder-identity-copy');
const commanderBuilderKeywordPanel = document.getElementById('commander-builder-keyword-panel');
const commanderBuilderKeywordSearch = document.getElementById('commander-builder-keyword-search');
const commanderBuilderKeywordStatus = document.getElementById('commander-builder-keyword-status');
const commanderBuilderKeywordGrid = document.getElementById('commander-builder-keyword-grid');
const commanderBuilderKeywordSelection = document.getElementById('commander-builder-keyword-selection');
const liveModalPrompt = document.getElementById('live-modal-prompt');
const liveModalTitle = document.getElementById('live-modal-title');
const liveModalCopy = document.getElementById('live-modal-copy');
const liveModalField = document.getElementById('live-modal-field');
const liveModalInput = document.getElementById('live-modal-input');
const liveModalTextarea = document.getElementById('live-modal-textarea');
const liveModalError = document.getElementById('live-modal-error');
const liveModalCancelButton = document.getElementById('live-modal-cancel');
const liveModalConfirmButton = document.getElementById('live-modal-confirm');
const cardImageModal = document.getElementById('card-image-modal');
const cardImageModalImg = document.getElementById('card-image-modal-img');
const cardImageModalCloseButton = document.querySelector('#card-image-modal .card-image-modal-close');
const cardImageModalOverlay = document.querySelector('#card-image-modal .card-image-modal-overlay');
const syncUserInput = document.getElementById('sync-user');
const syncTokenInput = document.getElementById('sync-token');
const syncConnectButton = document.getElementById('sync-connect');
const syncDisconnectButton = document.getElementById('sync-disconnect');
const syncNowButton = document.getElementById('sync-now');
const syncStatus = document.getElementById('sync-status');
const syncPanel = document.querySelector('.sync-panel');
const pageSwitch = document.querySelector('.page-switch');
const pageSwitchToggleButton = document.querySelector('.page-switch-toggle');
const pageSwitchPanel = document.querySelector('.page-switch-panel');

if (syncStatus) {
  syncStatus.setAttribute('role', 'status');
  syncStatus.setAttribute('aria-live', 'polite');
}

if (commanderBuilderStatus) {
  commanderBuilderStatus.setAttribute('role', 'status');
  commanderBuilderStatus.setAttribute('aria-live', 'polite');
}

let historySortKey = 'date';
let historySortDescending = true;
let editingGameId = null;
let knownPlayers = [];
let knownCommanders = [];
let commanderSortColumn = 'games';
let commanderSortDescending = true;
const tableSortState = {
  rankingsMain: { column: 'rank', descending: false },
  playerStats: { column: 'games', descending: true },
  commanderStats: { column: 'games', descending: true },
  recentPlayerTrends: { column: 'points', descending: true },
  recentCommanderTrends: { column: 'points', descending: true },
  playerStreaks: { column: 'currentWins', descending: true },
  commanderStreaks: { column: 'currentWins', descending: true },
  decks: { column: 'updatedAt', descending: true },
};
let appState = { games: [], powerLevels: {}, deckLists: [], decks: [], records: [], activeGame: null, activeGameUndo: [] };
let syncQueueTimer = null;
let syncRetryTimer = null;
let syncInFlight = false;
let syncBootstrapInFlight = false;
let syncRetryCount = 0;
let syncPendingChanges = false;
let syncLastSuccessAt = null;
let syncConnectionState = 'local';
let syncLastErrorMessage = '';
let syncCloudRevision = 0;
let syncCloudUpdatedAt = '';
let syncCloudUpdatedBy = '';
let syncHasLoadedCloudState = false;
let syncLocalContentVerified = false;
let syncQueuedGamesOnly = null;
let syncAuthenticatedUserId = '';
let syncAuthenticatedDisplayName = '';
let syncAuthenticatedRole = '';
let authAuditLogs = [];
let syncDebugLogs = [];
let authAuditLoading = false;
let authAuditLastLoadedAt = 0;
let authAuditLoadedForUserId = '';
let authAuditFilterState = { result: 'all', action: 'all', user: '', from: '', to: '' };
let authAuditPageSize = 10;
let authAuditCurrentPage = 1;
let registeredAccounts = [];
let registeredAccountsLoading = false;
let registeredAccountsLastLoadedAt = 0;
let registeredAccountsLoadedForUserId = '';
let syncConflictInfo = null;
let syncMetadataCheckInFlight = false;
let syncLastFreshnessCheckAt = 0;
let storageErrorMessage = '';
let syncStateMutationVersion = 0;
let syncDecksMutationVersion = 0;
let historyQueryFiltersApplied = false;
let deckSelectorSpinTimer = null;
let deckSelectorRotation = 0;
let commanderIdentityPrefetchInFlight = false;
let commanderIdentityPrefetchRequestId = 0;
let commanderBuilderQueryKey = '';
let commanderBuilderLoading = false;
let commanderBuilderRequestId = 0;
let commanderBuilderLastCardName = '';
let commanderBuilderMode = 'identity';
let commanderBuilderKeywordsLoading = false;
let commanderBuilderKeywordsLoaded = false;
let commanderBuilderKeywordErrorMessage = '';
let commanderBuilderKeywordCatalog = [];
let commanderBuilderKeywordSearchTerm = '';
let commanderBuilderSelectedKeywords = [];
let deckListOracleBackfillRan = false;
let deckLibraryPlayerFilterDefaulted = false;
let historyPlayerFilterDefaulted = false;
let rankingsCommanderIdentityLoading = false;
let rankingsCommanderIdentityRequestId = 0;
const commanderIdentityAttemptedKeys = new Set();
const rankingsCommanderIdentityAttemptedKeys = new Set();
const deckBuilderSearchCache = new Map();
const deckBuilderCardCache = new Map();
let activeGameState = null;
let activeGameUndoState = [];
let wakeLockSentinel = null;
let liveModalResolver = null;
let liveModalConfig = null;
let activeGamePersistTimer = null;
let decksPersistTimer = null;
let pageExitFlushStamp = 0;
const derivedGamesCache = new WeakMap();

const DEFAULT_RECORD_DEFINITIONS = [
  { key: 'earliest-turn-win', title: 'Earliest Turn Win', unit: 'turns' },
  { key: 'highest-damage-dealt', title: 'Highest Damage Dealt (Single Hit)', unit: 'damage' },
  { key: 'highest-hit-to-life', title: 'Highest Hit to Life', unit: 'damage' },
  { key: 'most-kills-one-game', title: 'Most Kills in One Game', unit: 'kills' },
  { key: 'highest-life-total', title: 'Highest Life Total', unit: 'life' },
  { key: 'most-commander-damage', title: 'Most Commander Damage', unit: 'damage' },
  { key: 'longest-game', title: 'Longest Game', unit: 'minutes' },
  { key: 'fastest-elimination', title: 'Fastest Elimination', unit: 'turns' },
];

const PLAYER_IDENTITY_KEYS = new Set(['player', 'holder', 'manualHolder', 'actorPlayer', 'targetPlayer', 'startingPlayer', 'owner']);
const COMMANDER_IDENTITY_KEYS = new Set(['commander', 'manualCommander', 'actorCommander']);
const PLAYER_IDENTITY_LIST_PARENTS = new Set(['players', 'finishOrder', 'killed']);
const RECOGNIZED_DECK_HOSTS = new Set(['moxfield.com', 'www.moxfield.com', 'archidekt.com', 'deckstats.net', 'www.deckstats.net', 'tappedout.net', 'www.tappedout.net', 'www.mtggoldfish.com', 'mtggoldfish.com', 'manabox.app', 'www.manabox.app', 'etherhub.io', 'www.etherhub.io']);
const COMMANDER_BUILDER_COLOR_OPTIONS = [
  { code: 'W', label: 'White' },
  { code: 'U', label: 'Blue' },
  { code: 'B', label: 'Black' },
  { code: 'R', label: 'Red' },
  { code: 'G', label: 'Green' },
];

function parseJsonSafe(value, fallback) {
  try {
    return JSON.parse(value);
  } catch (error) {
    return fallback;
  }
}

function getDerivedCacheBucket(games) {
  if (!Array.isArray(games)) {
    return {};
  }

  if (!derivedGamesCache.has(games)) {
    derivedGamesCache.set(games, {});
  }

  return derivedGamesCache.get(games);
}

function readLocalStorageValue(key) {
  storageErrorMessage = '';

  try {
    if (typeof globalThis === 'undefined' || !globalThis.localStorage) {
      return null;
    }

    const value = globalThis.localStorage.getItem(key);
    return value === null ? null : value;
  } catch (error) {
    storageErrorMessage = `Storage is unavailable: ${error.message}`;
    return null;
  }
}

function writeLocalStorageValue(key, value) {
  storageErrorMessage = '';

  try {
    if (typeof globalThis === 'undefined' || !globalThis.localStorage) {
      throw new Error('localStorage is unavailable');
    }

    globalThis.localStorage.setItem(key, String(value));
    return true;
  } catch (error) {
    storageErrorMessage = `Couldn’t save to storage: ${error.message}`;
    return false;
  }
}

function removeLocalStorageValue(key) {
  storageErrorMessage = '';

  try {
    if (typeof globalThis === 'undefined' || !globalThis.localStorage) {
      return true;
    }

    globalThis.localStorage.removeItem(key);
    return true;
  } catch (error) {
    storageErrorMessage = `Couldn’t clear storage: ${error.message}`;
    return false;
  }
}

function loadSyncPendingChangesState() {
  return readLocalStorageValue(SYNC_PENDING_CHANGES_STORAGE_KEY) === 'true';
}

function setSyncPendingChanges(value) {
  syncPendingChanges = Boolean(value);

  if (syncPendingChanges) {
    writeLocalStorageValue(SYNC_PENDING_CHANGES_STORAGE_KEY, 'true');
    return;
  }

  removeLocalStorageValue(SYNC_PENDING_CHANGES_STORAGE_KEY);
}

function hasLoadedCloudState() {
  return syncHasLoadedCloudState;
}

function getCloudStateLoadStatus() {
  return {
    loaded: syncHasLoadedCloudState,
    loading: syncBootstrapInFlight || syncInFlight || syncConnectionState === 'connecting',
    errorMessage: syncLastErrorMessage || (
      syncConnectionState === 'local' && hasSyncCredentials()
        ? 'Cloud session is not connected. Reconnect to load cloud decks.'
        : ''
    ),
  };
}

function getSyncCredentialAgeDays() {
  return null;
}

function getStorageWarningMessage() {
  return storageErrorMessage || '';
}

function getExactKnownCommander(rawCommander, knownCommanders) {
  const commanderKey = getIdentityKey(rawCommander);
  if (!commanderKey) {
    return '';
  }

  return (Array.isArray(knownCommanders) ? knownCommanders : [])
    .find((commander) => getIdentityKey(commander) === commanderKey) || '';
}

function getBestCommanderSuggestion(rawCommander, knownCommanders) {
  const inputKey = getIdentityKey(rawCommander);
  if (!inputKey) {
    return '';
  }

  let best = { score: 0, commander: '' };
  (Array.isArray(knownCommanders) ? knownCommanders : []).forEach((commander) => {
    const candidateKey = getIdentityKey(commander);
    const score = getStringSimilarity(inputKey, candidateKey);
    if (score > best.score) {
      best = { score, commander };
    }
  });

  return best.score >= 0.75 ? best.commander : '';
}

async function fetchDeckCardByNameWithFallback(name) {
  if (!name || !name.trim()) {
    return null;
  }

  try {
    const card = await fetchDeckCardByName(name);
    if (card) {
      return card;
    }
  } catch (error) {
    // Fall through to search fallback when exact lookup fails.
  }

  try {
    const results = await fetchDeckSearchResultsList(name);
    if (!results.length) {
      return null;
    }

    for (let i = 0; i < Math.min(results.length, 5); i += 1) {
      const candidate = results[i];
      if (!candidate) {
        continue;
      }

      try {
        const card = await fetchDeckCardByName(candidate);
        if (card) {
          return card;
        }
      } catch (error) {
        // Ignore individual candidate lookup failures.
      }
    }
  } catch (error) {
    // Ignore search fallback failures.
  }

  return null;
}

async function resolveCommanderInput(rawCommander, knownCommanders) {
  if (!rawCommander || !rawCommander.trim()) {
    return '';
  }

  try {
    const card = await fetchDeckCardByNameWithFallback(rawCommander);
    if (card?.name) {
      return card.name;
    }
  } catch (error) {
    // Ignore lookup failures during validation.
  }

  const exactKnown = getExactKnownCommander(rawCommander, knownCommanders);
  if (exactKnown) {
    return exactKnown;
  }

  return '';
}

async function canonicalizeCommanderInputValue(rawCommander) {
  const knownCommandersList = getKnownCommanderOptions();
  const resolvedCardName = await resolveCommanderInput(rawCommander, knownCommandersList);
  if (resolvedCardName) {
    return resolvedCardName;
  }

  const suggestion = getBestCommanderSuggestion(rawCommander, knownCommandersList);
  if (suggestion) {
    return suggestion;
  }

  return '';
}

async function validateCommanderEntries(rows, { allowExactCardLookup = true } = {}) {
  const knownCommandersList = Array.from(new Set(buildCanonicalIdentityMapFromValues(knownCommanders).values())).filter(Boolean);

  for (const row of rows) {
    const rawCommander = String(row.commanderRaw || row.commander || '').trim();
    if (!rawCommander) {
      return `Enter a commander for ${row.player}.`;
    }

    if (allowExactCardLookup) {
      const resolved = await resolveCommanderInput(rawCommander, knownCommandersList);
      if (resolved) {
        row.commander = resolved;
        continue;
      }
    }

    const exactMatch = getExactKnownCommander(rawCommander, knownCommandersList);
    if (exactMatch) {
      row.commander = exactMatch;
      continue;
    }

    const suggestion = getBestCommanderSuggestion(rawCommander, knownCommandersList);
    if (suggestion) {
      const confirmed = await promptLiveConfirm(
        `Commander "${rawCommander}" does not exactly match a known commander. Did you mean "${suggestion}"?`,
        { title: 'Confirm commander', confirmLabel: 'Use suggestion' },
      );
      if (confirmed) {
        row.commander = suggestion;
        continue;
      }
      return `Please correct the commander for ${row.player}: ${rawCommander}.`;
    }

    return `Commander "${rawCommander}" is not recognized. Enter the exact card name or choose a known commander.`;
  }

  return '';
}

function normalizeIdentityLabel(value) {
  return String(value || '')
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function getIdentityKey(value) {
  const normalizedValue = normalizeIdentityLabel(value);
  return normalizedValue ? normalizedValue.toLocaleLowerCase() : '';
}

function normalizeSyncUserKey(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[^a-z0-9]+/g, '');
}

function getIdentityDisplayScore(value) {
  const stringValue = String(value || '');
  const normalizedValue = normalizeIdentityLabel(value);
  if (!normalizedValue) {
    return 0;
  }

  let score = 0;
  if (normalizedValue !== normalizedValue.toLocaleLowerCase()) score += 2;
  if (/\b[A-Z]/.test(normalizedValue)) score += 1;
  if (/[^a-zA-Z0-9\s]/.test(stringValue)) score += 2;
  return score;
}

function recordIdentityVariant(bucketMap, value) {
  const normalizedValue = normalizeIdentityLabel(value);
  const identityKey = getIdentityKey(normalizedValue);
  const rawValue = String(value || '').trim();
  if (!identityKey || !rawValue) return;

  if (!bucketMap.has(identityKey)) {
    bucketMap.set(identityKey, new Map());
  }

  const variants = bucketMap.get(identityKey);
  variants.set(rawValue, (variants.get(rawValue) || 0) + 1);
}

function buildCanonicalIdentityMap(bucketMap) {
  const canonicalMap = new Map();

  bucketMap.forEach((variants, identityKey) => {
    let preferredValue = '';
    let preferredCount = -1;
    let preferredScore = -1;

    variants.forEach((count, value) => {
      const displayScore = getIdentityDisplayScore(value);
      const shouldReplace = displayScore > preferredScore
        || (displayScore === preferredScore && count > preferredCount)
        || (displayScore === preferredScore && count === preferredCount && value.localeCompare(preferredValue) < 0);

      if (shouldReplace) {
        preferredValue = value;
        preferredCount = count;
        preferredScore = displayScore;
      }
    });

    canonicalMap.set(identityKey, preferredValue);
  });

  return canonicalMap;
}

function buildCanonicalIdentityMapFromValues(values) {
  const bucketMap = new Map();
  (Array.isArray(values) ? values : []).forEach((value) => {
    recordIdentityVariant(bucketMap, value);
  });
  return buildCanonicalIdentityMap(bucketMap);
}

function canonicalizeIdentityValue(value, canonicalMap) {
  const rawValue = String(value || '').trim();
  if (!rawValue) return '';
  return canonicalMap?.get(getIdentityKey(rawValue)) || rawValue;
}

function getStringEditDistance(a, b) {
  const aLength = a.length;
  const bLength = b.length;
  if (!aLength) return bLength;
  if (!bLength) return aLength;

  const row = Array.from({ length: bLength + 1 }, (_, index) => index);
  let previousRow;

  for (let i = 1; i <= aLength; i += 1) {
    previousRow = row.slice();
    row[0] = i;
    for (let j = 1; j <= bLength; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      row[j] = Math.min(
        previousRow[j] + 1,
        row[j - 1] + 1,
        previousRow[j - 1] + cost,
      );
    }
  }

  return row[bLength];
}

function getStringSimilarity(a, b) {
  const normalizedA = String(a || '').trim().toLowerCase();
  const normalizedB = String(b || '').trim().toLowerCase();
  if (!normalizedA || !normalizedB) return 0;

  const distance = getStringEditDistance(normalizedA, normalizedB);
  return 1 - (distance / Math.max(normalizedA.length, normalizedB.length));
}

function normalizeIdentityList(value, canonicalMap) {
  const values = Array.isArray(value) ? value : String(value || '').split(',');
  return values
    .map((entry) => canonicalizeIdentityValue(entry, canonicalMap))
    .filter(Boolean);
}

function isPlayerIdentityField(key) {
  return PLAYER_IDENTITY_KEYS.has(key);
}

function isCommanderIdentityField(key) {
  return COMMANDER_IDENTITY_KEYS.has(key);
}

function isPlayerIdentityListParent(key) {
  return PLAYER_IDENTITY_LIST_PARENTS.has(key);
}

function scanIdentityValues(value, recordPlayer, recordCommander, parentKey = '') {
  if (Array.isArray(value)) {
    if (isPlayerIdentityListParent(parentKey)) {
      value.forEach(recordPlayer);
      return;
    }

    value.forEach((entry) => {
      scanIdentityValues(entry, recordPlayer, recordCommander, parentKey);
    });
    return;
  }

  if (!value || typeof value !== 'object') {
    return;
  }

  Object.entries(value).forEach(([key, nestedValue]) => {
    if (isPlayerIdentityField(key)) {
      recordPlayer(nestedValue);
      return;
    }

    if (isCommanderIdentityField(key)) {
      recordCommander(nestedValue);
      return;
    }

    scanIdentityValues(nestedValue, recordPlayer, recordCommander, key);
  });
}

function normalizeIdentityValues(value, { playerMap, commanderMap }, parentKey = '') {
  if (Array.isArray(value)) {
    if (isPlayerIdentityListParent(parentKey)) {
      return normalizeIdentityList(value, playerMap);
    }

    return value.map((entry) => normalizeIdentityValues(entry, { playerMap, commanderMap }, parentKey));
  }

  if (!value || typeof value !== 'object') {
    return value;
  }

  const normalizedValue = {};
  Object.entries(value).forEach(([key, nestedValue]) => {
    if (isPlayerIdentityField(key)) {
      normalizedValue[key] = canonicalizeIdentityValue(nestedValue, playerMap);
      return;
    }

    if (isCommanderIdentityField(key)) {
      normalizedValue[key] = canonicalizeIdentityValue(nestedValue, commanderMap);
      return;
    }

    normalizedValue[key] = normalizeIdentityValues(nestedValue, { playerMap, commanderMap }, key);
  });

  return normalizedValue;
}

function buildAppStateIdentityMaps(state) {
  const playerBuckets = new Map();
  const commanderBuckets = new Map();
  const recordPlayer = (value) => recordIdentityVariant(playerBuckets, value);
  const recordCommander = (value) => recordIdentityVariant(commanderBuckets, value);

  scanIdentityValues(Array.isArray(state?.games) ? state.games : [], recordPlayer, recordCommander, 'games');
  scanIdentityValues(Array.isArray(state?.deckLists) ? state.deckLists : [], recordPlayer, recordCommander, 'deckLists');
  scanIdentityValues(Array.isArray(state?.records) ? state.records : [], recordPlayer, recordCommander, 'records');
  Object.keys(state?.powerLevels && typeof state.powerLevels === 'object' ? state.powerLevels : {}).forEach(recordCommander);

  return {
    playerMap: buildCanonicalIdentityMap(playerBuckets),
    commanderMap: buildCanonicalIdentityMap(commanderBuckets),
  };
}

function getDeckCardPrimaryType(typeLine) {
  const normalizedTypeLine = String(typeLine || '').toLowerCase();
  const typePriority = ['creature', 'artifact', 'enchantment', 'planeswalker', 'battle', 'instant', 'sorcery', 'land'];
  const matchedType = typePriority.find((type) => normalizedTypeLine.includes(type));
  return matchedType ? matchedType.charAt(0).toUpperCase() + matchedType.slice(1) : 'Other';
}

function normalizeDeckCardFace(face) {
  if (!face || typeof face !== 'object') {
    return null;
  }

  const name = String(face.name || '').trim();
  const manaCost = String(face.manaCost || '').trim();
  const typeLine = String(face.typeLine || '').trim();
  const oracleText = String(face.oracleText || '').trim();
  const imageUri = String(face.imageUri || '').trim();
  const imageLargeUri = String(face.imageLargeUri || '').trim();
  const imagePngUri = String(face.imagePngUri || '').trim();
  const power = String(face.power || '').trim();
  const toughness = String(face.toughness || '').trim();
  const loyalty = String(face.loyalty || '').trim();
  const defense = String(face.defense || '').trim();

  if (!name && !manaCost && !typeLine && !oracleText && !imageUri && !imageLargeUri && !imagePngUri && !power && !toughness && !loyalty && !defense) {
    return null;
  }

  return {
    name,
    manaCost,
    typeLine,
    oracleText,
    imageUri,
    imageLargeUri,
    imagePngUri,
    power,
    toughness,
    loyalty,
    defense,
  };
}

function normalizeDeckCardEntry(card) {
  if (!card || typeof card !== 'object') {
    return null;
  }

  const name = String(card.name || '').trim();
  if (!name) {
    return null;
  }

  const typeLine = String(card.typeLine || '').trim();

  return {
    id: String(card.id || generateId()).trim(),
    oracleId: String(card.oracleId || '').trim(),
    name,
    manaCost: String(card.manaCost || '').trim(),
    typeLine,
    oracleText: String(card.oracleText || '').trim(),
    cardType: String(card.cardType || getDeckCardPrimaryType(typeLine)).trim() || 'Other',
    scryfallUri: String(card.scryfallUri || '').trim(),
    imageUri: String(card.imageUri || '').trim(),
    imageLargeUri: String(card.imageLargeUri || '').trim(),
    cardFaces: Array.isArray(card.cardFaces)
      ? card.cardFaces.map((face) => ({
        name: String(face?.name || '').trim(),
        manaCost: String(face?.manaCost || '').trim(),
        typeLine: String(face?.typeLine || '').trim(),
        oracleText: String(face?.oracleText || '').trim(),
        imageUri: String(face?.imageUri || '').trim(),
        imageLargeUri: String(face?.imageLargeUri || '').trim(),
        imagePngUri: String(face?.imagePngUri || '').trim(),
        power: String(face?.power || '').trim(),
        toughness: String(face?.toughness || '').trim(),
        loyalty: String(face?.loyalty || '').trim(),
        defense: String(face?.defense || '').trim(),
      })).filter((face) => face.name || face.oracleText || face.typeLine || face.manaCost || face.power || face.toughness || face.loyalty || face.defense)
      : [],
    colorIdentity: Array.isArray(card.colorIdentity) ? card.colorIdentity.map((value) => String(value || '').trim()).filter(Boolean) : [],
    power: String(card.power || '').trim(),
    toughness: String(card.toughness || '').trim(),
    loyalty: String(card.loyalty || '').trim(),
    defense: String(card.defense || '').trim(),
    count: Number.isFinite(Number(card.count)) ? Math.max(1, Number(card.count)) : 1,
    isBanned: Boolean(card.isBanned),
    isGameChanger: Boolean(card.isGameChanger),
    isCommanderLegal: Boolean(card.isCommanderLegal),
    isToken: Boolean(card.isToken),
  };
}

function mergeDeckBuilderTokenCards(cards) {
  const tokensByKey = new Map();
  cards.forEach((card) => {
    if (!card || !card.name) {
      return;
    }
    const key = getIdentityKey(card.name);
    if (!key) {
      return;
    }
    const count = Number.isFinite(Number(card.count)) ? Math.max(1, Number(card.count)) : 1;
    if (!tokensByKey.has(key)) {
      tokensByKey.set(key, { ...card, count, id: String(card.id || generateId()).trim() });
      return;
    }
    const existing = tokensByKey.get(key);
    existing.count = (Number.isFinite(Number(existing.count)) ? Math.max(1, Number(existing.count)) : 1) + count;
  });
  return [...tokensByKey.values()];
}

function normalizeDeckPowerLevel(value) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return null;
  }

  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return null;
  }

  return Math.min(10, Math.max(0, Math.round(parsed * 10) / 10));
}

function normalizeDeckRecord(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const name = String(entry.name || '').trim() || 'Untitled Deck';
  const owner = normalizeIdentityLabel(entry.owner || '');
  const ownerUserId = String(entry.ownerUserId || '').trim().toLowerCase();
  const commander = normalizeDeckCardEntry(entry.commander);
  // A secondCommander without a primary commander is invalid — clear it.
  const secondCommander = commander ? normalizeDeckCardEntry(entry.secondCommander) : null;
  const cards = Array.isArray(entry.cards)
    ? entry.cards.map(normalizeDeckCardEntry).filter(Boolean)
    : [];
  const maybeboard = Array.isArray(entry.maybeboard)
    ? entry.maybeboard.map(normalizeDeckCardEntry).filter(Boolean)
    : [];
  const tokens = Array.isArray(entry.tokens)
    ? mergeDeckBuilderTokenCards(entry.tokens.map(normalizeDeckCardEntry).filter(Boolean))
    : [];

  return {
    id: String(entry.id || generateId()).trim(),
    name,
    owner,
    ownerUserId,
    commander,
    secondCommander,
    cards,
    maybeboard,
    tokens,
    inRotation: entry.inRotation !== false,
    powerLevel: normalizeDeckPowerLevel(entry.powerLevel),
    createdAt: String(entry.createdAt || new Date().toISOString()).trim(),
    updatedAt: String(entry.updatedAt || entry.createdAt || new Date().toISOString()).trim(),
  };
}

function normalizeAppStateData(state) {
  const baseState = {
    games: Array.isArray(state?.games) ? state.games : [],
    powerLevels: state?.powerLevels && typeof state.powerLevels === 'object' ? state.powerLevels : {},
    deckLists: Array.isArray(state?.deckLists) ? state.deckLists : [],
    decks: Array.isArray(state?.decks) ? state.decks : [],
    records: Array.isArray(state?.records) ? state.records : [],
    activeGame: state?.activeGame && typeof state.activeGame === 'object' ? state.activeGame : null,
    activeGameUndo: Array.isArray(state?.activeGameUndo) ? state.activeGameUndo : [],
  };

  const identityMaps = buildAppStateIdentityMaps(baseState);
  const normalizedPowerLevels = {};
  Object.entries(baseState.powerLevels).forEach(([commander, value]) => {
    const canonicalCommander = canonicalizeIdentityValue(commander, identityMaps.commanderMap);
    if (!canonicalCommander || typeof value !== 'number' || Number.isNaN(value)) {
      return;
    }

    normalizedPowerLevels[canonicalCommander] = value;
  });

  return {
    games: (Array.isArray(baseState.games) ? baseState.games : []).map((game) => normalizeSavedGameEntry(game, baseState.decks)),
    powerLevels: normalizedPowerLevels,
    deckLists: normalizeIdentityValues(baseState.deckLists, identityMaps, 'deckLists'),
    decks: baseState.decks.map(normalizeDeckRecord).filter(Boolean),
    records: normalizeIdentityValues(baseState.records, identityMaps, 'records'),
    activeGame: normalizeActiveGameStateData(baseState.activeGame),
    activeGameUndo: baseState.activeGameUndo.map((entry) => normalizeActiveGameStateData(entry)).filter(Boolean),
  };
}

function normalizeActiveGameStateData(state) {
  if (!state || typeof state !== 'object') {
    return null;
  }

  const playerBuckets = new Map();
  const commanderBuckets = new Map();
  (Array.isArray(state.players) ? state.players : []).forEach((player) => {
    recordIdentityVariant(playerBuckets, player?.name);
    recordIdentityVariant(commanderBuckets, player?.commander);
    (Array.isArray(player?.killedPlayers) ? player.killedPlayers : []).forEach((killedPlayer) => {
      recordIdentityVariant(playerBuckets, killedPlayer);
    });
  });

  const playerMap = buildCanonicalIdentityMap(playerBuckets);
  const commanderMap = buildCanonicalIdentityMap(commanderBuckets);

  return {
    ...state,
    players: (Array.isArray(state.players) ? state.players : []).map((player) => ({
      ...player,
      name: canonicalizeIdentityValue(player?.name, playerMap),
      commander: canonicalizeIdentityValue(player?.commander, commanderMap),
      killedPlayers: normalizeIdentityList(player?.killedPlayers || [], playerMap),
    })),
  };
}

function renameIdentityValue(value, sourceKey, replacementValue) {
  const rawValue = String(value || '').trim();
  if (!rawValue) {
    return '';
  }

  return getIdentityKey(rawValue) === sourceKey ? replacementValue : rawValue;
}

function renameIdentityList(value, sourceKey, replacementValue) {
  const values = Array.isArray(value) ? value : String(value || '').split(',');
  return values
    .map((entry) => renameIdentityValue(entry, sourceKey, replacementValue))
    .filter(Boolean);
}

function renameIdentityValues(value, { sourceKey, replacementValue, type }, parentKey = '') {
  if (Array.isArray(value)) {
    if (type === 'player' && isPlayerIdentityListParent(parentKey)) {
      return renameIdentityList(value, sourceKey, replacementValue);
    }

    return value.map((entry) => renameIdentityValues(entry, { sourceKey, replacementValue, type }, parentKey));
  }

  if (!value || typeof value !== 'object') {
    return value;
  }

  const renamedValue = {};
  Object.entries(value).forEach(([key, nestedValue]) => {
    if (type === 'player' && isPlayerIdentityField(key)) {
      renamedValue[key] = renameIdentityValue(nestedValue, sourceKey, replacementValue);
      return;
    }

    if (type === 'commander' && isCommanderIdentityField(key)) {
      renamedValue[key] = renameIdentityValue(nestedValue, sourceKey, replacementValue);
      return;
    }

    renamedValue[key] = renameIdentityValues(nestedValue, { sourceKey, replacementValue, type }, key);
  });

  return renamedValue;
}

function mergeDeckListsByCommander(deckLists) {
  const mergedDeckLists = new Map();

  (Array.isArray(deckLists) ? deckLists : [])
    .map(normalizeDeckListEntry)
    .filter(Boolean)
    .forEach((entry) => {
      const commanderKey = getIdentityKey(entry.commander);
      if (!commanderKey) {
        return;
      }

      if (!mergedDeckLists.has(commanderKey)) {
        mergedDeckLists.set(commanderKey, entry);
        return;
      }

      const existingEntry = mergedDeckLists.get(commanderKey);
      mergedDeckLists.set(commanderKey, {
        ...existingEntry,
        commander: existingEntry.commander || entry.commander,
        owner: existingEntry.owner || entry.owner,
        url: existingEntry.url || entry.url,
      });
    });

  return Array.from(mergedDeckLists.values());
}

function renameCommanderPowerLevels(powerLevels, sourceKey, replacementValue) {
  const renamedPowerLevels = {};
  const replacementKey = getIdentityKey(replacementValue);
  let sourceValue = null;
  let replacementExistingValue = null;

  Object.entries(powerLevels && typeof powerLevels === 'object' ? powerLevels : {}).forEach(([commander, value]) => {
    if (typeof value !== 'number' || Number.isNaN(value)) {
      return;
    }

    const commanderKey = getIdentityKey(commander);
    if (!commanderKey) {
      return;
    }

    if (commanderKey === sourceKey) {
      sourceValue = value;
      return;
    }

    if (commanderKey === replacementKey) {
      replacementExistingValue = value;
      return;
    }

    renamedPowerLevels[normalizeIdentityLabel(commander)] = value;
  });

  if (replacementExistingValue !== null || sourceValue !== null) {
    renamedPowerLevels[replacementValue] = replacementExistingValue !== null ? replacementExistingValue : sourceValue;
  }

  return renamedPowerLevels;
}

function renameAppStateIdentity(state, { type, sourceKey, replacementValue }) {
  const renamedDeckLists = renameIdentityValues(Array.isArray(state?.deckLists) ? state.deckLists : [], {
    sourceKey,
    replacementValue,
    type,
  }, 'deckLists');

  return {
    games: renameIdentityValues(Array.isArray(state?.games) ? state.games : [], {
      sourceKey,
      replacementValue,
      type,
    }, 'games'),
    powerLevels: type === 'commander'
      ? renameCommanderPowerLevels(state?.powerLevels, sourceKey, replacementValue)
      : { ...(state?.powerLevels && typeof state.powerLevels === 'object' ? state.powerLevels : {}) },
    deckLists: type === 'commander' ? mergeDeckListsByCommander(renamedDeckLists) : renamedDeckLists,
    records: renameIdentityValues(Array.isArray(state?.records) ? state.records : [], {
      sourceKey,
      replacementValue,
      type,
    }, 'records'),
  };
}

function renameActiveGameStateIdentity(state, { type, sourceKey, replacementValue }) {
  if (!state || typeof state !== 'object') {
    return null;
  }

  return normalizeActiveGameStateData({
    ...state,
    players: (Array.isArray(state.players) ? state.players : []).map((player) => ({
      ...player,
      name: type === 'player' ? renameIdentityValue(player?.name, sourceKey, replacementValue) : player?.name,
      commander: type === 'commander' ? renameIdentityValue(player?.commander, sourceKey, replacementValue) : player?.commander,
      killedPlayers: type === 'player' ? renameIdentityList(player?.killedPlayers || [], sourceKey, replacementValue) : player?.killedPlayers || [],
    })),
  });
}

function getUniqueValuesBySimilarity(values, threshold = 0.95) {
  const buckets = [];

  (Array.isArray(values) ? values : []).forEach((value) => {
    const normalized = String(value || '').trim();
    if (!normalized) {
      return;
    }

    const normalizedKey = getIdentityKey(normalized);
    if (!normalizedKey) {
      return;
    }

    let bucket = buckets.find((entry) => getStringSimilarity(normalizedKey, entry.key) >= threshold);
    if (!bucket) {
      bucket = { key: normalizedKey, variants: new Map() };
      buckets.push(bucket);
    }

    bucket.variants.set(normalized, (bucket.variants.get(normalized) || 0) + 1);
  });

  const unique = buckets.map((bucket) => {
    let bestValue = '';
    let bestScore = -1;
    let bestCount = -1;

    bucket.variants.forEach((count, value) => {
      const score = getIdentityDisplayScore(value);
      if (
        score > bestScore
        || (score === bestScore && count > bestCount)
        || (score === bestScore && count === bestCount && value.localeCompare(bestValue) < 0)
      ) {
        bestValue = value;
        bestScore = score;
        bestCount = count;
      }
    });

    return bestValue;
  });

  return unique.sort((a, b) => a.localeCompare(b));
}

function getKnownPlayerOptions() {
  const players = [...knownPlayers];
  loadRecords().forEach((record) => {
    if (record?.holder) {
      players.push(record.holder);
    }
    if (record?.manualHolder) {
      players.push(record.manualHolder);
    }
  });
  (activeGameState?.players || []).forEach((player) => {
    if (player?.name) {
      players.push(player.name);
    }
  });
  return getUniqueValuesBySimilarity(Array.from(buildCanonicalIdentityMapFromValues(players).values()));
}

function getKnownCommanderOptions() {
  const commanders = [...knownCommanders, ...Object.keys(loadCommanderPowerLevels())];
  loadRecords().forEach((record) => {
    if (record?.commander) {
      commanders.push(record.commander);
    }
    if (record?.manualCommander) {
      commanders.push(record.manualCommander);
    }
  });
  (activeGameState?.players || []).forEach((player) => {
    if (player?.commander) {
      commanders.push(player.commander);
    }
  });
  return getUniqueValuesBySimilarity(Array.from(buildCanonicalIdentityMapFromValues(commanders).values()));
}

function renderIdentityRenameOptions() {
  buildDatalistOptions(playerRenameDatalist, getKnownPlayerOptions());
  buildDatalistOptions(commanderRenameDatalist, getKnownCommanderOptions());
  if (playerRenameMenu) {
    buildDropdownMenu(playerRenameMenu, getKnownPlayerOptions());
    attachLookupWrapperHandlers(playerRenameForm || document);
  }
  if (commanderRenameMenu) {
    buildDropdownMenu(commanderRenameMenu, getKnownCommanderOptions());
    attachLookupWrapperHandlers(commanderRenameForm || document);
  }
}

function getBestDeckListEntryForCommander(commanderName, entries) {
  if (!commanderName || !Array.isArray(entries) || !entries.length) {
    return null;
  }

  const inputKey = getIdentityKey(commanderName);
  if (!inputKey) {
    return null;
  }

  let best = { score: 0, entry: null };
  entries.forEach((entry) => {
    if (!entry?.commander) {
      return;
    }
    const entryKey = getIdentityKey(entry.commander);
    if (!entryKey) {
      return;
    }
    const score = getStringSimilarity(inputKey, entryKey);
    if (score > best.score) {
      best = { score, entry };
    }
  });

  return best.score >= 0.9 ? best.entry : null;
}

function getSavedDeckListEntryForCommander(commanderName) {
  const commanderKey = getIdentityKey(commanderName);
  if (!commanderKey) {
    return null;
  }

  const entries = loadDeckLists()
    .map(normalizeDeckListEntry)
    .filter(Boolean);

  const exactMatch = entries.find((entry) => getIdentityKey(entry.commander) === commanderKey) || null;
  if (exactMatch) {
    return exactMatch;
  }

  return getBestDeckListEntryForCommander(commanderName, entries);
}

function buildDeckListLinkOrText(label) {
  const displayLabel = String(label || '').trim();
  if (!displayLabel) {
    return '—';
  }

  // First try: find a deckList entry and resolve via that
  const deckListEntry = getSavedDeckListEntryForCommander(displayLabel);
  if (deckListEntry) {
    const linkedDeckId = resolveLinkedDeckIdForDeckList(deckListEntry);
    if (linkedDeckId) {
      return buildCommanderDisplayHtml(displayLabel, `<a class="history-drilldown-link" href="${escapeHtml(getDeckBuilderHref(linkedDeckId))}">${escapeHtml(displayLabel)}</a>`);
    }
  }

  // Second try: find a built deck directly by commander name (for games registered without a deck URL)
  const nameKey = getIdentityKey(displayLabel);
  if (nameKey) {
    const matchedDeck = loadDecks().find((deck) => {
      const deckNameKey = getIdentityKey(deck.commander?.name || '');
      return deckNameKey === nameKey;
    });
    if (matchedDeck) {
      return buildCommanderDisplayHtml(displayLabel, `<a class="history-drilldown-link" href="${escapeHtml(getDeckBuilderHref(matchedDeck.id))}">${escapeHtml(displayLabel)}</a>`);
    }
  }

  return buildCommanderDisplayHtml(displayLabel, escapeHtml(displayLabel));
}

function setIdentityRenameStatus(element, message, tone = 'muted') {
  if (!element) {
    return;
  }

  element.textContent = message;
  element.classList.remove('status-muted', 'status-success', 'status-error', 'status-neutral');
  element.classList.add(`status-${tone}`);
}

async function handleIdentityRenameSubmit({ type, currentInput, nextInput, statusElement }) {
  const subjectLabel = type === 'player' ? 'player' : 'commander';
  const currentValue = String(currentInput?.value || '').trim();
  const nextValue = String(nextInput?.value || '').trim();

  if (!currentValue || !nextValue) {
    setIdentityRenameStatus(statusElement, `Enter both the current ${subjectLabel} name and the new one.`, 'error');
    return;
  }

  const sourceKey = getIdentityKey(currentValue);
  const replacementKey = getIdentityKey(nextValue);
  const knownOptions = type === 'player' ? getKnownPlayerOptions() : getKnownCommanderOptions();
  const matchingCurrentValue = knownOptions.find((value) => getIdentityKey(value) === sourceKey) || '';
  const matchingReplacementValue = knownOptions.find((value) => getIdentityKey(value) === replacementKey) || '';

  if (!matchingCurrentValue) {
    setIdentityRenameStatus(statusElement, `Couldn't find that ${subjectLabel} in saved data.`, 'error');
    return;
  }

  const mergeNotice = matchingReplacementValue && replacementKey !== sourceKey
    ? ` Existing ${subjectLabel} data for ${matchingReplacementValue} will be merged into the renamed identity.`
    : '';
  const confirmed = await promptLiveConfirm(
    `Rename ${matchingCurrentValue} to ${nextValue}? This updates saved games, records, deck lists${type === 'commander' ? ', commander power levels,' : ''} and any live game saved on this device.${mergeNotice}`,
    {
      title: `Rename ${subjectLabel}`,
      confirmLabel: `Rename ${subjectLabel}`,
      cancelLabel: 'Cancel',
    },
  );

  if (!confirmed) {
    setIdentityRenameStatus(statusElement, `Rename cancelled. ${matchingCurrentValue} was not changed.`, 'muted');
    return;
  }

  appState = normalizeAppStateData(renameAppStateIdentity(appState, {
    type,
    sourceKey,
    replacementValue: nextValue,
  }));
  appState.records = mergeRecordsWithDefaults(appState.records, appState.games);
  persistLocalState(appState);

  persistActiveGameState(renameActiveGameStateIdentity(activeGameState, {
    type,
    sourceKey,
    replacementValue: nextValue,
  }));
  persistActiveGameUndoState(activeGameUndoState.map((snapshot) => renameActiveGameStateIdentity(snapshot, {
    type,
    sourceKey,
    replacementValue: nextValue,
  })).filter(Boolean));

  queueCloudSync();
  refresh();

  if (currentInput) {
    currentInput.value = nextValue;
  }
  if (nextInput) {
    nextInput.value = '';
  }

  setIdentityRenameStatus(statusElement, `${matchingCurrentValue} is now stored as ${nextValue}.`, 'success');
}

function buildPersistedAppStatePayload(state) {
  const normalizedState = normalizeAppStateData(state);
  return {
    games: normalizedState.games,
    powerLevels: normalizedState.powerLevels,
    deckLists: normalizedState.deckLists,
    decks: normalizedState.decks,
    records: normalizedState.records,
    activeGame: normalizedState.activeGame,
    activeGameUndo: normalizedState.activeGameUndo,
  };
}

function persistAppStateToStorage(state) {
  const normalizedState = normalizeAppStateData(state);
  const payload = buildPersistedAppStatePayload(normalizedState);

  writeLocalStorageValue(STORAGE_KEY, JSON.stringify(payload));
  writeLocalStorageValue(EXPECTED_POWER_STORAGE_KEY, JSON.stringify(normalizedState.powerLevels));
  writeLocalStorageValue(DECK_LIST_STORAGE_KEY, JSON.stringify(normalizedState.deckLists));
  writeLocalStorageValue(DECKS_STORAGE_KEY, JSON.stringify(normalizedState.decks));
  writeLocalStorageValue(RECORDS_STORAGE_KEY, JSON.stringify(normalizedState.records));

  if (normalizedState.activeGame) {
    writeLocalStorageValue(ACTIVE_GAME_STORAGE_KEY, JSON.stringify(normalizedState.activeGame));
  } else {
    removeLocalStorageValue(ACTIVE_GAME_STORAGE_KEY);
  }

  if (normalizedState.activeGameUndo.length) {
    writeLocalStorageValue(ACTIVE_GAME_UNDO_STORAGE_KEY, JSON.stringify(normalizedState.activeGameUndo));
  } else {
    removeLocalStorageValue(ACTIVE_GAME_UNDO_STORAGE_KEY);
  }

  return normalizedState;
}

function loadLocalState() {
  const storagePayload = parseJsonSafe(readLocalStorageValue(STORAGE_KEY), null);
  const legacyState = Array.isArray(storagePayload)
    ? { games: storagePayload }
    : storagePayload && typeof storagePayload === 'object'
      ? storagePayload
      : null;

  if (!legacyState) {
    return normalizeAppStateData(appState);
  }

  const fallbackState = {
    games: Array.isArray(legacyState.games) ? legacyState.games : [],
    powerLevels: legacyState.powerLevels && typeof legacyState.powerLevels === 'object' ? legacyState.powerLevels : {},
    deckLists: Array.isArray(legacyState.deckLists) ? legacyState.deckLists : [],
    decks: Array.isArray(legacyState.decks) ? legacyState.decks : [],
    records: Array.isArray(legacyState.records) ? legacyState.records : [],
    activeGame: legacyState.activeGame && typeof legacyState.activeGame === 'object' ? legacyState.activeGame : null,
    activeGameUndo: Array.isArray(legacyState.activeGameUndo) ? legacyState.activeGameUndo : [],
  };

  return normalizeAppStateData(fallbackState);
}

function persistLocalState(state) {
  appState = normalizeAppStateData({
    ...appState,
    ...(state && typeof state === 'object' ? state : {}),
  });
  persistAppStateToStorage(appState);
}

function hasAnyStateData(state = appState) {
  const normalizedState = normalizeAppStateData(state && typeof state === 'object' ? state : {});
  return (
    (Array.isArray(normalizedState.games) && normalizedState.games.length > 0)
    || (normalizedState.powerLevels && Object.keys(normalizedState.powerLevels).length > 0)
    || (Array.isArray(normalizedState.deckLists) && normalizedState.deckLists.length > 0)
    || (Array.isArray(normalizedState.decks) && normalizedState.decks.length > 0)
    || (Array.isArray(normalizedState.records) && normalizedState.records.length > 0)
    || Boolean(normalizedState.activeGame)
    || (Array.isArray(normalizedState.activeGameUndo) && normalizedState.activeGameUndo.length > 0)
  );
}

function loadActiveGameState() {
  return normalizeActiveGameStateData(appState?.activeGame);
}

function loadActiveGameUndoState() {
  return Array.isArray(appState?.activeGameUndo)
    ? appState.activeGameUndo.map((entry) => normalizeActiveGameStateData(entry)).filter(Boolean)
    : [];
}

function queueActiveGameCloudSync(delay = ACTIVE_GAME_PERSIST_DEBOUNCE_MS) {
  if (activeGamePersistTimer) {
    clearTimeout(activeGamePersistTimer);
    activeGamePersistTimer = null;
  }

  if (Number(delay) <= 0) {
    queueCloudSync(0, { gamesOnly: true });
    return;
  }

  activeGamePersistTimer = setTimeout(() => {
    activeGamePersistTimer = null;
    queueCloudSync(0, { gamesOnly: true });
  }, delay);
}

function persistActiveGameUndoState(state, { syncDelay = ACTIVE_GAME_PERSIST_DEBOUNCE_MS } = {}) {
  const normalizedState = Array.isArray(state)
    ? state.map((entry) => normalizeActiveGameStateData(cloneActiveGameState(entry))).filter(Boolean)
    : state
      ? [normalizeActiveGameStateData(cloneActiveGameState(state))]
      : [];

  activeGameUndoState = normalizedState;
  persistLocalState({ ...appState, activeGameUndo: normalizedState });
  queueActiveGameCloudSync(syncDelay);
}

function cloneActiveGameState(state) {
  return state ? parseJsonSafe(JSON.stringify(state), null) : null;
}

function saveUndoSnapshot() {
  if (!activeGameState) {
    persistActiveGameUndoState(null, { syncDelay: 0 });
    return;
  }

  persistActiveGameUndoState([
    ...activeGameUndoState,
    cloneActiveGameState(activeGameState),
  ], {
    syncDelay: ACTIVE_GAME_PERSIST_DEBOUNCE_MS,
  });
}

function persistActiveGameState(state, { syncDelay = ACTIVE_GAME_PERSIST_DEBOUNCE_MS } = {}) {
  const normalizedState = normalizeActiveGameStateData(state);
  activeGameState = normalizedState;
  persistLocalState({ ...appState, activeGame: normalizedState });
  queueActiveGameCloudSync(syncDelay);
}

function queueActiveGameStatePersist(state, delay = ACTIVE_GAME_PERSIST_DEBOUNCE_MS) {
  const normalizedState = normalizeActiveGameStateData(state);
  activeGameState = normalizedState;
  persistLocalState({ ...appState, activeGame: normalizedState });
  queueActiveGameCloudSync(delay);
}

function flushQueuedActiveGamePersist() {
  if (!activeGamePersistTimer) {
    return;
  }

  clearTimeout(activeGamePersistTimer);
  activeGamePersistTimer = null;
  queueCloudSync(0, { gamesOnly: true });
}

function buildLocalSyncAuthSnapshot(user, { role = '' } = {}) {
  const displayName = String(user || '').trim();
  const normalizedUserId = normalizeSyncUserKey(displayName);
  const resolvedRole = String(role || '').trim().toLowerCase() || (
    normalizedUserId === BUILTIN_ADMIN_USER_KEY
      ? 'admin'
      : 'member'
  );

  return {
    userId: normalizedUserId,
    displayName,
    role: resolvedRole,
  };
}

function applyStoredSyncSession(user, token = '', { role = '' } = {}) {
  const normalizedUser = String(user || '').trim();
  const normalizedToken = String(token || '').trim();

  if (!normalizedUser) {
    return null;
  }

  if (normalizedToken) {
    persistSyncCredentials(normalizedUser, normalizedToken);
  }

  const auth = buildLocalSyncAuthSnapshot(normalizedUser, { role });
  updateSyncAuthenticatedUser(auth);
  syncConnectionState = 'connected';
  syncLastErrorMessage = '';
  refreshSyncStatus();
  return auth;
}

function getSyncCredentials() {
  const storedUser = String(readLocalStorageValue(SYNC_USER_STORAGE_KEY) || '').trim();
  const storedToken = String(readLocalStorageValue(SYNC_TOKEN_STORAGE_KEY) || '').trim();
  return {
    user: String(syncAuthenticatedDisplayName || syncUserInput?.value || storedUser || '').trim(),
    token: String(syncTokenInput?.value || storedToken || '').trim(),
  };
}

function hasSyncCredentials() {
  const credentials = getSyncCredentials();
  return Boolean(getCurrentSyncUserId() || (credentials.user && credentials.token));
}

function persistSyncCredentials(user, token) {
  const normalizedUser = String(user || '').trim();
  const normalizedToken = String(token || '').trim();

  if (!normalizedUser || !normalizedToken) {
    return;
  }

  writeLocalStorageValue(SYNC_USER_STORAGE_KEY, normalizedUser);
  writeLocalStorageValue(SYNC_TOKEN_STORAGE_KEY, normalizedToken);
  writeLocalStorageValue(SYNC_CREDENTIAL_SET_AT_STORAGE_KEY, new Date().toISOString());
}

function clearPersistedSyncCredentials() {
  removeLocalStorageValue(SYNC_USER_STORAGE_KEY);
  removeLocalStorageValue(SYNC_TOKEN_STORAGE_KEY);
  removeLocalStorageValue(SYNC_CREDENTIAL_SET_AT_STORAGE_KEY);
}

function formatSyncTimestamp(value) {
  if (!value) {
    return '';
  }

  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }

  return date.toLocaleTimeString([], {
    hour: 'numeric',
    minute: '2-digit',
  });
}

function updateSyncMetadata({ revision = 0, updatedAt = '', updatedBy = '' } = {}) {
  syncCloudRevision = Number.isFinite(Number(revision)) ? Number(revision) : 0;
  syncCloudUpdatedAt = String(updatedAt || '').trim();
  syncCloudUpdatedBy = String(updatedBy || '').trim();

  // Persist the newest cloud revision we have seen (plus when we saw it) so a stale
  // edge-cached read on the NEXT page load cannot silently overwrite newer local data
  // (Cloudflare KV reads can be served stale for a while after a write).
  if (syncCloudRevision > 0) {
    writeLocalStorageValue(SYNC_REVISION_STORAGE_KEY, JSON.stringify({
      revision: syncCloudRevision,
      recordedAt: Date.now(),
    }));
  } else {
    removeLocalStorageValue(SYNC_REVISION_STORAGE_KEY);
  }
}

function readPersistedSyncRevisionEntry() {
  const rawValue = String(readLocalStorageValue(SYNC_REVISION_STORAGE_KEY) || '').trim();
  if (!rawValue) {
    return null;
  }

  // Backward compatible with the original plain-number format. Those entries have no
  // trustworthy timestamp, so report them as expired (recordedAt 0) — this lets any
  // client currently stuck on local cache resync from the cloud immediately.
  const plainNumber = Number.parseInt(rawValue, 10);
  if (Number.isFinite(plainNumber) && String(plainNumber) === rawValue) {
    return { revision: plainNumber, recordedAt: 0 };
  }

  try {
    const parsed = JSON.parse(rawValue);
    const revision = Number.parseInt(String(parsed?.revision ?? ''), 10);
    const recordedAt = Number(parsed?.recordedAt);
    if (Number.isFinite(revision) && revision > 0) {
      return { revision, recordedAt: Number.isFinite(recordedAt) ? recordedAt : 0 };
    }
  } catch (error) {
    // Fall through.
  }
  return null;
}

function restorePersistedSyncRevision() {
  const entry = readPersistedSyncRevisionEntry();
  if (entry && entry.revision > syncCloudRevision) {
    syncCloudRevision = entry.revision;
  }
}

// The stale-read guard is only trustworthy while a KV edge-cache read could still be
// serving a pre-write value. Past that window, a cloud revision lower than our stored
// one means a genuine desync (e.g. a backup restore) and the cloud must win so the
// client self-heals instead of showing local cache forever.
function isStaleReadGuardActive() {
  const entry = readPersistedSyncRevisionEntry();
  if (!entry || entry.revision <= 0) {
    return false;
  }
  return Date.now() - entry.recordedAt < 5 * 60 * 1000;
}

function updateSyncAuthenticatedUser(auth = null) {
  const previousUserId = String(syncAuthenticatedUserId || '').trim().toLowerCase();
  const nextUserId = String(auth?.userId || '').trim().toLowerCase();

  syncAuthenticatedUserId = String(auth?.userId || '').trim().toLowerCase();
  syncAuthenticatedDisplayName = String(auth?.displayName || '').trim();
  syncAuthenticatedRole = String(auth?.role || '').trim().toLowerCase();

  if (previousUserId !== nextUserId) {
    window.CommanderDeckLibrary?.resetFilterForUserChange({ clearSelection: true });
    window.CommanderHistory?.resetPlayerFilterForUserChange();
  }

  updateAdminLogsShortcutVisibility();
  document.dispatchEvent(new CustomEvent('commander:auth-changed'));
}

function clearSyncAuthenticatedUser() {
  updateSyncAuthenticatedUser(null);
  authAuditLogs = [];
  authAuditLoading = false;
  authAuditLastLoadedAt = 0;
  authAuditLoadedForUserId = '';
  registeredAccounts = [];
  registeredAccountsLoading = false;
  registeredAccountsLastLoadedAt = 0;
  registeredAccountsLoadedForUserId = '';
  window.CommanderDeckLibrary?.resetFilterForUserChange();
  window.CommanderHistory?.resetPlayerFilterForUserChange();
}

function getCurrentSyncUserId() {
  return String(syncAuthenticatedUserId || '').trim().toLowerCase();
}

function getCurrentSyncDisplayName() {
  return String(syncAuthenticatedDisplayName || '').trim() || normalizeIdentityLabel(syncUserInput?.value || '');
}

function isCurrentSyncUserAdmin() {
  if (syncAuthenticatedRole === 'admin') {
    return true;
  }

  // Only trust server-authenticated identity signals for admin detection.
  // Do not trust the locally typed credential username.
  const candidates = [
    syncAuthenticatedUserId,
    syncAuthenticatedDisplayName,
  ];

  return candidates.some((value) => normalizeSyncUserKey(value) === BUILTIN_ADMIN_USER_KEY);
}

function clearSyncConflict() {
  syncConflictInfo = null;
}

function hasNewerCloudRevision(metadata) {
  const revision = Number(metadata?.revision);
  return Number.isFinite(revision) && revision > syncCloudRevision;
}

// True when the cloud state is on a different baseline than this device's local
// state, meaning a blind push of local state could overwrite newer cloud data.
// Covers both revision directions: cloud ahead (another device synced) and cloud
// rolled BACK (a KV backup restore lowers the revision below what this device last
// saw). Also catches the corner case where the revision matches but the cloud copy
// was rewritten after we last saw it — impossible through normal writes, since
// every write bumps the revision, so it only happens after a restore/rollback.
function isCloudStateDesynced(metadata) {
  if (!metadata) {
    return false;
  }

  const revision = Number(metadata.revision);
  if (Number.isFinite(revision) && revision !== syncCloudRevision) {
    return true;
  }

  const updatedAtMs = Date.parse(String(metadata.updatedAt || ''));
  const entry = readPersistedSyncRevisionEntry();
  return Number.isFinite(updatedAtMs)
    && Boolean(entry?.recordedAt)
    && updatedAtMs > entry.recordedAt + 60000;
}

function describeSyncConflict(conflict = syncConflictInfo) {
  if (!conflict) {
    return 'Cloud state changed on another device.';
  }

  const byText = conflict.updatedBy ? ` by ${conflict.updatedBy}` : '';
  const atText = conflict.updatedAt ? ` at ${formatSyncTimestamp(conflict.updatedAt)}` : '';
  return `Cloud state changed on another device${byText}${atText}.`;
}

function clearSyncRetryTimer() {
  if (syncRetryTimer) {
    clearTimeout(syncRetryTimer);
    syncRetryTimer = null;
  }
}

function getSyncStatusSnapshot() {
  const credentials = getSyncCredentials();
  const lastSyncedText = syncLastSuccessAt ? formatSyncTimestamp(syncLastSuccessAt) : '';
  const storageWarning = getStorageWarningMessage();
  const credentialAgeDays = getSyncCredentialAgeDays();

  if (storageWarning) {
    return {
      tone: 'error',
      message: storageWarning,
    };
  }

  if (!hasSyncCredentials()) {
    return {
      tone: 'muted',
      message: 'Cloud sync not connected.',
    };
  }

  if (syncBootstrapInFlight) {
    return {
      tone: 'neutral',
      message: 'Syncing cloud data…',
    };
  }

  if (!navigator.onLine) {
    if (syncPendingChanges) {
      return {
        tone: 'error',
        message: 'Offline. Local changes are saved on this device and will sync when you reconnect.',
      };
    }

    return {
      tone: 'muted',
      message: lastSyncedText
        ? `Offline. Last synced at ${lastSyncedText}.`
        : 'Offline. Cloud sync will resume when you reconnect.',
    };
  }

  if (syncInFlight) {
    return {
      tone: 'neutral',
      message: syncConnectionState === 'connecting'
        ? 'Connecting to cloud and loading latest state...'
        : 'Syncing local changes to cloud...',
    };
  }

  if (syncLastErrorMessage) {
    return {
      tone: 'error',
      message: syncRetryTimer
        ? `Sync failed: ${syncLastErrorMessage}. Local changes are still saved and will retry automatically.`
        : `Sync failed: ${syncLastErrorMessage}. Local changes are still saved on this device.`,
    };
  }

  if (syncConflictInfo) {
    return {
      tone: 'error',
      message: `${describeSyncConflict(syncConflictInfo)} Local changes are still on this device and will not sync until you resolve the conflict. Use Sync now to review the latest cloud state.`,
    };
  }

  if (hasSyncCredentials() && !syncHasLoadedCloudState) {
    return {
      tone: 'neutral',
      message: 'Cloud sync is connected but waiting to load latest cloud state before sending edits.',
    };
  }

  if (syncPendingChanges) {
    return {
      tone: 'neutral',
      message: 'Local changes are queued and will sync automatically.',
    };
  }

  if (syncConnectionState === 'connected') {
    const connectedAs = getCurrentSyncDisplayName() || credentials.user;
    return {
      tone: credentialAgeDays !== null && credentialAgeDays >= 7 ? 'neutral' : 'success',
      message: credentialAgeDays !== null && credentialAgeDays >= 7
        ? `Cloud sync is still connected as ${connectedAs}. Shared device? Disconnect when you're done.`
        : (lastSyncedText
          ? `Cloud sync connected. Last synced at ${lastSyncedText}.`
          : `Cloud sync connected as ${connectedAs}.`),
    };
  }

  return {
    tone: 'muted',
    message: 'Cloud sync is configured. Pulling latest state when needed.',
  };
}

function refreshSyncStatus() {
  const snapshot = getSyncStatusSnapshot();
  setSyncStatus(snapshot.message, snapshot.tone);
  const syncNeedsAttention = syncInFlight
    || syncBootstrapInFlight
    || syncConnectionState === 'connecting'
    || syncPendingChanges
    || Boolean(syncQueueTimer)
    || Boolean(syncRetryTimer)
    || Boolean(syncConflictInfo)
    || snapshot.tone === 'error'
    || (hasSyncCredentials() && !syncHasLoadedCloudState);

  if (syncNeedsAttention) {
    setSyncUiCollapsed(false);
  } else if (syncConnectionState === 'connected') {
    setSyncUiCollapsed(true);
  }
  updateAdminLogsShortcutVisibility();
}

function shouldAutoPullCloudState() {
  return !syncConflictInfo && (!syncPendingChanges || !syncHasLoadedCloudState);
}

function setSyncBootstrapInFlight(isInFlight) {
  syncBootstrapInFlight = Boolean(isInFlight);
  refreshSyncStatus();
}

async function retryCloudBootstrap({ maxAttempts = 4, initialDelayMs = 1200 } = {}) {
  if (!hasSyncCredentials() || !navigator.onLine) {
    return false;
  }

  syncLastErrorMessage = '';
  setSyncBootstrapInFlight(true);
  window.CommanderDeckLibrary?.render?.();

  let lastError = null;
  try {
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      try {
        await pullCloudState();
        return true;
      } catch (error) {
        lastError = error;
        if (attempt >= maxAttempts) {
          break;
        }

        const delayMs = initialDelayMs * attempt;
        await new Promise((resolve) => {
          window.setTimeout(resolve, delayMs);
        });
      }
    }
  } finally {
    setSyncBootstrapInFlight(false);
  }

  if (lastError) {
    syncConnectionState = lastError.status === 401 ? 'local' : 'configured';
    syncLastErrorMessage = lastError.status === 401 ? '' : `${lastError.message}.`;
    refreshSyncStatus();
    window.CommanderDeckLibrary?.render?.();
  }

  return false;
}

function setSyncStatus(message, tone = 'neutral') {
  if (!syncStatus) {
    return;
  }

  syncStatus.textContent = message;
  syncStatus.classList.remove('status-neutral', 'status-success', 'status-error', 'status-muted');
  syncStatus.classList.add(`status-${tone}`);
}

function setSyncPanelVisible(isVisible) {
  if (!syncPanel) {
    return;
  }

  syncPanel.hidden = !isVisible;
}

function setSyncUiCollapsed(isCollapsed) {
  setSyncPanelVisible(!isCollapsed);
}

function updateSyncControls() {
  if (!syncUserInput || !syncTokenInput) {
    return;
  }

  const hasCredentials = hasSyncCredentials();
  const canAttemptConnect = Boolean(syncUserInput.value.trim() && syncTokenInput.value.trim());
  const isBusy = syncInFlight || syncConnectionState === 'connecting';
  if (syncConnectButton) {
    syncConnectButton.disabled = isBusy || !canAttemptConnect;
  }
  if (syncDisconnectButton) {
    syncDisconnectButton.disabled = !hasCredentials || isBusy;
  }
  if (syncNowButton) {
    syncNowButton.disabled = !hasCredentials || isBusy;
  }
}

async function cloudRequest(path, options = {}) {
  const headers = new Headers(options.headers || {});
  const credentials = getSyncCredentials();
  const timeoutMs = Number.isFinite(Number(options.timeoutMs)) ? Number(options.timeoutMs) : 0;
  const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
  const signal = options.signal || controller?.signal;
  const timeoutId = timeoutMs > 0 && controller
    ? window.setTimeout(() => {
      controller.abort();
    }, timeoutMs)
    : null;

  if (!headers.has('X-User-Name') && credentials.user) {
    headers.set('X-User-Name', credentials.user);
  }

  if (!headers.has('X-Pod-Token') && credentials.token) {
    headers.set('X-Pod-Token', credentials.token);
  }

  if (!headers.has('Content-Type') && options.method !== 'GET') {
    headers.set('Content-Type', 'application/json');
  }

  try {
    const response = await fetch(path, {
      ...options,
      headers,
      cache: 'no-store',
      credentials: 'same-origin',
      signal,
    });

    if (!response.ok) {
      let message = `Request failed (${response.status})`;
      let errorBody = null;
      try {
        errorBody = await response.json();
        if (errorBody && errorBody.error) {
          message = errorBody.error;
        }
      } catch (error) {
        // Keep default message.
      }

      const requestError = new Error(message);
      requestError.status = response.status;
      requestError.body = errorBody;
      throw requestError;
    }

    return response.json();
  } catch (error) {
    if (error?.name === 'AbortError') {
      const timeoutError = new Error(`Request timed out after ${timeoutMs}ms.`);
      timeoutError.status = 504;
      throw timeoutError;
    }

    throw error;
  } finally {
    if (timeoutId) {
      window.clearTimeout(timeoutId);
    }
  }
}

async function resolveSyncConflict(conflict = syncConflictInfo) {
  syncConflictInfo = conflict || {
    revision: syncCloudRevision,
    updatedAt: syncCloudUpdatedAt,
    updatedBy: syncCloudUpdatedBy,
  };
  syncConnectionState = 'configured';
  syncLastErrorMessage = '';
  clearSyncRetryTimer();
  syncRetryCount = 0;
  refreshSyncStatus();

  const shouldPullLatest = await promptLiveConfirm(
    `${describeSyncConflict(syncConflictInfo)} Pull the latest cloud state now? This will replace unsynced local changes on this device.`,
    {
      title: 'Cloud sync conflict',
      confirmLabel: 'Pull latest cloud state',
      cancelLabel: 'Keep local changes',
    },
  );

  if (!shouldPullLatest) {
    refreshSyncStatus();
    return;
  }

  setSyncPendingChanges(false);
  clearSyncConflict();
  refreshSyncStatus();

  try {
    await pullCloudState({ force: true });
    setSyncUiCollapsed(true);
    refreshSyncStatus();
  } catch (error) {
    syncConflictInfo = conflict || syncConflictInfo;
    syncLastErrorMessage = `${error.message}. Unable to pull the latest cloud state.`;
    refreshSyncStatus();
  }
}

async function pullCloudState({ force = false } = {}) {
  syncConnectionState = 'connecting';
  syncLastErrorMessage = '';
  updateSyncControls();
  refreshSyncStatus();

  try {
    // The full state payload is large (games + decks + decklists with card data),
    // and the abort stays armed through body download. A 1.5s budget aborted
    // legitimate pulls on slower connections, leaving every page on stale local
    // cache indefinitely. Give the pull a budget comparable to the push (10s).
    const payload = await cloudRequest(CLOUD_SYNC_ENDPOINT, { method: 'GET', timeoutMs: 15000 });
    const statePayload = payload?.state && typeof payload.state === 'object'
      ? payload.state
      : payload;
    updateSyncAuthenticatedUser(payload?.auth || statePayload?.auth || null);

    const pulledRevision = Number.isFinite(Number(payload?.revision ?? statePayload?.revision))
      ? Number(payload?.revision ?? statePayload?.revision)
      : 0;

    // Guard against stale edge-cached KV reads: if the cloud just returned an OLDER
    // revision than the newest one this device saw moments ago, applying it would
    // wipe newer local data (the "my update disappeared for minutes" loop). The guard
    // is time-bounded — past the KV staleness window a lower cloud revision means a
    // real desync (e.g. backup restore) and the cloud state is applied so the client
    // resyncs instead of showing local cache indefinitely.
    if (!force && pulledRevision < syncCloudRevision && isStaleReadGuardActive()) {
      if (window.CommanderDeckLibrary?.hasView?.()) {
        throw new Error('Cloud returned a stale state snapshot. Retrying the pull.');
      }
      syncConnectionState = 'connected';
      syncHasLoadedCloudState = true;
      syncLastErrorMessage = '';
      refreshSyncStatus();
      refresh();
      syncDebugLogs.unshift({
        timestamp: new Date().toISOString(),
        event: 'Skipped stale cloud read',
        decks: Array.isArray(appState.decks) ? appState.decks.length : 0,
        lists: Array.isArray(appState.deckLists) ? appState.deckLists.length : 0,
        user: `cloud rev ${pulledRevision} < local rev ${syncCloudRevision}`,
      });
      renderSyncDebugLog();
      return;
    }

    const games = Array.isArray(statePayload?.games) ? statePayload.games : [];
    const powerLevels = statePayload?.powerLevels && typeof statePayload.powerLevels === 'object' ? statePayload.powerLevels : {};
    const deckLists = Array.isArray(statePayload?.deckLists) ? statePayload.deckLists : [];
    const decks = Array.isArray(statePayload?.decks) ? statePayload.decks : [];
    const records = Array.isArray(statePayload?.records) ? statePayload.records : [];
    const activeGame = statePayload?.activeGame && typeof statePayload.activeGame === 'object' ? statePayload.activeGame : null;
    const activeGameUndo = Array.isArray(statePayload?.activeGameUndo) ? statePayload.activeGameUndo : [];
    updateSyncMetadata({
      revision: payload?.revision ?? statePayload?.revision,
      updatedAt: payload?.updatedAt ?? statePayload?.updatedAt,
      updatedBy: payload?.updatedBy ?? statePayload?.updatedBy,
    });
    clearSyncConflict();
    appState = normalizeAppStateData({ games, powerLevels, deckLists, decks, records, activeGame, activeGameUndo });
    persistLocalState(appState);
    activeGameState = loadActiveGameState();
    activeGameUndoState = loadActiveGameUndoState();
    if (activeGameState) {
      acquireWakeLock();
    } else {
      releaseWakeLock();
    }
    setSyncPendingChanges(false);
    syncHasLoadedCloudState = true;
    syncLocalContentVerified = true;
    syncConnectionState = 'connected';
    syncLastErrorMessage = '';
    syncLastSuccessAt = new Date().toISOString();
    // A pull can also surface empty "Untitled Deck" records that were pushed before
    // the page-exit leak was fixed; prune them and let the deferred sync clean KV.
    pruneStaleEmptyDeckRecords({ deferCloudSync: true });
    refresh();
    syncDebugLogs.unshift({
      timestamp: new Date().toISOString(),
      event: 'Cloud sync',
      decks: decks.length,
      lists: deckLists.length,
      user: getCurrentSyncUserId() || 'Unknown',
    });
    renderSyncDebugLog();
  } catch (error) {
    if (error.status === 401) {
      clearSyncAuthenticatedUser();
      syncHasLoadedCloudState = false;
      syncLocalContentVerified = false;
    }
    syncConnectionState = getCurrentSyncUserId() ? 'configured' : 'local';
    throw error;
  } finally {
    updateSyncControls();
  }
}

async function refreshSessionStatus() {
  const storedUser = String(readLocalStorageValue(SYNC_USER_STORAGE_KEY) || '').trim();
  const storedToken = String(readLocalStorageValue(SYNC_TOKEN_STORAGE_KEY) || '').trim();

  if (storedUser && storedToken) {
    applyStoredSyncSession(storedUser, storedToken);
  }

  try {
    const payload = await cloudRequest('/api/session', { method: 'GET', timeoutMs: 10000 });
    if (payload?.auth) {
      updateSyncAuthenticatedUser(payload.auth || null);
      syncConnectionState = 'connected';
      syncLastErrorMessage = '';
      refreshSyncStatus();
    }
    return payload.auth || null;
  } catch (error) {
    if (error.status === 401) {
      if (!hasSyncCredentials()) {
        clearSyncAuthenticatedUser();
      }
      syncConnectionState = 'local';
      refreshSyncStatus();
      return null;
    }

    if (error?.status === 504 || error?.name === 'AbortError') {
      return null;
    }

    throw error;
  }
}

function pushCloudStateKeepalive({ force = false } = {}) {
  if (!hasSyncCredentials() || !navigator.onLine || syncConflictInfo) {
    return;
  }

  // Never keepalive-push before this session has loaded cloud state — not even on
  // a forced page-exit flush. The persisted local cache can be older than the
  // cloud (e.g. after a KV backup restore), and if its stored revision happens to
  // match the restored one, the worker would accept the stale state and silently
  // overwrite newer cloud data. Unsynced changes stay in local storage and
  // reconcile on the next page load instead.
  if (!syncHasLoadedCloudState) {
    return;
  }

  if (!force && !syncPendingChanges && !decksPersistTimer) {
    return;
  }

  const headers = new Headers();
  headers.set('Content-Type', 'application/json');
  headers.set('X-State-Revision', String(syncCloudRevision));

  fetch(CLOUD_SYNC_ENDPOINT, {
    method: 'PUT',
    headers,
    body: JSON.stringify({
      games: appState.games,
      powerLevels: appState.powerLevels,
      deckLists: appState.deckLists,
      decks: appState.decks,
      records: appState.records,
      activeGame: appState.activeGame,
      activeGameUndo: appState.activeGameUndo,
    }),
    keepalive: true,
    credentials: 'same-origin',
  }).catch(() => {
    // Best-effort unload flush; normal retries still run after navigation.
  });
}

async function fetchCloudStateMetadata() {
  // Tiny payload, but keep the timeout generous: this check gates every
  // fast-path decision, and a false abort reads as "no update" and leaves
  // pages stale until the next refresh.
  const payload = await cloudRequest(CLOUD_SYNC_METADATA_ENDPOINT, { method: 'GET', timeoutMs: 10000 });
  updateSyncAuthenticatedUser(payload.auth || null);
  return {
    revision: Number.isFinite(Number(payload?.revision)) ? Number(payload.revision) : 0,
    updatedAt: String(payload?.updatedAt || '').trim(),
    updatedBy: String(payload?.updatedBy || '').trim(),
    counts: payload?.counts && typeof payload.counts === 'object' ? payload.counts : null,
  };
}

async function checkCloudStateFreshness({ autoPull = false, force = false } = {}) {
  if (!hasSyncCredentials() || !navigator.onLine || syncInFlight || syncMetadataCheckInFlight) {
    return;
  }

  const now = Date.now();
  if (!force && now - syncLastFreshnessCheckAt < 15000) {
    return;
  }

  syncMetadataCheckInFlight = true;
  syncLastFreshnessCheckAt = now;

  try {
    const metadata = await fetchCloudStateMetadata();
    if (!hasNewerCloudRevision(metadata)) {
      return;
    }

    updateSyncMetadata(metadata);

    if (syncHasLoadedCloudState && syncPendingChanges) {
      clearSyncRetryTimer();
      if (syncQueueTimer) {
        clearTimeout(syncQueueTimer);
        syncQueueTimer = null;
      }
      // The cloud moved ahead while this device has unsynced changes. When no
      // fragile in-progress work exists (deck draft / live game), heal right away
      // by pulling the authoritative cloud state instead of freezing sync behind
      // a conflict notice.
      if (canAutoPullOverLocalChanges()) {
        await autoResolveConflictByPulling(metadata, 'cloud-moved-ahead');
        return;
      }
      syncConflictInfo = metadata;
      syncLastErrorMessage = '';
      refreshSyncStatus();
      return;
    }

    if (autoPull) {
      await pullCloudState();
      setSyncUiCollapsed(true);
      refreshSyncStatus();
      return;
    }

    syncConnectionState = 'configured';
    syncLastErrorMessage = '';
    refreshSyncStatus();
  } catch (error) {
    if (force) {
      syncLastErrorMessage = error.message;
      refreshSyncStatus();
    }
  } finally {
    syncMetadataCheckInFlight = false;
  }
}

function scheduleSyncRetry(errorMessage) {
  clearSyncRetryTimer();

  if (!hasSyncCredentials() || !navigator.onLine) {
    return;
  }

  syncRetryCount += 1;
  const retryDelay = Math.min(30000, 1000 * (2 ** Math.min(syncRetryCount - 1, 4)));
  syncLastErrorMessage = errorMessage;
  refreshSyncStatus();

  syncRetryTimer = setTimeout(() => {
    syncRetryTimer = null;
    pushCloudState();
  }, retryDelay);
}

// Revision equality does not prove content equality: background metadata polls
// advance syncCloudRevision without downloading state, so a device can hold a
// stale local snapshot while its revision marker looks current. Before pushing
// pending changes from an unverified baseline, confirm the local state actually
// contains every item the cloud has — otherwise the push would overwrite the
// cloud with the stale snapshot.
function getLocalStateCounts() {
  return {
    games: Array.isArray(appState.games) ? appState.games.length : 0,
    deckLists: Array.isArray(appState.deckLists) ? appState.deckLists.length : 0,
    decks: Array.isArray(appState.decks) ? appState.decks.length : 0,
    records: Array.isArray(appState.records) ? appState.records.length : 0,
    powerLevels: appState.powerLevels && typeof appState.powerLevels === 'object' && !Array.isArray(appState.powerLevels)
      ? Object.keys(appState.powerLevels).length
      : 0,
  };
}

function localStateMatchesCloudCounts(counts) {
  if (!counts || typeof counts !== 'object') {
    // Older worker without counts — fall back to revision-only trust.
    return true;
  }

  const local = getLocalStateCounts();
  return ['games', 'deckLists', 'decks', 'records', 'powerLevels'].every((key) => {
    const cloudCount = Number(counts[key]);
    return !Number.isFinite(cloudCount) || cloudCount === local[key];
  });
}

// When the cloud is authoritative (user is connected), local state should never
// win silently. The only exceptions are in-progress work that a pull would
// destroy mid-flight: an unsaved deck-builder draft or an active live game.
function canAutoPullOverLocalChanges() {
  return !hasUnsavedDeckBuilderChanges() && !activeGameState;
}

// Best-effort snapshot of the local state before an automatic pull replaces it,
// so an auto-heal never destroys local edits irrecoverably.
function stashLocalStateBeforeAutoPull(reason) {
  try {
    writeLocalStorageValue(SYNC_CONFLICT_BACKUP_STORAGE_KEY, JSON.stringify({
      reason: String(reason || ''),
      stashedAt: new Date().toISOString(),
      state: appState,
    }));
  } catch (error) {
    // localStorage quota is tight with a multi-MB state — the pull still proceeds.
  }
}

async function autoResolveConflictByPulling(conflict, reason) {
  stashLocalStateBeforeAutoPull(reason);
  clearSyncConflict();
  syncConnectionState = 'connecting';
  refreshSyncStatus();

  try {
    await pullCloudState({ force: true });
    setSyncUiCollapsed(true);
    syncDebugLogs.unshift({
      timestamp: new Date().toISOString(),
      event: 'Auto-healed from cloud',
      decks: Array.isArray(appState.decks) ? appState.decks.length : 0,
      lists: Array.isArray(appState.deckLists) ? appState.deckLists.length : 0,
      user: String(reason || ''),
    });
    renderSyncDebugLog();
  } catch (error) {
    // Pull failed — fall back to the manual conflict state so sync stays blocked
    // rather than risking a blind push of the stale local state.
    syncConflictInfo = conflict || syncConflictInfo;
    syncLastErrorMessage = `${error.message}. Unable to pull the latest cloud state.`;
  }

  refreshSyncStatus();
}

function countCloudItemsMissingFromLocal(cloudEntries, localEntries) {
  if (!Array.isArray(cloudEntries) || !cloudEntries.length) {
    return 0;
  }

  const localIds = new Set(
    (Array.isArray(localEntries) ? localEntries : [])
      .map((entry) => String(entry?.id || '').trim())
      .filter(Boolean),
  );
  return cloudEntries.reduce((missing, entry) => {
    const id = String(entry?.id || '').trim();
    return id && !localIds.has(id) ? missing + 1 : missing;
  }, 0);
}

async function verifyPendingStateCoversCloud() {
  let payload;
  try {
    payload = await cloudRequest(CLOUD_SYNC_ENDPOINT, { method: 'GET', timeoutMs: 15000 });
  } catch (error) {
    return { outcome: 'unknown', error };
  }

  const statePayload = payload?.state && typeof payload.state === 'object'
    ? payload.state
    : payload;
  const metadata = {
    revision: Number.isFinite(Number(payload?.revision ?? statePayload?.revision))
      ? Number(payload?.revision ?? statePayload?.revision)
      : 0,
    updatedAt: String((payload?.updatedAt ?? statePayload?.updatedAt) || '').trim(),
    updatedBy: String((payload?.updatedBy ?? statePayload?.updatedBy) || '').trim(),
  };
  const missingFromLocal = countCloudItemsMissingFromLocal(statePayload?.games, appState.games)
    + countCloudItemsMissingFromLocal(statePayload?.deckLists, appState.deckLists)
    + countCloudItemsMissingFromLocal(statePayload?.decks, appState.decks)
    + countCloudItemsMissingFromLocal(statePayload?.records, appState.records);

  // Small gaps are treated as intentional local deletes (the app deletes entries
  // one at a time); the worker-side regression guard remains the backstop for
  // anything larger.
  return {
    outcome: missingFromLocal > 3 ? 'pull' : 'push',
    metadata,
    missingFromLocal,
  };
}

async function pushCloudState() {
  if (!hasSyncCredentials() || syncInFlight || syncConflictInfo) {
    return;
  }

  if (!syncLocalContentVerified) {
    if (!syncPendingChanges) {
      try {
        await pullCloudState();
      } catch (error) {
        setSyncPendingChanges(true);
        syncLastErrorMessage = `${error.message}. Sync is blocked until latest cloud state loads.`;
        refreshSyncStatus();
        return;
      }

      if (!syncLocalContentVerified) {
        // The pull was skipped by the stale-read guard — without verified content
        // this push could overwrite the cloud with a stale snapshot. Stay queued.
        return;
      }
    } else {
      // Hold syncInFlight during verification so a concurrent queueCloudSync
      // can't slip a second unverified push through the GET window; any mutation
      // queued meanwhile bumps the version counters and triggers a follow-up sync.
      syncInFlight = true;
      updateSyncControls();
      let verification;
      try {
        verification = await verifyPendingStateCoversCloud();
      } finally {
        syncInFlight = false;
        updateSyncControls();
      }
      if (verification.outcome === 'pull') {
        if (verification.metadata.revision > 0) {
          updateSyncMetadata(verification.metadata);
        }
        syncDebugLogs.unshift({
          timestamp: new Date().toISOString(),
          event: 'Blocked stale push',
          decks: Array.isArray(appState.decks) ? appState.decks.length : 0,
          lists: Array.isArray(appState.deckLists) ? appState.deckLists.length : 0,
          user: `cloud holds ${verification.missingFromLocal} entries this device is missing`,
        });
        renderSyncDebugLog();
        // The cloud provably holds entries this device lacks — the local baseline
        // is stale. Heal automatically by pulling the authoritative state; only an
        // in-progress deck draft or live game keeps the manual conflict prompt.
        if (canAutoPullOverLocalChanges()) {
          await autoResolveConflictByPulling(verification.metadata, 'stale-local-content');
          return;
        }
        await resolveSyncConflict(verification.metadata);
        return;
      }
      if (verification.outcome !== 'push') {
        scheduleSyncRetry(verification.error?.message || 'Unable to verify cloud state before syncing.');
        return;
      }
      syncLocalContentVerified = true;
    }
  }

  const requestStateMutationVersion = syncStateMutationVersion;
  const requestDeckMutationVersion = syncDecksMutationVersion;
  const requestGamesOnly = syncQueuedGamesOnly;
  const localStateHasData = hasAnyStateData(appState);

  // Safety guard: never send a full empty-state overwrite before we've loaded
  // cloud state for this session.
  if (!requestGamesOnly && !syncHasLoadedCloudState && !localStateHasData) {
    try {
      await pullCloudState();
      return;
    } catch (error) {
      setSyncPendingChanges(true);
      syncLastErrorMessage = `${error.message}. Blocked empty local state push until cloud data can be loaded.`;
      refreshSyncStatus();
      return;
    }
  }

  let shouldQueueFollowUpSync = false;
  syncInFlight = true;
  syncConnectionState = 'connected';
  syncLastErrorMessage = '';
  updateSyncControls();
  refreshSyncStatus();

  try {
    const payload = await cloudRequest(CLOUD_SYNC_ENDPOINT, {
      method: 'PUT',
      timeoutMs: 10000,
      headers: {
        'X-State-Revision': String(syncCloudRevision),
      },
      body: JSON.stringify(requestGamesOnly
        ? {
          games: appState.games,
          activeGame: appState.activeGame,
          activeGameUndo: appState.activeGameUndo,
        }
        : {
          games: appState.games,
          powerLevels: appState.powerLevels,
          deckLists: appState.deckLists,
          decks: appState.decks,
          records: appState.records,
          activeGame: appState.activeGame,
          activeGameUndo: appState.activeGameUndo,
        }),
    });
    updateSyncAuthenticatedUser(payload.auth || null);
    updateSyncMetadata({
      revision: payload.revision,
      updatedAt: payload.updatedAt,
      updatedBy: payload.updatedBy,
    });
    const stateChangedDuringFlight = syncStateMutationVersion !== requestStateMutationVersion;
    const decksChangedDuringFlight = syncDecksMutationVersion !== requestDeckMutationVersion;
    // If the server returned normalized decks (e.g. with resolved ownerUserId), adopt them.
    if (!requestGamesOnly && Array.isArray(payload.decks) && !decksChangedDuringFlight) {
      appState = normalizeAppStateData({ ...appState, decks: payload.decks });
      persistLocalState(appState);
      // Re-render so activeDeckBuilderRecord picks up the server-corrected values.
      refresh();
    }
    setSyncPendingChanges(stateChangedDuringFlight);
    syncHasLoadedCloudState = true;
    syncLocalContentVerified = true;
    shouldQueueFollowUpSync = stateChangedDuringFlight;
    syncRetryCount = 0;
    syncLastErrorMessage = '';
    clearSyncConflict();
    syncLastSuccessAt = new Date().toISOString();
    clearSyncRetryTimer();
    refreshSyncStatus();
  } catch (error) {
    if (error.status === 409) {
      updateSyncMetadata({
        revision: error.body?.conflict?.revision,
        updatedAt: error.body?.conflict?.updatedAt,
        updatedBy: error.body?.conflict?.updatedBy,
      });
      // A 409 means this device's baseline diverged from the cloud. The cloud is
      // authoritative whenever connected, so heal automatically — only a deck
      // draft or live game in progress keeps the manual conflict prompt.
      if (canAutoPullOverLocalChanges()) {
        await autoResolveConflictByPulling(error.body?.conflict || null, 'push-revision-conflict');
        return;
      }
      await resolveSyncConflict(error.body?.conflict || null);
    } else {
      scheduleSyncRetry(error.message);
    }
  } finally {
    syncInFlight = false;
    updateSyncControls();
    refreshSyncStatus();
    if (shouldQueueFollowUpSync) {
      queueCloudSync(0, { gamesOnly: syncQueuedGamesOnly === true });
    } else {
      syncQueuedGamesOnly = null;
    }
  }
}

function queueCloudSync(delay = 500, { gamesOnly = false } = {}) {
  if (!hasSyncCredentials()) {
    return;
  }

  syncStateMutationVersion += 1;
  if (!gamesOnly) {
    syncQueuedGamesOnly = false;
  } else if (syncQueuedGamesOnly !== false) {
    syncQueuedGamesOnly = true;
  }
  setSyncPendingChanges(true);
  syncLastErrorMessage = '';
  refreshSyncStatus();

  if (syncConflictInfo) {
    return;
  }

  // Never auto-push before first successful cloud bootstrap for this session.
  // This prevents local pre-refresh mutations from overwriting cloud state.
  if (!syncHasLoadedCloudState) {
    return;
  }

  if (syncQueueTimer) {
    clearTimeout(syncQueueTimer);
  }

  clearSyncRetryTimer();

  if (Number(delay) <= 0) {
    pushCloudState();
    return;
  }

  syncQueueTimer = setTimeout(() => {
    syncQueueTimer = null;
    pushCloudState();
  }, delay);
}

function generateId() {
  return crypto.randomUUID?.() || `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

function getQueryParam(name) {
  return new URLSearchParams(window.location.search).get(name);
}

function getCurrentHistoryQueryState() {
  return {
    ...(window.CommanderHistory?.getCurrentFilters() || {}),
    sort: historySortKey || 'date',
    order: historySortDescending ? 'desc' : 'asc',
  };
}

function buildHistoryFilterHref(filters = {}, options = {}) {
  const query = new URLSearchParams();
  const includeSortState = options.includeSortState !== false;
  const combinedFilters = includeSortState
    ? { ...getCurrentHistoryQueryState(), ...filters }
    : filters;

  Object.entries(combinedFilters).forEach(([key, value]) => {
    const normalizedValue = String(value || '').trim();
    if (normalizedValue) {
      query.set(key, normalizedValue);
    }
  });

  const queryString = query.toString();
  return `history.html${queryString ? `?${queryString}` : ''}`;
}

function buildHistoryFilterLink(label, filters = {}) {
  const href = buildHistoryFilterHref(filters);
  return `<a class="history-drilldown-link" href="${escapeHtml(href)}">${escapeHtml(label || '—')}</a>`;
}

function getTableSort(tableKey, fallbackColumn, fallbackDescending) {
  const state = tableSortState[tableKey];
  if (!state) {
    return { column: fallbackColumn, descending: fallbackDescending };
  }

  if (!state.column) {
    state.column = fallbackColumn;
  }

  if (typeof state.descending !== 'boolean') {
    state.descending = fallbackDescending;
  }

  return state;
}

function compareTextValues(a, b) {
  return String(a || '').localeCompare(String(b || ''), undefined, { numeric: true, sensitivity: 'base' });
}

function compareNumberValues(a, b) {
  const first = Number.isFinite(a) ? a : Number.NEGATIVE_INFINITY;
  const second = Number.isFinite(b) ? b : Number.NEGATIVE_INFINITY;
  return first - second;
}

function compareDateValues(a, b) {
  return compareTextValues(a || '', b || '');
}

function finalizeSortResult(result, descending) {
  return descending ? -result : result;
}

function getSortableTableHeaders(table) {
  return Array.from(table?.querySelectorAll('.sortable-header') || []);
}

function ensureMobileSortControls(table) {
  const tableKey = table?.dataset.sortTable;
  const wrapper = table?.closest('.player-table-wrapper');
  const headers = getSortableTableHeaders(table);
  if (!tableKey || !wrapper || !headers.length) {
    return null;
  }

  let controls = wrapper.querySelector(`.mobile-sort-controls[data-sort-table="${tableKey}"]`);
  if (!controls) {
    controls = document.createElement('div');
    controls.className = 'mobile-sort-controls';
    controls.dataset.sortTable = tableKey;
    controls.innerHTML = `
      <label class="mobile-sort-field">
        <span class="mobile-sort-label">Sort by</span>
        <select class="mobile-sort-select"></select>
      </label>
      <button type="button" class="secondary-button mobile-sort-direction"></button>
    `;

    const select = controls.querySelector('.mobile-sort-select');
    const directionButton = controls.querySelector('.mobile-sort-direction');

    select?.addEventListener('change', (event) => {
      const column = event.target.value;
      if (!column) {
        return;
      }

      const state = tableSortState[tableKey];
      if (!state) {
        return;
      }

      state.column = column;
      state.descending = getTableSortDefaultDescending(tableKey, column);
      rerenderSortedTable(tableKey);
    });

    directionButton?.addEventListener('click', () => {
      const state = tableSortState[tableKey];
      if (!state) {
        return;
      }

      state.descending = !state.descending;
      rerenderSortedTable(tableKey);
    });

    wrapper.insertBefore(controls, table);
  }

  const select = controls.querySelector('.mobile-sort-select');
  const state = tableSortState[tableKey];
  if (select && state) {
    select.innerHTML = headers
      .map((header) => {
        const column = header.dataset.column || '';
        const label = header.textContent.replace(/▲|▼/g, '').trim();
        const selected = state.column === column ? ' selected' : '';
        return `<option value="${escapeHtml(column)}"${selected}>${escapeHtml(label)}</option>`;
      })
      .join('');
    select.value = state.column;
  }

  return controls;
}

function updateMobileSortControls(tableKey) {
  const table = document.querySelector(`[data-sort-table="${tableKey}"]`);
  const controls = table ? ensureMobileSortControls(table) : null;
  const state = tableSortState[tableKey];
  if (!controls || !state) {
    return;
  }

  const directionButton = controls.querySelector('.mobile-sort-direction');
  const select = controls.querySelector('.mobile-sort-select');
  if (select) {
    select.value = state.column;
  }

  if (directionButton) {
    directionButton.textContent = state.descending ? 'Descending' : 'Ascending';
    directionButton.setAttribute('aria-label', `${state.descending ? 'Descending' : 'Ascending'} sort order`);
  }
}

function initializeMobileSortControls() {
  document.querySelectorAll('[data-sort-table]').forEach((table) => {
    ensureMobileSortControls(table);
    updateMobileSortControls(table.dataset.sortTable || '');
  });
}

function updateSortableTableIndicators(tableKey) {
  const state = tableSortState[tableKey];
  const headers = document.querySelectorAll(`[data-sort-table="${tableKey}"] .sortable-header`);
  headers.forEach((header) => {
    const column = header.dataset.column;
    const indicator = header.querySelector('.sort-indicator');
    const isActive = Boolean(state) && column === state.column;

    header.setAttribute('aria-sort', isActive ? (state.descending ? 'descending' : 'ascending') : 'none');

    if (!indicator) {
      return;
    }

    if (isActive) {
      indicator.textContent = state.descending ? ' ▼' : ' ▲';
      indicator.style.opacity = '1';
    } else {
      indicator.textContent = '';
      indicator.style.opacity = '0.3';
    }
  });

  updateMobileSortControls(tableKey);
}

function toggleTableSort(tableKey, column, fallbackDescending = true) {
  const state = getTableSort(tableKey, column, fallbackDescending);
  if (state.column === column) {
    state.descending = !state.descending;
  } else {
    state.column = column;
    state.descending = fallbackDescending;
  }
}

function getTableSortDefaultDescending(tableKey, column) {
  const ascendingColumns = new Set([
    'rank',
    'player',
    'commander',
    'favoriteCommander',
    'bestDeck',
    'nemesis',
    'victim',
    'owner',
    'url',
    'avgPlace',
  ]);

  if (column === 'lastWin') {
    return true;
  }

  return !ascendingColumns.has(column);
}

function rerenderSortedTable(tableKey) {
  const games = loadGames();
  switch (tableKey) {
    case 'rankingsMain':
      window.CommanderRankings?.renderMain(games);
      break;
    case 'playerStats':
      window.CommanderPlayerStats?.render(games);
      break;
    case 'commanderStats':
      window.CommanderStats?.render(games);
      break;
    case 'recentPlayerTrends':
    case 'recentCommanderTrends':
    case 'playerStreaks':
    case 'commanderStreaks':
      renderRankingsPage(games);
      break;
    default:
      break;
  }

  applyResponsiveTableLabels();
}

function loadGames() {
  const games = Array.isArray(appState.games) ? appState.games : [];
  let updated = false;

  games.forEach((game) => {
    if (!game.id) {
      game.id = generateId();
      updated = true;
    }
  });

  if (updated) {
    saveGames(games);
  }

  return games;
}

function saveGames(games, { queueSync = true } = {}) {
  appState = normalizeAppStateData({
    ...appState,
    games: Array.isArray(games) ? games : [],
  });
  persistLocalState(appState);

  // Storage events do not fire in the same tab, so update the deck-builder
  // performance panel immediately after game saves.
  if (window.CommanderDeckBuilder?.hasView?.()) {
    const activeDeck = applyDeckBuilderDraftMeta(ensureActiveDeckBuilderRecord() || null);
    if (activeDeck) {
      renderDeckBuilderBreakdown(activeDeck);
    }
  }

  if (queueSync) {
    queueCloudSync(0, { gamesOnly: true });
  }
}

function loadDeckLists() {
  return Array.isArray(appState.deckLists) ? appState.deckLists : [];
}

function loadDecks() {
  return Array.isArray(appState.decks) ? appState.decks : [];
}

function saveDeckLists(deckLists) {
  appState = normalizeAppStateData({
    ...appState,
    deckLists: Array.isArray(deckLists) ? deckLists : [],
  });
  persistLocalState(appState);
  queueCloudSync(0);
}

function saveDecks(decks, options = {}) {
  const { deferPersist = false, delay = DECKS_RAPID_ACTION_PERSIST_DEBOUNCE_MS } = options;
  syncDecksMutationVersion += 1;
  appState = normalizeAppStateData({
    ...appState,
    decks: Array.isArray(decks) ? decks : [],
  });

  if (!deferPersist) {
    if (decksPersistTimer) {
      clearTimeout(decksPersistTimer);
      decksPersistTimer = null;
    }

    persistLocalState(appState);
    queueCloudSync(0);
    return;
  }

  if (Number(delay) <= 0) {
    if (decksPersistTimer) {
      clearTimeout(decksPersistTimer);
      decksPersistTimer = null;
    }

    persistLocalState(appState);
    queueCloudSync(0);
    return;
  }

  if (decksPersistTimer) {
    clearTimeout(decksPersistTimer);
  }

  decksPersistTimer = setTimeout(() => {
    decksPersistTimer = null;
    persistLocalState(appState);
    queueCloudSync(0);
  }, delay);
}

function flushQueuedDeckPersist({ force = false, queueSync = true } = {}) {
  if (!decksPersistTimer && !force) {
    return;
  }

  if (decksPersistTimer) {
    clearTimeout(decksPersistTimer);
  }
  decksPersistTimer = null;
  persistLocalState(appState);

  if (queueSync) {
    queueCloudSync(0);
    return;
  }

  // During page transitions, avoid a blocking full sync call.
  // Keepalive upload handles best-effort cloud flush.
  if (hasSyncCredentials()) {
    setSyncPendingChanges(true);
    syncLastErrorMessage = '';
    refreshSyncStatus();
  }
}

const STALE_EMPTY_DECK_PRUNE_AGE_MS = 15 * 60 * 1000;

function isEmptyUntitledDeckRecord(deck) {
  const name = String(deck?.name || '').trim();
  if (name && name !== 'Untitled Deck') {
    return false;
  }

  return !deck?.commander
    && !deck?.secondCommander
    && !(Array.isArray(deck?.cards) && deck.cards.length)
    && !(Array.isArray(deck?.maybeboard) && deck.maybeboard.length)
    && !(Array.isArray(deck?.tokens) && deck.tokens.length);
}

// Removes empty "Untitled Deck" records the user never chose to save. Older builds
// leaked these into local storage when a page exited with unsaved deck-builder
// staging; they then rendered on the Decks page indefinitely even though they were
// never synced to the cloud KV. Only records old enough to be abandoned are touched
// and the deck currently open in the builder is never pruned. Returns true when any
// records were removed.
function pruneStaleEmptyDeckRecords({ deferCloudSync = false } = {}) {
  const decks = loadDecks();
  if (!decks.length) {
    return false;
  }

  const currentUserId = getCurrentSyncUserId();
  const currentOwnerKey = getIdentityKey(getCurrentSyncDisplayName());
  const hasSyncIdentity = Boolean(currentUserId || currentOwnerKey);
  const now = Date.now();
  const activeDeckId = getDeckBuilderSyncState().activeDeckId;

  const keptDecks = decks.filter((deck) => {
    if (!deck || deck.id === activeDeckId || !isEmptyUntitledDeckRecord(deck)) {
      return true;
    }

    if (hasSyncIdentity) {
      const ownerUserId = String(deck?.ownerUserId || '').trim().toLowerCase();
      const ownedByCurrentUser = ownerUserId
        ? ownerUserId === currentUserId
        : Boolean(currentOwnerKey) && getIdentityKey(deck?.owner || '') === currentOwnerKey;
      if (!ownedByCurrentUser) {
        return true;
      }
    }

    const createdAtMs = Date.parse(deck?.createdAt || '');
    const ageMs = Number.isFinite(createdAtMs) ? now - createdAtMs : Number.POSITIVE_INFINITY;
    return ageMs < STALE_EMPTY_DECK_PRUNE_AGE_MS;
  });

  if (keptDecks.length === decks.length) {
    return false;
  }

  saveDecks(keptDecks, deferCloudSync
    ? { deferPersist: true, delay: DECKS_PERSIST_DEBOUNCE_MS }
    : { deferPersist: false, delay: 0 });
  return true;
}

function normalizeList(value) {
  return value
    .split(',')
    .map((item) => normalizeIdentityLabel(item))
    .filter(Boolean);
}

function formatPlayerCommanders(list) {
  return list.map(({ player, commander }) => `${player}: ${commander || '—'}`).join(', ');
}

function getActiveLiveModalControl() {
  if (!liveModalField || liveModalField.hidden) {
    return null;
  }
  return liveModalConfig?.multiline ? liveModalTextarea : liveModalInput;
}

function resetLiveModalState() {
  if (liveModalField) {
    liveModalField.hidden = true;
  }
  if (liveModalInput) {
    liveModalInput.hidden = true;
    liveModalInput.value = '';
    liveModalInput.placeholder = '';
    liveModalInput.type = 'text';
    liveModalInput.inputMode = '';
  }
  if (liveModalTextarea) {
    liveModalTextarea.hidden = true;
    liveModalTextarea.value = '';
    liveModalTextarea.placeholder = '';
  }
  if (liveModalError) {
    liveModalError.hidden = true;
    liveModalError.textContent = '';
  }
  if (liveModalCancelButton) {
    liveModalCancelButton.hidden = false;
    liveModalCancelButton.textContent = 'Cancel';
  }
  if (liveModalConfirmButton) {
    liveModalConfirmButton.textContent = 'Confirm';
  }
  liveModalConfig = null;
}

function hideLiveModal(result = null) {
  if (liveModalPrompt) {
    liveModalPrompt.hidden = true;
  }

  const resolver = liveModalResolver;
  liveModalResolver = null;
  resetLiveModalState();

  if (resolver) {
    resolver(result);
  }
}

function openCardImageModal(src, alt = 'Card image') {
  const imageSrc = String(src || '').trim();
  if (!cardImageModal || !cardImageModalImg || !imageSrc) {
    return;
  }

  cardImageModalImg.src = imageSrc;
  cardImageModalImg.alt = String(alt || 'Card image').trim() || 'Card image';
  cardImageModal.hidden = false;
}

function closeCardImageModal() {
  if (!cardImageModal || !cardImageModalImg) {
    return;
  }

  cardImageModal.hidden = true;
  cardImageModalImg.src = '';
  cardImageModalImg.alt = 'Card image';
}

function setLiveModalError(message = '') {
  if (!liveModalError) {
    return;
  }
  liveModalError.textContent = message;
  liveModalError.hidden = !message;
}

function showLiveModal({
  title,
  description = '',
  confirmLabel = 'Confirm',
  cancelLabel = 'Cancel',
  showCancel = true,
  showInput = false,
  defaultValue = '',
  placeholder = '',
  inputType = 'text',
  inputMode = '',
  multiline = false,
  validate = null,
}) {
  if (!liveModalPrompt || !liveModalTitle || !liveModalCopy || !liveModalConfirmButton || !liveModalCancelButton) {
    return Promise.resolve(null);
  }

  resetLiveModalState();
  liveModalConfig = { validate, multiline };
  liveModalTitle.textContent = title;
  liveModalCopy.textContent = description;
  liveModalConfirmButton.textContent = confirmLabel;
  liveModalCancelButton.textContent = cancelLabel;
  liveModalCancelButton.hidden = !showCancel;

  const activeControl = showInput ? (multiline ? liveModalTextarea : liveModalInput) : null;
  if (activeControl) {
    if (liveModalField) {
      liveModalField.hidden = false;
    }
    activeControl.hidden = false;
    activeControl.value = defaultValue;
    activeControl.placeholder = placeholder;
    if (!multiline) {
      activeControl.type = inputType;
      activeControl.inputMode = inputMode;
    }
  }

  liveModalPrompt.hidden = false;
  setLiveModalError('');

  requestAnimationFrame(() => {
    const control = getActiveLiveModalControl();
    if (control) {
      control.focus();
      if (typeof control.select === 'function') {
        control.select();
      }
      return;
    }

    liveModalConfirmButton.focus();
  });

  return new Promise((resolve) => {
    liveModalResolver = resolve;
  });
}

async function promptLiveAlert(message, title = 'Notice') {
  await showLiveModal({
    title,
    description: message,
    confirmLabel: 'OK',
    showInput: false,
    showCancel: false,
  });
}

async function promptLiveConfirm(message, {
  title = 'Confirm action',
  confirmLabel = 'Confirm',
  cancelLabel = 'Cancel',
} = {}) {
  const result = await showLiveModal({
    title,
    description: message,
    confirmLabel,
    cancelLabel,
    showInput: false,
    showCancel: true,
  });
  return result === true;
}

function parseLiveModalInteger(rawValue, validatorMessage) {
  const parsedValue = parseInt(String(rawValue || '').trim(), 10);
  if (!Number.isFinite(parsedValue)) {
    return { error: validatorMessage };
  }
  return { value: parsedValue };
}

async function promptLiveText(message, {
  title = 'Enter value',
  confirmLabel = 'Save',
  cancelLabel = 'Cancel',
  defaultValue = '',
  placeholder = '',
  multiline = false,
  inputType = 'text',
  inputMode = '',
  validate = null,
} = {}) {
  return showLiveModal({
    title,
    description: message,
    confirmLabel,
    cancelLabel,
    showCancel: true,
    showInput: true,
    defaultValue,
    placeholder,
    multiline,
    inputType,
    inputMode,
    validate,
  });
}

function getLiveCommanderOptionValues() {
  return getUniqueValuesBySimilarity(
    loadDecks()
      .map(normalizeDeckRecord)
      .filter((deck) => Boolean(deck?.commander?.name))
      .filter((deck) => deck.inRotation !== false)
      .map((deck) => deck.commander.name),
  );
}

async function acquireWakeLock() {
  if (!('wakeLock' in navigator)) {
    return;
  }
  try {
    wakeLockSentinel = await navigator.wakeLock.request('screen');
    wakeLockSentinel.addEventListener('release', () => {
      wakeLockSentinel = null;
    });
  } catch {
    wakeLockSentinel = null;
  }
}

function releaseWakeLock() {
  if (wakeLockSentinel) {
    wakeLockSentinel.release();
    wakeLockSentinel = null;
  }
}

function getCurrentPageName() {
  const pathName = String(window.location.pathname || '').trim();
  const segments = pathName.split('/').filter(Boolean);
  return (segments[segments.length - 1] || 'index.html').toLowerCase();
}

function closePrimaryMenu() {
  if (!pageSwitch) {
    return;
  }

  pageSwitch.classList.remove('is-open');
  if (pageSwitchToggleButton) {
    pageSwitchToggleButton.setAttribute('aria-expanded', 'false');
  }
}

function togglePrimaryMenu(forceOpen) {
  if (!pageSwitch) {
    return;
  }

  const nextOpen = typeof forceOpen === 'boolean'
    ? forceOpen
    : !pageSwitch.classList.contains('is-open');
  pageSwitch.classList.toggle('is-open', nextOpen);
  if (pageSwitchToggleButton) {
    pageSwitchToggleButton.setAttribute('aria-expanded', String(nextOpen));
  }
}

function initializePrimaryMenu() {
  if (!pageSwitch || !pageSwitchToggleButton || !pageSwitchPanel) {
    return;
  }

  const currentPageName = getCurrentPageName();
  Array.from(pageSwitchPanel.querySelectorAll('.page-link')).forEach((link) => {
    const href = String(link.getAttribute('href') || '').trim().toLowerCase();
    const isCurrent = href === currentPageName;
    link.classList.toggle('is-current', isCurrent);
    if (isCurrent) {
      link.setAttribute('aria-current', 'page');
    } else {
      link.removeAttribute('aria-current');
    }
  });

  pageSwitchToggleButton.addEventListener('click', () => {
    togglePrimaryMenu();
  });

  pageSwitchPanel.addEventListener('click', (event) => {
    if (event.target.closest('.page-link')) {
      closePrimaryMenu();
    }
  });

  updateAdminLogsShortcutVisibility();
}

function updateAdminLogsShortcutVisibility() {
  if (!pageSwitchPanel) {
    return;
  }

  let adminLinks = Array.from(pageSwitchPanel.querySelectorAll('.page-link')).filter((link) => {
    const rawHref = String(link.getAttribute('href') || '').trim().toLowerCase();
    return rawHref === 'admin-logs.html' || rawHref.endsWith('/admin-logs.html');
  });
  if (!adminLinks.length) {
    const pageSwitchGrid = pageSwitchPanel.querySelector('.page-switch-grid');
    if (!pageSwitchGrid) {
      return;
    }
    const adminLink = document.createElement('a');
    adminLink.href = 'admin-logs.html';
    adminLink.className = 'page-link';
    adminLink.textContent = 'Admin Logs';
    pageSwitchGrid.appendChild(adminLink);
    adminLinks = [adminLink];
  }

  const canViewShortcut = hasSyncCredentials() && isCurrentSyncUserAdmin();
  adminLinks.forEach((link) => {
    link.classList.toggle('admin-link-visible', canViewShortcut);
    link.hidden = !canViewShortcut;
    link.setAttribute('aria-hidden', canViewShortcut ? 'false' : 'true');
    if (canViewShortcut) {
      link.removeAttribute('tabindex');
    } else {
      link.setAttribute('tabindex', '-1');
    }
  });
}

function createPlayerRow(data = {}) {
  const row = document.createElement('tr');
  const killedValue = Array.isArray(data.killed) ? data.killed.join(', ') : data.killed || '';

  row.innerHTML = `
    <td class="lookup-field-cell">
      <div class="combined-input-wrapper">
        <input class="lookup-input" type="text" name="player" list="player-list" placeholder="Player" required value="${escapeHtml(data.player || '')}" autocomplete="off" autocapitalize="none" autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" />
        <button type="button" class="dropdown-button player-dropdown-button" title="Show options">▼</button>
        <div class="dropdown-menu player-dropdown-menu"></div>
      </div>
    </td>
    <td class="lookup-field-cell">
      <div class="combined-input-wrapper">
        <input class="lookup-input" type="text" name="commander" list="commander-list" placeholder="Commander" value="${escapeHtml(data.commander || '')}" autocomplete="off" autocapitalize="none" autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" />
        <button type="button" class="dropdown-button commander-dropdown-button" title="Show options">▼</button>
        <div class="dropdown-menu commander-dropdown-menu"></div>
      </div>
    </td>
    <td><input type="number" name="place" min="1" step="1" value="${escapeHtml(data.place || '')}" placeholder="Place" /></td>
    <td><input type="number" name="kills" min="0" step="1" value="${escapeHtml(data.kills || 0)}" placeholder="Kills" /></td>
    <td><input type="number" name="turnKilled" min="1" step="1" value="${escapeHtml(data.turnKilled || '')}" placeholder="Turn" /></td>
    <td><textarea name="killed" placeholder="Killed">${escapeHtml(killedValue)}</textarea></td>
  `;

  attachLookupWrapperHandlers(row);
  return row;
}

function addPlayerRow(data = {}) {
  const row = createPlayerRow(data);
  playerTableBody.appendChild(row);
  populateRowSelectors(row);
}

function removePlayerRow() {
  const rows = playerTableBody.querySelectorAll('tr');
  if (rows.length > 2) {
    rows[rows.length - 1].remove();
    playerTableBody.querySelectorAll('tr').forEach((r) => populateRowSelectors(r));
  }
}

function resetPlayerTable() {
  playerTableBody.innerHTML = '';
  for (let i = 0; i < 4; i += 1) {
    addPlayerRow();
  }
}

function getPlayerRows() {
  const playerMap = buildCanonicalIdentityMapFromValues(knownPlayers);
  const commanderMap = buildCanonicalIdentityMapFromValues(knownCommanders);

  return Array.from(playerTableBody.querySelectorAll('tr')).map((row) => {
    const player = canonicalizeIdentityValue(row.querySelector('[name="player"]')?.value || '', playerMap);
    const commanderRaw = row.querySelector('[name="commander"]')?.value || '';
    const commander = canonicalizeIdentityValue(commanderRaw, commanderMap);
    const placeRaw = row.querySelector('input[name="place"]').value.trim();
    const killsRaw = row.querySelector('input[name="kills"]').value.trim();
    const turnKilledRaw = row.querySelector('input[name="turnKilled"]').value.trim();
    const killed = normalizeList(row.querySelector('[name="killed"]')?.value || '');

    return {
      player,
      commanderRaw,
      commander,
      placeRaw,
      killsRaw,
      turnKilledRaw,
      place: parseOptionalPositiveInteger(placeRaw),
      kills: killsRaw ? parseOptionalNonNegativeInteger(killsRaw) : 0,
      turnKilled: parseOptionalPositiveInteger(turnKilledRaw),
      killed,
    };
  }).filter((row) => row.player);
}

function parseOptionalPositiveInteger(value) {
  const normalizedValue = String(value || '').trim();
  if (!normalizedValue) {
    return null;
  }

  const parsedValue = Number(normalizedValue);
  return Number.isInteger(parsedValue) && parsedValue > 0 ? parsedValue : null;
}

function parseOptionalNonNegativeInteger(value) {
  const normalizedValue = String(value || '').trim();
  if (!normalizedValue) {
    return null;
  }

  const parsedValue = Number(normalizedValue);
  return Number.isInteger(parsedValue) && parsedValue >= 0 ? parsedValue : null;
}

function sanitizeManualGameRows(rows) {
  const playerBuckets = new Map();
  const commanderBuckets = new Map();

  rows.forEach((row) => {
    recordIdentityVariant(playerBuckets, row.player);
    recordIdentityVariant(commanderBuckets, row.commander);
    row.killed.forEach((player) => {
      recordIdentityVariant(playerBuckets, player);
    });
  });

  const playerMap = buildCanonicalIdentityMap(playerBuckets);
  const commanderMap = buildCanonicalIdentityMap(commanderBuckets);

  return rows.map(({ placeRaw, killsRaw, turnKilledRaw, ...row }) => ({
    ...row,
    player: canonicalizeIdentityValue(row.player, playerMap),
    commander: canonicalizeIdentityValue(row.commander, commanderMap),
    killed: normalizeIdentityList(row.killed, playerMap),
  })).map((row) => enrichSavedGamePlayerRow(row));
}

function enrichSavedGamePlayerRow(row = {}, decks = loadDecks()) {
  const existingDeckId = String(row?.deckId || '').trim();
  const existingDeckName = String(row?.deckName || '').trim();
  const existingDeckCommander = String(row?.deckCommander || '').trim();
  const player = String(row?.player || '').trim();
  const commander = String(row?.commander || '').trim();

  if (!commander && !player) {
    return {
      ...row,
      deckId: existingDeckId,
      deckName: existingDeckName,
      deckCommander: existingDeckCommander,
    };
  }

  if (existingDeckId && existingDeckName) {
    return {
      ...row,
      deckId: existingDeckId,
      deckName: existingDeckName,
      deckCommander: existingDeckCommander,
    };
  }

  const ownerKey = getIdentityKey(player);
  const commanderKey = getIdentityKey(commander);
  if (!commanderKey) {
    return {
      ...row,
      deckId: existingDeckId,
      deckName: existingDeckName,
      deckCommander: existingDeckCommander,
    };
  }

  const exactDeckMatch = decks.find((deck) => {
    const deckCommanderKey = getIdentityKey(deck?.commander?.name || '');
    if (!deckCommanderKey || deckCommanderKey !== commanderKey) {
      return false;
    }
    if (!ownerKey) {
      return true;
    }
    return getIdentityKey(deck?.owner || '') === ownerKey;
  });

  if (exactDeckMatch) {
    return {
      ...row,
      deckId: String(exactDeckMatch.id || '').trim(),
      deckName: String(exactDeckMatch.name || exactDeckMatch?.commander?.name || '').trim(),
      deckCommander: String(exactDeckMatch?.commander?.name || '').trim(),
    };
  }

  const fallbackDeckMatch = decks.find((deck) => {
    const deckCommanderKey = getIdentityKey(deck?.commander?.name || '');
    if (!deckCommanderKey) {
      return false;
    }
    if (!ownerKey) {
      return getStringSimilarity(commanderKey, deckCommanderKey) >= 0.9;
    }
    return getStringSimilarity(commanderKey, deckCommanderKey) >= 0.9
      && getIdentityKey(deck?.owner || '') === ownerKey;
  });

  if (fallbackDeckMatch) {
    return {
      ...row,
      deckId: String(fallbackDeckMatch.id || '').trim(),
      deckName: String(fallbackDeckMatch.name || fallbackDeckMatch?.commander?.name || '').trim(),
      deckCommander: String(fallbackDeckMatch?.commander?.name || '').trim(),
    };
  }

  return {
    ...row,
    deckId: existingDeckId,
    deckName: existingDeckName,
    deckCommander: existingDeckCommander,
  };
}

function normalizeSavedGameEntry(game = {}, decks = []) {
  const baseGame = game && typeof game === 'object' ? game : {};
  const finishOrder = Array.isArray(baseGame.finishOrder) && baseGame.finishOrder.length
    ? baseGame.finishOrder.filter(Boolean)
    : [];
  const rawRows = Array.isArray(baseGame.playerRows) && baseGame.playerRows.length
    ? baseGame.playerRows
    : (Array.isArray(baseGame.playerCommanders) && baseGame.playerCommanders.length
      ? baseGame.playerCommanders
      : (Array.isArray(baseGame.players) && baseGame.players.length
        ? baseGame.players.map((player) => ({ player }))
        : []));

  const rows = rawRows
    .map((row) => {
      const player = String(row?.player || '').trim();
      if (!player) {
        return null;
      }

      const explicitPlace = typeof row?.place === 'number'
        ? row.place
        : parseOptionalPositiveInteger(row?.place || '');
      const place = explicitPlace ?? (finishOrder.length ? finishOrder.indexOf(player) + 1 : null);
      const kills = typeof row?.kills === 'number'
        ? row.kills
        : parseOptionalNonNegativeInteger(row?.kills || '');
      const killed = Array.isArray(row?.killed)
        ? row.killed
        : normalizeList(row?.killed || '');
      const turnKilled = typeof row?.turnKilled === 'number'
        ? row.turnKilled
        : parseOptionalPositiveInteger(row?.turnKilled || '');

      return enrichSavedGamePlayerRow({
        ...row,
        player,
        commander: String(row?.commander || '').trim(),
        place: place ?? null,
        kills: typeof kills === 'number' ? kills : 0,
        killed,
        turnKilled: typeof turnKilled === 'number' ? turnKilled : null,
      }, decks);
    })
    .filter(Boolean);

  const normalizedFinishOrder = finishOrder.length
    ? finishOrder
    : rows
      .slice()
      .sort((a, b) => ((a.place || 999) - (b.place || 999)))
      .map((row) => row.player);

  const liveSummary = baseGame.liveSummary && typeof baseGame.liveSummary === 'object' ? baseGame.liveSummary : {};
  const normalizedLiveSummary = {
    startingPlayer: String(liveSummary.startingPlayer || '').trim(),
    alternateWinCondition: String(liveSummary.alternateWinCondition || '').trim(),
    alternateLoseConditions: Array.isArray(liveSummary.alternateLoseConditions)
      ? liveSummary.alternateLoseConditions
        .map((entry) => ({
          player: String(entry?.player || '').trim(),
          details: String(entry?.details || '').trim(),
        }))
        .filter((entry) => entry.player || entry.details)
      : [],
    durationMinutes: typeof liveSummary.durationMinutes === 'number' ? liveSummary.durationMinutes : null,
    firstBlood: liveSummary.firstBlood && typeof liveSummary.firstBlood === 'object'
      ? {
        actorPlayer: String(liveSummary.firstBlood.actorPlayer || '').trim(),
        actorCommander: String(liveSummary.firstBlood.actorCommander || '').trim(),
        targetPlayer: String(liveSummary.firstBlood.targetPlayer || '').trim(),
        turnNumber: typeof liveSummary.firstBlood.turnNumber === 'number' ? liveSummary.firstBlood.turnNumber : null,
      }
      : null,
    recordStats: liveSummary.recordStats && typeof liveSummary.recordStats === 'object' ? liveSummary.recordStats : null,
    turnNumber: typeof liveSummary.turnNumber === 'number' ? liveSummary.turnNumber : null,
    eventCount: typeof liveSummary.eventCount === 'number' ? liveSummary.eventCount : null,
  };

  return {
    ...baseGame,
    id: String(baseGame.id || generateId()).trim(),
    date: String(baseGame.date || new Date().toISOString().slice(0, 10)).trim(),
    playerRows: rows,
    players: rows.map((row) => row.player),
    playerCommanders: rows.map((row) => ({ player: row.player, commander: row.commander })),
    finishOrder: normalizedFinishOrder,
    notes: String(baseGame.notes || '').trim(),
    liveSummary: normalizedLiveSummary,
  };
}

function buildCompactSavedGameEntry(game = {}) {
  return normalizeSavedGameEntry(game, loadDecks());
}

function validateManualGameEntry(rows, { firstBloodPlayer, firstBloodTurn, winningTurn }) {
  if (rows.length < 2) {
    return 'Please add at least two players before saving a game.';
  }

  const playerNameLookup = new Map();

  for (const row of rows) {
    const normalizedPlayer = row.player.toLocaleLowerCase();
    if (playerNameLookup.has(normalizedPlayer)) {
      return `Player names must be unique within a game. Duplicate entry: ${row.player}.`;
    }
    playerNameLookup.set(normalizedPlayer, row.player);

    if (row.placeRaw && row.place === null) {
      return `Place for ${row.player} must be a whole number greater than 0.`;
    }

    if (row.killsRaw && row.kills === null) {
      return `Kills for ${row.player} must be a whole number 0 or greater.`;
    }

    if (row.turnKilledRaw && row.turnKilled === null) {
      return `Turn killed for ${row.player} must be a whole number greater than 0.`;
    }
  }

  if (!rows.every((row) => row.place !== null)) {
    return 'Enter a finish place for every player before saving the game.';
  }

  const sortedPlaces = rows.map((row) => row.place).slice().sort((a, b) => a - b);
  for (let index = 0; index < sortedPlaces.length; index += 1) {
    if (sortedPlaces[index] !== index + 1) {
      return `Finish places must be unique and run from 1 through ${rows.length}.`;
    }
  }

  if (firstBloodPlayer && !playerNameLookup.has(firstBloodPlayer.toLocaleLowerCase())) {
    return 'First blood must match one of the players in this game.';
  }

  if (firstBloodTurn && !firstBloodPlayer) {
    return 'Enter the player who got first blood before adding a first blood turn.';
  }

  if (firstBloodTurn && winningTurn && firstBloodTurn > winningTurn) {
    return 'First blood turn cannot be later than the winning turn.';
  }

  const rowsByPlayerName = new Map(
    rows.map((row) => [row.player.toLocaleLowerCase(), row]),
  );
  const winnerRow = rows.find((row) => row.place === 1) || null;
  const eliminatedRows = rows
    .filter((row) => row.place !== 1)
    .slice()
    .sort((a, b) => b.place - a.place);

  for (const row of eliminatedRows) {
    if (row.turnKilled === null) {
      return `${row.player} finished in place ${row.place} and must have an elimination turn.`;
    }
  }

  for (let index = 1; index < eliminatedRows.length; index += 1) {
    const earlierElimination = eliminatedRows[index - 1];
    const laterElimination = eliminatedRows[index];

    if (laterElimination.turnKilled < earlierElimination.turnKilled) {
      return `${laterElimination.player} finished ahead of ${earlierElimination.player} but is marked eliminated earlier.`;
    }
  }

  if (winningTurn && eliminatedRows.some((row) => row.turnKilled !== null && row.turnKilled > winningTurn)) {
    return 'Winning turn cannot be earlier than any recorded elimination turn.';
  }

  if (firstBloodPlayer && firstBloodTurn) {
    const firstBloodRow = rowsByPlayerName.get(firstBloodPlayer.toLocaleLowerCase()) || null;
    if (firstBloodRow?.turnKilled !== null && firstBloodTurn > firstBloodRow.turnKilled) {
      return `${firstBloodPlayer} cannot be credited with first blood on turn ${firstBloodTurn} after dying on turn ${firstBloodRow.turnKilled}.`;
    }
  }

  const victimToKiller = new Map();
  const maximumCreditedKills = winnerRow ? rows.length - 1 : Math.max(0, rows.length - 1);
  let totalCreditedKills = 0;

  for (const row of rows) {
    const kills = typeof row.kills === 'number' ? row.kills : 0;
    const killedPlayers = getCleanKilledList(row.killed);
    const normalizedKilled = new Set();
    const maxPossibleKills = rows.length - row.place;

    if (kills > rows.length - 1) {
      return `${row.player} cannot record more kills than there were opponents in the game.`;
    }

    if (kills > maxPossibleKills) {
      return `${row.player} cannot record ${kills} kills from place ${row.place}.`;
    }

    if (row.place === rows.length && kills > 0) {
      return `${row.player} finished last and must have 0 kills.`;
    }

    if (row.place === 1 && row.turnKilled !== null) {
      return `${row.player} is marked as the winner and cannot also have a turn killed value.`;
    }

    for (const killedPlayer of killedPlayers) {
      const normalizedKilledPlayer = killedPlayer.toLocaleLowerCase();

      if (normalizedKilledPlayer === row.player.toLocaleLowerCase()) {
        return `${row.player} cannot list themselves in the killed field.`;
      }

      if (!playerNameLookup.has(normalizedKilledPlayer)) {
        return `${row.player} lists ${killedPlayer} as killed, but that player is not in this game.`;
      }

      const killedRow = rowsByPlayerName.get(normalizedKilledPlayer);
      if (!killedRow) {
        return `${row.player} lists ${killedPlayer} as killed, but that player is not in this game.`;
      }

      if (victimToKiller.has(normalizedKilledPlayer)) {
        return `${killedPlayer} is already credited as a kill for ${victimToKiller.get(normalizedKilledPlayer)}.`;
      }

      if (killedRow.place <= row.place) {
        return `${row.player} cannot be credited with killing ${killedPlayer} because ${killedPlayer} finished ahead of them.`;
      }

      if (row.turnKilled !== null && killedRow.turnKilled !== null && killedRow.turnKilled > row.turnKilled) {
        return `${row.player} cannot be credited with killing ${killedPlayer} on turn ${killedRow.turnKilled} after dying on turn ${row.turnKilled}.`;
      }

      if (normalizedKilled.has(normalizedKilledPlayer)) {
        return `${row.player} lists ${killedPlayer} more than once in the killed field.`;
      }

      normalizedKilled.add(normalizedKilledPlayer);
      victimToKiller.set(normalizedKilledPlayer, row.player);
    }

    if (killedPlayers.length > kills) {
      return `${row.player} lists more killed players than their kill total.`;
    }

    totalCreditedKills += kills;
  }

  if (winnerRow && winnerRow.turnKilled !== null) {
    return `${winnerRow.player} is marked as the winner and cannot also have an elimination turn.`;
  }

  if (totalCreditedKills > maximumCreditedKills) {
    return `This game has ${totalCreditedKills} credited kills, but only ${maximumCreditedKills} eliminations were possible.`;
  }

  return '';
}

function getManualGameAdvisories(rows, { firstBloodPlayer, firstBloodTurn, winningTurn }) {
  const advisories = [];
  const rowsWithMissingCommander = rows.filter((row) => !normalizeIdentityLabel(row.commander));
  const totalCreditedKills = rows.reduce((sum, row) => sum + (typeof row.kills === 'number' ? row.kills : 0), 0);
  const winnerRow = rows.find((row) => row.place === 1) || null;

  if (rowsWithMissingCommander.length) {
    advisories.push(`Missing commander for ${rowsWithMissingCommander.map((row) => row.player).join(', ')}.`);
  }

  if (winnerRow && !winnerRow.commander) {
    advisories.push(`Winner ${winnerRow.player} has no commander listed.`);
  }

  if (!firstBloodPlayer && !firstBloodTurn) {
    advisories.push('No first blood was entered for this game.');
  }

  if (!winningTurn) {
    advisories.push('No winning turn was entered for this game.');
  }

  if (rows.length >= 3 && totalCreditedKills === 0) {
    advisories.push('No kill credit was recorded for any elimination.');
  }

  return advisories;
}

function buildManualGameRecordStats(rows, gameDate) {
  const eliminatedRows = rows.filter((row) => typeof row.turnKilled === 'number' && row.turnKilled > 0);
  if (!eliminatedRows.length) {
    return null;
  }

  const fastestEliminationRow = eliminatedRows.slice().sort((a, b) => a.turnKilled - b.turnKilled || a.player.localeCompare(b.player))[0];
  return {
    fastestElimination: createRecordCandidate({
      value: fastestEliminationRow.turnKilled,
      holder: '',
      commander: '',
      date: gameDate,
      notes: `${fastestEliminationRow.player} was eliminated on turn ${fastestEliminationRow.turnKilled}.`,
    }),
  };
}

function buildManualGameLiveSummary(rows, gameDate, firstBloodPlayer, firstBloodTurn, winningTurn) {
  const normalizedFirstBloodPlayer = String(firstBloodPlayer || '').trim();
  const normalizedFirstBloodTurn = parseOptionalPositiveInteger(firstBloodTurn);
  const normalizedWinningTurn = parseOptionalPositiveInteger(winningTurn);
  const firstBloodCommander = normalizedFirstBloodPlayer
    ? rows.find((row) => row.player === normalizedFirstBloodPlayer)?.commander || ''
    : '';
  const recordStats = buildManualGameRecordStats(rows, gameDate);

  if (!normalizedFirstBloodPlayer && !normalizedWinningTurn && !recordStats) {
    return null;
  }

  return {
    firstBlood: normalizedFirstBloodPlayer
      ? {
        actorPlayer: normalizedFirstBloodPlayer,
        actorCommander: firstBloodCommander,
        targetPlayer: '',
        turnNumber: normalizedFirstBloodTurn,
      }
      : null,
    recordStats,
    turnNumber: normalizedWinningTurn,
  };
}

function mergeEditedGameLiveSummary(existingLiveSummary, manualLiveSummary) {
  if (!manualLiveSummary) {
    return null;
  }

  const existingFirstBlood = existingLiveSummary?.firstBlood;
  const manualFirstBlood = manualLiveSummary.firstBlood;

  return {
    ...manualLiveSummary,
    firstBlood: manualFirstBlood
      ? {
        ...existingFirstBlood,
        ...manualFirstBlood,
        targetPlayer: manualFirstBlood.targetPlayer || existingFirstBlood?.targetPlayer || '',
      }
      : null,
    recordStats: manualLiveSummary.recordStats || existingLiveSummary?.recordStats || null,
  };
}

function createStatCard(title, body) {
  const card = document.createElement('div');
  card.className = 'stats-card';
  card.innerHTML = `<h3>${title}</h3><p>${body}</p>`;
  return card;
}

function getCleanKilledList(killed) {
  if (Array.isArray(killed)) {
    return killed.map((name) => (name || '').trim()).filter(Boolean);
  }
  return normalizeList(String(killed || ''));
}

function ensurePlayerStats(stats, player) {
  if (!stats[player]) {
    stats[player] = {
      games: 0,
      wins: 0,
      kills: 0,
      firstBloods: 0,
      winTurnTotal: 0,
      winTurnCount: 0,
      winTurns: [],
      commanders: {},
      commanderStats: {},
      killerCounts: {},
      victimCounts: {},
    };
  }
  return stats[player];
}

function getMaxCountKey(counts) {
  let bestKey = '';
  let bestCount = -1;
  Object.entries(counts).forEach(([key, count]) => {
    if (count > bestCount || (count === bestCount && key < bestKey)) {
      bestCount = count;
      bestKey = key;
    }
  });
  return bestCount > 0 ? bestKey : '—';
}

function formatPercent(value) {
  return `${Number.isFinite(value) ? value.toFixed(1) : 0}%`;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function getMean(values) {
  if (!values.length) {
    return 0;
  }
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function getStandardDeviation(values, meanValue = getMean(values)) {
  if (values.length < 2) {
    return 0;
  }

  const variance = values.reduce((sum, value) => sum + ((value - meanValue) ** 2), 0) / values.length;
  return Math.sqrt(variance);
}

function getCommanderConfidence(games) {
  if (games >= 15) {
    return { label: 'High', levelClass: 'high', score: 3, percent: 100 };
  }
  if (games >= 6) {
    return { label: 'Medium', levelClass: 'medium', score: 2, percent: 65 };
  }
  return { label: 'Low', levelClass: 'low', score: 1, percent: 35 };
}

function loadCommanderPowerLevels() {
  return appState.powerLevels && typeof appState.powerLevels === 'object' ? appState.powerLevels : {};
}

function saveCommanderPowerLevels(levels) {
  appState = normalizeAppStateData({
    ...appState,
    powerLevels: levels && typeof levels === 'object' ? levels : {},
  });
  persistLocalState(appState);
  queueCloudSync(0);
}

function setCommanderExpectedPower(commander, value) {
  const levels = loadCommanderPowerLevels();
  const canonicalCommander = canonicalizeIdentityValue(commander, buildCanonicalIdentityMapFromValues(Object.keys(levels).concat(knownCommanders)));
  if (!canonicalCommander) {
    return;
  }

  if (typeof value !== 'number' || Number.isNaN(value)) {
    delete levels[canonicalCommander];
  } else {
    levels[canonicalCommander] = Math.min(10, Math.max(0, Math.round(value * 10) / 10));
  }
  saveCommanderPowerLevels(levels);
}

function getCommanderExpectedPower(commander) {
  const levels = loadCommanderPowerLevels();
  const canonicalCommander = canonicalizeIdentityValue(commander, buildCanonicalIdentityMapFromValues(Object.keys(levels).concat(knownCommanders)));
  return typeof levels[canonicalCommander] === 'number' ? levels[canonicalCommander] : '';
}

function normalizeRecordEntry(entry, index = 0) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const playerMap = buildCanonicalIdentityMapFromValues(knownPlayers);
  const commanderMap = buildCanonicalIdentityMapFromValues(knownCommanders.concat(Object.keys(loadCommanderPowerLevels())));

  const isCustom = Boolean(entry.isCustom);
  const key = isCustom ? '' : String(entry.key || '').trim();
  const title = String(entry.title || '').trim();
  if (!title) {
    return null;
  }

  const fallbackId = key || `${title.toLowerCase().replace(/[^a-z0-9]+/g, '-') || 'record'}-${index}`;
  const resolvedValue = String(entry.value || '').trim();
  const resolvedHolder = canonicalizeIdentityValue(entry.holder, playerMap);
  const resolvedCommander = canonicalizeIdentityValue(entry.commander, commanderMap);
  const resolvedDate = String(entry.date || '').trim();
  const resolvedNotes = String(entry.notes || '').trim();

  return {
    id: String(entry.id || fallbackId),
    key,
    title,
    unit: String(entry.unit || '').trim(),
    value: resolvedValue,
    holder: resolvedHolder,
    commander: resolvedCommander,
    date: resolvedDate,
    notes: resolvedNotes,
    manualValue: isCustom ? '' : String(typeof entry.manualValue !== 'undefined' ? entry.manualValue : resolvedValue).trim(),
    manualHolder: isCustom ? '' : canonicalizeIdentityValue(typeof entry.manualHolder !== 'undefined' ? entry.manualHolder : resolvedHolder, playerMap),
    manualCommander: isCustom ? '' : canonicalizeIdentityValue(typeof entry.manualCommander !== 'undefined' ? entry.manualCommander : resolvedCommander, commanderMap),
    manualDate: isCustom ? '' : String(typeof entry.manualDate !== 'undefined' ? entry.manualDate : resolvedDate).trim(),
    manualNotes: isCustom ? '' : String(typeof entry.manualNotes !== 'undefined' ? entry.manualNotes : resolvedNotes).trim(),
    isCustom,
  };
}

function getDefaultRecords() {
  return DEFAULT_RECORD_DEFINITIONS.map((definition, index) => normalizeRecordEntry({
    id: definition.key,
    key: definition.key,
    title: definition.title,
    unit: definition.unit,
    value: '',
    holder: '',
    commander: '',
    date: '',
    notes: '',
    manualValue: '',
    manualHolder: '',
    manualCommander: '',
    manualDate: '',
    manualNotes: '',
    isCustom: false,
  }, index));
}

function parseRecordNumericValue(value) {
  const parsedValue = Number.parseFloat(String(value || '').trim());
  return Number.isFinite(parsedValue) ? parsedValue : null;
}

function formatRecordNumericValue(value) {
  if (!Number.isFinite(value)) {
    return '';
  }

  return Number.isInteger(value) ? String(value) : String(Math.round(value * 100) / 100);
}

function isLowerRecordBetter(key) {
  return key === 'earliest-turn-win' || key === 'fastest-elimination';
}

function chooseBetterRecordCandidate(currentCandidate, nextCandidate, key) {
  if (!nextCandidate || nextCandidate.numericValue === null) {
    return currentCandidate;
  }

  if (!currentCandidate || currentCandidate.numericValue === null) {
    return nextCandidate;
  }

  if (nextCandidate.numericValue === currentCandidate.numericValue) {
    if (currentCandidate.source === 'manual' && nextCandidate.source !== 'manual') {
      return currentCandidate;
    }
    if (nextCandidate.source === 'manual' && currentCandidate.source !== 'manual') {
      return nextCandidate;
    }
    return currentCandidate;
  }

  const lowerIsBetter = isLowerRecordBetter(key);
  const isBetter = lowerIsBetter
    ? nextCandidate.numericValue < currentCandidate.numericValue
    : nextCandidate.numericValue > currentCandidate.numericValue;
  return isBetter ? nextCandidate : currentCandidate;
}

function createRecordCandidate({ value, holder = '', commander = '', date = '', notes = '', source = 'game' }) {
  const numericValue = parseRecordNumericValue(value);
  if (numericValue === null) {
    return null;
  }

  return {
    numericValue,
    value: formatRecordNumericValue(numericValue),
    holder: String(holder || '').trim(),
    commander: String(commander || '').trim(),
    date: String(date || '').trim(),
    notes: String(notes || '').trim(),
    source,
  };
}

function buildManualRecordCandidate(entry) {
  if (!entry || entry.isCustom || !entry.key) {
    return null;
  }

  return createRecordCandidate({
    value: entry.manualValue,
    holder: entry.manualHolder,
    commander: entry.manualCommander,
    date: entry.manualDate,
    notes: entry.manualNotes,
    source: 'manual',
  });
}

function getWinningRow(game) {
  const rows = getGameRows(game).slice().sort((a, b) => (a.place || 999) - (b.place || 999));
  return rows.find((row) => row.place === 1) || rows[0] || null;
}

function getGameRecordCandidates(game) {
  const candidates = new Map();
  const rows = getGameRows(game);
  const winningRow = getWinningRow(game);
  const liveSummary = game?.liveSummary && typeof game.liveSummary === 'object' ? game.liveSummary : null;
  const recordStats = liveSummary?.recordStats && typeof liveSummary.recordStats === 'object' ? liveSummary.recordStats : null;

  const setCandidate = (key, candidate) => {
    if (!candidate) {
      return;
    }
    candidates.set(key, chooseBetterRecordCandidate(candidates.get(key) || null, candidate, key));
  };

  if (winningRow && typeof liveSummary?.turnNumber === 'number') {
    setCandidate('earliest-turn-win', createRecordCandidate({
      value: liveSummary.turnNumber,
      holder: winningRow.player,
      commander: winningRow.commander,
      date: game.date,
      notes: liveSummary.alternateWinCondition ? `Alternate win: ${liveSummary.alternateWinCondition}` : '',
    }));
  }

  rows.forEach((row) => {
    if (typeof row.kills === 'number' && row.kills > 0) {
      setCandidate('most-kills-one-game', createRecordCandidate({
        value: row.kills,
        holder: row.player,
        commander: row.commander,
        date: game.date,
        notes: Array.isArray(row.killed) && row.killed.length ? `Killed: ${row.killed.join(', ')}` : '',
      }));
    }
  });

  if (typeof liveSummary?.durationMinutes === 'number') {
    setCandidate('longest-game', createRecordCandidate({
      value: liveSummary.durationMinutes,
      holder: winningRow?.player || '',
      commander: winningRow?.commander || '',
      date: game.date,
      notes: `${liveSummary.durationMinutes} minute${liveSummary.durationMinutes === 1 ? '' : 's'}`,
    }));
  }

  if (recordStats?.highestSingleHit) {
    setCandidate('highest-damage-dealt', createRecordCandidate(recordStats.highestSingleHit));
  }

  if (recordStats?.highestLifeLossHit) {
    setCandidate('highest-hit-to-life', createRecordCandidate(recordStats.highestLifeLossHit));
  }

  if (recordStats?.highestLifeTotal) {
    setCandidate('highest-life-total', createRecordCandidate(recordStats.highestLifeTotal));
  }

  if (recordStats?.mostCommanderDamage) {
    setCandidate('most-commander-damage', createRecordCandidate(recordStats.mostCommanderDamage));
  }

  if (recordStats?.fastestElimination) {
    setCandidate('fastest-elimination', createRecordCandidate(recordStats.fastestElimination));
  }

  return candidates;
}

function resolvePredefinedRecord(entry, games) {
  const manualCandidate = buildManualRecordCandidate(entry);
  let bestCandidate = manualCandidate;

  games.forEach((game) => {
    const gameCandidate = getGameRecordCandidates(game).get(entry.key) || null;
    bestCandidate = chooseBetterRecordCandidate(bestCandidate, gameCandidate, entry.key);
  });

  if (!bestCandidate) {
    return {
      ...entry,
      value: '',
      holder: '',
      commander: '',
      date: '',
      notes: '',
    };
  }

  return {
    ...entry,
    value: bestCandidate.value,
    holder: bestCandidate.holder,
    commander: bestCandidate.commander,
    date: bestCandidate.date,
    notes: bestCandidate.notes,
  };
}

function mergeRecordsWithDefaults(records, games = appState.games) {
  const normalized = Array.isArray(records)
    ? records.map((entry, index) => normalizeRecordEntry(entry, index)).filter(Boolean)
    : [];

  const predefinedByKey = new Map(
    normalized
      .filter((entry) => !entry.isCustom && entry.key)
      .map((entry) => [entry.key, entry]),
  );

  const mergedDefaults = getDefaultRecords().map((entry) => ({
    ...entry,
    ...(predefinedByKey.get(entry.key) || {}),
    id: entry.key,
    key: entry.key,
    title: entry.title,
    unit: entry.unit,
    isCustom: false,
  })).map((entry) => resolvePredefinedRecord(entry, Array.isArray(games) ? games : []));

  const customRecords = normalized
    .filter((entry) => entry.isCustom)
    .sort((a, b) => a.title.localeCompare(b.title));

  return [...mergedDefaults, ...customRecords];
}

function loadRecords() {
  return mergeRecordsWithDefaults(appState.records, appState.games);
}

function saveRecords(records) {
  appState = normalizeAppStateData({
    ...appState,
    records: Array.isArray(records) ? records : [],
  });
  appState.records = mergeRecordsWithDefaults(appState.records, appState.games);
  persistLocalState(appState);
  queueCloudSync();
}

function getGameById(id) {
  const games = loadGames();
  return games.find((game) => game.id === id) || null;
}

function setEditMode(game) {
  if (!form || !game) {
    return;
  }

  history.replaceState(null, '', 'index.html');
  editingGameId = game.id;

  dateInput.value = game.date || new Date().toISOString().slice(0, 10);
  notesInput.value = game.notes || '';
  if (firstBloodPlayerInput) {
    firstBloodPlayerInput.value = game.liveSummary?.firstBlood?.actorPlayer || '';
  }
  if (firstBloodTurnInput) {
    firstBloodTurnInput.value = game.liveSummary?.firstBlood?.turnNumber || '';
  }
  if (winningTurnInput) {
    winningTurnInput.value = game.liveSummary?.turnNumber || '';
  }

  resetPlayerTable();
  const existingRows = getGameRows(game);

  playerTableBody.innerHTML = '';
  existingRows.forEach((row) => {
    addPlayerRow(row);
  });

  const submitButton = form.querySelector('button[type="submit"]');
  if (submitButton) {
    submitButton.textContent = 'Save changes';
  }
}

function resetEditMode() {
  editingGameId = null;
  if (!form) {
    return;
  }
  const submitButton = form.querySelector('button[type="submit"]');
  if (submitButton) {
    submitButton.textContent = 'Save game';
  }
}

function deleteGame(gameId) {
  const games = loadGames();
  const remaining = games.filter((game) => game.id !== gameId);
  saveGames(remaining);
  refresh();
}

function setAuthAuditStatus(message, tone = 'muted') {
  if (!authAuditStatus) {
    return;
  }

  authAuditStatus.textContent = message;
  authAuditStatus.classList.remove('status-neutral', 'status-success', 'status-error', 'status-muted');
  authAuditStatus.classList.add(`status-${tone}`);
}

function formatAuthAuditTimestamp(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return 'Unknown time';
  }

  return date.toLocaleString([], {
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    second: '2-digit',
  });
}

function normalizeAuditDateInputValue(value) {
  const normalized = String(value || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(normalized) ? normalized : '';
}

function updateAuthAuditFilterStateFromInputs() {
  authAuditFilterState = {
    result: authAuditFilterResultSelect?.value || 'all',
    action: authAuditFilterActionSelect?.value || 'all',
    user: String(authAuditFilterUserInput?.value || '').trim().toLowerCase(),
    from: normalizeAuditDateInputValue(authAuditFilterFromInput?.value || ''),
    to: normalizeAuditDateInputValue(authAuditFilterToInput?.value || ''),
  };
  authAuditCurrentPage = 1;
}

function resetAuthAuditFilters() {
  if (authAuditFilterResultSelect) {
    authAuditFilterResultSelect.value = 'all';
  }
  if (authAuditFilterActionSelect) {
    authAuditFilterActionSelect.value = 'all';
  }
  if (authAuditFilterUserInput) {
    authAuditFilterUserInput.value = '';
  }
  if (authAuditFilterFromInput) {
    authAuditFilterFromInput.value = '';
  }
  if (authAuditFilterToInput) {
    authAuditFilterToInput.value = '';
  }
  updateAuthAuditFilterStateFromInputs();
}

function getAuthAuditAction(entry) {
  const explicitAction = String(entry?.action || '').trim().toLowerCase();
  if (explicitAction) {
    return explicitAction;
  }

  const path = String(entry?.path || '').trim().toLowerCase();
  const reason = String(entry?.reason || '').trim().toLowerCase();
  const status = Number(entry?.status || 0);

  if (!entry?.success || status >= 400 || reason.includes('error') || reason.includes('failed') || reason.includes('invalid') || reason.includes('unable')) {
    return 'error';
  }

  if (reason.startsWith('saved game:')) {
    return 'game-save';
  }
  if (reason.startsWith('updated game:')) {
    return 'game-update';
  }
  if (reason.startsWith('deleted game:')) {
    return 'game-delete';
  }
  if (reason.startsWith('saved deck:')) {
    return 'deck-save';
  }
  if (reason.startsWith('updated deck:')) {
    return 'deck-update';
  }
  if (reason.startsWith('deleted deck:')) {
    return 'deck-delete';
  }

  if (path.startsWith('/api/session') || reason.includes('authenticated') || reason.includes('authentication') || reason.includes('incorrect pod access code') || reason.includes('display name is required') || reason.includes('pod access code is required')) {
    return 'authentication';
  }

  return 'authentication';
}

function updateAuthAuditPaginationState() {
  const parsedPageSize = Number.parseInt(String(authAuditPageSizeSelect?.value || authAuditPageSize || '10'), 10);
  authAuditPageSize = Number.isFinite(parsedPageSize) && parsedPageSize > 0 ? parsedPageSize : 10;
  authAuditCurrentPage = Math.max(1, authAuditCurrentPage);
}

function getFilteredAuthAuditLogs(logs) {
  if (!Array.isArray(logs) || !logs.length) {
    return [];
  }

  const { result, action, user, from, to } = authAuditFilterState;
  return logs.filter((entry) => {
    const isSuccess = Boolean(entry?.success);
    if (result === 'failed' && isSuccess) {
      return false;
    }
    if (result === 'success' && !isSuccess) {
      return false;
    }

    if (action !== 'all' && getAuthAuditAction(entry) !== action) {
      return false;
    }

    if (user) {
      const haystack = [
        String(entry?.auth?.displayName || ''),
        String(entry?.auth?.userId || ''),
        String(entry?.user || ''),
        String(entry?.normalizedUser || ''),
      ].join(' ').toLowerCase();
      if (!haystack.includes(user)) {
        return false;
      }
    }

    const isoDate = String(entry?.timestamp || '').slice(0, 10);
    if (from && (!isoDate || isoDate < from)) {
      return false;
    }
    if (to && (!isoDate || isoDate > to)) {
      return false;
    }

    return true;
  });
}

function getAuthAuditStatusSummary() {
  const visibleCount = getFilteredAuthAuditLogs(authAuditLogs).length;
  return `Showing ${visibleCount} of ${authAuditLogs.length} stored authentication events.`;
}

function getPagedAuthAuditLogs(logs) {
  const visibleLogs = getFilteredAuthAuditLogs(logs);
  const totalPages = Math.max(1, Math.ceil(visibleLogs.length / authAuditPageSize));
  authAuditCurrentPage = Math.min(Math.max(1, authAuditCurrentPage), totalPages);
  const startIndex = (authAuditCurrentPage - 1) * authAuditPageSize;
  const endIndex = startIndex + authAuditPageSize;

  return {
    visibleLogs,
    pagedLogs: visibleLogs.slice(startIndex, endIndex),
    totalPages,
    startIndex,
    endIndex: Math.min(endIndex, visibleLogs.length),
  };
}

function updateAuthAuditPaginationControls(totalVisible = 0, totalPages = 1, startIndex = 0, endIndex = 0) {
  if (authAuditPageSizeSelect) {
    authAuditPageSizeSelect.value = String(authAuditPageSize);
  }

  if (authAuditPagePrevButton) {
    authAuditPagePrevButton.disabled = authAuditCurrentPage <= 1 || totalVisible === 0;
  }

  if (authAuditPageNextButton) {
    authAuditPageNextButton.disabled = authAuditCurrentPage >= totalPages || totalVisible === 0;
  }

  if (authAuditPageSummary) {
    authAuditPageSummary.textContent = totalVisible
      ? `Page ${authAuditCurrentPage} of ${totalPages} · ${startIndex + 1}-${endIndex} of ${totalVisible}`
      : 'Page 1 of 1';
  }
}

function updateAuthAuditStatusSummary(tone = 'neutral') {
  if (!authAuditSection || authAuditSection.hidden) {
    return;
  }

  setAuthAuditStatus(getAuthAuditStatusSummary(), tone);
}

function setRegisteredAccountsStatus(message, tone = 'muted') {
  if (!registeredAccountsStatus) {
    return;
  }

  registeredAccountsStatus.textContent = message;
  registeredAccountsStatus.classList.remove('status-neutral', 'status-success', 'status-error', 'status-muted');
  registeredAccountsStatus.classList.add(`status-${tone}`);
}

function renderRegisteredAccounts() {
  if (!registeredAccountsSection || !registeredAccountsBody) {
    return;
  }

  if (!hasSyncCredentials() || !isCurrentSyncUserAdmin()) {
    registeredAccountsSection.hidden = true;
    registeredAccountsBody.innerHTML = '';
    return;
  }

  registeredAccountsSection.hidden = false;

  if (!registeredAccounts.length) {
    registeredAccountsBody.innerHTML = '<tr><td colspan="4">No registered accounts found yet.</td></tr>';
    applyResponsiveTableLabels();
    return;
  }

  registeredAccountsBody.innerHTML = registeredAccounts.map((entry) => {
    const displayName = entry?.displayName || 'Unknown';
    const userId = entry?.userId || '';
    const role = entry?.role || 'member';
    const source = entry?.source || 'unknown';
    return `
      <tr>
        <td>${escapeHtml(displayName)}</td>
        <td>${escapeHtml(userId)}</td>
        <td>${escapeHtml(role)}</td>
        <td>${escapeHtml(source)}</td>
      </tr>`;
  }).join('');

  applyResponsiveTableLabels();
}

async function refreshRegisteredAccounts({ force = false } = {}) {
  if (!registeredAccountsSection || !registeredAccountsBody || !registeredAccountsStatus) {
    return;
  }

  if (!hasSyncCredentials() || !isCurrentSyncUserAdmin()) {
    registeredAccountsSection.hidden = true;
    registeredAccounts = [];
    registeredAccountsLoadedForUserId = '';
    registeredAccountsLastLoadedAt = 0;
    setRegisteredAccountsStatus('Registered accounts are only visible to admins.', 'muted');
    renderRegisteredAccounts();
    return;
  }

  const currentUserId = getCurrentSyncUserId();
  const now = Date.now();
  const shouldRefresh = force
    || registeredAccountsLoadedForUserId !== currentUserId
    || !registeredAccounts.length
    || now - registeredAccountsLastLoadedAt >= AUTH_AUDIT_REFRESH_INTERVAL_MS;

  registeredAccountsSection.hidden = false;
  renderRegisteredAccounts();
  if (!shouldRefresh || registeredAccountsLoading) {
    return;
  }

  registeredAccountsLoading = true;
  setRegisteredAccountsStatus('Loading registered accounts...', 'neutral');

  try {
    const payload = await cloudRequest(REGISTERED_ACCOUNTS_ENDPOINT, { method: 'GET' });
    registeredAccounts = Array.isArray(payload?.accounts) ? payload.accounts : [];
    registeredAccountsLoadedForUserId = currentUserId;
    registeredAccountsLastLoadedAt = Date.now();
    setRegisteredAccountsStatus(`Showing ${registeredAccounts.length} registered account${registeredAccounts.length === 1 ? '' : 's'}.`, 'success');
    renderRegisteredAccounts();
  } catch (error) {
    if (error?.status === 403) {
      registeredAccounts = [];
      renderRegisteredAccounts();
      setRegisteredAccountsStatus('Connected as non-admin on the server. Sign in as Baily to view registered accounts.', 'error');
      return;
    }

    setRegisteredAccountsStatus(`Unable to load registered accounts: ${error.message}`, 'error');
  } finally {
    registeredAccountsLoading = false;
  }
}

function renderSyncDebugLog() {
  if (!syncDebugList || !syncDebugStatus) {
    return;
  }

  if (!syncDebugLogs.length) {
    syncDebugStatus.textContent = 'No sync events yet.';
    syncDebugStatus.className = 'sync-status status-muted';
    syncDebugList.innerHTML = '<tr><td colspan="5" class="history-empty-state">No sync events yet.</td></tr>';
    return;
  }

  syncDebugStatus.textContent = `Last sync: ${syncDebugLogs[0]?.timestamp || 'unknown'}`;
  syncDebugStatus.className = 'sync-status status-success';
  syncDebugList.innerHTML = syncDebugLogs.map((entry) => {
    const timestamp = entry?.timestamp ? new Date(entry.timestamp).toLocaleString() : 'Unknown';
    const event = entry?.event || 'Unknown event';
    const decks = entry?.decks ?? '-';
    const lists = entry?.lists ?? '-';
    const user = entry?.user || 'Unknown';
    return `
      <tr>
        <td>${escapeHtml(timestamp)}</td>
        <td>${escapeHtml(event)}</td>
        <td>${escapeHtml(String(decks))}</td>
        <td>${escapeHtml(String(lists))}</td>
        <td>${escapeHtml(user)}</td>
      </tr>`;
  }).join('');

  applyResponsiveTableLabels();
}

function buildAuthAuditExportBaseName() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');
  return `auth-audit-${y}${m}${d}-${hh}${mm}${ss}`;
}

function triggerFileDownload(filename, mimeType, content) {
  const blob = new Blob([content], { type: mimeType });
  const blobUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = blobUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(blobUrl);
}

function buildCsvCell(value) {
  const text = String(value ?? '');
  const escaped = text.replace(/"/g, '""');
  return `"${escaped}"`;
}

function exportFilteredAuthAuditAsJson() {
  const logs = getFilteredAuthAuditLogs(authAuditLogs);
  if (!logs.length) {
    updateAuthAuditStatusSummary('muted');
    return;
  }

  const payload = {
    exportedAt: new Date().toISOString(),
    filters: { ...authAuditFilterState },
    count: logs.length,
    totalAvailable: authAuditLogs.length,
    logs,
  };
  triggerFileDownload(`${buildAuthAuditExportBaseName()}.json`, 'application/json;charset=utf-8', `${JSON.stringify(payload, null, 2)}\n`);
  setAuthAuditStatus(`Exported ${logs.length} authentication events to JSON.`, 'success');
}

function exportFilteredAuthAuditAsCsv() {
  const logs = getFilteredAuthAuditLogs(authAuditLogs);
  if (!logs.length) {
    updateAuthAuditStatusSummary('muted');
    return;
  }

  const header = [
    'timestamp',
    'action',
    'result',
    'reason',
    'statusCode',
    'method',
    'path',
    'displayName',
    'userId',
    'rawUser',
    'normalizedUser',
    'ip',
    'tokenHint',
    'authMode',
    'userAgent',
  ];
  const lines = [header.map((value) => buildCsvCell(value)).join(',')];

  logs.forEach((entry) => {
    const row = [
      entry?.timestamp || '',
      getAuthAuditAction(entry),
      entry?.success ? 'success' : 'failed',
      entry?.reason || '',
      Number.isFinite(Number(entry?.status)) ? Number(entry.status) : '',
      entry?.method || '',
      entry?.path || '',
      entry?.auth?.displayName || '',
      entry?.auth?.userId || '',
      entry?.user || '',
      entry?.normalizedUser || '',
      entry?.ip || '',
      entry?.tokenHint || '',
      entry?.auth?.mode || '',
      entry?.userAgent || '',
    ];
    lines.push(row.map((value) => buildCsvCell(value)).join(','));
  });

  triggerFileDownload(`${buildAuthAuditExportBaseName()}.csv`, 'text/csv;charset=utf-8', `${lines.join('\n')}\n`);
  setAuthAuditStatus(`Exported ${logs.length} authentication events to CSV.`, 'success');
}

function renderAuthAuditLogs() {
  if (!authAuditSection || !authAuditList) {
    return;
  }

  if (!hasSyncCredentials() || !isCurrentSyncUserAdmin()) {
    authAuditSection.hidden = true;
    authAuditList.innerHTML = '';
    return;
  }

  authAuditSection.hidden = false;
  updateAuthAuditPaginationState();
  const { visibleLogs, pagedLogs, totalPages, startIndex, endIndex } = getPagedAuthAuditLogs(authAuditLogs);
  updateAuthAuditPaginationControls(visibleLogs.length, totalPages, startIndex, endIndex);

  if (!visibleLogs.length) {
    authAuditList.innerHTML = '<p class="history-empty-state">No login attempts logged yet.</p>';
    return;
  }

  authAuditList.innerHTML = pagedLogs.map((entry) => {
    const status = entry?.success ? 'success' : 'failed';
    const action = getAuthAuditAction(entry);
    const reason = entry?.reason || (entry?.success ? 'Authenticated.' : 'Authentication failed.');
    const who = entry?.auth?.displayName || entry?.user || 'Unknown user';
    const userId = entry?.auth?.userId || entry?.normalizedUser || '';
    const ip = entry?.ip || 'Unknown IP';
    const method = entry?.method || 'UNKNOWN';
    const path = entry?.path || '/api/state';
    const statusCode = Number.isFinite(Number(entry?.status)) ? Number(entry.status) : 0;
    const tokenHint = entry?.tokenHint || '';

    return `
      <article class="auth-audit-item auth-audit-item--${entry?.success ? 'success' : 'failed'}">
        <div class="auth-audit-item-topline">
          <strong>${escapeHtml(status.toUpperCase())}</strong>
          <span>${escapeHtml(formatAuthAuditTimestamp(entry?.timestamp))}</span>
        </div>
        <p class="auth-audit-item-reason">${escapeHtml(reason)}</p>
        <div class="auth-audit-item-meta">
          <span class="history-item-fact">Action: ${escapeHtml(action)}</span>
          <span class="history-item-fact">User: ${escapeHtml(who)}</span>
          ${userId ? `<span class="history-item-fact">ID: ${escapeHtml(userId)}</span>` : ''}
          <span class="history-item-fact">HTTP: ${escapeHtml(method)} ${escapeHtml(path)}</span>
          ${statusCode ? `<span class="history-item-fact">Status: ${escapeHtml(String(statusCode))}</span>` : ''}
          <span class="history-item-fact">IP: ${escapeHtml(ip)}</span>
          ${tokenHint ? `<span class="history-item-fact">Token: ${escapeHtml(tokenHint)}</span>` : ''}
        </div>
      </article>`;
  }).join('');
}

async function refreshAuthAuditLogs({ force = false } = {}) {
  if (!authAuditSection || !authAuditList || !authAuditStatus) {
    return;
  }

  if (!hasSyncCredentials() || !isCurrentSyncUserAdmin()) {
    authAuditSection.hidden = true;
    authAuditLogs = [];
    authAuditLoadedForUserId = '';
    authAuditLastLoadedAt = 0;
    setAuthAuditStatus('Authentication audit logs are only visible to admins.', 'muted');
    renderAuthAuditLogs();
    return;
  }

  const currentUserId = getCurrentSyncUserId();
  const now = Date.now();
  const shouldRefresh = force
    || authAuditLoadedForUserId !== currentUserId
    || !authAuditLogs.length
    || now - authAuditLastLoadedAt >= AUTH_AUDIT_REFRESH_INTERVAL_MS;

  authAuditSection.hidden = false;
  renderAuthAuditLogs();
  if (!shouldRefresh || authAuditLoading) {
    return;
  }

  authAuditLoading = true;
  setAuthAuditStatus('Loading authentication logs...', 'neutral');

  try {
    const payload = await cloudRequest(`${AUTH_AUDIT_ENDPOINT}?limit=1000`, { method: 'GET' });
    authAuditLogs = Array.isArray(payload?.logs) ? payload.logs : [];
    authAuditCurrentPage = 1;
    authAuditLoadedForUserId = currentUserId;
    authAuditLastLoadedAt = Date.now();
    updateAuthAuditStatusSummary('success');
    renderAuthAuditLogs();
  } catch (error) {
    if (error?.status === 403) {
      authAuditSection.hidden = false;
      authAuditLogs = [];
      renderAuthAuditLogs();
      setAuthAuditStatus('Connected as non-admin on the server. Sign in as Baily with the matching pod credentials, then refresh logs.', 'error');
      return;
    }

    setAuthAuditStatus(`Unable to load authentication logs: ${error.message}`, 'error');
  } finally {
    authAuditLoading = false;
  }
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getUniqueValues(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function getCanonicalValues(values) {
  return getUniqueValuesBySimilarity(Array.from(buildCanonicalIdentityMapFromValues(values).values()));
}

if (typeof window !== 'undefined' && !window.getCanonicalValues) {
  window.getCanonicalValues = getCanonicalValues;
}

function buildDatalistOptions(element, values) {
  if (!element) {
    return;
  }

  const normalized = getUniqueValuesBySimilarity(values);
  element.innerHTML = normalized
    .map((value) => `<option value="${escapeHtml(value)}"></option>`)
    .join('');
}

function buildSelectOptions(element, values, selectedValue, placeholder) {
  if (!element) {
    return;
  }

  const normalized = getUniqueValuesBySimilarity(values);
  element.innerHTML = [
    `<option value="">${escapeHtml(placeholder)}</option>`,
    ...normalized.map((value) => {
      const selected = selectedValue && value === selectedValue ? ' selected' : '';
      return `<option value="${escapeHtml(value)}"${selected}>${escapeHtml(value)}</option>`;
    }),
  ].join('');
}

function getUsedPlayersInTable(tableBody, excludeRow) {
  if (!tableBody) return new Set();
  const used = new Set();
  Array.from(tableBody.querySelectorAll('tr')).forEach((row) => {
    if (row === excludeRow) return;
    const input = row.querySelector('input[name="player"]');
    const val = getIdentityKey(input?.value || '');
    if (val) used.add(val);
  });
  return used;
}

function populateRowSelectors(row) {
  if (!row) {
    return;
  }

  const tableBody = row.closest('tbody');
  const isLiveSetupRow = tableBody?.id === 'live-game-player-body';
  const usedPlayers = getUsedPlayersInTable(tableBody, row);
  const availablePlayers = usedPlayers.size
    ? knownPlayers.filter((p) => !usedPlayers.has(getIdentityKey(p)))
    : knownPlayers;
  const availableCommanders = isLiveSetupRow ? getLiveCommanderOptionValues() : knownCommanders;

  const playerMenu = row.querySelector('.player-dropdown-menu');
  const commanderMenu = row.querySelector('.commander-dropdown-menu');

  if (playerMenu) {
    buildDropdownMenu(playerMenu, availablePlayers);
  }

  if (commanderMenu) {
    buildDropdownMenu(commanderMenu, availableCommanders);
  }
}

function buildDropdownMenu(menuElement, values) {
  if (!menuElement) {
    return;
  }
  
  const normalized = getUniqueValuesBySimilarity(values);
  if (!normalized.length) {
    menuElement.innerHTML = '<div class="dropdown-empty">No saved options yet</div>';
    syncMobileLookupSelect(menuElement.closest('.combined-input-wrapper'));
    return;
  }

  menuElement.innerHTML = normalized
    .map((value) => `<div class="dropdown-item" data-value="${escapeHtml(value)}">${escapeHtml(value)}</div>`)
    .join('');

  syncMobileLookupSelect(menuElement.closest('.combined-input-wrapper'));
}

function isMobileDropdownMode() {
  return window.matchMedia('(max-width: 900px)').matches;
}

function syncMobileLookupSelect(wrapper) {
  if (!wrapper) {
    return;
  }

  const menu = wrapper.querySelector('.dropdown-menu');
  if (!menu) {
    return;
  }

  let mobileSelect = wrapper.querySelector('.mobile-native-picker');
  if (!mobileSelect) {
    mobileSelect = document.createElement('select');
    mobileSelect.className = 'mobile-native-picker';
    mobileSelect.setAttribute('aria-label', 'Choose saved option');
    mobileSelect.tabIndex = -1;

    mobileSelect.addEventListener('change', () => {
      if (mobileSelect.value) {
        applyLookupSelection(wrapper, mobileSelect.value);
      }
      mobileSelect.selectedIndex = 0;
    });

    wrapper.appendChild(mobileSelect);
  }

  const optionValues = getUniqueValuesBySimilarity(
    Array.from(menu.querySelectorAll('.dropdown-item')).map((item) => item.dataset.value || '').filter(Boolean),
  );

  const previousValue = mobileSelect.value;
  mobileSelect.innerHTML = [
    `<option value="">${optionValues.length ? 'Choose saved option' : 'No saved options yet'}</option>`,
    ...optionValues.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`),
  ].join('');
  mobileSelect.disabled = optionValues.length === 0;
  if (previousValue && optionValues.includes(previousValue)) {
    mobileSelect.value = previousValue;
  }
}

function openLookupOptions(wrapper) {
  if (!wrapper) {
    return;
  }

  if (isMobileDropdownMode()) {
    const mobileSelect = wrapper.querySelector('.mobile-native-picker');
    if (mobileSelect && !mobileSelect.disabled && typeof mobileSelect.showPicker === 'function') {
      mobileSelect.showPicker();
      return;
    }
  }

  toggleLookupMenu(wrapper);
}

function closeAllDropdownMenus(exceptMenu = null) {
  document.querySelectorAll('.dropdown-menu.active').forEach((activeMenu) => {
    if (activeMenu !== exceptMenu) {
      activeMenu.classList.remove('active');
      activeMenu.closest('.combined-input-wrapper')?.classList.remove('dropdown-open');
    }
  });
  updateDropdownLayeringState();
}

function toggleLookupMenu(wrapper) {
  if (!wrapper) {
    return;
  }

  const menu = wrapper.querySelector('.dropdown-menu');
  if (!menu) {
    return;
  }

  const shouldOpen = !menu.classList.contains('active');
  closeAllDropdownMenus(menu);
  menu.classList.toggle('active', shouldOpen);
  wrapper.classList.toggle('dropdown-open', shouldOpen);
  updateDropdownLayeringState();
}

function applyLookupSelection(wrapper, value) {
  if (!wrapper) {
    return;
  }

  const input = wrapper.querySelector('.lookup-input');
  const menu = wrapper.querySelector('.dropdown-menu');
  if (!input) {
    return;
  }

  input.value = value || '';
  input.dispatchEvent(new Event('input', { bubbles: true }));
  input.dispatchEvent(new Event('change', { bubbles: true }));

  if (menu) {
    menu.classList.remove('active');
  }
  wrapper.classList.remove('dropdown-open');

  input.focus();
  updateDropdownLayeringState();
}

function attachLookupWrapperHandlers(scope = document) {
  scope.querySelectorAll('.combined-input-wrapper').forEach((wrapper) => {
    if (wrapper.dataset.dropdownHandlersAttached) {
      return;
    }

    const button = wrapper.querySelector('.dropdown-button');
    const menu = wrapper.querySelector('.dropdown-menu');
    const input = wrapper.querySelector('.lookup-input');
    if (!button || !menu || !input) {
      return;
    }

    wrapper.dataset.dropdownHandlersAttached = 'true';

    syncMobileLookupSelect(wrapper);

    input.addEventListener('click', (event) => {
      event.stopPropagation();
      openLookupOptions(wrapper);
    });

    button.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();
      openLookupOptions(wrapper);
    });

    input.addEventListener('blur', async () => {
      if (input.name === 'commander') {
        const canonicalValue = await canonicalizeCommanderInputValue(input.value);
        if (canonicalValue && canonicalValue !== input.value.trim()) {
          input.value = canonicalValue;
          input.dispatchEvent(new Event('input', { bubbles: true }));
        }
      }
      // When a player field loses focus, refresh sibling rows so they exclude this player.
      if (input.name === 'player') {
        const tableBody = wrapper.closest('tbody');
        const currentRow = wrapper.closest('tr');
        if (tableBody) {
          Array.from(tableBody.querySelectorAll('tr')).forEach((siblingRow) => {
            if (siblingRow !== currentRow) populateRowSelectors(siblingRow);
          });
        }
      }
    });

    menu.addEventListener('mousedown', (event) => {
      event.preventDefault();
    });

    menu.addEventListener('click', (event) => {
      const item = event.target.closest('.dropdown-item');
      if (!item) {
        return;
      }

      event.preventDefault();
      event.stopPropagation();
      applyLookupSelection(wrapper, item.dataset.value || '');

      // If this was a player selection, refresh sibling rows to exclude the chosen player.
      if (input.name === 'player') {
        const tableBody = wrapper.closest('tbody');
        const currentRow = wrapper.closest('tr');
        if (tableBody) {
          Array.from(tableBody.querySelectorAll('tr')).forEach((siblingRow) => {
            if (siblingRow !== currentRow) populateRowSelectors(siblingRow);
          });
        }
      }
    });
  });
}

function updateDropdownLayeringState() {
  const hasActiveDropdown = document.querySelector('.dropdown-menu.active');
  document.querySelectorAll('.player-table-wrapper').forEach((wrapper) => {
    wrapper.classList.toggle('dropdown-open', Boolean(hasActiveDropdown));
  });
}

// Global outside-click handler — registered once
document.addEventListener('click', (e) => {
  if (!e.target.closest('.combined-input-wrapper')) {
    closeAllDropdownMenus();
  }
});

function refreshRowSelectors() {
  Array.from(document.querySelectorAll('tr')).forEach((row) => {
    populateRowSelectors(row);
    attachLookupWrapperHandlers(row);
  });
}

function updateFormDatalists(games) {
  const players = [];
  const commanders = [];
  const deckLists = loadDeckLists();
  const decks = loadDecks();

  games.forEach((game) => {
    getGameRows(game).forEach((row) => {
      if (row.player) {
        players.push(row.player);
      }
      if (row.commander) {
        commanders.push(row.commander);
      }
    });
  });

  deckLists.forEach((entry) => {
    const commander = (entry?.commander || '').trim();
    const owner = (entry?.owner || '').trim();
    if (owner) {
      players.push(owner);
    }
    if (commander) {
      commanders.push(commander);
    }
  });

  decks.forEach((deck) => {
    if (deck?.owner) {
      players.push(deck.owner);
    }
    if (deck?.commander?.name) {
      commanders.push(deck.commander.name);
    }
  });

  knownPlayers = getUniqueValuesBySimilarity(Array.from(buildCanonicalIdentityMapFromValues(players).values()));
  knownCommanders = getUniqueValuesBySimilarity(Array.from(buildCanonicalIdentityMapFromValues(commanders).values()));

  if (playerDatalist) {
    buildDatalistOptions(playerDatalist, knownPlayers);
  }
  if (commanderDatalist) {
    buildDatalistOptions(commanderDatalist, window.CommanderLiveGame?.hasView?.() ? getLiveCommanderOptionValues() : knownCommanders);
  }
  if (firstBloodPlayerMenu) {
    buildDropdownMenu(firstBloodPlayerMenu, knownPlayers);
  }

  refreshRowSelectors();
  attachLookupWrapperHandlers(form || document);
  populateDeckBuilderLookupMenus();
  window.CommanderRecords?.populateLookupMenus();
}

function normalizeDeckListEntry(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const playerMap = buildCanonicalIdentityMapFromValues(knownPlayers);
  const commanderMap = buildCanonicalIdentityMapFromValues(knownCommanders.concat(Object.keys(loadCommanderPowerLevels())));
  const commander = canonicalizeIdentityValue(entry.commander, commanderMap);
  const commanderOracleId = String(entry.commanderOracleId || '').trim();
  const owner = canonicalizeIdentityValue(entry.owner, playerMap);
  const url = String(entry.url || '').trim();
  if (!commander || !url) {
    return null;
  }

  return {
    id: entry.id || generateId(),
    commander,
    commanderOracleId,
    owner,
    url,
    deckId: String(entry.deckId || '').trim(),
  };
}

function getBestMatchingDeckForCommander(commanderName, decks, ownerKey = '') {
  if (!commanderName || !Array.isArray(decks) || !decks.length) {
    return null;
  }

  const inputKey = getIdentityKey(commanderName);
  if (!inputKey) {
    return null;
  }

  let best = { score: 0, deck: null };
  decks.forEach((deck) => {
    if (!deck?.commander?.name) {
      return;
    }

    if (ownerKey && getIdentityKey(deck.owner || '') !== ownerKey) {
      return;
    }

    const deckCommanderKey = getIdentityKey(deck.commander.name);
    if (!deckCommanderKey) {
      return;
    }

    const score = getStringSimilarity(inputKey, deckCommanderKey);
    if (score > best.score) {
      best = { score, deck };
    }
  });

  return best.score >= 0.9 ? best.deck : null;
}

function getCommanderEquivalenceKey({ name = '', oracleId = '' } = {}) {
  const normalizedOracleId = String(oracleId || '').trim().toLowerCase();
  if (normalizedOracleId) {
    return `oracle:${normalizedOracleId}`;
  }

  const normalizedNameKey = getIdentityKey(name);
  return normalizedNameKey ? `name:${normalizedNameKey}` : '';
}

function resolveLinkedDeckIdForDeckList(entry, decks = loadDecks()) {
  if (!entry) {
    return '';
  }

  const normalizedDeckId = String(entry.deckId || '').trim();
  if (normalizedDeckId && decks.some((deck) => deck.id === normalizedDeckId)) {
    return normalizedDeckId;
  }

  const commanderKey = getCommanderEquivalenceKey({
    name: entry.commander,
    oracleId: entry.commanderOracleId,
  });
  if (!commanderKey) {
    return '';
  }

  const ownerKey = getIdentityKey(entry.owner || '');
  const exactMatch = decks.find((deck) => {
    const deckCommanderKey = getCommanderEquivalenceKey({
      name: deck.commander?.name || '',
      oracleId: deck.commander?.oracleId || '',
    });
    if (deckCommanderKey !== commanderKey) {
      return false;
    }
    if (!ownerKey) {
      return true;
    }
    return getIdentityKey(deck.owner || '') === ownerKey;
  });

  if (exactMatch) {
    return exactMatch.id;
  }

  const fallbackDeck = getBestMatchingDeckForCommander(entry.commander, decks, ownerKey);
  return fallbackDeck?.id || '';
}

function linkDeckListToDeck(deck) {
  const commanderKey = getCommanderEquivalenceKey({
    name: deck?.commander?.name || '',
    oracleId: deck?.commander?.oracleId || '',
  });
  if (!commanderKey) {
    return;
  }

  const deckLists = loadDeckLists().map(normalizeDeckListEntry).filter(Boolean);
  let changed = false;
  const linkedDeckLists = deckLists.map((entry) => {
    const entryCommanderKey = getCommanderEquivalenceKey({
      name: entry.commander,
      oracleId: entry.commanderOracleId,
    });
    if (entryCommanderKey !== commanderKey) {
      return entry;
    }
    if (entry.deckId === deck.id) {
      return entry;
    }
    changed = true;
    return {
      ...entry,
      deckId: deck.id,
    };
  });

  if (changed) {
    saveDeckLists(linkedDeckLists);
  }
}

function getNormalizedDeckUrl(url) {
  try {
    const parsedUrl = new URL(String(url || '').trim());
    if (parsedUrl.protocol !== 'https:') {
      return { error: 'Deck URL must use https.' };
    }

    parsedUrl.hash = '';
    return { value: parsedUrl.toString() };
  } catch (error) {
    return { error: 'Please enter a valid deck URL.' };
  }
}

function getChangedRecordTitlesAfterGameRemoval(gameId) {
  const currentGames = loadGames();
  const nextGames = currentGames.filter((game) => game.id !== gameId);
  const currentRecords = mergeRecordsWithDefaults(appState.records, currentGames);
  const nextRecords = mergeRecordsWithDefaults(appState.records, nextGames);
  const nextRecordsById = new Map(nextRecords.map((record) => [record.id, record]));

  return currentRecords
    .filter((record) => {
      const nextRecord = nextRecordsById.get(record.id);
      return !nextRecord
        || nextRecord.value !== record.value
        || nextRecord.holder !== record.holder
        || nextRecord.commander !== record.commander
        || nextRecord.date !== record.date;
    })
    .map((record) => record.title);
}

function getSortedDeckLists() {
  return loadDeckLists()
    .map(normalizeDeckListEntry)
    .filter(Boolean)
    .sort((a, b) => a.commander.localeCompare(b.commander));
}

async function backfillDeckListCommanderOracleIds() {
  if (deckListOracleBackfillRan) {
    return;
  }

  deckListOracleBackfillRan = true;
  const deckLists = loadDeckLists().map(normalizeDeckListEntry).filter(Boolean);
  const entriesMissingOracleId = deckLists.filter((entry) => !entry.commanderOracleId && entry.commander);
  if (!entriesMissingOracleId.length) {
    return;
  }

  const commanderNames = getUniqueValues(entriesMissingOracleId.map((entry) => entry.commander));
  try {
    const cardsByQuery = await fetchDeckCardsByNamesBulk(commanderNames);
    let changed = false;

    const nextDeckLists = deckLists.map((entry) => {
      if (entry.commanderOracleId || !entry.commander) {
        return entry;
      }

      const card = cardsByQuery.get(getIdentityKey(entry.commander));
      const oracleId = String(card?.oracleId || '').trim();
      if (!oracleId) {
        return entry;
      }

      changed = true;
      return {
        ...entry,
        commanderOracleId: oracleId,
      };
    });

    if (!changed) {
      return;
    }

    saveDeckLists(nextDeckLists);
    refresh();
  } catch (error) {
    console.warn('Unable to backfill deck list oracle IDs:', error);
  }
}

function populateDeckBuilderLookupMenus() {
  window.CommanderDeckBuilder?.populateDeckBuilderLookupMenus?.();
  attachLookupWrapperHandlers(document);
}

function getDeckBuilderHref(deckId) {
  const query = new URLSearchParams();
  if (deckId) {
    query.set('deckId', deckId);
  }

  const queryString = query.toString();
  return `deckbuilder.html${queryString ? `?${queryString}` : ''}`;
}

// --- Color identity helpers ---

const COLOR_LETTER_NAMES = { W: 'White', U: 'Blue', B: 'Black', R: 'Red', G: 'Green' };

function extractColorsFromManaSymbols(text) {
  const colors = new Set();
  const symbols = String(text || '').match(/\{[^}]+\}/g) || [];
  symbols.forEach((sym) => {
    const letters = sym.match(/[WUBRG]/gi) || [];
    letters.forEach((l) => colors.add(l.toUpperCase()));
  });
  return colors;
}

function getCardEffectiveColorIdentity(card) {
  const colors = new Set(
    (Array.isArray(card?.colorIdentity) ? card.colorIdentity : [])
      .map((c) => String(c).toUpperCase())
      .filter((c) => 'WUBRG'.includes(c))
  );
  // Also scan mana cost and oracle text on the card and all faces so phyrexian
  // mana symbols and similar edge cases can't be missed.
  const textsToScan = [card?.manaCost, card?.oracleText];
  if (Array.isArray(card?.cardFaces)) {
    card.cardFaces.forEach((face) => {
      textsToScan.push(face?.manaCost, face?.oracleText);
    });
  }
  textsToScan.forEach((t) => extractColorsFromManaSymbols(t).forEach((c) => colors.add(c)));
  return colors;
}

function getDeckValidationSummary(deck) {
  const normalizedDeck = deck?.cards && deck?.commander ? deck : normalizeDeckRecord(deck);
  const commander = normalizedDeck?.commander || null;
  const secondCommander = normalizedDeck?.secondCommander || null;

  // Build a flat list of all names across both commanders + main deck (for duplicate detection)
  const cardNames = [];
  if (commander?.name) cardNames.push(commander.name);
  if (secondCommander?.name) cardNames.push(secondCommander.name);
  normalizedDeck.cards.forEach((card) => {
    if (card?.name) cardNames.push(card.name);
  });

  // Cards whose oracle text says "a deck can have any number of cards named …" are exempt
  // from the singleton rule, as are basic lands (detected by name key).
  const BASIC_LAND_KEYS = new Set(['plains', 'island', 'swamp', 'mountain', 'forest', 'wastes',
    'snow-covered plains', 'snow-covered island', 'snow-covered swamp', 'snow-covered mountain', 'snow-covered forest']);
  // Build a set of unlimited-copy card name keys from actual card objects
  const unlimitedCopyKeys = new Set();
  [...normalizedDeck.cards, commander, secondCommander]
    .filter(Boolean)
    .forEach((c) => { if (window.CommanderDeckRules.isUnlimitedCopiesCard(c)) unlimitedCopyKeys.add(getIdentityKey(c.name)); });

  const duplicates = [];
  const seenCount = new Map();
  cardNames.forEach((name) => {
    const key = getIdentityKey(name);
    if (!key || BASIC_LAND_KEYS.has(key) || unlimitedCopyKeys.has(key)) return;
    const count = (seenCount.get(key) || 0) + 1;
    seenCount.set(key, count);
    if (count === 2) duplicates.push(name);
  });

  const bannedCards = [];
  if (commander?.isBanned) bannedCards.push(commander.name);
  if (secondCommander?.isBanned) bannedCards.push(secondCommander.name);
  normalizedDeck.cards.forEach((card) => {
    if (card.isBanned) bannedCards.push(card.name);
  });

  const commanderCount = (commander ? 1 : 0) + (secondCommander ? 1 : 0);
  const totalCards = normalizedDeck.cards.length + commanderCount;

  // Color identity — union of both commanders when a partner pair is present
  const colorViolations = [];
  if (commander) {
    const commanderIdentity = getCardEffectiveColorIdentity(commander);
    if (secondCommander) {
      getCardEffectiveColorIdentity(secondCommander).forEach((c) => commanderIdentity.add(c));
    }
    normalizedDeck.cards.forEach((card) => {
      const cardIdentity = getCardEffectiveColorIdentity(card);
      const illegal = [...cardIdentity].filter((c) => !commanderIdentity.has(c));
      if (illegal.length) {
        const illegalNames = illegal.map((c) => COLOR_LETTER_NAMES[c] || c).join('/');
        colorViolations.push(`${card.name} (${illegalNames})`);
      }
    });
  }

  // A valid deck has exactly 1 or 2 commanders (partner) and 100 total cards
  const validCommanderCount = commanderCount === 1 || commanderCount === 2;
  return {
    commanderCount,
    totalCards,
    duplicates,
    bannedCards,
    colorViolations,
    isValid: validCommanderCount && totalCards === 100 && !duplicates.length && !colorViolations.length,
  };
}

function canCurrentUserEditDeck(deck) {
  const ownerUserId = String(deck?.ownerUserId || '').trim().toLowerCase();
  if (!ownerUserId) {
    return true;
  }

  const currentUserId = getCurrentSyncUserId();
  if (!currentUserId) {
    return true;
  }

  return isCurrentSyncUserAdmin() || currentUserId === ownerUserId;
}

function getDeckReadOnlyMessage(deck) {
  const ownerLabel = normalizeIdentityLabel(deck?.owner || '') || 'the deck owner';
  return `Read-only: only ${ownerLabel} can edit this deck.`;
}

function applyDeckOwnership(nextDeck) {
  const normalizedDeck = normalizeDeckRecord(nextDeck);
  if (!normalizedDeck) {
    return null;
  }

  const currentUserId = getCurrentSyncUserId();
  if (!currentUserId || normalizedDeck.ownerUserId) {
    return normalizedDeck;
  }

  return normalizeDeckRecord({
    ...normalizedDeck,
    ownerUserId: currentUserId,
    owner: normalizedDeck.owner || getCurrentSyncDisplayName(),
  });
}

function isDeckUsedInGame(deck) {
  if (!deck?.id) return false;
  const deckOracleId = String(deck.commander?.oracleId || '').trim().toLowerCase();
  const deckNameKey = getIdentityKey(deck.commander?.name || '');
  const ownerKey = getIdentityKey(deck.owner || '');
  if (!deckNameKey) return false;

  const commanderMatchesDl = (dl) => {
    const dlOracleId = String(dl.commanderOracleId || '').trim().toLowerCase();
    const dlNameKey = getIdentityKey(dl.commander || '');
    return (deckOracleId && dlOracleId && deckOracleId === dlOracleId)
      || (deckNameKey && dlNameKey && deckNameKey === dlNameKey);
  };

  // Check deckLists (games registered with a deck URL)
  if (loadDeckLists().some((dl) => commanderMatchesDl(dl) && (!ownerKey || getIdentityKey(dl.owner || '') === ownerKey))) {
    return true;
  }

  // Check game rows (games registered without a deck URL)
  return loadGames().some((game) =>
    getGameRows(game).some((row) => {
      const rowNameKey = getIdentityKey(row.commander || '');
      const rowOwnerKey = getIdentityKey(row.player || '');
      return rowNameKey === deckNameKey && (!ownerKey || rowOwnerKey === ownerKey);
    })
  );
}

function buildDeckUsageLookup(deckLists = loadDeckLists(), games = loadGames()) {
  const anyCommanderOracleIds = new Set();
  const anyCommanderNameKeys = new Set();
  const ownerCommanderOracleIds = new Set();
  const ownerCommanderNameKeys = new Set();

  const addCommanderUsage = (name, owner, oracleId = '') => {
    const ownerKey = getIdentityKey(owner || '');
    const commanderNameKey = getIdentityKey(name || '');
    const commanderOracleId = String(oracleId || '').trim().toLowerCase();

    if (commanderOracleId) {
      anyCommanderOracleIds.add(commanderOracleId);
      if (ownerKey) {
        ownerCommanderOracleIds.add(`${ownerKey}::${commanderOracleId}`);
      }
    }

    if (commanderNameKey) {
      anyCommanderNameKeys.add(commanderNameKey);
      if (ownerKey) {
        ownerCommanderNameKeys.add(`${ownerKey}::${commanderNameKey}`);
      }
    }
  };

  deckLists.forEach((deckList) => {
    addCommanderUsage(deckList?.commander, deckList?.owner, deckList?.commanderOracleId);
  });

  games.forEach((game) => {
    getGameRows(game).forEach((row) => {
      addCommanderUsage(row?.commander, row?.player);
    });
  });

  return {
    anyCommanderOracleIds,
    anyCommanderNameKeys,
    ownerCommanderOracleIds,
    ownerCommanderNameKeys,
  };
}

function isDeckUsedInGameFromLookup(deck, usageLookup) {
  if (!deck?.id || !usageLookup) {
    return false;
  }

  const deckOracleId = String(deck.commander?.oracleId || '').trim().toLowerCase();
  const deckNameKey = getIdentityKey(deck.commander?.name || '');
  const ownerKey = getIdentityKey(deck.owner || '');

  if (deckOracleId) {
    if (ownerKey) {
      if (usageLookup.ownerCommanderOracleIds.has(`${ownerKey}::${deckOracleId}`)) {
        return true;
      }
    } else if (usageLookup.anyCommanderOracleIds.has(deckOracleId)) {
      return true;
    }
  }

  if (!deckNameKey) {
    return false;
  }

  if (ownerKey) {
    return usageLookup.ownerCommanderNameKeys.has(`${ownerKey}::${deckNameKey}`);
  }

  return usageLookup.anyCommanderNameKeys.has(deckNameKey);
}

function getDeckSummaryLabel(deck, summary = null) {
  const resolvedSummary = summary || getDeckValidationSummary(deck);
  if (!resolvedSummary.commanderCount) {
    return 'Commander missing';
  }
  if (resolvedSummary.totalCards !== 100) {
    return `${resolvedSummary.totalCards}/100 cards`;
  }
  if (resolvedSummary.duplicates.length) {
    return 'Duplicate cards found';
  }
  return 'Ready';
}

function formatDeckUpdatedAt(value) {
  const timestamp = Date.parse(String(value || '').trim());
  if (!Number.isFinite(timestamp)) {
    return '—';
  }

  return new Date(timestamp).toLocaleString();
}

function ensureActiveDeckBuilderRecord(options = {}) {
  return window.CommanderDeckBuilder?.ensureActiveDeckBuilderRecord?.(options) || null;
}

function getDeckBuilderSyncState() {
  return window.CommanderDeckBuilder?.getSyncState?.() || { activeDeckId: '', hasUnsavedChanges: false };
}

function hasUnsavedDeckBuilderChanges() {
  return window.CommanderDeckBuilder?.hasUnsavedDeckBuilderChanges?.() || false;
}

function applyDeckBuilderDraftMeta(deck) {
  return window.CommanderDeckBuilder?.applyDeckBuilderDraftMeta?.(deck) ?? deck;
}

// Returns true if the card has the generic "Partner" keyword (NOT "Partner with X").
async function fetchDeckSearchResultsList(query) {
  const normalizedQuery = String(query || '').trim().toLowerCase();
  if (!normalizedQuery) {
    return [];
  }

  if (deckBuilderSearchCache.has(normalizedQuery)) {
    return deckBuilderSearchCache.get(normalizedQuery);
  }

  const response = await fetch(`${DECK_SEARCH_ENDPOINT}?q=${encodeURIComponent(query)}`, {
    cache: 'no-store',
  });

  if (!response.ok) {
    let message = `Request failed (${response.status})`;
    try {
      const payload = await response.json();
      if (payload?.detail) {
        message = `${payload.error || message} ${payload.detail}`.trim();
      } else if (payload?.error) {
        message = payload.error;
      }
    } catch (error) {
      // Keep default message.
    }
    throw new Error(message);
  }

  const payload = await response.json();
  const results = Array.isArray(payload?.results) ? payload.results.map((value) => String(value || '').trim()).filter(Boolean) : [];
  deckBuilderSearchCache.set(normalizedQuery, results);
  return results;
}

async function fetchDeckCardByName(name) {
  const normalizedName = String(name || '').trim().toLowerCase();
  if (!normalizedName) {
    return null;
  }

  if (deckBuilderCardCache.has(normalizedName)) {
    const cached = deckBuilderCardCache.get(normalizedName);
    if (cached.manaCost || cached.cardFaces?.length) {
      return cached;
    }
    deckBuilderCardCache.delete(normalizedName);
  }

  const response = await fetch(`${DECK_CARD_ENDPOINT}&name=${encodeURIComponent(name)}`, {
    cache: 'no-store',
  });

  if (!response.ok) {
    let message = `Request failed (${response.status})`;
    try {
      const payload = await response.json();
      if (payload?.detail) {
        message = `${payload.error || message} ${payload.detail}`.trim();
      } else if (payload?.error) {
        message = payload.error;
      }
    } catch (error) {
      // Keep default message.
    }
    throw new Error(message);
  }

  const payload = await response.json();
  const card = normalizeDeckCardEntry(payload?.card);
  if (card) {
    deckBuilderCardCache.set(normalizedName, card);
  }
  return card;
}

async function fetchDeckCardsByNamesBulk(names) {
  const normalizedNames = Array.isArray(names)
    ? names.map((value) => String(value || '').trim()).filter(Boolean)
    : [];

  if (!normalizedNames.length) {
    return new Map();
  }

  const resultMap = new Map();
  const wait = (ms) => new Promise((resolve) => window.setTimeout(resolve, ms));

  for (let offset = 0; offset < normalizedNames.length; offset += DECK_CARDS_BULK_REQUEST_CHUNK_SIZE) {
    const batch = normalizedNames.slice(offset, offset + DECK_CARDS_BULK_REQUEST_CHUNK_SIZE);
    const response = await fetch(DECK_CARDS_BULK_ENDPOINT, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ names: batch }),
      cache: 'no-store',
    });

    if (!response.ok) {
      let message = `Request failed (${response.status})`;
      try {
        const payload = await response.json();
        if (payload?.detail) {
          message = `${payload.error || message} ${payload.detail}`.trim();
        } else if (payload?.error) {
          message = payload.error;
        }
      } catch (error) {
        // Keep default message.
      }
      throw new Error(message);
    }

    const payload = await response.json();
    const cards = Array.isArray(payload?.cards) ? payload.cards : [];
    cards.forEach((entry) => {
      const queryName = String(entry?.name || '').trim();
      if (!queryName) {
        return;
      }

      const normalizedCard = normalizeDeckCardEntry(entry?.card);
      if (normalizedCard) {
        deckBuilderCardCache.set(queryName.toLowerCase(), normalizedCard);
        deckBuilderCardCache.set(normalizedCard.name.toLowerCase(), normalizedCard);
      }
      resultMap.set(getIdentityKey(queryName), normalizedCard || null);
      if (normalizedCard) {
        const cardNameKey = getIdentityKey(normalizedCard.name);
        if (cardNameKey && !resultMap.has(cardNameKey)) {
          resultMap.set(cardNameKey, normalizedCard);
        }
      }
    });

    if (offset + DECK_CARDS_BULK_REQUEST_CHUNK_SIZE < normalizedNames.length) {
      await wait(DECK_CARDS_BULK_REQUEST_GAP_MS);
    }
  }

  return resultMap;
}

function getDeckCardRulesText(card) {
  if (!card || typeof card !== 'object') {
    return '';
  }

  const directRulesText = String(card.oracleText || '').trim();
  if (directRulesText) {
    return directRulesText;
  }

  const firstFace = Array.isArray(card.cardFaces) ? card.cardFaces[0] : null;
  return String(firstFace?.oracleText || '').trim();
}

function getDeckCardStatLine(card) {
  const firstFace = Array.isArray(card.cardFaces) ? card.cardFaces[0] : null;
  const power = String(card.power || firstFace?.power || '').trim();
  const toughness = String(card.toughness || firstFace?.toughness || '').trim();
  const loyalty = String(card.loyalty || firstFace?.loyalty || '').trim();
  const defense = String(card.defense || firstFace?.defense || '').trim();

  if (power && toughness) {
    return `${power} / ${toughness}`;
  }
  if (loyalty) {
    return `Loyalty: ${loyalty}`;
  }
  if (defense) {
    return `Defense: ${defense}`;
  }
  return '';
}

function renderDeckBuilderBreakdown(deck) {
  return window.CommanderDeckBuilder?.renderBreakdown?.(deck);
}

function getDeckOwnerGroups() {
  const deckLists = loadDeckLists().map(normalizeDeckListEntry).filter(Boolean);
  const deckListByDeckId = new Map();
  const deckListByOwnerCommander = new Map();

  deckLists.forEach((entry) => {
    const deckId = String(entry.deckId || '').trim();
    const ownerKey = getIdentityKey(entry.owner || '');
    const commanderKey = getIdentityKey(entry.commander || '');

    if (deckId && !deckListByDeckId.has(deckId)) {
      deckListByDeckId.set(deckId, entry);
    }

    if (ownerKey && commanderKey) {
      const compositeKey = `${ownerKey}::${commanderKey}`;
      if (!deckListByOwnerCommander.has(compositeKey)) {
        deckListByOwnerCommander.set(compositeKey, entry);
      }
    }
  });

  const selectorDecks = loadDecks()
    .map(normalizeDeckRecord)
    .filter((deck) => Boolean(deck?.commander?.name && deck?.owner))
    .filter((deck) => deck.inRotation !== false)
    .map((deck) => {
      const ownerKey = getIdentityKey(deck.owner || '');
      const commanderKey = getIdentityKey(deck.commander?.name || '');
      const linkedDeckList = deckListByDeckId.get(deck.id)
        || (ownerKey && commanderKey ? deckListByOwnerCommander.get(`${ownerKey}::${commanderKey}`) : null);

      return {
        id: deck.id,
        deckId: deck.id,
        owner: deck.owner,
        commander: deck.commander.name,
        url: String(linkedDeckList?.url || '').trim(),
      };
    });

  return selectorDecks.reduce((groups, entry) => {
    const owner = (entry.owner || '').trim();
    if (!owner) {
      return groups;
    }

    if (!groups[owner]) {
      groups[owner] = [];
    }

    groups[owner].push(entry);
    return groups;
  }, {});
}

function getCommanderBuilderInputs() {
  if (!commanderBuilderForm) {
    return [];
  }

  return Array.from(commanderBuilderForm.querySelectorAll('input[name="commander-color"]'));
}

function getCommanderBuilderMode() {
  const selectedInput = commanderBuilderModeInputs.find((input) => input.checked);
  const value = String(selectedInput?.value || '').trim().toLowerCase();
  return value === 'keywords' || value === 'both' ? value : 'identity';
}

function normalizeCommanderBuilderKeyword(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function getSelectedCommanderBuilderKeywords() {
  return commanderBuilderSelectedKeywords.slice();
}

function getSelectedCommanderBuilderIdentity() {
  const selectedValues = getCommanderBuilderInputs()
    .filter((input) => input.checked)
    .map((input) => String(input.value || '').trim().toUpperCase())
    .filter(Boolean);

  if (!selectedValues.length) {
    return '';
  }

  // Colorless is exclusive. If mixed with colors, treat as invalid selection.
  if (selectedValues.includes('C')) {
    return selectedValues.length === 1 ? 'c' : '';
  }

  const order = ['W', 'U', 'B', 'R', 'G'];
  const normalized = order.filter((code) => selectedValues.includes(code)).join('');
  return normalized.toLowerCase();
}

function getCommanderBuilderQuery() {
  const mode = getCommanderBuilderMode();
  if (mode === 'keywords') {
    const keywords = getSelectedCommanderBuilderKeywords();
    return keywords.length ? { mode, keywords } : null;
  }

  if (mode === 'both') {
    const keywords = getSelectedCommanderBuilderKeywords();
    if (!keywords.length) {
      return null;
    }

    const identity = getSelectedCommanderBuilderIdentity();
    return identity ? { mode, keywords, identity } : null;
  }

  const identity = getSelectedCommanderBuilderIdentity();
  return identity ? { mode, identity } : null;
}

function getCommanderBuilderQueryKey(query) {
  if (!query || typeof query !== 'object') {
    return '';
  }

  if (query.mode === 'keywords' || query.mode === 'both') {
    const keywordKey = (query.keywords || []).map((keyword) => normalizeCommanderBuilderKeyword(keyword).toLowerCase()).join('|');
    const identityKey = String(query.identity || '').trim().toLowerCase();
    return `${query.mode}:${keywordKey}:identity:${identityKey || 'any'}`;
  }

  return `identity:${String(query.identity || '').trim().toLowerCase()}`;
}

function getCommanderBuilderCriteriaLabel(query) {
  if (!query || typeof query !== 'object') {
    return '';
  }

  if (query.mode === 'keywords' || query.mode === 'both') {
    const keywordLabel = (query.keywords || []).join(', ');
    const identityLabel = query.identity ? getCommanderIdentityLabel(query.identity) : '';
    return identityLabel ? `${keywordLabel} within ${identityLabel}` : keywordLabel;
  }

  return getCommanderIdentityLabel(query.identity);
}

function getCommanderBuilderTagLabel(query) {
  if (!query || typeof query !== 'object') {
    return '';
  }

  return query.mode === 'keywords' || query.mode === 'both'
    ? `Keywords: ${getCommanderBuilderCriteriaLabel(query)}`
    : getCommanderBuilderCriteriaLabel(query);
}

function getCommanderBuilderIdlePlaceholder(mode = getCommanderBuilderMode()) {
  if (mode === 'keywords') {
    return 'Select at least one keyword ability to get started.';
  }

  if (mode === 'both') {
    return 'Choose a color identity and at least one keyword ability to get started.';
  }

  return 'Choose a color identity to get started.';
}

function getCommanderBuilderIdleCount(mode = getCommanderBuilderMode()) {
  if (mode === 'keywords') {
    return 'Choose keyword abilities to load the pool.';
  }

  if (mode === 'both') {
    return 'Choose colors and keyword abilities to load the pool.';
  }

  return 'Choose colors to load the pool.';
}

function getCommanderBuilderIdleStatus(mode = getCommanderBuilderMode()) {
  if (mode === 'keywords') {
    return 'Select at least one keyword ability, then roll for a commander.';
  }

  if (mode === 'both') {
    return 'Choose an exact color identity and at least one keyword ability, then roll for a commander.';
  }

  return 'Choose at least one color, or select Colorless, then roll for a commander.';
}

function getCommanderBuilderReadyPlaceholder(query) {
  if (!query || typeof query !== 'object') {
    return getCommanderBuilderIdlePlaceholder();
  }

  return query.mode === 'keywords' || query.mode === 'both'
    ? `Roll to load a commander that matches ${getCommanderBuilderCriteriaLabel(query)}.`
    : `Roll to load a random ${getCommanderBuilderCriteriaLabel(query)} commander.`;
}

function getCommanderBuilderReadyCount(query) {
  if (!query || typeof query !== 'object') {
    return getCommanderBuilderIdleCount();
  }

  return query.mode === 'keywords' || query.mode === 'both'
    ? `Ready to search ${query.keywords.length} keyword abil${query.keywords.length === 1 ? 'ity' : 'ities'}.`
    : `Ready to load ${getCommanderBuilderCriteriaLabel(query)} commanders.`;
}

function getCommanderBuilderReadyStatus(query) {
  if (!query || typeof query !== 'object') {
    return getCommanderBuilderIdleStatus();
  }

  return query.mode === 'keywords' || query.mode === 'both'
    ? `Roll to load commanders matching ${getCommanderBuilderCriteriaLabel(query)}.`
    : `Roll to load the ${getCommanderBuilderCriteriaLabel(query)} pool.`;
}

function buildCommanderBuilderKeywordCatalog(payload) {
  const seen = new Set();

  return (Array.isArray(payload?.keywordAbilities) ? payload.keywordAbilities : [])
    .map((name) => ({ name: normalizeCommanderBuilderKeyword(name), category: 'ability' }))
    .filter(({ name }) => name)
    .filter(({ name }) => {
      const key = name.toLowerCase();
      if (seen.has(key)) {
        return false;
      }

      seen.add(key);
      return true;
    })
    .sort((left, right) => compareTextValues(left.name, right.name));
}

function getCommanderBuilderKeywordCategoryLabel(category) {
  if (category === 'action') {
    return 'Action';
  }

  if (category === 'word') {
    return 'Ability Word';
  }

  return 'Ability';
}

function renderCommanderBuilderKeywordSelection() {
  if (!commanderBuilderKeywordSelection) {
    return;
  }

  if (!commanderBuilderSelectedKeywords.length) {
    commanderBuilderKeywordSelection.innerHTML = '<p class="commander-builder-keyword-empty status-muted">No keyword abilities selected yet.</p>';
    return;
  }

  commanderBuilderKeywordSelection.innerHTML = commanderBuilderSelectedKeywords.map((keyword) => `
    <button type="button" class="commander-builder-selected-keyword" data-remove-commander-keyword="${escapeHtml(keyword)}" aria-label="Remove ${escapeHtml(keyword)}">
      <span>${escapeHtml(keyword)}</span>
      <span aria-hidden="true">×</span>
    </button>`).join('');
}

function renderCommanderBuilderKeywordCatalog() {
  if (!commanderBuilderKeywordGrid || !commanderBuilderKeywordStatus) {
    return;
  }

  const normalizedSearch = commanderBuilderKeywordSearchTerm.toLowerCase();
  const filteredKeywords = commanderBuilderKeywordCatalog.filter(({ name }) => name.toLowerCase().includes(normalizedSearch));

  if (commanderBuilderKeywordsLoading && !commanderBuilderKeywordCatalog.length) {
    commanderBuilderKeywordStatus.textContent = 'Loading keyword ability catalog...';
    commanderBuilderKeywordStatus.className = 'commander-builder-keyword-status status-muted';
    commanderBuilderKeywordGrid.innerHTML = '<p class="status-muted">Loading keyword ability catalog...</p>';
    renderCommanderBuilderKeywordSelection();
    return;
  }

  if (commanderBuilderKeywordErrorMessage && !commanderBuilderKeywordCatalog.length) {
    commanderBuilderKeywordStatus.textContent = commanderBuilderKeywordErrorMessage;
    commanderBuilderKeywordStatus.className = 'commander-builder-keyword-status status-error';
    commanderBuilderKeywordGrid.innerHTML = '<p class="status-error">Keyword ability search is unavailable right now.</p>';
    renderCommanderBuilderKeywordSelection();
    return;
  }

  if (!filteredKeywords.length) {
    commanderBuilderKeywordStatus.textContent = commanderBuilderKeywordCatalog.length
      ? 'No keyword abilities match your search.'
      : 'No keyword abilities are available right now.';
    commanderBuilderKeywordStatus.className = `commander-builder-keyword-status ${commanderBuilderKeywordCatalog.length ? 'status-muted' : 'status-error'}`;
    commanderBuilderKeywordGrid.innerHTML = '<p class="status-muted">Try a different keyword ability search.</p>';
    renderCommanderBuilderKeywordSelection();
    return;
  }

  commanderBuilderKeywordStatus.textContent = `Showing ${filteredKeywords.length} of ${commanderBuilderKeywordCatalog.length} keyword abilities.`;
  commanderBuilderKeywordStatus.className = 'commander-builder-keyword-status status-muted';
  commanderBuilderKeywordGrid.innerHTML = filteredKeywords.map(({ name, category }) => {
    const isSelected = commanderBuilderSelectedKeywords.includes(name);
    return `
      <label class="commander-builder-keyword-option${isSelected ? ' is-selected' : ''}">
        <input type="checkbox" name="commander-builder-keyword" value="${escapeHtml(name)}"${isSelected ? ' checked' : ''} />
        <span class="commander-builder-keyword-name">${escapeHtml(name)}</span>
        <span class="commander-builder-keyword-badge commander-builder-keyword-badge-${escapeHtml(category)}">${escapeHtml(getCommanderBuilderKeywordCategoryLabel(category))}</span>
      </label>`;
  }).join('');
  renderCommanderBuilderKeywordSelection();
}

async function loadCommanderBuilderKeywordCatalog() {
  if (!commanderBuilderKeywordGrid || commanderBuilderKeywordsLoading || commanderBuilderKeywordsLoaded) {
    return;
  }

  commanderBuilderKeywordsLoading = true;
  commanderBuilderKeywordErrorMessage = '';
  renderCommanderBuilderKeywordCatalog();

  try {
    const requestUrl = new URL(KEYWORDS_ENDPOINT, window.location.origin);
    requestUrl.searchParams.set('_', String(Date.now()));
    const response = await fetch(requestUrl.toString(), {
      cache: 'no-store',
    });

    if (!response.ok) {
      throw new Error(`Unable to load keywords (${response.status}).`);
    }

    const payload = await response.json();
    commanderBuilderKeywordCatalog = buildCommanderBuilderKeywordCatalog(payload);
    commanderBuilderKeywordsLoaded = true;
  } catch (error) {
    commanderBuilderKeywordErrorMessage = error instanceof Error ? error.message : 'Unable to load keyword catalog right now.';
  } finally {
    commanderBuilderKeywordsLoading = false;
    renderCommanderBuilderKeywordCatalog();
    updateCommanderBuilderControls();
  }
}

function toggleCommanderBuilderKeyword(keyword, isSelected) {
  const normalizedKeyword = normalizeCommanderBuilderKeyword(keyword);
  if (!normalizedKeyword) {
    return;
  }

  if (isSelected) {
    if (!commanderBuilderSelectedKeywords.includes(normalizedKeyword)) {
      commanderBuilderSelectedKeywords = [...commanderBuilderSelectedKeywords, normalizedKeyword];
    }
  } else {
    commanderBuilderSelectedKeywords = commanderBuilderSelectedKeywords.filter((value) => value !== normalizedKeyword);
  }

  renderCommanderBuilderKeywordCatalog();
}

function updateCommanderBuilderModeUi() {
  commanderBuilderMode = getCommanderBuilderMode();
  const showIdentityPanel = commanderBuilderMode === 'identity' || commanderBuilderMode === 'both';
  const showKeywordPanel = commanderBuilderMode === 'keywords' || commanderBuilderMode === 'both';

  if (commanderBuilderIdentityPanel) {
    commanderBuilderIdentityPanel.hidden = !showIdentityPanel;
  }

  if (commanderBuilderIdentityTitle) {
    commanderBuilderIdentityTitle.textContent = commanderBuilderMode === 'both'
      ? 'Pick an Exact Color Identity'
      : 'Pick an Exact Color Identity';
  }

  if (commanderBuilderIdentityCopy) {
    commanderBuilderIdentityCopy.textContent = commanderBuilderMode === 'both'
      ? 'This stays exact. Choosing blue and black only returns Dimir commanders that also match your selected keyword abilities.'
      : 'The selection is exact. Choosing blue and black only returns Dimir commanders, not mono-blue or Grixis options.';
  }

  if (commanderBuilderKeywordPanel) {
    commanderBuilderKeywordPanel.hidden = !showKeywordPanel;
  }

  if (showKeywordPanel) {
    void loadCommanderBuilderKeywordCatalog();
  }
}

function syncCommanderBuilderPreviewState() {
  const mode = getCommanderBuilderMode();
  const query = getCommanderBuilderQuery();
  const queryKey = getCommanderBuilderQueryKey(query);

  if (!query) {
    renderCommanderBuilderPlaceholder(getCommanderBuilderIdlePlaceholder(mode));
    setCommanderBuilderCount(getCommanderBuilderIdleCount(mode));
    setCommanderBuilderStatus(getCommanderBuilderIdleStatus(mode), 'muted');
    updateCommanderBuilderControls();
    return;
  }

  if (queryKey !== commanderBuilderQueryKey) {
    renderCommanderBuilderPlaceholder(getCommanderBuilderReadyPlaceholder(query));
    setCommanderBuilderCount(getCommanderBuilderReadyCount(query));
    setCommanderBuilderStatus(getCommanderBuilderReadyStatus(query), 'neutral');
  }

  updateCommanderBuilderControls();
}

// Renders a list of deck card rows, grouping unlimited-copy cards into stacked rows with +/- counters.
function getCommanderIdentityLabel(identity) {
  if (identity === 'c') {
    return 'Colorless';
  }

  const codes = new Set(String(identity || '').toUpperCase().split(''));
  return COMMANDER_BUILDER_COLOR_OPTIONS
    .filter(({ code }) => codes.has(code))
    .map(({ label }) => label)
    .join(' / ');
}

function setCommanderBuilderStatus(message, tone = 'muted') {
  if (!commanderBuilderStatus) {
    return;
  }

  commanderBuilderStatus.textContent = message;
  commanderBuilderStatus.classList.remove('status-neutral', 'status-success', 'status-error', 'status-muted');
  commanderBuilderStatus.classList.add(`status-${tone}`);
}

function setCommanderBuilderCount(message) {
  if (!commanderBuilderCount) {
    return;
  }

  commanderBuilderCount.textContent = message;
}

function loadCommanderBuilderCache() {
  const raw = readLocalStorageValue(COMMANDER_BUILDER_CACHE_STORAGE_KEY);
  const cache = parseJsonSafe(raw || '{}', {});
  return cache && typeof cache === 'object' ? cache : {};
}

function saveCommanderBuilderCache(identity, cards) {
  if (!identity) {
    return;
  }

  const cache = loadCommanderBuilderCache();
  cache[identity] = {
    fetchedAt: Date.now(),
    totalCards: Number.isFinite(Number(cards?.totalCards)) ? Number(cards.totalCards) : 0,
  };
  writeLocalStorageValue(COMMANDER_BUILDER_CACHE_STORAGE_KEY, JSON.stringify(cache));
}

function getCommanderBuilderCachedCards(identity, { allowExpired = false } = {}) {
  if (!identity) {
    return null;
  }

  const entry = loadCommanderBuilderCache()[identity];
  if (!entry || !Number.isFinite(Number(entry.totalCards))) {
    return null;
  }

  const fetchedAt = Number(entry.fetchedAt || 0);
  const isFresh = Number.isFinite(fetchedAt) && (Date.now() - fetchedAt) < COMMANDER_BUILDER_CACHE_TTL_MS;
  if (!allowExpired && !isFresh) {
    return null;
  }

  return {
    totalCards: Number(entry.totalCards),
  };
}

async function fetchCommanderBuilderSelection(query) {
  const queryKey = getCommanderBuilderQueryKey(query);
  const cachedSummary = getCommanderBuilderCachedCards(queryKey);
  const staleSummary = getCommanderBuilderCachedCards(queryKey, { allowExpired: true });
  let response;
  try {
    const requestUrl = new URL(COMMANDER_BUILDER_ENDPOINT, window.location.origin);
    requestUrl.searchParams.set('mode', query?.mode === 'keywords' || query?.mode === 'both' ? 'keywords' : 'identity');
    if (query?.mode === 'keywords' || query?.mode === 'both') {
      (query.keywords || []).forEach((keyword) => {
        requestUrl.searchParams.append('keyword', keyword);
      });
      if (query.identity) {
        requestUrl.searchParams.set('identity', query.identity);
      }
    } else {
      requestUrl.searchParams.set('identity', query?.identity || '');
    }
    requestUrl.searchParams.set('_', String(Date.now()));
    response = await fetch(requestUrl.toString(), {
      cache: 'no-store',
    });
  } catch (error) {
    if (staleSummary) {
      throw new Error('Unable to refresh the commander pick right now. Try again in a moment.');
    }
    throw error;
  }

  if (!response.ok) {
    let message = `Request failed (${response.status})`;
    try {
      const errorBody = await response.json();
      if (errorBody?.detail) {
        message = `${errorBody.error || 'Request failed'} ${errorBody.detail}`.trim();
      } else if (errorBody?.error) {
        message = errorBody.error;
      }
    } catch (error) {
      // Keep default message.
    }

    throw new Error(message);
  }

  const payload = await response.json();
  const selectedCard = payload?.card && payload.card.name && payload.card.scryfallUri ? payload.card : null;
  const totalCards = Number.isFinite(Number(payload?.totalCards))
    ? Number(payload.totalCards)
    : Number(cachedSummary?.totalCards || staleSummary?.totalCards || 0);

  saveCommanderBuilderCache(queryKey, { totalCards });

  return {
    card: selectedCard,
    totalCards,
    source: 'network',
    identity: String(payload?.identity || query?.identity || '').trim().toLowerCase(),
    keywords: Array.isArray(payload?.keywords) ? payload.keywords : (query?.keywords || []),
    mode: query?.mode || (payload?.mode === 'keywords' ? 'keywords' : 'identity'),
  };
}

function updateCommanderBuilderControls() {
  const query = getCommanderBuilderQuery();
  const queryKey = getCommanderBuilderQueryKey(query);
  const submitButton = commanderBuilderForm?.querySelector('button[type="submit"]');

  if (submitButton) {
    submitButton.disabled = commanderBuilderLoading || !query;
  }

  if (commanderBuilderRerollButton) {
    commanderBuilderRerollButton.disabled = commanderBuilderLoading || !query || commanderBuilderQueryKey !== queryKey;
  }
}

function renderCommanderBuilderPlaceholder(message) {
  if (!commanderBuilderResult) {
    return;
  }

  commanderBuilderResult.innerHTML = `<p>${escapeHtml(message)}</p>`;
}

function formatCommanderBuilderRichText(text) {
  return escapeHtml(String(text || '').trim()).replace(/\n/g, '<br />');
}

function getCommanderBuilderStatsLine(card) {
  if (!card || typeof card !== 'object') {
    return '';
  }

  const parts = [];
  const power = String(card.power || '').trim();
  const toughness = String(card.toughness || '').trim();
  const loyalty = String(card.loyalty || '').trim();
  const defense = String(card.defense || '').trim();

  if (power || toughness) {
    parts.push(`${escapeHtml(power || '?')}/${escapeHtml(toughness || '?')}`);
  }

  if (loyalty) {
    parts.push(`Loyalty ${escapeHtml(loyalty)}`);
  }

  if (defense) {
    parts.push(`Defense ${escapeHtml(defense)}`);
  }

  return parts.join(' · ');
}

function getCommanderBuilderImageSources(card) {
  if (!card || typeof card !== 'object') {
    return [];
  }

  const faceCards = Array.isArray(card.cardFaces) ? card.cardFaces : [];
  const sources = [
    card.imageUri,
    card.imageLargeUri,
    card.imagePngUri,
    ...faceCards.flatMap((face) => [face.imageUri, face.imageLargeUri, face.imagePngUri]),
  ];

  return [...new Set(sources.map((value) => String(value || '').trim()).filter(Boolean))];
}

function attachCommanderBuilderImageFallback(card) {
  if (!commanderBuilderResult || !card) {
    return;
  }

  const image = commanderBuilderResult.querySelector('.commander-builder-image');
  if (!(image instanceof HTMLImageElement)) {
    return;
  }

  const sources = getCommanderBuilderImageSources(card);
  if (!sources.length) {
    image.closest('.commander-builder-media')?.remove();
    return;
  }

  let sourceIndex = sources.findIndex((source) => source === image.src);
  if (sourceIndex < 0) {
    sourceIndex = 0;
    image.src = sources[0];
  }

  image.addEventListener('error', () => {
    sourceIndex += 1;
    if (sourceIndex < sources.length) {
      image.src = sources[sourceIndex];
      return;
    }

    image.closest('.commander-builder-media')?.remove();
  }, { once: false });
}

function renderCommanderBuilderResultCard(card, criteriaLabel, totalCards, source) {
  if (!commanderBuilderResult || !card) {
    return;
  }

  const safeName = escapeHtml(card.name);
  const safeTypeLine = escapeHtml(card.typeLine || 'Commander-eligible card');
  const safeManaCost = escapeHtml(card.manaCost || 'No mana cost');
  const safeCriteriaLabel = escapeHtml(criteriaLabel || 'Commander');
  const safePoolSize = escapeHtml(String(totalCards));
  const imageSources = getCommanderBuilderImageSources(card);
  const safeImageUri = escapeHtml(imageSources[0] || '');
  const faceCards = Array.isArray(card.cardFaces) ? card.cardFaces : [];
  const useFaceSections = faceCards.length > 1;
  const fallbackFace = faceCards[0] || null;
  const rulesText = String(card.oracleText || fallbackFace?.oracleText || '').trim();
  const statsLine = getCommanderBuilderStatsLine(card) || getCommanderBuilderStatsLine(fallbackFace);
  const sourceLabel = source === 'network'
    ? 'Live Scryfall pool'
    : source === 'cache'
      ? 'Cached pool from this device'
      : 'Cached pool from this device due to a fetch issue';
  const imageMarkup = safeImageUri
    ? `<div class="commander-builder-media"><img class="commander-builder-image" src="${safeImageUri}" alt="${safeName}" loading="lazy" /></div>`
    : '';
  const rulesMarkup = useFaceSections
    ? faceCards.map((face) => {
      const faceName = escapeHtml(face.name || card.name);
      const faceTypeLine = escapeHtml(face.typeLine || '');
      const faceManaCost = escapeHtml(face.manaCost || 'No mana cost');
      const faceStats = getCommanderBuilderStatsLine(face);
      const faceRulesText = String(face.oracleText || '').trim();

      return `
        <section class="commander-builder-rules-block commander-builder-face-block">
          <div class="commander-builder-rules-header">
            <h4>${faceName}</h4>
            <p class="commander-builder-meta">${faceTypeLine || 'Commander face'}</p>
          </div>
          <p class="commander-builder-meta">Mana cost: ${faceManaCost}</p>
          ${faceStats ? `<p class="commander-builder-stats">Stats: ${faceStats}</p>` : ''}
          <div class="commander-builder-rules-text">${formatCommanderBuilderRichText(faceRulesText || 'No rules text available.')}</div>
        </section>`;
    }).join('')
    : `
      <section class="commander-builder-rules-block">
        <div class="commander-builder-rules-header">
          <h4>Card text</h4>
        </div>
        ${statsLine ? `<p class="commander-builder-stats">Stats: ${statsLine}</p>` : ''}
        <div class="commander-builder-rules-text">${formatCommanderBuilderRichText(rulesText || 'No rules text available.')}</div>
      </section>`;

  commanderBuilderResult.innerHTML = `
    <article class="commander-builder-result-card">
      ${imageMarkup}
      <div class="commander-builder-details">
        <p class="commander-builder-tag">${safeCriteriaLabel}</p>
        <h3>${buildCommanderDisplayHtml(card.name, safeName, card)}</h3>
        <p class="commander-builder-meta">${safeTypeLine}</p>
        <p class="commander-builder-meta">Mana cost: ${safeManaCost}</p>
        <p class="commander-builder-pool">Pulled from ${safePoolSize} eligible commanders.</p>
        ${rulesMarkup}
        <p class="commander-builder-source">${escapeHtml(sourceLabel)}</p>
        <div class="actions">
          <a href="${escapeHtml(card.scryfallUri)}" target="_blank" rel="noopener noreferrer">View on Scryfall</a>
          <a href="${escapeHtml(`deckbuilder.html?new=1&commander=${encodeURIComponent(card.name)}`)}">Build This Deck</a>
        </div>
      </div>
    </article>`;

  attachCommanderBuilderImageFallback(card);
}

function syncCommanderBuilderExclusiveSelection(changedInput) {
  const allInputs = getCommanderBuilderInputs();
  if (!changedInput || !allInputs.length) {
    return;
  }

  if (changedInput.value === 'C' && changedInput.checked) {
    allInputs.forEach((input) => {
      if (input !== changedInput) {
        input.checked = false;
      }
    });
    return;
  }

  if (changedInput.checked) {
    const colorlessInput = allInputs.find((input) => input.value === 'C');
    if (colorlessInput) {
      colorlessInput.checked = false;
    }
  }
}

async function runCommanderBuilderRoll() {
  const mode = getCommanderBuilderMode();
  const query = getCommanderBuilderQuery();
  if (!query) {
    renderCommanderBuilderPlaceholder(getCommanderBuilderIdlePlaceholder(mode));
    setCommanderBuilderCount(getCommanderBuilderIdleCount(mode));
    setCommanderBuilderStatus(getCommanderBuilderIdleStatus(mode), 'error');
    updateCommanderBuilderControls();
    return;
  }

  const queryKey = getCommanderBuilderQueryKey(query);
  const criteriaLabel = getCommanderBuilderCriteriaLabel(query);

  const requestId = commanderBuilderRequestId + 1;
  commanderBuilderRequestId = requestId;
  commanderBuilderLoading = true;
  updateCommanderBuilderControls();
  renderCommanderBuilderPlaceholder('Loading commanders...');
  setCommanderBuilderStatus(
    query.mode === 'keywords' || query.mode === 'both'
      ? `Loading commanders matching ${criteriaLabel} from Scryfall...`
      : `Loading ${criteriaLabel} commanders from Scryfall...`,
    'neutral',
  );

  try {
    const payload = await fetchCommanderBuilderSelection(query);
    if (requestId !== commanderBuilderRequestId) {
      return;
    }

    commanderBuilderQueryKey = queryKey;
    setCommanderBuilderCount(`${payload.totalCards} eligible commanders`);

    if (!payload.card || !payload.totalCards) {
      commanderBuilderLastCardName = '';
      renderCommanderBuilderPlaceholder(
        query.mode === 'keywords' || query.mode === 'both'
          ? `No commander-eligible cards were found matching ${criteriaLabel}.`
          : `No commander-eligible cards were found for ${criteriaLabel}.`,
      );
      setCommanderBuilderStatus(
        query.mode === 'keywords' || query.mode === 'both'
          ? 'No commanders found for that keyword combination.'
          : 'No commanders found for that exact color identity.',
        'error',
      );
      return;
    }

    const selectedCard = payload.card;
    commanderBuilderLastCardName = selectedCard?.name || '';
    renderCommanderBuilderResultCard(selectedCard, getCommanderBuilderTagLabel(query), payload.totalCards, payload.source);
    setCommanderBuilderStatus(`${commanderBuilderLastCardName} selected for ${criteriaLabel}.`, 'success');
  } catch (error) {
    if (requestId !== commanderBuilderRequestId) {
      return;
    }

    commanderBuilderQueryKey = '';
    commanderBuilderLastCardName = '';
    renderCommanderBuilderPlaceholder('Unable to load commanders right now. Try again in a moment.');
    setCommanderBuilderCount('Commander pool unavailable right now.');
    setCommanderBuilderStatus(error instanceof Error ? error.message : 'Unable to load commanders right now.', 'error');
  } finally {
    if (requestId === commanderBuilderRequestId) {
      commanderBuilderLoading = false;
      updateCommanderBuilderControls();
    }
  }
}

function rerollCommanderBuilderCard() {
  const query = getCommanderBuilderQuery();
  const queryKey = getCommanderBuilderQueryKey(query);
  if (!query || commanderBuilderQueryKey !== queryKey) {
    setCommanderBuilderStatus('Load a commander pool for the current filter before rerolling.', 'error');
    return;
  }

  runCommanderBuilderRoll();
}

function renderCommanderBuilder() {
  if (!commanderBuilderForm || !commanderBuilderResult) {
    return;
  }

  updateCommanderBuilderModeUi();
  renderCommanderBuilderKeywordCatalog();

  if (!commanderBuilderResult.dataset.initialized) {
    renderCommanderBuilderPlaceholder(getCommanderBuilderIdlePlaceholder());
    commanderBuilderResult.dataset.initialized = 'true';
  }

  if (!commanderBuilderStatus?.dataset.initialized) {
    setCommanderBuilderStatus(getCommanderBuilderIdleStatus(), 'muted');
    commanderBuilderStatus.dataset.initialized = 'true';
  }

  if (!commanderBuilderCount?.dataset.initialized) {
    setCommanderBuilderCount(getCommanderBuilderIdleCount());
    commanderBuilderCount.dataset.initialized = 'true';
  }

  if (commanderBuilderMode === 'keywords' || commanderBuilderMode === 'both') {
    void loadCommanderBuilderKeywordCatalog();
  }

  syncCommanderBuilderPreviewState();
}

function getGameWinner(game) {
  if (Array.isArray(game.finishOrder) && game.finishOrder.length) {
    return game.finishOrder[0];
  }

  const rows = getGameRows(game);
  const placeOne = rows.find((row) => row.place === 1);
  return placeOne ? placeOne.player : (rows[0] ? rows[0].player : '');
}

function getGameTotalKills(game) {
  return getGameRows(game).reduce((sum, row) => {
    const kills = typeof row.kills === 'number' && !Number.isNaN(row.kills) ? row.kills : 0;
    return sum + kills;
  }, 0);
}

function getGamePlayerCount(game) {
  return getGameRows(game).length;
}

function getGameRows(game) {
  const normalizedGame = normalizeSavedGameEntry(game, loadDecks());

  if (Array.isArray(normalizedGame.playerRows) && normalizedGame.playerRows.length) {
    return normalizedGame.playerRows;
  }

  if (Array.isArray(normalizedGame.playerCommanders) && normalizedGame.playerCommanders.length) {
    return normalizedGame.playerCommanders.map((entry) => ({
      player: entry.player || '',
      commander: entry.commander || '',
      place: typeof entry.place !== 'undefined'
        ? entry.place
        : (Array.isArray(normalizedGame.finishOrder) ? normalizedGame.finishOrder.indexOf(entry.player) + 1 : null),
      kills: typeof entry.kills === 'number' ? entry.kills : 0,
      killed: entry.killed ?? [],
    }));
  }

  if (Array.isArray(normalizedGame.players) && normalizedGame.players.length) {
    return normalizedGame.players.map((player) => ({
      player,
      commander: '',
      place: Array.isArray(normalizedGame.finishOrder) ? normalizedGame.finishOrder.indexOf(player) + 1 : null,
      kills: 0,
      killed: [],
    }));
  }

  return [];
}

function getGameFirstBloodInfo(game) {
  const firstBlood = game?.liveSummary?.firstBlood;
  if (!firstBlood || typeof firstBlood !== 'object') {
    return null;
  }

  const actorPlayer = String(firstBlood.actorPlayer || '').trim();
  const actorCommander = String(firstBlood.actorCommander || '').trim();
  const targetPlayer = String(firstBlood.targetPlayer || '').trim();
  const turnNumber = typeof firstBlood.turnNumber === 'number' ? firstBlood.turnNumber : null;

  if (!actorPlayer) {
    return null;
  }

  return {
    actorPlayer,
    actorCommander,
    targetPlayer,
    turnNumber,
  };
}

function getGameFirstBloodLabel(game) {
  const firstBlood = getGameFirstBloodInfo(game);
  if (!firstBlood) {
    return 'None';
  }

  const commanderText = firstBlood.actorCommander ? ` (${firstBlood.actorCommander})` : '';
  const turnText = firstBlood.turnNumber ? ` on turn ${firstBlood.turnNumber}` : '';
  if (!firstBlood.targetPlayer) {
    return `${firstBlood.actorPlayer}${commanderText}${turnText}`;
  }
  return `${firstBlood.actorPlayer}${commanderText} on ${firstBlood.targetPlayer}${turnText}`;
}

function getGamesSortedByDateAscending(games) {
  const cacheBucket = getDerivedCacheBucket(games);
  if (cacheBucket.sortedByDateAscending) {
    return cacheBucket.sortedByDateAscending;
  }

  const sortedGames = [...games]
    .map((game, index) => ({ game, index }))
    .sort((a, b) => {
      const dateResult = (a.game.date || '').localeCompare(b.game.date || '');
      if (dateResult) {
        return dateResult;
      }
      return a.index - b.index;
    })
    .map(({ game }) => game);

  cacheBucket.sortedByDateAscending = sortedGames;
  return sortedGames;
}

function getGameRowPlace(row, game) {
  if (typeof row?.place === 'number' && row.place > 0) {
    return row.place;
  }

  if (Array.isArray(game?.finishOrder) && row?.player) {
    const finishIndex = game.finishOrder.indexOf(row.player);
    if (finishIndex >= 0) {
      return finishIndex + 1;
    }
  }

  return null;
}

function getPlacementPoints(place) {
  if (place === 1) {
    return 5;
  }
  if (place === 2) {
    return 3;
  }
  if (place === 3) {
    return 2;
  }
  if (place === 4) {
    return 1;
  }
  return 0;
}

function getPairwiseActualScore(placeA, placeB) {
  if (!placeA || !placeB) {
    return 0.5;
  }
  if (placeA < placeB) {
    return 1;
  }
  if (placeA > placeB) {
    return 0;
  }
  return 0.5;
}

function getExpectedEloScore(ratingA, ratingB) {
  return 1 / (1 + (10 ** ((ratingB - ratingA) / 400)));
}

function getNormalizedTurnWinScore(turnNumber, slowTurnThreshold = 18) {
  const normalizedTurn = Number(turnNumber);
  if (!Number.isFinite(normalizedTurn) || normalizedTurn <= 0) {
    return 0;
  }

  const cappedThreshold = Math.max(2, Number(slowTurnThreshold) || 18);
  const clampedTurn = clamp(normalizedTurn, 1, cappedThreshold);
  const linearScore = (cappedThreshold - clampedTurn) / (cappedThreshold - 1);
  return linearScore ** 1.2;
}

function getNeutralTurnWinScore() {
  return getNormalizedTurnWinScore(12);
}

function getSingleGameWinTurnEloBonus(turnNumber) {
  const baseline = getNeutralTurnWinScore();
  const speedScore = getNormalizedTurnWinScore(turnNumber);
  return (speedScore - baseline) * 14;
}

function getAverageWinTurnEloBonus(averageTurnToWin, winCount) {
  const normalizedWins = clamp((Number(winCount) || 0) / 6, 0, 1);
  if (!normalizedWins) {
    return 0;
  }

  const baseline = getNeutralTurnWinScore();
  const speedScore = getNormalizedTurnWinScore(averageTurnToWin);
  return (speedScore - baseline) * 30 * normalizedWins;
}

function getRecentWinTurnEloBonus(winTurns, recentLimit = 3) {
  if (!Array.isArray(winTurns) || !winTurns.length) {
    return 0;
  }

  const recentTurns = winTurns.slice(-Math.max(1, recentLimit)).reverse();
  const recentWeights = [0.5, 0.3, 0.2];
  const weightedSpeed = recentTurns.reduce((sum, turn, index) => {
    const weight = recentWeights[index] || 0.1;
    return sum + (getNormalizedTurnWinScore(turn) * weight);
  }, 0);
  const weightTotal = recentTurns.reduce((sum, _, index) => sum + (recentWeights[index] || 0.1), 0);
  if (!weightTotal) {
    return 0;
  }

  const baseline = getNeutralTurnWinScore();
  const weightedAverageSpeed = weightedSpeed / weightTotal;
  return (weightedAverageSpeed - baseline) * 22;
}

function getRowKills(row) {
  const killedList = getCleanKilledList(row?.killed);
  if (typeof row?.kills === 'number' && !Number.isNaN(row.kills)) {
    return row.kills;
  }
  return killedList.length;
}

function formatAveragePlace(value) {
  return Number.isFinite(value) ? value.toFixed(2) : '—';
}

function formatPointsPerGame(value) {
  return Number.isFinite(value) ? value.toFixed(2) : '0.00';
}

function formatRating(value) {
  return Number.isFinite(value) ? String(Math.round(value)) : '1500';
}

function getSampleSizeLabel(count, highConfidenceThreshold = 10, mediumConfidenceThreshold = 4) {
  if (count >= highConfidenceThreshold) {
    return 'Robust sample';
  }
  if (count >= mediumConfidenceThreshold) {
    return 'Moderate sample';
  }
  return 'Small sample';
}

function getRecentWindowSummary(games, limit = 10) {
  const recentGames = getGamesSortedByDateAscending(games).slice(-limit);
  if (!recentGames.length) {
    return 'Recent views are based on game dates once enough history exists.';
  }

  return `Recent views use the ${Math.min(limit, recentGames.length)} most recent game dates, from ${recentGames[0].date || 'unknown'} to ${recentGames[recentGames.length - 1].date || 'unknown'}. Backfilled older dates can move games in or out of this window.`;
}

function compareAggregateEntries(a, b) {
  if (b.points !== a.points) {
    return b.points - a.points;
  }
  if (b.wins !== a.wins) {
    return b.wins - a.wins;
  }
  if (b.pointsPerGame !== a.pointsPerGame) {
    return b.pointsPerGame - a.pointsPerGame;
  }
  if (b.kills !== a.kills) {
    return b.kills - a.kills;
  }
  if ((a.avgPlace || 999) !== (b.avgPlace || 999)) {
    return (a.avgPlace || 999) - (b.avgPlace || 999);
  }
  if (b.games !== a.games) {
    return b.games - a.games;
  }
  return a.name.localeCompare(b.name);
}

function compareRankingsEntries(a, b) {
  if (b.rating !== a.rating) {
    return b.rating - a.rating;
  }
  if (b.pointsPerGame !== a.pointsPerGame) {
    return b.pointsPerGame - a.pointsPerGame;
  }
  if (b.winRate !== a.winRate) {
    return b.winRate - a.winRate;
  }
  if ((a.avgPlace || 999) !== (b.avgPlace || 999)) {
    return (a.avgPlace || 999) - (b.avgPlace || 999);
  }
  if (b.games !== a.games) {
    return b.games - a.games;
  }
  return a.name.localeCompare(b.name);
}

function getPlayerGameWindowLookup(games, gameLimit = 10) {
  const appearancesByPlayer = new Map();

  getGamesSortedByDateAscending(games).forEach((game) => {
    getGameRows(game).forEach((row) => {
      const player = String(row.player || '').trim();
      if (!player) {
        return;
      }

      if (!appearancesByPlayer.has(player)) {
        appearancesByPlayer.set(player, []);
      }

      appearancesByPlayer.get(player).push(game.id);
    });
  });

  return new Map(
    Array.from(appearancesByPlayer.entries()).map(([player, gameIds]) => [
      player,
      new Set(gameIds.slice(-gameLimit)),
    ]),
  );
}

function buildPlayerRankingEntries(games) {
  const cacheBucket = getDerivedCacheBucket(games);
  if (cacheBucket.playerRankingEntries) {
    return cacheBucket.playerRankingEntries;
  }

  const rawCommanders = [];
  getGamesSortedByDateAscending(games).forEach((game) => {
    getGameRows(game).forEach((row) => {
      const commander = String(row.commander || '').trim();
      if (commander) {
        rawCommanders.push(commander);
      }
    });
  });

  const commanderMap = buildCanonicalIdentityMapFromValues([
    ...getKnownCommanderOptions(),
    ...rawCommanders,
  ]);

  const stats = {};
  const eloKFactor = 28;
  const eloKillBonus = 2;
  const streakBonusStep = 3;
  const maxStreakBonus = 18;
  const playerGameWindowLookup = getPlayerGameWindowLookup(games, 10);

  getGamesSortedByDateAscending(games).forEach((game) => {
    const rows = getGameRows(game);
    const firstBlood = getGameFirstBloodInfo(game);
    const winningTurn = parseOptionalPositiveInteger(game?.liveSummary?.turnNumber);
    const participants = [];

    rows.forEach((row) => {
      const player = String(row.player || '').trim();
      if (!player) {
        return;
      }

      if (!stats[player]) {
        stats[player] = {
          name: player,
          games: 0,
          wins: 0,
          kills: 0,
          firstBloods: 0,
          points: 0,
          rating: 1500,
          currentWinStreak: 0,
          placementTotal: 0,
          placementGames: 0,
          winTurnTotal: 0,
          winTurnCount: 0,
          winTurns: [],
          commanders: {},
        };
      }

      const entry = stats[player];
      const place = getGameRowPlace(row, game);
      const kills = getRowKills(row);
      const commander = canonicalizeIdentityValue(String(row.commander || '').trim(), commanderMap);
      const firstBloodBonus = firstBlood?.actorPlayer === player ? 1 : 0;
      const isInPlayerWindow = playerGameWindowLookup.get(player)?.has(game.id) ?? true;

      if (isInPlayerWindow) {
        entry.games += 1;
        entry.kills += kills;
        entry.firstBloods += firstBloodBonus;
        entry.points += getPlacementPoints(place) + kills + firstBloodBonus;
        if (place === 1) {
          entry.wins += 1;
          if (winningTurn) {
            entry.winTurnTotal += winningTurn;
            entry.winTurnCount += 1;
            entry.winTurns.push(winningTurn);
          }
        }
        if (place) {
          entry.placementTotal += place;
          entry.placementGames += 1;
        }
        if (commander) {
          entry.commanders[commander] = (entry.commanders[commander] || 0) + 1;
        }
      }

      participants.push({
        player,
        place,
        kills,
        isInPlayerWindow,
      });
    });

    if (participants.length < 2) {
      return;
    }

    const preGameRatings = Object.fromEntries(
      participants.map(({ player }) => [player, stats[player].rating]),
    );

    participants.forEach(({ player, place, kills, isInPlayerWindow }) => {
      if (!isInPlayerWindow) {
        return;
      }

      let ratingDelta = 0;

      participants.forEach((opponent) => {
        if (opponent.player === player) {
          return;
        }

        const actualScore = getPairwiseActualScore(place, opponent.place);
        const expectedScore = getExpectedEloScore(preGameRatings[player], preGameRatings[opponent.player]);
        ratingDelta += actualScore - expectedScore;
      });

      const baseRatingChange = eloKFactor * (ratingDelta / (participants.length - 1));
      if (place === 1) {
        stats[player].currentWinStreak += 1;
      } else {
        stats[player].currentWinStreak = 0;
      }

      const streakBonus = stats[player].currentWinStreak > 1
        ? Math.min(maxStreakBonus, (stats[player].currentWinStreak - 1) * streakBonusStep)
        : 0;

      const singleGameTurnEloBonus = place === 1 && winningTurn
        ? getSingleGameWinTurnEloBonus(winningTurn)
        : 0;

      stats[player].rating += baseRatingChange + streakBonus + (kills * eloKillBonus) + singleGameTurnEloBonus;
    });
  });

  const playerRankingEntries = Object.values(stats)
    .filter((entry) => entry.games > 0)
    .map((entry) => {
      const avgWinTurn = entry.winTurnCount ? (entry.winTurnTotal / entry.winTurnCount) : 0;
      const averageWinTurnEloBonus = getAverageWinTurnEloBonus(avgWinTurn, entry.winTurnCount);
      const recentWinTurnEloBonus = getRecentWinTurnEloBonus(entry.winTurns || []);
      return {
        ...entry,
        rating: Math.round(entry.rating + averageWinTurnEloBonus + recentWinTurnEloBonus),
        avgWinTurn,
        averageWinTurnEloBonus,
        recentWinTurnEloBonus,
        avgPlace: entry.placementGames ? entry.placementTotal / entry.placementGames : null,
        pointsPerGame: entry.games ? entry.points / entry.games : 0,
        winRate: entry.games ? (entry.wins / entry.games) * 100 : 0,
        favoriteCommander: getMaxCountKey(entry.commanders),
      };
    })
    .sort(compareRankingsEntries);

  cacheBucket.playerRankingEntries = playerRankingEntries;
  return playerRankingEntries;
}

function buildCommanderRankingEntries(games) {
  const cacheBucket = getDerivedCacheBucket(games);
  if (cacheBucket.commanderRankingEntries) {
    return cacheBucket.commanderRankingEntries;
  }

  const rawCommanders = [];
  games.forEach((game) => {
    getGameRows(game).forEach((row) => {
      const commander = String(row.commander || '').trim();
      if (commander) {
        rawCommanders.push(commander);
      }
    });
  });

  const commanderMap = buildCanonicalIdentityMapFromValues([
    ...getKnownCommanderOptions(),
    ...rawCommanders,
  ]);

  const stats = {};

  games.forEach((game) => {
    const rows = getGameRows(game);
    const firstBlood = getGameFirstBloodInfo(game);

    rows.forEach((row) => {
      const commander = canonicalizeIdentityValue(String(row.commander || '').trim(), commanderMap);
      if (!commander) {
        return;
      }

      if (!stats[commander]) {
        stats[commander] = {
          name: commander,
          games: 0,
          wins: 0,
          kills: 0,
          firstBloods: 0,
          points: 0,
          placementTotal: 0,
          placementGames: 0,
        };
      }

      const entry = stats[commander];
      const place = getGameRowPlace(row, game);
      const kills = getRowKills(row);
      const firstBloodBonus = firstBlood?.actorPlayer === row.player ? 1 : 0;

      entry.games += 1;
      entry.kills += kills;
      entry.firstBloods += firstBloodBonus;
      entry.points += getPlacementPoints(place) + kills + firstBloodBonus;
      if (place === 1) {
        entry.wins += 1;
      }
      if (place) {
        entry.placementTotal += place;
        entry.placementGames += 1;
      }
    });
  });

  const commanderRankingEntries = Object.values(stats)
    .map((entry) => ({
      ...entry,
      avgPlace: entry.placementGames ? entry.placementTotal / entry.placementGames : null,
      pointsPerGame: entry.games ? entry.points / entry.games : 0,
      winRate: entry.games ? (entry.wins / entry.games) * 100 : 0,
    }))
    .sort(compareAggregateEntries);

  cacheBucket.commanderRankingEntries = commanderRankingEntries;
  return commanderRankingEntries;
}

function buildStreakEntries(games, keySelector, cacheKey = '') {
  const cacheBucket = getDerivedCacheBucket(games);
  if (cacheKey && cacheBucket[cacheKey]) {
    return cacheBucket[cacheKey];
  }

  const appearances = {};

  getGamesSortedByDateAscending(games).forEach((game) => {
    getGameRows(game).forEach((row) => {
      const key = String(keySelector(row) || '').trim();
      if (!key) {
        return;
      }

      if (!appearances[key]) {
        appearances[key] = [];
      }

      appearances[key].push({
        date: String(game.date || '').trim(),
        win: getGameRowPlace(row, game) === 1,
      });
    });
  });

  const streakEntries = Object.entries(appearances)
    .map(([name, gamesPlayed]) => {
      let currentWinStreak = 0;
      for (let index = gamesPlayed.length - 1; index >= 0; index -= 1) {
        if (!gamesPlayed[index].win) {
          break;
        }
        currentWinStreak += 1;
      }

      let bestWinStreak = 0;
      let runningWinStreak = 0;
      gamesPlayed.forEach((entry) => {
        if (entry.win) {
          runningWinStreak += 1;
          bestWinStreak = Math.max(bestWinStreak, runningWinStreak);
          return;
        }
        runningWinStreak = 0;
      });

      const lastWinIndex = gamesPlayed.map((entry) => entry.win).lastIndexOf(true);
      const drought = lastWinIndex >= 0 ? (gamesPlayed.length - 1) - lastWinIndex : gamesPlayed.length;
      const lastWinDate = lastWinIndex >= 0 ? gamesPlayed[lastWinIndex].date : '';

      return {
        name,
        games: gamesPlayed.length,
        currentWinStreak,
        bestWinStreak,
        drought,
        lastWinDate,
      };
    })
    .sort((a, b) => {
      if (b.currentWinStreak !== a.currentWinStreak) {
        return b.currentWinStreak - a.currentWinStreak;
      }
      if (b.bestWinStreak !== a.bestWinStreak) {
        return b.bestWinStreak - a.bestWinStreak;
      }
      if (b.games !== a.games) {
        return b.games - a.games;
      }
      return a.name.localeCompare(b.name);
    });

  if (cacheKey) {
    cacheBucket[cacheKey] = streakEntries;
  }

  return streakEntries;
}

function renderStatCardGroup(container, cards) {
  if (!container) {
    return;
  }

  container.innerHTML = '';
  cards.forEach(({ title, body }) => {
    container.appendChild(createStatCard(title, body));
  });
}

function buildSkeletonTableRows(columnCount, rowCount = 4) {
  return Array.from({ length: rowCount }, () => `
      <tr>
        ${Array.from({ length: columnCount }, () => '<td><span class="table-skeleton"></span></td>').join('')}
      </tr>`).join('');
}

function getCachedCommanderCardByName(name) {
  const normalizedName = String(name || '').trim();
  const normalizedKey = getIdentityKey(normalizedName);
  if (!normalizedKey) {
    return null;
  }

  const cachedCard = deckBuilderCardCache.get(normalizedName.toLowerCase());
  if (cachedCard) {
    return cachedCard;
  }

  const savedDeckMatch = loadDecks().find((deck) => getIdentityKey(deck?.commander?.name || '') === normalizedKey);
  return savedDeckMatch?.commander || null;
}

function getCommanderIdentitySymbols(name, cardOverride = null) {
  const card = cardOverride || getCachedCommanderCardByName(name);
  if (!card) {
    return null;
  }

  const identity = Array.from(getCardEffectiveColorIdentity(card));
  const orderedIdentity = ['W', 'U', 'B', 'R', 'G'].filter((color) => identity.includes(color));
  return orderedIdentity.length ? orderedIdentity : ['C'];
}

function buildCommanderIdentityPips(name, cardOverride = null) {
  const identitySymbols = getCommanderIdentitySymbols(name, cardOverride);
  if (!identitySymbols) {
    return '';
  }

  return `<span class="commander-identity-pips" aria-label="Color identity ${identitySymbols.map((symbol) => COLOR_LETTER_NAMES[symbol] || 'Colorless').join(', ')}">${identitySymbols.map((symbol) => `<span class="commander-identity-pip commander-identity-pip--${symbol.toLowerCase()}" title="${escapeHtml(COLOR_LETTER_NAMES[symbol] || 'Colorless')}">${escapeHtml(symbol)}</span>`).join('')}</span>`;
}

function buildCommanderDisplayHtml(name, contentHtml, cardOverride = null) {
  const displayName = String(name || '').trim();
  if (!displayName) {
    return '—';
  }

  return `<span class="commander-identity-display">${contentHtml}${buildCommanderIdentityPips(displayName, cardOverride)}</span>`;
}

function buildCommanderTextHtml(name, cardOverride = null) {
  const displayName = String(name || '').trim();
  if (!displayName) {
    return '—';
  }

  return buildCommanderDisplayHtml(displayName, escapeHtml(displayName), cardOverride);
}

function collectCommanderDisplayNames(games = []) {
  const commanderNames = new Set();

  (Array.isArray(games) ? games : []).forEach((game) => {
    getGameRows(game).forEach((row) => {
      const commander = String(row?.commander || '').trim();
      if (commander) {
        commanderNames.add(commander);
      }
    });
  });

  loadRecords().forEach((record) => {
    const commander = String(record?.commander || record?.manualCommander || '').trim();
    if (commander) {
      commanderNames.add(commander);
    }
  });

  loadDeckLists().forEach((entry) => {
    const commander = String(entry?.commander || '').trim();
    if (commander) {
      commanderNames.add(commander);
    }
  });

  loadDecks().forEach((deck) => {
    const primaryCommander = String(deck?.commander?.name || '').trim();
    const secondCommander = String(deck?.secondCommander?.name || '').trim();
    if (primaryCommander) {
      commanderNames.add(primaryCommander);
    }
    if (secondCommander) {
      commanderNames.add(secondCommander);
    }
  });

  (activeGameState?.players || []).forEach((player) => {
    const commander = String(player?.commander || '').trim();
    if (commander) {
      commanderNames.add(commander);
    }
  });

  return Array.from(commanderNames);
}

function warmCommanderIdentityDisplayCache(games = []) {
  if (commanderIdentityPrefetchInFlight) {
    return;
  }

  const pendingNames = collectCommanderDisplayNames(games).filter((name) => {
    const identityKey = getIdentityKey(name);
    return identityKey && !getCachedCommanderCardByName(name) && !commanderIdentityAttemptedKeys.has(identityKey);
  });

  if (!pendingNames.length) {
    return;
  }

  pendingNames.forEach((name) => commanderIdentityAttemptedKeys.add(getIdentityKey(name)));
  commanderIdentityPrefetchInFlight = true;
  const requestId = ++commanderIdentityPrefetchRequestId;

  fetchDeckCardsByNamesBulk(pendingNames)
    .catch(() => new Map())
    .finally(() => {
      if (requestId !== commanderIdentityPrefetchRequestId) {
        return;
      }
      commanderIdentityPrefetchInFlight = false;
      refresh();
    });
}

function collectRankingsCommanderNames(games) {
  const commanderNames = new Set();

  buildPlayerRankingEntries(games).forEach((entry) => {
    if (entry.favoriteCommander) {
      commanderNames.add(entry.favoriteCommander);
    }
  });

  buildCommanderRankingEntries(games).forEach((entry) => {
    if (entry.name) {
      commanderNames.add(entry.name);
    }
  });

  buildStreakEntries(games, (row) => row.commander, 'commanderStreakEntries').forEach((entry) => {
    if (entry.name) {
      commanderNames.add(entry.name);
    }
  });

  return Array.from(commanderNames);
}

function getPendingRankingsCommanderNames(games) {
  return collectRankingsCommanderNames(games).filter((name) => {
    const identityKey = getIdentityKey(name);
    return identityKey && !getCachedCommanderCardByName(name) && !rankingsCommanderIdentityAttemptedKeys.has(identityKey);
  });
}

function ensureRankingsCommanderIdentities(games) {
  const pendingNames = getPendingRankingsCommanderNames(games);
  if (!pendingNames.length) {
    return false;
  }

  pendingNames.forEach((name) => rankingsCommanderIdentityAttemptedKeys.add(getIdentityKey(name)));
  rankingsCommanderIdentityLoading = true;
  window.CommanderRankings?.renderLoading();

  const requestId = ++rankingsCommanderIdentityRequestId;
  fetchDeckCardsByNamesBulk(pendingNames)
    .catch(() => new Map())
    .finally(() => {
      if (requestId !== rankingsCommanderIdentityRequestId) {
        return;
      }
      rankingsCommanderIdentityLoading = false;
      renderRankingsPage(loadGames());
      applyResponsiveTableLabels();
    });

  return true;
}

function renderRankingsPage(games) {
  if (rankingsCommanderIdentityLoading) {
    window.CommanderRankings?.renderLoading();
    return;
  }

  if (games.length && ensureRankingsCommanderIdentities(games)) {
    return;
  }

  window.CommanderRankings?.renderAll(games);
}

function applyResponsiveTableLabels() {
  const tables = document.querySelectorAll('.player-table, .history-game-table');

  tables.forEach((table) => {
    const headers = Array.from(table.querySelectorAll('thead th')).map((header) => header.textContent.trim());
    if (!headers.length) {
      return;
    }

    table.querySelectorAll('tbody tr').forEach((row) => {
      const cells = Array.from(row.children).filter((cell) => cell.tagName === 'TD');
      cells.forEach((cell, index) => {
        const label = headers[index] || '';
        if (label) {
          cell.dataset.label = label;
        }
      });
    });
  });
}

let activeTableColumnResize = null;

function resetTableColumnWidths(table) {
  if (!table) {
    return;
  }

  table.querySelectorAll('th, td').forEach((cell) => {
    cell.style.width = '';
    cell.style.minWidth = '';
  });
  table.querySelectorAll('.table-column-resize-handle').forEach((handle) => {
    handle.remove();
  });
}

function initializeTableColumnResizers() {
  document.querySelectorAll('.player-table').forEach((table) => {
    resetTableColumnWidths(table);

    const headers = Array.from(table.querySelectorAll(':scope > thead > tr:first-child > th'));
    if (!headers.length) {
      return;
    }

    headers.forEach((header, index) => {
      if (index === headers.length - 1) {
        return;
      }

      const handle = document.createElement('span');
      handle.className = 'table-column-resize-handle';
      handle.setAttribute('role', 'separator');
      handle.setAttribute('aria-label', `Resize ${header.textContent.trim()} column`);
      handle.addEventListener('mousedown', (event) => {
        event.preventDefault();
        event.stopPropagation();
        const tableRect = table.getBoundingClientRect();
        activeTableColumnResize = {
          table,
          index,
          startX: event.clientX,
          startWidth: header.getBoundingClientRect().width,
          tableLeft: tableRect.left,
        };
        document.body.classList.add('is-resizing-table-column');
      });
      handle.addEventListener('click', (event) => {
        event.preventDefault();
        event.stopPropagation();
      });
      header.appendChild(handle);
    });
  });
}

document.addEventListener('mousemove', (event) => {
  if (!activeTableColumnResize) {
    return;
  }

  const { table, index, startX, startWidth } = activeTableColumnResize;
  const tableWidth = table.getBoundingClientRect().width;
  const nextWidth = Math.max(64, startWidth + (event.clientX - startX));
  const cells = table.querySelectorAll(`tr > :nth-child(${index + 1})`);
  cells.forEach((cell) => {
    cell.style.width = `${Math.min(nextWidth, Math.max(64, tableWidth - 64))}px`;
    cell.style.minWidth = `${Math.min(nextWidth, Math.max(64, tableWidth - 64))}px`;
  });
});

document.addEventListener('mouseup', () => {
  if (!activeTableColumnResize) {
    return;
  }
  activeTableColumnResize = null;
  document.body.classList.remove('is-resizing-table-column');
});

function handleSortableHeaderClick(event) {
  const header = event.target.closest('.sortable-header');
  if (!header) {
    return;
  }

  const table = header.closest('[data-sort-table]');
  const tableKey = table?.dataset.sortTable;
  const column = header.dataset.column;
  if (!tableKey || !column) {
    return;
  }

  toggleTableSort(tableKey, column, getTableSortDefaultDescending(tableKey, column));
  rerenderSortedTable(tableKey);
}

function handleSortableHeaderKeydown(event) {
  if (event.key !== 'Enter' && event.key !== ' ') {
    return;
  }

  event.preventDefault();
  handleSortableHeaderClick(event);
}

function renderSummary(games) {
  if (!summaryEl) {
    return;
  }

  summaryEl.innerHTML = '';

  if (!games.length) {
    summaryEl.appendChild(createStatCard('No games yet', 'Add a game to start tracking commander results.'));
    return;
  }

  const total = games.length;
  const mostRecent = games
    .slice()
    .sort((a, b) => b.date.localeCompare(a.date))[0].date;

  const playerWins = {};
  const commanderCount = {};

  games.forEach((game) => {
    if (game.finishOrder && game.finishOrder.length) {
      const winner = game.finishOrder[0];
      playerWins[winner] = (playerWins[winner] || 0) + 1;
    }

    const playerCommandersList = getGameRows(game).map(({ player, commander }) => ({ player, commander }));
    playerCommandersList.forEach(({ commander }) => {
      if (commander) {
        commanderCount[commander] = (commanderCount[commander] || 0) + 1;
      }
    });
  });

  const bestPlayers = Object.entries(playerWins)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([player, wins]) => `${player} (${wins} wins)`)
    .join(', ');

  const bestCommanders = Object.entries(commanderCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([commander, count]) => `${buildCommanderTextHtml(commander)} (${count})`)
    .join('<br />');

  summaryEl.appendChild(createStatCard('Total games', total));
  summaryEl.appendChild(createStatCard('Most recent game', mostRecent));
  summaryEl.appendChild(createStatCard('Top winners', bestPlayers || 'No winners yet')); 
  summaryEl.appendChild(createStatCard('Popular commanders', bestCommanders || 'No commander data yet'));
}

function refresh() {
  const games = loadGames();
  const shouldUpdateLookupCaches = Boolean(
    form
    || window.CommanderLiveGame?.hasView?.()
    || window.CommanderRecords?.hasView()
    || window.CommanderDeckBuilder?.hasView?.()
    || playerRenameForm
    || commanderRenameForm
    || window.CommanderHistory?.hasView()
  );
  const hasRankingsView = Boolean(window.CommanderRankings?.hasView());
  const hasHistoryView = Boolean(window.CommanderHistory?.hasView());
  const hasCommanderStatsView = Boolean(window.CommanderStats?.hasView());
  const hasDeckLibraryView = Boolean(window.CommanderDeckLibrary?.hasView());
  const hasDeckSelectorView = Boolean(window.CommanderDeckSelector?.hasView());
  const hasDeckBuilderView = Boolean(window.CommanderDeckBuilder?.hasView?.());
  const hasRecordsView = Boolean(window.CommanderRecords?.hasView());
  const hasLiveView = Boolean(window.CommanderLiveGame?.hasView?.());
  const hasAdminLogsView = Boolean(authAuditSection || registeredAccountsSection || document.body.classList.contains('page-admin-logs'));
  const shouldWarmCommanderIdentityCache = Boolean(
    hasRankingsView
    || hasHistoryView
    || hasCommanderStatsView
    || hasDeckSelectorView
  );

  if (shouldUpdateLookupCaches) {
    updateFormDatalists(games);
    renderIdentityRenameOptions();
  }

  if (summaryEl) {
    renderSummary(games);
  }

  if (hasRankingsView) {
    renderRankingsPage(games);
  }

  if (window.CommanderPlayerStats?.hasView()) {
    window.CommanderPlayerStats.render(games);
  }

  if (hasHistoryView) {
    window.CommanderHistory.updateFilters(games);
    window.CommanderHistory.applyQueryFilters();
    window.CommanderHistory.render(games);
  }

  if (hasCommanderStatsView) {
    window.CommanderStats.render(games);
  }

  if (hasDeckLibraryView) {
    window.CommanderDeckLibrary.render();
  }

  if (hasDeckSelectorView) {
    window.CommanderDeckSelector.render();
  }

  if (shouldWarmCommanderIdentityCache) {
    warmCommanderIdentityDisplayCache(games);
  }

  if (hasDeckBuilderView) {
    window.CommanderDeckBuilder?.render?.();
  }

  if (hasRecordsView) {
    window.CommanderRecords.render();
  }

  if (hasLiveView) {
    window.CommanderLiveGame.refresh();
  }

  initializeMobileSortControls();
  applyResponsiveTableLabels();
  initializeTableColumnResizers();

  if (hasHistoryView || hasCommanderStatsView || hasRankingsView || hasDeckSelectorView) {
    void backfillDeckListCommanderOracleIds();
  }
}

function resetForm() {
  form.reset();
  resetPlayerTable();
  dateInput.valueAsDate = new Date();
}

if (liveModalCancelButton) {
  liveModalCancelButton.addEventListener('click', () => {
    hideLiveModal(null);
  });
}

if (liveModalConfirmButton) {
  liveModalConfirmButton.addEventListener('click', () => {
    if (!liveModalConfig) {
      hideLiveModal(true);
      return;
    }

    const activeControl = getActiveLiveModalControl();
    if (!activeControl) {
      hideLiveModal(true);
      return;
    }

    const rawValue = activeControl.value;
    const validationResult = typeof liveModalConfig.validate === 'function'
      ? liveModalConfig.validate(rawValue)
      : { value: rawValue };

    if (validationResult && validationResult.error) {
      setLiveModalError(validationResult.error);
      activeControl.focus();
      return;
    }

    hideLiveModal(validationResult && Object.prototype.hasOwnProperty.call(validationResult, 'value') ? validationResult.value : rawValue);
  });
}

if (liveModalInput) {
  liveModalInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      liveModalConfirmButton?.click();
    }
  });
}

if (cardImageModalCloseButton) {
  cardImageModalCloseButton.addEventListener('click', () => {
    closeCardImageModal();
  });
}

if (cardImageModalOverlay) {
  cardImageModalOverlay.addEventListener('click', () => {
    closeCardImageModal();
  });
}

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && cardImageModal && !cardImageModal.hidden) {
    closeCardImageModal();
  }
});

document.addEventListener('click', (event) => {
  const rulingsCardImage = event.target.closest('.rulings-card-image');
  if (!rulingsCardImage || !cardImageModal) {
    return;
  }

  openCardImageModal(rulingsCardImage.currentSrc || rulingsCardImage.src, rulingsCardImage.alt || 'Card image');
});

document.addEventListener('click', (event) => {
  if (pageSwitch && pageSwitch.classList.contains('is-open') && !event.target.closest('.page-switch')) {
    closePrimaryMenu();
  }

});

document.addEventListener('keydown', (event) => {
  if (event.key !== 'Escape') {
    return;
  }

  if (pageSwitch && pageSwitch.classList.contains('is-open')) {
    closePrimaryMenu();
    return;
  }

  if (liveModalPrompt && !liveModalPrompt.hidden) {
    hideLiveModal(null);
    return;
  }

});

const playerRenameToggleButton = document.getElementById('player-rename-toggle');

if (playerRenameToggleButton && playerRenameForm) {
  playerRenameToggleButton.addEventListener('click', () => {
    const isOpening = playerRenameForm.hidden;
    playerRenameForm.hidden = !isOpening;
    playerRenameToggleButton.setAttribute('aria-expanded', String(isOpening));
    playerRenameToggleButton.textContent = isOpening ? 'Hide rename player' : 'Rename player…';
    if (isOpening) {
      playerRenameCurrentInput?.focus();
    }
  });
}

if (playerRenameForm) {
  if (playerRenameStatus) {
    playerRenameStatus.setAttribute('role', 'status');
    playerRenameStatus.setAttribute('aria-live', 'polite');
  }

  playerRenameForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    await handleIdentityRenameSubmit({
      type: 'player',
      currentInput: playerRenameCurrentInput,
      nextInput: playerRenameNextInput,
      statusElement: playerRenameStatus,
    });
  });
}

if (commanderRenameForm) {
  if (commanderRenameStatus) {
    commanderRenameStatus.setAttribute('role', 'status');
    commanderRenameStatus.setAttribute('aria-live', 'polite');
  }

  commanderRenameForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    await handleIdentityRenameSubmit({
      type: 'commander',
      currentInput: commanderRenameCurrentInput,
      nextInput: commanderRenameNextInput,
      statusElement: commanderRenameStatus,
    });
  });
}

document.querySelectorAll('[data-sort-table]').forEach((table) => {
  table.addEventListener('click', handleSortableHeaderClick);
  table.addEventListener('keydown', handleSortableHeaderKeydown);
});

window.CommanderHistory?.updateSortOrderLabel();

window.addEventListener('storage', (event) => {
  const isStateStorageKey = (
    event.key === STORAGE_KEY
    || event.key === EXPECTED_POWER_STORAGE_KEY
    || event.key === DECK_LIST_STORAGE_KEY
    || event.key === DECKS_STORAGE_KEY
    || event.key === RECORDS_STORAGE_KEY
  );

  const shouldIgnoreStateStorageReload = isStateStorageKey && hasSyncCredentials() && syncHasLoadedCloudState;

  if (
    event.key === STORAGE_KEY
    || event.key === EXPECTED_POWER_STORAGE_KEY
    || event.key === DECK_LIST_STORAGE_KEY
    || event.key === DECKS_STORAGE_KEY
    || event.key === RECORDS_STORAGE_KEY
    || event.key === ACTIVE_GAME_STORAGE_KEY
    || event.key === ACTIVE_GAME_UNDO_STORAGE_KEY
    || event.key === SYNC_USER_STORAGE_KEY
    || event.key === SYNC_TOKEN_STORAGE_KEY
    || event.key === SYNC_CREDENTIAL_SET_AT_STORAGE_KEY
    || event.key === SYNC_PENDING_CHANGES_STORAGE_KEY
  ) {
    if (!shouldIgnoreStateStorageReload) {
      appState = loadLocalState();
    }
    activeGameState = loadActiveGameState();
    activeGameUndoState = loadActiveGameUndoState();
    setSyncPendingChanges(loadSyncPendingChangesState());
    const credentials = getSyncCredentials();
    if (syncUserInput) {
      syncUserInput.value = credentials.user;
    }
    if (syncTokenInput) {
      syncTokenInput.value = credentials.token;
    }
    window.CommanderHistory?.applyQueryFilters(true);
    refresh();

    if (shouldIgnoreStateStorageReload) {
      checkCloudStateFreshness({ autoPull: true, force: true });
    }
  }
});

window.addEventListener('popstate', () => {
  window.CommanderHistory?.applyQueryFilters(true);
  window.CommanderHistory?.render(loadGames());
});

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible' && activeGameState && !wakeLockSentinel) {
    acquireWakeLock();
  }
});

function flushStateForPageExit({ forceSync = false } = {}) {
  const now = Date.now();
  if (now - pageExitFlushStamp < 250 && !forceSync) {
    return;
  }

  pageExitFlushStamp = now;

  // Discard uncommitted deck-builder staging BEFORE anything below persists state.
  // The builder stages edits into appState.decks in memory only — leaving the page
  // without Save Deck is supposed to discard them — but this exit flush was writing
  // them into localStorage. That leak is how never-saved "Untitled Deck" records
  // ended up on the Decks page even though they were never synced to the cloud KV.
  const deckBuilderSyncState = getDeckBuilderSyncState();
  if (deckBuilderSyncState.hasUnsavedChanges && deckBuilderSyncState.activeDeckId) {
    const storedDecks = parseJsonSafe(readLocalStorageValue(DECKS_STORAGE_KEY) || '', []);
    const storedDecksById = new Map(
      (Array.isArray(storedDecks) ? storedDecks : [])
        .map((deck) => [String(deck?.id || ''), deck])
        .filter(([deckId]) => deckId),
    );
    appState = normalizeAppStateData({
      ...appState,
      decks: (Array.isArray(appState.decks) ? appState.decks : [])
        .map((deck) => (deck && deck.id === deckBuilderSyncState.activeDeckId ? storedDecksById.get(deck.id) || null : deck))
        .filter(Boolean),
    });
  }

  // Only keepalive-push on exit when there were genuinely unsynced local changes.
  // Pushing unchanged state still bumps the cloud revision, but the keepalive
  // response is ignored — so the client never learns the new revision, and the next
  // save fails with a phantom 409 "changed on another device" conflict that freezes
  // sync and leaves pages showing stale local data.
  const hadPendingLocalChanges = syncPendingChanges
    || Boolean(decksPersistTimer)
    || Boolean(activeGamePersistTimer);

  flushQueuedActiveGamePersist();
  flushQueuedDeckPersist({ force: true, queueSync: false });
  persistLocalState(appState);

  if (!hadPendingLocalChanges) {
    // The flushes above mark pending unconditionally; undo that when nothing was
    // actually waiting to sync, so the next page load doesn't attempt a phantom push.
    setSyncPendingChanges(false);
    return;
  }

  if (hasSyncCredentials()) {
    setSyncPendingChanges(true);
    syncLastErrorMessage = '';
    refreshSyncStatus();
    pushCloudStateKeepalive({ force: true });
  }
}

function handleInternalPageNavigation(event) {
  const anchor = event.target instanceof Element ? event.target.closest('a[href]') : null;
  if (!anchor || anchor.hasAttribute('download')) {
    return;
  }

  const href = anchor.getAttribute('href') || '';
  const normalizedHref = String(href || '').trim();
  if (!normalizedHref || normalizedHref.startsWith('#') || normalizedHref.startsWith('mailto:') || normalizedHref.startsWith('tel:') || normalizedHref.startsWith('javascript:')) {
    return;
  }

  const target = anchor.getAttribute('target') || '';
  if (target === '_blank' || target === '_parent' || target === '_top') {
    return;
  }

  try {
    const absoluteUrl = new URL(normalizedHref, window.location.href);
    if (absoluteUrl.origin !== window.location.origin) {
      return;
    }
  } catch (error) {
    return;
  }

  event.preventDefault();
  flushStateForPageExit({ forceSync: true });
  window.setTimeout(() => {
    window.location.assign(normalizedHref);
  }, 80);
}

document.addEventListener('click', handleInternalPageNavigation, true);


window.addEventListener('resize', () => {
  closePrimaryMenu();
  window.CommanderLiveGame?.handleViewportChange?.();
});

window.addEventListener('orientationchange', () => {
  closePrimaryMenu();
  window.setTimeout(() => {
    window.CommanderLiveGame?.handleViewportChange?.();
  }, 100);
});

window.addEventListener('load', () => {
  window.CommanderLiveGame?.refreshViewport?.();
});

window.visualViewport?.addEventListener('resize', () => {
  window.CommanderLiveGame?.refreshViewport?.();
});

window.addEventListener('online', async () => {
  refreshSyncStatus();
  await checkCloudStateFreshness({ autoPull: shouldAutoPullCloudState(), force: true });
  if (!syncConflictInfo && hasSyncCredentials()) {
    void retryCloudBootstrap();
  }
  if (!syncConflictInfo && hasSyncCredentials() && (syncPendingChanges || syncLastErrorMessage)) {
    queueCloudSync(250);
  }
});

window.addEventListener('offline', () => {
  clearSyncRetryTimer();
  refreshSyncStatus();
});

document.addEventListener('visibilitychange', async () => {
  if (document.visibilityState === 'hidden') {
    flushStateForPageExit();
  }

  if (document.visibilityState === 'visible') {
    // On mobile the tab is suspended/resumed rather than reloaded, so bypass the
    // 15s freshness cache (force) and auto-pull any newer cloud revision when local
    // has no unsynced pending changes AND no unsaved deck-builder draft in progress.
    // This keeps resumed data accurate without requiring the app to be closed and
    // reopened, while never pulling over in-progress edits.
    const hasLocalChanges = syncPendingChanges || hasUnsavedDeckBuilderChanges();
    await checkCloudStateFreshness({ autoPull: !hasLocalChanges, force: true });
    if (hasSyncCredentials() && !syncPendingChanges) {
      await retryCloudBootstrap({ maxAttempts: 3, initialDelayMs: 800 });
    } else if (hasSyncCredentials()) {
      // Local has unsynced changes — push them rather than pulling over them.
      queueCloudSync(0);
    }
  }
});

// Poll cloud freshness while the page stays open so changes made on another
// device (or tab) show up without a manual refresh. Each tick is a tiny metadata
// fetch; the full multi-MB state is only downloaded when the cloud revision
// actually moved. Never auto-pull while the user has unsynced local changes, a
// deck-builder session open, or focus inside a form field — a completed pull
// re-renders the page via refresh(), which rebuilds form rows and would destroy
// unsaved input mid-edit.
window.setInterval(() => {
  if (
    document.visibilityState !== 'visible'
    || !navigator.onLine
    || !hasSyncCredentials()
    || !syncHasLoadedCloudState
    || syncConflictInfo
  ) {
    return;
  }

  const focusedElement = document.activeElement;
  const userIsEditing = focusedElement instanceof HTMLElement
    && focusedElement.matches('input:not([type="checkbox"]):not([type="radio"]):not([type="button"]):not([type="submit"]), textarea, select, [contenteditable="true"]');

  const hasLocalChanges = syncPendingChanges
    || hasUnsavedDeckBuilderChanges()
    || Boolean(getDeckBuilderSyncState().activeDeckId)
    || userIsEditing;

  void checkCloudStateFreshness({ autoPull: !hasLocalChanges });
}, 30000);

window.addEventListener('pagehide', () => {
  flushStateForPageExit({ forceSync: true });
});

window.addEventListener('focus', async () => {
  // Focus fires on every desktop alt-tab/window click, so use the cached (non-forced)
  // freshness check here — only a real tab suspend/resume (visibilitychange) forces an
  // immediate check. This avoids a burst of metadata pulls while rapidly editing decks.
  const hasLocalChanges = syncPendingChanges || hasUnsavedDeckBuilderChanges();
  checkCloudStateFreshness({ autoPull: !hasLocalChanges });
  if (hasSyncCredentials() && !syncPendingChanges) {
    await retryCloudBootstrap({ maxAttempts: 3, initialDelayMs: 800 });
  } else if (hasSyncCredentials()) {
    queueCloudSync(0);
  }
});

function setupSyncUi() {
  if (!syncUserInput || !syncTokenInput) {
    return;
  }

  const storedUser = String(readLocalStorageValue(SYNC_USER_STORAGE_KEY) || '').trim();
  const storedToken = String(readLocalStorageValue(SYNC_TOKEN_STORAGE_KEY) || '').trim();
  syncUserInput.value = storedUser;
  syncTokenInput.value = storedToken;
  if (storedUser && storedToken) {
    applyStoredSyncSession(storedUser, storedToken);
  } else {
    syncConnectionState = 'local';
  }
  syncHasLoadedCloudState = false;
  syncLocalContentVerified = false;
  updateSyncControls();
  refreshSyncStatus();

  const handleSyncConnect = async (event) => {
    event?.preventDefault?.();

    const user = syncUserInput.value.trim();
    const token = syncTokenInput.value.trim();

    if (!user || !token) {
      setSyncStatus('Enter both display name and pod access code.', 'error');
      return;
    }

    setSyncUiCollapsed(false);
    clearSyncRetryTimer();
    syncRetryCount = 0;
    setSyncPendingChanges(false);
    syncLastSuccessAt = null;
    syncLastErrorMessage = '';
    syncHasLoadedCloudState = false;
    syncLocalContentVerified = false;
    syncInFlight = true;
    clearSyncAuthenticatedUser();
    updateSyncMetadata();
    clearSyncConflict();
    syncConnectionState = 'connecting';
    updateSyncControls();
    refreshSyncStatus();
    window.CommanderDeckLibrary?.render?.();

    try {
      applyStoredSyncSession(user, token);
      syncConnectionState = 'connecting';
      setSyncUiCollapsed(false);
      refreshSyncStatus();

      if (!navigator.onLine) {
        throw new Error('You appear to be offline. Check your connection and try again.');
      }

      const sessionPayload = await cloudRequest('/api/session', {
        method: 'POST',
        body: JSON.stringify({ user, token }),
        timeoutMs: 5000,
      });
      if (!sessionPayload?.auth) {
        throw new Error('The server did not confirm this login. Check the display name and access code.');
      }

      updateSyncAuthenticatedUser(sessionPayload.auth);
      await pullCloudState();
      setSyncUiCollapsed(true);
    } catch (error) {
      syncConnectionState = 'configured';
      clearSyncAuthenticatedUser();
      syncLastErrorMessage = error.message;
      setSyncUiCollapsed(false);
      refreshSyncStatus();
      window.CommanderDeckLibrary?.render?.();
    } finally {
      syncInFlight = false;
      updateSyncControls();
      refreshSyncStatus();
    }
  };

  if (syncConnectButton) {
    syncConnectButton.addEventListener('click', handleSyncConnect);
  }

  [syncUserInput, syncTokenInput].forEach((input) => {
    input?.addEventListener('input', updateSyncControls);
    input?.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        handleSyncConnect(event);
      }
    });
  });

  if (syncDisconnectButton) {
    syncDisconnectButton.addEventListener('click', () => {
      clearSyncRetryTimer();
      if (syncQueueTimer) {
        clearTimeout(syncQueueTimer);
        syncQueueTimer = null;
      }
      setSyncPendingChanges(false);
      syncRetryCount = 0;
      syncLastSuccessAt = null;
      syncLastErrorMessage = '';
      syncHasLoadedCloudState = false;
      syncLocalContentVerified = false;
      updateSyncMetadata();
      appState = normalizeAppStateData({ games: [], powerLevels: {}, deckLists: [], decks: [], records: [], activeGame: null, activeGameUndo: [] });
      activeGameState = null;
      activeGameUndoState = [];
      clearSyncAuthenticatedUser();
      clearSyncConflict();
      syncConnectionState = 'local';
      syncUserInput.value = '';
      syncTokenInput.value = '';
      clearPersistedSyncCredentials();
      void cloudRequest('/api/session', { method: 'DELETE' }).catch(() => null);
      updateSyncControls();
      refreshSyncStatus();
      refresh();
    });
  }

  if (syncNowButton) {
    syncNowButton.addEventListener('click', async () => {
      if (syncConflictInfo) {
        await resolveSyncConflict(syncConflictInfo);
        return;
      }

      clearSyncRetryTimer();
      syncRetryCount = 0;
      syncLastErrorMessage = '';
      refreshSyncStatus();

      // Reconcile instead of blindly pushing: if the cloud baseline has moved in
      // either direction since our last sync (another device pushed, or a KV
      // restore rolled the revision back), local state is stale — pull the
      // authoritative cloud state rather than overwriting it.
      try {
        const metadata = await fetchCloudStateMetadata();
        if (isCloudStateDesynced(metadata)) {
          updateSyncMetadata(metadata);
          await pullCloudState({ force: true });
          return;
        }
      } catch (error) {
        // Metadata check failed; fall through to the push — the worker's revision
        // guard still protects the cloud (a 409 conflict prompts a pull instead).
      }

      setSyncPendingChanges(true);
      await pushCloudState();
    });
  }
}

async function restorePersistedSyncSession() {
  const storedUser = String(readLocalStorageValue(SYNC_USER_STORAGE_KEY) || '').trim();
  const storedToken = String(readLocalStorageValue(SYNC_TOKEN_STORAGE_KEY) || '').trim();

  if (!storedUser || !storedToken) {
    return false;
  }

  if (getCurrentSyncUserId() && syncHasLoadedCloudState) {
    return true;
  }

  applyStoredSyncSession(storedUser, storedToken);

  if (!navigator.onLine) {
    return true;
  }

  void cloudRequest('/api/session', {
    method: 'POST',
    body: JSON.stringify({ user: storedUser, token: storedToken }),
    timeoutMs: 10000,
  }).then((sessionPayload) => {
    if (sessionPayload?.auth) {
      updateSyncAuthenticatedUser(sessionPayload.auth || null);
      refreshSyncStatus();
    }
    // No pull here — the page-load bootstrap owns pulling (metadata-first), so a
    // second full-state download from this callback would just double bandwidth.
  }).catch(() => null);

  return true;
}

async function initializeApp() {
  appState = loadLocalState();
  restorePersistedSyncRevision();
  setSyncPendingChanges(loadSyncPendingChangesState());
  activeGameState = loadActiveGameState();
  activeGameUndoState = loadActiveGameUndoState();
  initializePrimaryMenu();
  setupSyncUi();

  // Capture before refresh() or replaceState can strip URL params.
  // Do NOT set deckBuilderCommanderPrefill yet — pullCloudState() overwrites appState
  // and calls refresh(), which would wipe the newly created deck. We apply it after sync.
  const pendingDeckCommanderPrefill = window.CommanderDeckBuilder?.hasView?.() && !getQueryParam('deckId')
    ? getQueryParam('commander') || ''
    : ''

  if (form) {
    resetForm();
    const editId = getQueryParam('editId');
    if (editId) {
      const game = getGameById(editId);
      if (game) {
        setEditMode(game);
      }
    }
  }

  window.CommanderLiveGame?.init?.();

  refresh();

  const restoreEditModeIfNeeded = () => {
    if (!form) {
      return;
    }

    const editId = getQueryParam('editId');
    if (!editId) {
      return;
    }

    const game = getGameById(editId);
    if (game) {
      setEditMode(game);
    }
  };

  restoreEditModeIfNeeded();
  refreshSyncStatus();

  void (async () => {
    // Restore the prior session from stored credentials in the background so page-to-page
    // navigation does not force a fresh login prompt.
    if (hasSyncCredentials()) {
      void restorePersistedSyncSession().catch(() => null);
    }

    // Clean up empty "Untitled Deck" records that leaked into storage from unsaved
    // deck-builder sessions. Stored sync credentials are applied synchronously by the
    // restore call above, so ownership matching works here even before the cloud
    // session resolves.
    if (pruneStaleEmptyDeckRecords()) {
      refresh();
    }

    if (!hasSyncCredentials() || !navigator.onLine) {
      return;
    }

    // If this device has unsynced local changes, reconcile them rather than blindly
    // pulling cloud state over them (a pull replaces appState wholesale). A keepalive
    // push on the previous page's exit may already have landed and bumped the cloud
    // revision — its response is ignored, so we can't tell locally. Check metadata:
    // if the cloud moved ahead, the changes are already up there and we just pull to
    // resync; otherwise we push the pending local state.
    if (syncPendingChanges) {
      try {
        const metadata = await fetchCloudStateMetadata();
        if (isCloudStateDesynced(metadata)) {
          // The cloud baseline moved in EITHER direction — another device pushed,
          // or a KV restore rolled the revision back below ours. Either way the
          // pending local state was built against a different cloud baseline, so
          // pushing it would silently overwrite whatever is up there now. Pull the
          // authoritative cloud state instead (forced, so a rolled-back revision
          // can't trip the stale-read guard).
          updateSyncMetadata(metadata);
          await pullCloudState({ force: true });
        } else {
          await pushCloudState();
        }
      } catch (error) {
        // Leave the pending flag set so the next opportunity retries.
        refreshSyncStatus();
        return;
      }

      restoreEditModeIfNeeded();
      setSyncUiCollapsed(false);
      refreshSyncStatus();
      return;
    }

    // No pending local changes — before downloading the full (multi-MB) state
    // payload, check the tiny metadata endpoint. Trust the local cache only when
    // BOTH the revision matches AND the collection sizes match: revision equality
    // alone cannot prove content equality (metadata polls advance a device's
    // revision without downloading state), which is how the 2026-10-12 stale push
    // wiped the cloud. A count mismatch falls through to the full pull below, so
    // stale devices heal automatically instead of silently using local data.
    try {
      const metadata = await fetchCloudStateMetadata();
      if (!isCloudStateDesynced(metadata) && hasAnyStateData(appState) && localStateMatchesCloudCounts(metadata?.counts)) {
        if (!window.CommanderDeckLibrary?.hasView?.()) {
          updateSyncMetadata(metadata);
          syncHasLoadedCloudState = true;
          syncConnectionState = 'connected';
          syncLastErrorMessage = '';
          restoreEditModeIfNeeded();
          setSyncUiCollapsed(false);
          refreshSyncStatus();
          return;
        }
      }

      // Cloud moved — record the revision now so a stale edge-cached body in the
      // pull below is rejected by the stale-read guard instead of wiping state.
      updateSyncMetadata(metadata);
    } catch (error) {
      // Metadata check failed — fall through to the pull-based bootstrap.
    }

    const succeeded = await retryCloudBootstrap();
    if (!succeeded) {
      return;
    }

    restoreEditModeIfNeeded();
    setSyncUiCollapsed(false);
    refreshSyncStatus();
  })();

  // Create and prefill the deck AFTER cloud sync settles so pullCloudState
  // cannot overwrite appState and wipe the newly created deck.
  if (pendingDeckCommanderPrefill) {
    window.CommanderDeckBuilder?.setCommanderPrefill?.(pendingDeckCommanderPrefill);
    window.CommanderDeckBuilder?.render?.();
  }

}

if (typeof window !== 'undefined') {
  window.CommanderBuilderPage = {
    renderCommanderBuilder,
    updateCommanderBuilderModeUi,
    syncCommanderBuilderPreviewState,
    syncCommanderBuilderExclusiveSelection,
    toggleCommanderBuilderKeyword,
    renderCommanderBuilderKeywordCatalog,
    setCommanderBuilderKeywordSearchTerm: (value) => {
      commanderBuilderKeywordSearchTerm = String(value || '').trim();
    },
    runCommanderBuilderRoll,
    rerollCommanderBuilderCard,
  };

  window.CommanderAdminLogs = {
    refreshAuthAuditLogs,
    refreshRegisteredAccounts,
    renderAuthAuditLogs,
    renderRegisteredAccounts,
    renderSyncDebugLog,
    updateAuthAuditFilterStateFromInputs,
    updateAuthAuditPaginationState,
    updateAuthAuditStatusSummary,
    resetAuthAuditFilters,
    exportFilteredAuthAuditAsJson,
    exportFilteredAuthAuditAsCsv,
    getAuthAuditState: () => ({
      currentPage: authAuditCurrentPage,
      setCurrentPage: (page) => {
        authAuditCurrentPage = Math.max(1, Number(page) || 1);
      },
      clearSyncDebugLogs: () => {
        syncDebugLogs = [];
      },
    }),
  };
}

initializeApp();