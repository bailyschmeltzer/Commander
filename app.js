const STORAGE_KEY = 'commanderTrackerGames';
const EXPECTED_POWER_STORAGE_KEY = 'commanderExpectedPowerLevels';
const DECK_LIST_STORAGE_KEY = 'commanderDeckLists';
const RECORDS_STORAGE_KEY = 'commanderTrackerRecords';
const ACTIVE_GAME_STORAGE_KEY = 'commanderTrackerActiveGame';
const ACTIVE_GAME_UNDO_STORAGE_KEY = 'commanderTrackerActiveGameUndo';
const SYNC_USER_STORAGE_KEY = 'commanderTrackerSyncUser';
const SYNC_TOKEN_STORAGE_KEY = 'commanderTrackerSyncToken';
const CLOUD_SYNC_ENDPOINT = '/api/state';
const CLOUD_SYNC_METADATA_ENDPOINT = '/api/state?meta=1';
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
const playerStatsTableBody = document.getElementById('player-stats-body');
const playerRenameForm = document.getElementById('player-rename-form');
const playerRenameCurrentInput = document.getElementById('player-rename-current');
const playerRenameNextInput = document.getElementById('player-rename-next');
const playerRenameStatus = document.getElementById('player-rename-status');
const playerRenameDatalist = document.getElementById('player-rename-list');
const rankingsSummary = document.getElementById('rankings-summary');
const rankingsTableBody = document.getElementById('rankings-table-body');
const recentTrendsSummary = document.getElementById('recent-trends-summary');
const recentPlayerTrendsBody = document.getElementById('recent-player-trends-body');
const recentCommanderTrendsBody = document.getElementById('recent-commander-trends-body');
const streaksSummary = document.getElementById('streaks-summary');
const playerStreaksBody = document.getElementById('player-streaks-body');
const commanderStreaksBody = document.getElementById('commander-streaks-body');
const historyList = document.getElementById('history-list');
const historySortSelect = document.getElementById('history-sort');
const historySortOrderButton = document.getElementById('history-sort-order');
const historyFilterWinner = document.getElementById('history-filter-winner');
const historyFilterCommander = document.getElementById('history-filter-commander');
const historyFilterPlayer = document.getElementById('history-filter-player');
const historyFilterDateFrom = document.getElementById('history-filter-date-from');
const historyFilterDateTo = document.getElementById('history-filter-date-to');
const historyResetFiltersButton = document.getElementById('history-reset-filters');
const historyActiveFilters = document.getElementById('history-active-filters');
const commanderSearch = document.getElementById('commander-search');
const commanderStatsTableBody = document.getElementById('commander-stats-body');
const commanderRenameForm = document.getElementById('commander-rename-form');
const commanderRenameCurrentInput = document.getElementById('commander-rename-current');
const commanderRenameNextInput = document.getElementById('commander-rename-next');
const commanderRenameStatus = document.getElementById('commander-rename-status');
const commanderRenameDatalist = document.getElementById('commander-rename-list');
const removePlayerRowButton = document.getElementById('remove-player-row');
const deckListForm = document.getElementById('deck-list-form');
const deckCommanderInput = document.getElementById('deck-commander');
const deckCommanderMenu = document.getElementById('deck-commander-menu');
const deckCommanderDropdownButton = document.getElementById('deck-commander-dropdown');
const deckOwnerInput = document.getElementById('deck-owner');
const deckOwnerMenu = document.getElementById('deck-owner-menu');
const deckUrlInput = document.getElementById('deck-url');
const deckListTableBody = document.getElementById('deck-list-body');
const deckListCancelButton = document.getElementById('deck-list-cancel');
const deckListSubmitButton = document.querySelector('#deck-list-form button[type="submit"]');
const deckLookupSelect = document.getElementById('deck-lookup-commander');
const deckLookupResult = document.getElementById('deck-lookup-result');
const deckSelectorForm = document.getElementById('deck-selector-form');
const deckSelectorOwnerList = document.getElementById('deck-selector-owner-list');
const deckSelectorWheel = document.getElementById('deck-selector-wheel');
const deckSelectorWheelDisc = document.getElementById('deck-selector-wheel-disc');
const deckSelectorWheelStatus = document.getElementById('deck-selector-wheel-status');
const deckSelectorResults = document.getElementById('deck-selector-results');
const deckSelectorSubmitButton = document.querySelector('#deck-selector-form button[type="submit"]');
const recordsForm = document.getElementById('records-form');
const recordsTableBody = document.getElementById('records-table-body');
const customRecordForm = document.getElementById('custom-record-form');
const customRecordTitleInput = document.getElementById('custom-record-title');
const customRecordUnitInput = document.getElementById('custom-record-unit');
const customRecordValueInput = document.getElementById('custom-record-value');
const customRecordHolderInput = document.getElementById('custom-record-holder');
const customRecordCommanderInput = document.getElementById('custom-record-commander');
const customRecordDateInput = document.getElementById('custom-record-date');
const customRecordNotesInput = document.getElementById('custom-record-notes');
const customRecordHolderMenu = document.getElementById('custom-record-holder-menu');
const customRecordCommanderMenu = document.getElementById('custom-record-commander-menu');
const liveGameForm = document.getElementById('live-game-form');
const liveGameDateInput = document.getElementById('live-game-date');
const liveStartingLifeInput = document.getElementById('live-starting-life');
const liveGamePlayerBody = document.getElementById('live-game-player-body');
const liveAddPlayerButton = document.getElementById('live-add-player');
const liveRemovePlayerButton = document.getElementById('live-remove-player');
const liveRandomizeFirstButton = document.getElementById('live-randomize-first');
const liveOrderPreview = document.getElementById('live-order-preview');
const liveActiveCard = document.getElementById('live-active-card');
const liveGameStatus = document.getElementById('live-game-status');
const livePlayerGrid = document.getElementById('live-player-grid');
const liveEventLog = document.getElementById('live-event-log');
const liveActiveActions = document.getElementById('live-active-actions');
const liveActionsToggleButton = document.getElementById('live-actions-toggle');
const liveBackButton = document.getElementById('live-back');
const liveUndoButton = document.getElementById('live-undo');
const liveSwapLifeButton = document.getElementById('live-swap-life');
const liveFinishGameButton = document.getElementById('live-finish-game');
const liveAbandonGameButton = document.getElementById('live-abandon-game');
const liveSourcePrompt = document.getElementById('live-source-prompt');
const liveSourceTitle = document.getElementById('live-source-title');
const liveSourceCopy = document.getElementById('live-source-copy');
const liveSourceOptions = document.getElementById('live-source-options');
const liveSourceCancelButton = document.getElementById('live-source-cancel');
const liveModalPrompt = document.getElementById('live-modal-prompt');
const liveModalTitle = document.getElementById('live-modal-title');
const liveModalCopy = document.getElementById('live-modal-copy');
const liveModalField = document.getElementById('live-modal-field');
const liveModalInput = document.getElementById('live-modal-input');
const liveModalTextarea = document.getElementById('live-modal-textarea');
const liveModalError = document.getElementById('live-modal-error');
const liveModalCancelButton = document.getElementById('live-modal-cancel');
const liveModalConfirmButton = document.getElementById('live-modal-confirm');

const syncUserInput = document.getElementById('sync-user');
const syncTokenInput = document.getElementById('sync-token');
const syncConnectButton = document.getElementById('sync-connect');
const syncDisconnectButton = document.getElementById('sync-disconnect');
const syncNowButton = document.getElementById('sync-now');
const syncStatus = document.getElementById('sync-status');
const syncPanel = document.querySelector('.sync-panel');

if (syncStatus) {
  syncStatus.setAttribute('role', 'status');
  syncStatus.setAttribute('aria-live', 'polite');
}

if (liveGameStatus) {
  liveGameStatus.setAttribute('role', 'status');
  liveGameStatus.setAttribute('aria-live', 'polite');
}

if (liveEventLog) {
  liveEventLog.setAttribute('role', 'log');
  liveEventLog.setAttribute('aria-live', 'polite');
  liveEventLog.setAttribute('aria-relevant', 'additions text');
}

if (deckLookupResult) {
  deckLookupResult.setAttribute('role', 'status');
  deckLookupResult.setAttribute('aria-live', 'polite');
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
  deckLists: { column: 'commander', descending: false },
};
let appState = { games: [], powerLevels: {}, deckLists: [], records: [] };
let editingDeckListId = null;
let syncQueueTimer = null;
let syncRetryTimer = null;
let syncInFlight = false;
let syncRetryCount = 0;
let syncPendingChanges = false;
let syncLastSuccessAt = null;
let syncConnectionState = 'local';
let syncLastErrorMessage = '';
let syncCloudRevision = 0;
let syncCloudUpdatedAt = '';
let syncCloudUpdatedBy = '';
let syncConflictInfo = null;
let syncMetadataCheckInFlight = false;
let syncLastFreshnessCheckAt = 0;
let historyQueryFiltersApplied = false;
let deckSelectorSpinTimer = null;
let deckSelectorRotation = 0;
let activeGameState = null;
let activeGameUndoState = [];
let liveSetupFirstPlayerId = null;
let liveSourcePromptResolver = null;
let liveModalResolver = null;
let liveModalConfig = null;
let liveHoldTimerId = null;
let liveHoldIntervalId = null;
let liveHoldRepeated = false;
let liveMeasurementTimerId = null;

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

function parseJsonSafe(value, fallback) {
  try {
    return JSON.parse(value);
  } catch (error) {
    return fallback;
  }
}

function normalizeIdentityLabel(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function getIdentityKey(value) {
  const normalizedValue = normalizeIdentityLabel(value);
  return normalizedValue ? normalizedValue.toLocaleLowerCase() : '';
}

function getIdentityDisplayScore(value) {
  const normalizedValue = normalizeIdentityLabel(value);
  if (!normalizedValue) {
    return 0;
  }

  let score = 0;
  if (normalizedValue !== normalizedValue.toLocaleLowerCase()) {
    score += 2;
  }
  if (/\b[A-Z]/.test(normalizedValue)) {
    score += 1;
  }
  return score;
}

function recordIdentityVariant(bucketMap, value) {
  const normalizedValue = normalizeIdentityLabel(value);
  const identityKey = getIdentityKey(normalizedValue);
  if (!identityKey) {
    return;
  }

  if (!bucketMap.has(identityKey)) {
    bucketMap.set(identityKey, new Map());
  }

  const variants = bucketMap.get(identityKey);
  variants.set(normalizedValue, (variants.get(normalizedValue) || 0) + 1);
}

function buildCanonicalIdentityMap(bucketMap) {
  const canonicalMap = new Map();

  bucketMap.forEach((variants, identityKey) => {
    let preferredValue = '';
    let preferredCount = -1;
    let preferredScore = -1;

    variants.forEach((count, value) => {
      const displayScore = getIdentityDisplayScore(value);
      const shouldReplace = count > preferredCount
        || (count === preferredCount && displayScore > preferredScore)
        || (count === preferredCount && displayScore === preferredScore && value.localeCompare(preferredValue) < 0);

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
  const normalizedValue = normalizeIdentityLabel(value);
  if (!normalizedValue) {
    return '';
  }

  const identityKey = getIdentityKey(normalizedValue);
  return canonicalMap?.get(identityKey) || normalizedValue;
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

function normalizeAppStateData(state) {
  const baseState = {
    games: Array.isArray(state?.games) ? state.games : [],
    powerLevels: state?.powerLevels && typeof state.powerLevels === 'object' ? state.powerLevels : {},
    deckLists: Array.isArray(state?.deckLists) ? state.deckLists : [],
    records: Array.isArray(state?.records) ? state.records : [],
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
    games: normalizeIdentityValues(baseState.games, identityMaps, 'games'),
    powerLevels: normalizedPowerLevels,
    deckLists: normalizeIdentityValues(baseState.deckLists, identityMaps, 'deckLists'),
    records: normalizeIdentityValues(baseState.records, identityMaps, 'records'),
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
  const normalizedValue = normalizeIdentityLabel(value);
  if (!normalizedValue) {
    return '';
  }

  return getIdentityKey(normalizedValue) === sourceKey ? replacementValue : normalizedValue;
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
  return getUniqueValues(players.map((value) => normalizeIdentityLabel(value)).filter(Boolean));
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
  return getUniqueValues(commanders.map((value) => normalizeIdentityLabel(value)).filter(Boolean));
}

function renderIdentityRenameOptions() {
  buildDatalistOptions(playerRenameDatalist, getKnownPlayerOptions());
  buildDatalistOptions(commanderRenameDatalist, getKnownCommanderOptions());
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
  const currentValue = normalizeIdentityLabel(currentInput?.value || '');
  const nextValue = normalizeIdentityLabel(nextInput?.value || '');

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

function loadLocalState() {
  const games = parseJsonSafe(localStorage.getItem(STORAGE_KEY) || '[]', []);
  const powerLevels = parseJsonSafe(localStorage.getItem(EXPECTED_POWER_STORAGE_KEY) || '{}', {});
  const deckLists = parseJsonSafe(localStorage.getItem(DECK_LIST_STORAGE_KEY) || '[]', []);
  const records = parseJsonSafe(localStorage.getItem(RECORDS_STORAGE_KEY) || '[]', []);
  const rawState = {
    games: Array.isArray(games) ? games : [],
    powerLevels: powerLevels && typeof powerLevels === 'object' ? powerLevels : {},
    deckLists: Array.isArray(deckLists) ? deckLists : [],
    records: Array.isArray(records) ? records : [],
  };
  const normalizedState = normalizeAppStateData(rawState);

  if (JSON.stringify(rawState) !== JSON.stringify(normalizedState)) {
    persistLocalState(normalizedState);
  }

  return normalizedState;
}

function persistLocalState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.games || []));
  localStorage.setItem(EXPECTED_POWER_STORAGE_KEY, JSON.stringify(state.powerLevels || {}));
  localStorage.setItem(DECK_LIST_STORAGE_KEY, JSON.stringify(state.deckLists || []));
  localStorage.setItem(RECORDS_STORAGE_KEY, JSON.stringify(state.records || []));
}

function loadActiveGameState() {
  const storedState = parseJsonSafe(localStorage.getItem(ACTIVE_GAME_STORAGE_KEY) || 'null', null);
  const normalizedState = normalizeActiveGameStateData(storedState);

  if (JSON.stringify(storedState) !== JSON.stringify(normalizedState)) {
    persistActiveGameState(normalizedState);
  }

  return normalizedState;
}

function loadActiveGameUndoState() {
  const storedState = parseJsonSafe(localStorage.getItem(ACTIVE_GAME_UNDO_STORAGE_KEY) || 'null', null);
  if (Array.isArray(storedState)) {
    return storedState.map((entry) => normalizeActiveGameStateData(entry)).filter(Boolean);
  }
  return storedState ? [normalizeActiveGameStateData(storedState)] : [];
}

function persistActiveGameUndoState(state) {
  const normalizedState = Array.isArray(state)
    ? state.map((entry) => normalizeActiveGameStateData(cloneActiveGameState(entry))).filter(Boolean)
    : state
      ? [normalizeActiveGameStateData(cloneActiveGameState(state))]
      : [];

  activeGameUndoState = normalizedState;
  if (!normalizedState.length) {
    localStorage.removeItem(ACTIVE_GAME_UNDO_STORAGE_KEY);
    return;
  }

  localStorage.setItem(ACTIVE_GAME_UNDO_STORAGE_KEY, JSON.stringify(normalizedState));
}

function cloneActiveGameState(state) {
  return state ? parseJsonSafe(JSON.stringify(state), null) : null;
}

function saveUndoSnapshot() {
  if (!activeGameState) {
    persistActiveGameUndoState(null);
    return;
  }

  persistActiveGameUndoState([
    ...activeGameUndoState,
    cloneActiveGameState(activeGameState),
  ]);
}

function persistActiveGameState(state) {
  const normalizedState = normalizeActiveGameStateData(state);
  activeGameState = normalizedState;
  if (!normalizedState) {
    localStorage.removeItem(ACTIVE_GAME_STORAGE_KEY);
    return;
  }

  localStorage.setItem(ACTIVE_GAME_STORAGE_KEY, JSON.stringify(normalizedState));
}

function getSyncCredentials() {
  return {
    user: (localStorage.getItem(SYNC_USER_STORAGE_KEY) || '').trim(),
    token: (localStorage.getItem(SYNC_TOKEN_STORAGE_KEY) || '').trim(),
  };
}

function hasSyncCredentials() {
  const credentials = getSyncCredentials();
  return Boolean(credentials.user && credentials.token);
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
}

function clearSyncConflict() {
  syncConflictInfo = null;
}

function hasNewerCloudRevision(metadata) {
  const revision = Number(metadata?.revision);
  return Number.isFinite(revision) && revision > syncCloudRevision;
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

  if (!credentials.user || !credentials.token) {
    return {
      tone: 'muted',
      message: 'Cloud sync not connected. Data stays local until you connect.',
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

  if (syncPendingChanges) {
    return {
      tone: 'neutral',
      message: 'Local changes are queued and will sync automatically.',
    };
  }

  if (syncConnectionState === 'connected') {
    return {
      tone: 'success',
      message: lastSyncedText
        ? `Cloud sync connected. Last synced at ${lastSyncedText}.`
        : `Cloud sync connected as ${credentials.user}.`,
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
  const isBusy = syncInFlight || syncConnectionState === 'connecting';
  if (syncConnectButton) {
    syncConnectButton.disabled = isBusy;
  }
  if (syncDisconnectButton) {
    syncDisconnectButton.disabled = !hasCredentials || isBusy;
  }
  if (syncNowButton) {
    syncNowButton.disabled = !hasCredentials || isBusy;
  }
}

async function cloudRequest(path, options = {}) {
  const credentials = getSyncCredentials();
  if (!credentials.user || !credentials.token) {
    throw new Error('Missing sync credentials');
  }

  const headers = new Headers(options.headers || {});
  headers.set('Content-Type', 'application/json');
  headers.set('X-User-Name', credentials.user);
  headers.set('X-Pod-Token', credentials.token);

  const response = await fetch(path, {
    ...options,
    headers,
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

  syncPendingChanges = false;
  clearSyncConflict();
  refreshSyncStatus();

  try {
    await pullCloudState();
    setSyncUiCollapsed(true);
    refreshSyncStatus();
  } catch (error) {
    syncConflictInfo = conflict || syncConflictInfo;
    syncLastErrorMessage = `${error.message}. Unable to pull the latest cloud state.`;
    refreshSyncStatus();
  }
}

async function pullCloudState() {
  syncConnectionState = 'connecting';
  syncLastErrorMessage = '';
  updateSyncControls();
  refreshSyncStatus();

  const payload = await cloudRequest(CLOUD_SYNC_ENDPOINT, { method: 'GET' });
  const games = Array.isArray(payload.games) ? payload.games : [];
  const powerLevels = payload.powerLevels && typeof payload.powerLevels === 'object' ? payload.powerLevels : {};
  const deckLists = Array.isArray(payload.deckLists) ? payload.deckLists : [];
  const records = Array.isArray(payload.records) ? payload.records : [];
  updateSyncMetadata({
    revision: payload.revision,
    updatedAt: payload.updatedAt,
    updatedBy: payload.updatedBy,
  });
  clearSyncConflict();
  appState = normalizeAppStateData({ games, powerLevels, deckLists, records });
  persistLocalState(appState);
  syncConnectionState = 'connected';
  syncLastErrorMessage = '';
  syncLastSuccessAt = new Date().toISOString();
  refresh();
}

async function fetchCloudStateMetadata() {
  const payload = await cloudRequest(CLOUD_SYNC_METADATA_ENDPOINT, { method: 'GET' });
  return {
    revision: Number.isFinite(Number(payload?.revision)) ? Number(payload.revision) : 0,
    updatedAt: String(payload?.updatedAt || '').trim(),
    updatedBy: String(payload?.updatedBy || '').trim(),
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

    if (syncPendingChanges) {
      clearSyncRetryTimer();
      if (syncQueueTimer) {
        clearTimeout(syncQueueTimer);
        syncQueueTimer = null;
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

async function pushCloudState() {
  if (!hasSyncCredentials() || syncInFlight || syncConflictInfo) {
    return;
  }

  syncInFlight = true;
  syncConnectionState = 'connected';
  syncLastErrorMessage = '';
  updateSyncControls();
  refreshSyncStatus();

  try {
    const payload = await cloudRequest(CLOUD_SYNC_ENDPOINT, {
      method: 'PUT',
      headers: {
        'X-State-Revision': String(syncCloudRevision),
      },
      body: JSON.stringify({
        games: appState.games,
        powerLevels: appState.powerLevels,
        deckLists: appState.deckLists,
        records: appState.records,
      }),
    });
    updateSyncMetadata({
      revision: payload.revision,
      updatedAt: payload.updatedAt,
      updatedBy: payload.updatedBy,
    });
    syncPendingChanges = false;
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
      await resolveSyncConflict(error.body?.conflict || null);
    } else {
      scheduleSyncRetry(error.message);
    }
  } finally {
    syncInFlight = false;
    updateSyncControls();
    refreshSyncStatus();
  }
}

function queueCloudSync(delay = 500) {
  if (!hasSyncCredentials()) {
    return;
  }

  syncPendingChanges = true;
  syncLastErrorMessage = '';
  refreshSyncStatus();

  if (syncConflictInfo) {
    return;
  }

  if (syncQueueTimer) {
    clearTimeout(syncQueueTimer);
  }

  clearSyncRetryTimer();

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

function buildHistoryFilterHref(filters = {}) {
  const query = new URLSearchParams();

  Object.entries(filters).forEach(([key, value]) => {
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

  if (tableKey === 'deckLists') {
    return false;
  }

  return !ascendingColumns.has(column);
}

function rerenderSortedTable(tableKey) {
  const games = loadGames();
  switch (tableKey) {
    case 'rankingsMain':
      renderPodRankings(games);
      break;
    case 'playerStats':
      renderPlayerStats(games);
      break;
    case 'commanderStats':
      renderCommanderStats(games);
      break;
    case 'recentPlayerTrends':
    case 'recentCommanderTrends':
    case 'playerStreaks':
    case 'commanderStreaks':
      renderRankingsPage(games);
      break;
    case 'deckLists':
      renderDeckLists();
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

function saveGames(games) {
  appState = normalizeAppStateData({
    ...appState,
    games: Array.isArray(games) ? games : [],
  });
  appState.records = mergeRecordsWithDefaults(appState.records, appState.games);
  persistLocalState(appState);
  queueCloudSync();
}

function loadDeckLists() {
  return Array.isArray(appState.deckLists) ? appState.deckLists : [];
}

function saveDeckLists(deckLists) {
  appState = normalizeAppStateData({
    ...appState,
    deckLists: Array.isArray(deckLists) ? deckLists : [],
  });
  persistLocalState(appState);
  queueCloudSync();
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

function getPlayerNameById(playerId, activeGame = activeGameState) {
  return activeGame?.players?.find((player) => player.id === playerId)?.name || 'Unknown player';
}

function getActiveAlivePlayers(activeGame = activeGameState) {
  return (activeGame?.players || []).filter((player) => !player.eliminatedAt);
}

function getCurrentTurnPlayer(activeGame = activeGameState) {
  if (!activeGame || !Array.isArray(activeGame.turnOrder) || !activeGame.turnOrder.length) {
    return null;
  }

  const currentPlayerId = activeGame.turnOrder[activeGame.currentTurnIndex] || '';
  return activeGame.players.find((player) => player.id === currentPlayerId) || null;
}

function getLiveMobileTablePlayerCount(activeGame = activeGameState) {
  if (!window.matchMedia('(max-width: 900px)').matches || !activeGame) {
    return 0;
  }

  const playerCount = (activeGame.players || []).length;
  return playerCount === 4 || playerCount === 5 ? playerCount : 0;
}

function isLiveMobileTableMode() {
  return getLiveMobileTablePlayerCount() > 0;
}

function updateLiveTableModeClass() {
  const mobileTablePlayerCount = getLiveMobileTablePlayerCount();
  document.body.classList.toggle('live-table-mode', mobileTablePlayerCount > 0);
  document.body.classList.toggle('live-table-mode-4', mobileTablePlayerCount === 4);
  document.body.classList.toggle('live-table-mode-5', mobileTablePlayerCount === 5);
}

function getLiveMobileSeatLayout(player, playerCount = getLiveMobileTablePlayerCount()) {
  const seat = Number.parseInt(`${player?.seat || 0}`, 10);
  if (playerCount === 5) {
    const orientationBySeat = {
      1: 'right',
      2: 'left',
      3: 'up',
      4: 'left',
      5: 'right',
    };

    return {
      seatClass: `live-seat-${seat}`,
      orientationClass: `live-orientation-${orientationBySeat[seat] || 'up'}`,
    };
  }

  return {
    seatClass: `live-seat-${seat}`,
    orientationClass: `live-orientation-${seat}`,
  };
}

function hideLiveSourcePrompt(selectedSourceId = null) {
  if (liveSourcePrompt) {
    liveSourcePrompt.hidden = true;
  }
  if (liveSourceOptions) {
    liveSourceOptions.innerHTML = '';
  }

  if (liveSourcePromptResolver) {
    const resolver = liveSourcePromptResolver;
    liveSourcePromptResolver = null;
    resolver(selectedSourceId);
  }
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

function promptForLiveSource({ targetPlayerId, title, description, allowSelf = true, requireOpponent = false }) {
  if (!liveSourcePrompt || !liveSourceOptions) {
    return Promise.resolve(null);
  }

  const targetPlayer = activeGameState?.players?.find((player) => player.id === targetPlayerId);
  if (!targetPlayer) {
    return Promise.resolve(null);
  }

  const options = [];
  if (allowSelf && !requireOpponent) {
    options.push({ id: targetPlayer.id, label: `${targetPlayer.name} (self)` });
  }

  activeGameState.players
    .filter((player) => player.id !== targetPlayer.id && !player.eliminatedAt)
    .forEach((player) => {
      options.push({ id: player.id, label: player.name });
    });

  liveSourceTitle.textContent = title;
  liveSourceCopy.textContent = description;
  liveSourceOptions.innerHTML = options
    .map((option) => `<button type="button" class="live-source-option" data-source-id="${escapeHtml(option.id)}">${escapeHtml(option.label)}</button>`)
    .join('');
  liveSourcePrompt.hidden = false;

  return new Promise((resolve) => {
    liveSourcePromptResolver = resolve;
  });
}

function promptForLivePlayerChoice({ title, description, excludePlayerIds = [] }) {
  if (!liveSourcePrompt || !liveSourceOptions || !activeGameState) {
    return Promise.resolve(null);
  }

  const excludedIds = new Set(excludePlayerIds.filter(Boolean));
  const options = activeGameState.players
    .filter((player) => !player.eliminatedAt && !excludedIds.has(player.id))
    .map((player) => ({ id: player.id, label: player.name }));

  if (!options.length) {
    return Promise.resolve(null);
  }

  liveSourceTitle.textContent = title;
  liveSourceCopy.textContent = description;
  liveSourceOptions.innerHTML = options
    .map((option) => `<button type="button" class="live-source-option" data-source-id="${escapeHtml(option.id)}">${escapeHtml(option.label)}</button>`)
    .join('');
  liveSourcePrompt.hidden = false;

  return new Promise((resolve) => {
    liveSourcePromptResolver = resolve;
  });
}

function shouldPromptForSource(targetPlayer, projectedLife, eventType, sourcePlayerId = '') {
  if (!activeGameState || !targetPlayer) {
    return false;
  }

  if (eventType === 'elimination' || eventType === 'commander-damage') {
    return true;
  }

  if (eventType === 'life-loss' && !activeGameState.firstBlood) {
    return true;
  }

  if (targetPlayer.cannotLoseTheGame) {
    return Boolean(activeGameState.shouldPromptForSource);
  }

  const currentCommanderDamage = sourcePlayerId ? (targetPlayer.commanderDamageTaken?.[sourcePlayerId] || 0) : 0;
  const commanderKill = eventType === 'commander-damage' && (currentCommanderDamage > 20);
  const lifeKill = projectedLife <= 0;
  return Boolean(activeGameState.shouldPromptForSource || lifeKill || commanderKill);
}

function maybeRecordFirstBlood(sourcePlayerId, targetPlayerId, turnNumber) {
  if (!activeGameState || activeGameState.firstBlood || !sourcePlayerId || sourcePlayerId === targetPlayerId) {
    return;
  }

  const sourcePlayer = activeGameState.players.find((player) => player.id === sourcePlayerId && !player.eliminatedAt);
  if (!sourcePlayer) {
    return;
  }

  activeGameState.firstBlood = {
    actorPlayerId: sourcePlayerId,
    targetPlayerId,
    turnNumber,
  };
}

function recordLiveEvent({ type, actorPlayerId = '', targetPlayerId = '', amount = 0, turnNumber, notes = '' }) {
  activeGameState.events.push({
    id: generateId(),
    type,
    actorPlayerId,
    targetPlayerId,
    amount,
    turnNumber,
    notes,
    timestamp: new Date().toISOString(),
  });
}

function getDuplicatePlayerName(rows) {
  const seenPlayers = new Map();

  for (const row of rows) {
    const playerName = String(row?.player || '').trim();
    if (!playerName) {
      continue;
    }

    const normalizedPlayerName = playerName.toLocaleLowerCase();
    if (seenPlayers.has(normalizedPlayerName)) {
      return playerName;
    }

    seenPlayers.set(normalizedPlayerName, playerName);
  }

  return '';
}

function normalizeEliminationSourcePlayerId(targetPlayerId, sourcePlayerId) {
  if (!activeGameState || !sourcePlayerId || sourcePlayerId === targetPlayerId) {
    return '';
  }

  const sourcePlayer = activeGameState.players.find((player) => player.id === sourcePlayerId && !player.eliminatedAt);
  return sourcePlayer ? sourcePlayer.id : '';
}

function eliminateLivePlayer(targetPlayer, sourcePlayerId, turnNumber, reason) {
  if (!targetPlayer || targetPlayer.eliminatedAt || targetPlayer.cannotLoseTheGame) {
    return false;
  }

  const aliveCount = getActiveAlivePlayers(activeGameState).length;
  const creditedSourcePlayerId = normalizeEliminationSourcePlayerId(targetPlayer.id, sourcePlayerId);
  targetPlayer.eliminatedAt = new Date().toISOString();
  targetPlayer.eliminatedTurnNumber = turnNumber;
  targetPlayer.eliminatedByPlayerId = creditedSourcePlayerId;
  targetPlayer.eliminationReason = reason;
  targetPlayer.place = aliveCount;

  if (creditedSourcePlayerId) {
    const sourcePlayer = activeGameState.players.find((player) => player.id === creditedSourcePlayerId);
    if (sourcePlayer) {
      sourcePlayer.kills += 1;
      sourcePlayer.killedPlayers = [...new Set([...(sourcePlayer.killedPlayers || []), targetPlayer.name])];
    }
  }

  activeGameState.shouldPromptForSource = true;
  return true;
}

function hasCommanderDamageLethal(player) {
  return Object.values(player?.commanderDamageTaken || {}).some((amount) => amount > 20);
}

function evaluateLiveElimination(targetPlayer, sourcePlayerId, turnNumber, eventType) {
  if (!targetPlayer || targetPlayer.cannotLoseTheGame) {
    return false;
  }

  if (targetPlayer.life <= 0) {
    return eliminateLivePlayer(targetPlayer, sourcePlayerId, turnNumber, 'life-total');
  }

  if (eventType === 'commander-damage' && sourcePlayerId && (targetPlayer.commanderDamageTaken?.[sourcePlayerId] || 0) > 20) {
    return eliminateLivePlayer(targetPlayer, sourcePlayerId, turnNumber, 'commander-damage');
  }

  return false;
}

async function resolveLiveSourceSelection({ targetPlayerId, eventType, amount, projectedLife }) {
  const targetPlayer = activeGameState?.players?.find((player) => player.id === targetPlayerId);
  if (!targetPlayer) {
    return null;
  }

  const needsPrompt = shouldPromptForSource(targetPlayer, projectedLife, eventType);
  if (!needsPrompt) {
    return '';
  }

  const title = eventType === 'elimination' ? 'Select the killer' : 'Select damage source';
  const description = eventType === 'commander-damage'
    ? `Who dealt ${amount} commander damage to ${targetPlayer.name}?`
    : eventType === 'elimination'
      ? `Who eliminated ${targetPlayer.name}? Choose self for no kill credit.`
      : `Who caused ${targetPlayer.name} to lose ${amount} life?`;
  const requireOpponent = eventType === 'commander-damage';
  const selectedSourceId = await promptForLiveSource({
    targetPlayerId,
    title,
    description,
    allowSelf: true,
    requireOpponent,
  });

  if (selectedSourceId === null) {
    return null;
  }

  if (eventType !== 'elimination') {
    const stillNeedsFirstBlood = eventType === 'life-loss'
      && !activeGameState.firstBlood
      && (!selectedSourceId || selectedSourceId === targetPlayerId);
    activeGameState.shouldPromptForSource = stillNeedsFirstBlood || (projectedLife <= 0 && !targetPlayer.cannotLoseTheGame);
  }

  return selectedSourceId || '';
}

async function swapLivePlayerLifeTotals() {
  if (!activeGameState) {
    return;
  }

  const eligiblePlayers = activeGameState.players.filter((player) => !player.eliminatedAt);
  if (eligiblePlayers.length < 2) {
    await promptLiveAlert('At least two active players are required to swap life totals.', 'Unable to swap life');
    return;
  }

  const firstPlayerId = await promptForLivePlayerChoice({
    title: 'Select first player',
    description: 'Choose the first player for the life-total swap.',
  });
  if (!firstPlayerId) {
    return;
  }

  const firstPlayer = activeGameState.players.find((player) => player.id === firstPlayerId && !player.eliminatedAt);
  if (!firstPlayer) {
    return;
  }

  const secondPlayerId = await promptForLivePlayerChoice({
    title: 'Select second player',
    description: `Choose who should swap life totals with ${firstPlayer.name}.`,
    excludePlayerIds: [firstPlayerId],
  });
  if (!secondPlayerId) {
    return;
  }

  const secondPlayer = activeGameState.players.find((player) => player.id === secondPlayerId && !player.eliminatedAt);
  if (!secondPlayer) {
    return;
  }

  const firstLifeBeforeSwap = firstPlayer.life;
  const secondLifeBeforeSwap = secondPlayer.life;

  saveUndoSnapshot();
  firstPlayer.life = secondLifeBeforeSwap;
  secondPlayer.life = firstLifeBeforeSwap;

  recordLiveEvent({
    type: 'life-swap',
    actorPlayerId: firstPlayer.id,
    targetPlayerId: secondPlayer.id,
    amount: 0,
    turnNumber: getLiveTrackedTurnNumber(),
    notes: `${firstPlayer.name} (${firstLifeBeforeSwap}) swapped life totals with ${secondPlayer.name} (${secondLifeBeforeSwap}).`,
  });

  persistActiveGameState(activeGameState);
  refreshLiveTrackerUi();
}

function stopLiveHoldRepeat() {
  if (liveHoldTimerId) {
    clearTimeout(liveHoldTimerId);
    liveHoldTimerId = null;
  }
  if (liveHoldIntervalId) {
    clearInterval(liveHoldIntervalId);
    liveHoldIntervalId = null;
  }
}

function startLiveHoldRepeat(button) {
  stopLiveHoldRepeat();

  const playerId = button.dataset.playerId || '';
  const delta = parseInt(button.dataset.delta || '0', 10);
  const player = activeGameState?.players?.find((entry) => entry.id === playerId);
  if (!player || Number.isNaN(delta) || delta === 0) {
    return;
  }

  if (delta < 0 && shouldPromptForSource(player, player.life + delta, 'life-loss')) {
    return;
  }

  liveHoldRepeated = false;
  liveHoldTimerId = setTimeout(() => {
    liveHoldRepeated = true;
    applyQuickLifeChange(playerId, delta);
    liveHoldIntervalId = setInterval(() => {
      applyQuickLifeChange(playerId, delta);
    }, 260);
  }, 350);
}

function getLiveTrackedTurnNumber() {
  const rawValue = parseInt(`${activeGameState?.turnNumber || 1}`, 10);
  return Math.max(1, Number.isNaN(rawValue) ? 1 : rawValue);
}

function syncActiveGameTurnFromInput() {
  const turnNumber = getLiveTrackedTurnNumber();
  if (!activeGameState) {
    return turnNumber;
  }

  activeGameState.turnNumber = turnNumber;
  persistActiveGameState(activeGameState);
  return turnNumber;
}

async function promptForTurnNumber(message, defaultValue = getLiveTrackedTurnNumber()) {
  const parsed = await promptLiveText(message, {
    title: 'Turn number',
    confirmLabel: 'Save',
    defaultValue: String(defaultValue),
    inputType: 'number',
    inputMode: 'numeric',
    validate: (rawValue) => {
      const result = parseLiveModalInteger(rawValue, 'Enter a whole number greater than 0.');
      if (result.error || result.value < 1) {
        return { error: 'Enter a whole number greater than 0.' };
      }
      return { value: result.value };
    },
  });
  if (parsed === null) {
    return null;
  }

  if (activeGameState) {
    activeGameState.turnNumber = parsed;
  }
  return parsed;
}

function getTurnOrderPreview(rows, firstPlayerId) {
  const normalizedRows = Array.isArray(rows) ? rows.filter((row) => row.player) : [];
  if (!normalizedRows.length) {
    return [];
  }

  const firstIndex = normalizedRows.findIndex((row) => row.id === firstPlayerId);
  if (firstIndex < 0) {
    return normalizedRows;
  }

  return [
    ...normalizedRows.slice(firstIndex),
    ...normalizedRows.slice(0, firstIndex),
  ];
}

function createLiveSetupRow(data = {}) {
  const row = document.createElement('tr');
  row.dataset.playerId = data.id || generateId();
  row.innerHTML = `
    <td><span class="live-seat-badge"></span></td>
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
  `;

  populateRowSelectors(row);
  attachLookupWrapperHandlers(row);
  return row;
}

function updateLiveSetupSeatLabels() {
  if (!liveGamePlayerBody) {
    return;
  }

  Array.from(liveGamePlayerBody.querySelectorAll('tr')).forEach((row, index) => {
    const badge = row.querySelector('.live-seat-badge');
    if (badge) {
      badge.textContent = `Seat ${index + 1}`;
    }
  });
}

function addLiveSetupRow(data = {}) {
  if (!liveGamePlayerBody) {
    return;
  }

  liveGamePlayerBody.appendChild(createLiveSetupRow(data));
  updateLiveSetupSeatLabels();
}

function removeLiveSetupRow() {
  if (!liveGamePlayerBody) {
    return;
  }

  const rows = liveGamePlayerBody.querySelectorAll('tr');
  if (rows.length > 2) {
    rows[rows.length - 1].remove();
  }
  updateLiveSetupSeatLabels();
}

function getLiveSetupRows() {
  if (!liveGamePlayerBody) {
    return [];
  }

  const playerMap = buildCanonicalIdentityMapFromValues(knownPlayers);
  const commanderMap = buildCanonicalIdentityMapFromValues(knownCommanders);

  return Array.from(liveGamePlayerBody.querySelectorAll('tr'))
    .map((row, index) => ({
      id: row.dataset.playerId || generateId(),
      player: canonicalizeIdentityValue(row.querySelector('[name="player"]')?.value || '', playerMap),
      commander: canonicalizeIdentityValue(row.querySelector('[name="commander"]')?.value || '', commanderMap),
      seat: index + 1,
    }))
    .filter((row) => row.player);
}

function renderLiveOrderPreview() {
  if (!liveOrderPreview) {
    return;
  }

  const rows = getLiveSetupRows();
  if (!rows.length) {
    liveOrderPreview.textContent = 'Add players to preview the randomized turn order.';
    return;
  }

  if (!liveSetupFirstPlayerId || !rows.some((row) => row.id === liveSetupFirstPlayerId)) {
    liveSetupFirstPlayerId = rows[0].id;
  }

  const orderedRows = getTurnOrderPreview(rows, liveSetupFirstPlayerId);
  liveOrderPreview.textContent = `First player: ${orderedRows[0]?.player || '—'}. Turn order: ${orderedRows.map((row) => row.player).join(' → ')}.`;
}

function randomizeLiveFirstPlayer() {
  const rows = getLiveSetupRows();
  if (!rows.length) {
    renderLiveOrderPreview();
    return;
  }

  const randomIndex = Math.floor(Math.random() * rows.length);
  liveSetupFirstPlayerId = rows[randomIndex].id;
  renderLiveOrderPreview();
}

function generateLiveEventDescription(event, activeGame = activeGameState) {
  const actor = event.actorPlayerId ? getPlayerNameById(event.actorPlayerId, activeGame) : 'Environment';
  const target = event.targetPlayerId ? getPlayerNameById(event.targetPlayerId, activeGame) : 'Unknown player';

  if (event.type === 'life-loss') {
    return `${target} lost ${event.amount} life${event.actorPlayerId ? ` from ${actor}` : ''}.`;
  }
  if (event.type === 'life-gain') {
    return `${target} gained ${event.amount} life.`;
  }
  if (event.type === 'commander-damage') {
    return `${actor} dealt ${event.amount} commander damage to ${target}.`;
  }
  if (event.type === 'elimination') {
    return `${actor} eliminated ${target}.`;
  }
  if (event.type === 'automatic-win') {
    return `${target} was marked as the winner.`;
  }
  if (event.type === 'life-swap') {
    return `${actor} and ${target} swapped life totals.`;
  }
  if (event.type === 'next-turn') {
    return `${target} started turn ${event.turnNumber}.`;
  }
  return 'Event recorded.';
}

function renderLivePlayerGrid() {
  if (!livePlayerGrid) {
    return;
  }

  if (!activeGameState) {
    livePlayerGrid.innerHTML = '<p>No game in progress.</p>';
    return;
  }

  const mobileTablePlayerCount = getLiveMobileTablePlayerCount();
  const isMobileTable = mobileTablePlayerCount > 0;
  const playersToRender = activeGameState.players;

  livePlayerGrid.innerHTML = playersToRender
    .slice()
    .sort((a, b) => a.seat - b.seat)
    .map((player) => {
      const playerName = escapeHtml(player.name);
      const playerCommander = escapeHtml(player.commander || 'No commander');
      const damageEntries = Object.entries(player.commanderDamageTaken || {}).filter(([, amount]) => amount > 0);
      const damageListMarkup = `<ul class="live-card-damage-list">${damageEntries.map(([sourceId, amount]) => `<li>${escapeHtml(getPlayerNameById(sourceId, activeGameState))}: <strong>${amount}</strong></li>`).join('')}</ul>`;
      const damageMarkup = damageEntries.length
        ? `
            <div class="live-player-damage-panel has-damage">
              <div class="live-player-damage-title">Cmdr Dmg</div>
              ${damageListMarkup}
            </div>`
        : '';
      const desktopDamageMarkup = damageEntries.length
        ? `
            <div class="live-player-damage-inline">
              <div>Commander damage received:</div>
              ${damageListMarkup}
            </div>`
        : '';
      const firstPlayerMarkup = player.id === activeGameState.startingPlayerId
        ? '<span class="live-first-player-indicator">First</span>'
        : '';
      const mobileSeatLayout = getLiveMobileSeatLayout(player, mobileTablePlayerCount);

      if (!isMobileTable) {
        return `
          <article class="live-player-card${player.eliminatedAt ? ' is-eliminated' : ''}" aria-label="${playerName}, ${playerCommander}">
            <div class="live-player-card-header">
              <div>
                <h3>${playerName}</h3>
                <p>${playerCommander}</p>
              </div>
              <div class="live-player-header-badges">
                ${firstPlayerMarkup}
              </div>
            </div>
            <div class="live-player-life">
              <button type="button" class="live-player-life-button live-player-life-button-standard" data-action="manual-life-entry" data-player-id="${escapeHtml(player.id)}" aria-label="Set life total for ${playerName}. Current life ${player.life}.">${player.life}</button>
            </div>
            <div class="live-quick-actions live-quick-actions-standard">
              <button type="button" class="live-quick-action is-negative" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="-1" aria-label="Subtract 1 life from ${playerName}">-1</button>
              <button type="button" class="live-quick-action is-positive" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="1" aria-label="Add 1 life to ${playerName}">+1</button>
              <button type="button" class="live-quick-action" data-action="manual-commander-damage" data-player-id="${escapeHtml(player.id)}" aria-label="Add commander damage to ${playerName}">Cmdr</button>
              <button type="button" class="live-quick-action" data-action="manual-eliminate" data-player-id="${escapeHtml(player.id)}" aria-label="Mark ${playerName} out of the game">Out</button>
              <button type="button" class="live-quick-action" data-action="auto-win" data-player-id="${escapeHtml(player.id)}" aria-label="Mark ${playerName} as the winner">Win</button>
            </div>
            <div class="live-player-meta live-player-meta-standard">
              <label class="live-player-toggle">
                <input type="checkbox" data-action="toggle-cannot-lose" data-player-id="${escapeHtml(player.id)}" aria-label="${playerName} cannot lose the game"${player.cannotLoseTheGame ? ' checked' : ''} />
                <span>Cannot lose the game</span>
              </label>
              ${desktopDamageMarkup}
            </div>
          </article>`;
      }

      return `
        <article class="live-player-card ${mobileSeatLayout.seatClass}${player.eliminatedAt ? ' is-eliminated' : ''}" aria-label="${playerName}, ${playerCommander}">
          <div class="live-player-card-body ${mobileSeatLayout.orientationClass}">
            <div class="live-player-card-topline">
              <div class="live-player-title-block">
                <h3>${playerName}</h3>
                <p>${playerCommander}</p>
              </div>
              ${firstPlayerMarkup}
            </div>
            <div class="live-player-main-layout">
              <div class="live-player-action-column live-player-action-column-loss">
                <button type="button" class="live-quick-action is-negative" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="-1" aria-label="Subtract 1 life from ${playerName}">-1</button>
                <button type="button" class="live-quick-action is-positive" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="1" aria-label="Add 1 life to ${playerName}">+1</button>
                <button type="button" class="live-quick-action" data-action="manual-commander-damage" data-player-id="${escapeHtml(player.id)}" aria-label="Add commander damage to ${playerName}">Cmdr</button>
                <button type="button" class="live-quick-action" data-action="manual-eliminate" data-player-id="${escapeHtml(player.id)}" aria-label="Mark ${playerName} out of the game">Out</button>
                <button type="button" class="live-quick-action" data-action="auto-win" data-player-id="${escapeHtml(player.id)}" aria-label="Mark ${playerName} as the winner">Win</button>
                <label class="live-player-toggle live-player-toggle-compact">
                  <input type="checkbox" data-action="toggle-cannot-lose" data-player-id="${escapeHtml(player.id)}" aria-label="${playerName} cannot lose the game"${player.cannotLoseTheGame ? ' checked' : ''} />
                  <span>No<br />lose</span>
                </label>
              </div>
              <div class="live-player-counter-column">
                <button type="button" class="live-player-life live-player-life-button" data-action="manual-life-entry" data-player-id="${escapeHtml(player.id)}" aria-label="Set life total for ${playerName}. Current life ${player.life}.">${player.life}</button>
              </div>
              <div class="live-player-damage-column">
                ${damageMarkup}
              </div>
            </div>
          </div>
        </article>`;
    })
    .join('');

  updateLivePlayerCardMeasurements();
}

function updateLivePlayerCardMeasurements() {
  if (!livePlayerGrid) {
    return;
  }

  const cards = livePlayerGrid.querySelectorAll('.live-player-card');
  if (!cards.length) {
    return;
  }

  const measureCards = () => {
    cards.forEach((card) => {
      const rect = card.getBoundingClientRect();
      card.style.setProperty('--live-card-width', `${rect.width}px`);
      card.style.setProperty('--live-card-height', `${rect.height}px`);
    });
  };

  measureCards();
  requestAnimationFrame(measureCards);

  if (liveMeasurementTimerId) {
    window.clearTimeout(liveMeasurementTimerId);
  }
  liveMeasurementTimerId = window.setTimeout(measureCards, 120);
}

function renderLiveEventLog() {
  if (!liveEventLog) {
    return;
  }

  if (!activeGameState?.events?.length) {
    liveEventLog.innerHTML = '<p>No events yet.</p>';
    return;
  }

  liveEventLog.innerHTML = activeGameState.events
    .slice()
    .reverse()
    .slice(0, 10)
    .map((event) => `
      <div class="live-event-item">
        <p>${escapeHtml(generateLiveEventDescription(event, activeGameState))}</p>
        <small>Turn ${event.turnNumber || activeGameState.turnNumber} · ${new Date(event.timestamp).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}</small>
      </div>`)
    .join('');
}

function renderLiveGameStatus() {
  if (!liveActiveCard) {
    return;
  }

  updateLiveTableModeClass();

  if (!activeGameState) {
    if (liveGameStatus) {
      liveGameStatus.textContent = 'No game in progress.';
    }
    if (liveUndoButton) {
      liveUndoButton.disabled = true;
    }
    renderLivePlayerGrid();
    renderLiveEventLog();
    return;
  }

  const elapsedMinutes = Math.max(1, Math.round((Date.now() - new Date(activeGameState.startedAt).getTime()) / 60000));

  if (liveGameStatus) {
    liveGameStatus.textContent = `${getActiveAlivePlayers(activeGameState).length} players alive · ${elapsedMinutes} minute${elapsedMinutes === 1 ? '' : 's'} elapsed.`;
  }
  if (liveUndoButton) {
    liveUndoButton.disabled = !activeGameUndoState.length;
  }
  renderLivePlayerGrid();
  renderLiveEventLog();
}

function refreshLiveTrackerUi() {
  if (!liveSourcePromptResolver) {
    hideLiveSourcePrompt();
  }
  updateLiveTableModeClass();
  renderLiveOrderPreview();
  renderLiveGameStatus();
}

function buildActiveGameSummary(gameState) {
  const killLeader = gameState.players.slice().sort((a, b) => b.kills - a.kills)[0];
  const finishOrderSummary = gameState.players
    .slice()
    .sort((a, b) => (a.place || 999) - (b.place || 999))
    .map((player) => `${player.place || '—'}. ${player.name}`)
    .join(' | ');
  const firstBloodSummary = gameState.firstBlood
    ? `${gameState.firstBlood.actorPlayerId ? getPlayerNameById(gameState.firstBlood.actorPlayerId, gameState) : 'Environment'} drew first blood on ${getPlayerNameById(gameState.firstBlood.targetPlayerId, gameState)} on turn ${gameState.firstBlood.turnNumber}`
    : 'No first blood was recorded';
  const alternateWinSummary = gameState.alternateWinCondition
    ? `Alternate win: ${gameState.alternateWinCondition}`
    : '';
  const alternateLoseSummary = gameState.players
    .filter((player) => player.eliminationDetails)
    .map((player) => `${player.name} lost via ${player.eliminationDetails}`)
    .join(' | ');

  return [
    `First player: ${getPlayerNameById(gameState.startingPlayerId, gameState)}`,
    firstBloodSummary,
    alternateWinSummary,
    alternateLoseSummary,
    `Kill leader: ${killLeader ? `${killLeader.name} (${killLeader.kills})` : 'None'}`,
    `Finish order: ${finishOrderSummary}`,
  ].filter(Boolean).join(' · ');
}

function buildLiveGameRecordStats(gameState, finalPlayers, durationMinutes) {
  const playersById = new Map((gameState.players || []).map((player) => [player.id, player]));
  const runningLifeTotals = new Map((gameState.players || []).map((player) => [player.id, player.startingLife]));
  const commanderDamageByActor = new Map();
  let highestSingleHit = null;
  let highestLifeLossHit = null;
  let highestLifeTotal = null;
  let fastestElimination = null;

  const createLiveCandidate = (value, holderId, notes = '') => createRecordCandidate({
    value,
    holder: holderId ? getPlayerNameById(holderId, gameState) : '',
    commander: holderId ? playersById.get(holderId)?.commander || '' : '',
    date: gameState.date,
    notes,
  });

  (gameState.players || []).forEach((player) => {
    const candidate = createLiveCandidate(player.startingLife, player.id, 'Starting life total.');
    highestLifeTotal = chooseBetterRecordCandidate(highestLifeTotal, candidate, 'highest-life-total');
  });

  (gameState.events || []).forEach((event) => {
    const targetId = event.targetPlayerId || '';
    const actorId = event.actorPlayerId || '';
    const amount = typeof event.amount === 'number' ? event.amount : Number.parseFloat(event.amount || '0');
    if (!Number.isFinite(amount) || amount <= 0) {
      return;
    }

    if ((event.type === 'life-loss' || event.type === 'commander-damage') && actorId) {
      const targetName = targetId ? getPlayerNameById(targetId, gameState) : 'Unknown player';
      const notes = `Hit ${targetName}${event.type === 'commander-damage' ? ' with commander damage' : ''}.`;
      highestSingleHit = chooseBetterRecordCandidate(highestSingleHit, createLiveCandidate(amount, actorId, notes), 'highest-damage-dealt');
    }

    if (event.type === 'life-loss' && actorId) {
      const targetName = targetId ? getPlayerNameById(targetId, gameState) : 'Unknown player';
      highestLifeLossHit = chooseBetterRecordCandidate(highestLifeLossHit, createLiveCandidate(amount, actorId, `Hit ${targetName} to life.`), 'highest-hit-to-life');
    }

    if (event.type === 'commander-damage' && actorId) {
      commanderDamageByActor.set(actorId, (commanderDamageByActor.get(actorId) || 0) + amount);
    }

    if (targetId && runningLifeTotals.has(targetId)) {
      const currentLife = runningLifeTotals.get(targetId) || 0;
      const updatedLife = event.type === 'life-gain'
        ? currentLife + amount
        : (event.type === 'life-loss' || event.type === 'commander-damage')
          ? currentLife - amount
          : currentLife;
      runningLifeTotals.set(targetId, updatedLife);

      const lifeCandidate = createLiveCandidate(updatedLife, targetId, 'Highest life reached in a live game.');
      highestLifeTotal = chooseBetterRecordCandidate(highestLifeTotal, lifeCandidate, 'highest-life-total');
    }
  });

  let mostCommanderDamage = null;
  commanderDamageByActor.forEach((totalDamage, actorId) => {
    mostCommanderDamage = chooseBetterRecordCandidate(
      mostCommanderDamage,
      createLiveCandidate(totalDamage, actorId, 'Total commander damage dealt in one live game.'),
      'most-commander-damage',
    );
  });

  finalPlayers.forEach((player) => {
    if (typeof player.eliminatedTurnNumber === 'number' && player.eliminatedTurnNumber > 0) {
      fastestElimination = chooseBetterRecordCandidate(
        fastestElimination,
        createRecordCandidate({
          value: player.eliminatedTurnNumber,
          holder: player.eliminatedByPlayerId ? getPlayerNameById(player.eliminatedByPlayerId, gameState) : '',
          commander: player.eliminatedByPlayerId ? playersById.get(player.eliminatedByPlayerId)?.commander || '' : '',
          date: gameState.date,
          notes: `Eliminated ${player.name}.`,
        }),
        'fastest-elimination',
      );
    }
  });

  return {
    durationMinutes,
    highestSingleHit,
    highestLifeLossHit,
    highestLifeTotal,
    mostCommanderDamage,
    fastestElimination,
  };
}

function closeLiveActionsMenu() {
  if (!liveActiveActions) {
    return;
  }

  liveActiveActions.classList.remove('is-open');
  if (liveActionsToggleButton) {
    liveActionsToggleButton.setAttribute('aria-expanded', 'false');
  }
}

function toggleLiveActionsMenu() {
  if (!liveActiveActions) {
    return;
  }

  const nextOpen = !liveActiveActions.classList.contains('is-open');
  liveActiveActions.classList.toggle('is-open', nextOpen);
  if (liveActionsToggleButton) {
    liveActionsToggleButton.setAttribute('aria-expanded', String(nextOpen));
  }
}

async function startLiveGame() {
  const players = getLiveSetupRows();
  if (players.length < 2) {
    await promptLiveAlert('Add at least two players to start a live game.', 'Unable to start game');
    return;
  }

  const duplicatePlayerName = getDuplicatePlayerName(players);
  if (duplicatePlayerName) {
    await promptLiveAlert(`Each player can only appear once per match. Duplicate player: ${duplicatePlayerName}.`, 'Unable to start game');
    return;
  }

  const startingLife = parseInt(liveStartingLifeInput?.value || '40', 10);
  if (!startingLife || startingLife < 1) {
    await promptLiveAlert('Enter a valid starting life total.', 'Unable to start game');
    return;
  }

  if (!liveSetupFirstPlayerId || !players.some((player) => player.id === liveSetupFirstPlayerId)) {
    liveSetupFirstPlayerId = players[Math.floor(Math.random() * players.length)].id;
  }

  const turnOrder = getTurnOrderPreview(players, liveSetupFirstPlayerId).map((player) => player.id);
  persistActiveGameUndoState(null);
  persistActiveGameState({
    id: generateId(),
    date: liveGameDateInput?.value || new Date().toISOString().slice(0, 10),
    startedAt: new Date().toISOString(),
    turnNumber: 1,
    alternateWinCondition: '',
    startingPlayerId: liveSetupFirstPlayerId,
    turnOrder,
    firstBlood: null,
    shouldPromptForSource: true,
    events: [],
    players: players.map((player) => ({
      id: player.id,
      name: player.player,
      commander: player.commander,
      seat: player.seat,
      startingLife,
      life: startingLife,
      kills: 0,
      killedPlayers: [],
      commanderDamageTaken: {},
      cannotLoseTheGame: false,
      eliminatedAt: '',
      eliminatedByPlayerId: '',
      eliminationDetails: '',
      place: null,
    })),
  });
  closeLiveActionsMenu();
  refreshLiveTrackerUi();
}

async function promptForPositiveNumber(message, defaultValue = 1) {
  return promptLiveText(message, {
    title: 'Enter amount',
    confirmLabel: 'Apply',
    defaultValue: String(defaultValue),
    inputType: 'number',
    inputMode: 'numeric',
    validate: (rawValue) => {
      const result = parseLiveModalInteger(rawValue, 'Enter a whole number greater than 0.');
      if (result.error || result.value < 1) {
        return { error: 'Enter a whole number greater than 0.' };
      }
      return { value: result.value };
    },
  });
}

async function promptForSignedNumber(message, defaultValue = -1) {
  return promptLiveText(message, {
    title: 'Enter life change',
    confirmLabel: 'Apply',
    defaultValue: String(defaultValue),
    inputType: 'number',
    inputMode: 'numeric',
    validate: (rawValue) => {
      const result = parseLiveModalInteger(rawValue, 'Enter a whole number greater than 0 or less than 0.');
      if (result.error || result.value === 0) {
        return { error: 'Enter a whole number greater than 0 or less than 0.' };
      }
      return { value: result.value };
    },
  });
}

async function applyCommanderDamageToPlayer(targetPlayerId) {
  if (!activeGameState) {
    return;
  }

  const targetPlayer = activeGameState.players.find((player) => player.id === targetPlayerId);
  if (!targetPlayer || targetPlayer.eliminatedAt) {
    return;
  }

  const amount = await promptForPositiveNumber(`How much commander damage did ${targetPlayer.name} take?`, 1);
  if (amount === null) {
    return;
  }

  const sourcePlayerId = await resolveLiveSourceSelection({
    targetPlayerId,
    eventType: 'commander-damage',
    amount,
    projectedLife: targetPlayer.life - amount,
  });
  if (sourcePlayerId === null) {
    return;
  }

  const projectedLife = targetPlayer.life - amount;
  const projectedCommanderDamage = (targetPlayer.commanderDamageTaken[sourcePlayerId] || 0) + amount;
  const needsFirstBloodTurn = Boolean(sourcePlayerId && sourcePlayerId !== targetPlayerId && !activeGameState.firstBlood);
  const needsKillTurn = !targetPlayer.cannotLoseTheGame && (projectedLife <= 0 || projectedCommanderDamage > 20);
  let eventTurnNumber = getLiveTrackedTurnNumber();
  if (needsFirstBloodTurn || needsKillTurn) {
    const promptedTurn = await promptForTurnNumber(`What turn was this commander damage on?`, eventTurnNumber);
    if (promptedTurn === null) {
      return;
    }
    eventTurnNumber = promptedTurn;
  }

  saveUndoSnapshot();
  targetPlayer.life -= amount;
  targetPlayer.commanderDamageTaken[sourcePlayerId] = (targetPlayer.commanderDamageTaken[sourcePlayerId] || 0) + amount;
  maybeRecordFirstBlood(sourcePlayerId, targetPlayerId, eventTurnNumber);
  evaluateLiveElimination(targetPlayer, sourcePlayerId, eventTurnNumber, 'commander-damage');
  recordLiveEvent({
    type: 'commander-damage',
    actorPlayerId: sourcePlayerId,
    targetPlayerId,
    amount,
    turnNumber: eventTurnNumber,
  });

  const alivePlayers = getActiveAlivePlayers(activeGameState);
  if (alivePlayers.length === 1) {
    alivePlayers[0].place = 1;
  }

  persistActiveGameState(activeGameState);
  refreshLiveTrackerUi();

  if (alivePlayers.length === 1 && await promptLiveConfirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`, {
    title: 'Finish game?',
    confirmLabel: 'Finish and save',
  })) {
    await completeActiveGame();
  }
}

async function manuallyEliminatePlayer(targetPlayerId) {
  if (!activeGameState) {
    return;
  }

  const targetPlayer = activeGameState.players.find((player) => player.id === targetPlayerId);
  if (!targetPlayer || targetPlayer.eliminatedAt) {
    return;
  }

  const sourcePlayerId = await resolveLiveSourceSelection({
    targetPlayerId,
    eventType: 'elimination',
    amount: 0,
    projectedLife: targetPlayer.life,
  });
  if (sourcePlayerId === null) {
    return;
  }

  const eventTurnNumber = await promptForTurnNumber(`What turn was ${targetPlayer.name} eliminated on?`);
  if (eventTurnNumber === null) {
    return;
  }

  const eliminationDetails = await promptLiveText(`Enter ${targetPlayer.name}'s alternate lose condition for the notes. Leave blank for a normal elimination.`, {
    title: 'Alternate lose condition',
    confirmLabel: 'Save',
    defaultValue: targetPlayer.eliminationDetails || '',
    placeholder: 'Optional details',
  });
  if (eliminationDetails === null) {
    return;
  }

  saveUndoSnapshot();
  targetPlayer.eliminationDetails = eliminationDetails.trim();
  eliminateLivePlayer(targetPlayer, sourcePlayerId, eventTurnNumber, 'manual');
  recordLiveEvent({
    type: 'elimination',
    actorPlayerId: sourcePlayerId,
    targetPlayerId,
    amount: 0,
    turnNumber: eventTurnNumber,
    notes: targetPlayer.eliminationDetails ? `Alternate lose: ${targetPlayer.eliminationDetails}` : '',
  });

  const alivePlayers = getActiveAlivePlayers(activeGameState);
  if (alivePlayers.length === 1) {
    alivePlayers[0].place = 1;
  }

  persistActiveGameState(activeGameState);
  refreshLiveTrackerUi();

  if (alivePlayers.length === 1 && await promptLiveConfirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`, {
    title: 'Finish game?',
    confirmLabel: 'Finish and save',
  })) {
    await completeActiveGame();
  }
}

async function applyQuickLifeChange(playerId, delta) {
  if (!activeGameState) {
    return;
  }

  const player = activeGameState.players.find((entry) => entry.id === playerId);
  if (!player || player.eliminatedAt) {
    return;
  }

  const projectedLife = player.life + delta;
  const sourcePlayerId = delta < 0
    ? await resolveLiveSourceSelection({
      targetPlayerId: playerId,
      eventType: 'life-loss',
      amount: Math.abs(delta),
      projectedLife,
    })
    : '';
  if (sourcePlayerId === null) {
    return;
  }

  const needsFirstBloodTurn = delta < 0 && Boolean(sourcePlayerId && sourcePlayerId !== playerId && !activeGameState.firstBlood);
  const needsKillTurn = delta < 0 && !player.cannotLoseTheGame && projectedLife <= 0;
  let turnNumber = getLiveTrackedTurnNumber();
  if (needsFirstBloodTurn || needsKillTurn) {
    const promptedTurn = await promptForTurnNumber(`What turn was this damage on?`, turnNumber);
    if (promptedTurn === null) {
      return;
    }
    turnNumber = promptedTurn;
  }

  saveUndoSnapshot();
  player.life += delta;

  const eventType = delta < 0 ? 'life-loss' : 'life-gain';
  if (delta < 0) {
    maybeRecordFirstBlood(sourcePlayerId, playerId, turnNumber);
    evaluateLiveElimination(player, sourcePlayerId, turnNumber, eventType);
  }

  recordLiveEvent({
    type: eventType,
    actorPlayerId: sourcePlayerId,
    targetPlayerId: playerId,
    amount: Math.abs(delta),
    turnNumber,
  });

  const alivePlayers = getActiveAlivePlayers(activeGameState);
  if (alivePlayers.length === 1) {
    alivePlayers[0].place = 1;
  }

  persistActiveGameState(activeGameState);
  refreshLiveTrackerUi();

  if (alivePlayers.length === 1 && await promptLiveConfirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`, {
    title: 'Finish game?',
    confirmLabel: 'Finish and save',
  })) {
    await completeActiveGame();
  }
}

async function applyManualLifeEntry(playerId) {
  if (!activeGameState) {
    return;
  }

  const player = activeGameState.players.find((entry) => entry.id === playerId);
  if (!player || player.eliminatedAt) {
    return;
  }

  const delta = await promptForSignedNumber(`Enter life change for ${player.name}. Use negative for damage and positive for life gain.`, -1);
  if (delta === null) {
    return;
  }

  await applyQuickLifeChange(playerId, delta);
}

async function setPlayerCannotLoseState(playerId, isEnabled) {
  if (!activeGameState) {
    return;
  }

  const player = activeGameState.players.find((entry) => entry.id === playerId);
  if (!player || player.eliminatedAt) {
    refreshLiveTrackerUi();
    return;
  }

  if (isEnabled) {
    saveUndoSnapshot();
    player.cannotLoseTheGame = true;
    persistActiveGameState(activeGameState);
    refreshLiveTrackerUi();
    return;
  }

  const lethalWhileDisabled = player.life <= 0 || hasCommanderDamageLethal(player);
  if (!lethalWhileDisabled) {
    saveUndoSnapshot();
    player.cannotLoseTheGame = false;
    persistActiveGameState(activeGameState);
    refreshLiveTrackerUi();
    return;
  }

  const sourcePlayerId = await resolveLiveSourceSelection({
    targetPlayerId: player.id,
    eventType: 'elimination',
    amount: 0,
    projectedLife: player.life,
  });
  if (sourcePlayerId === null) {
    refreshLiveTrackerUi();
    return;
  }

  const turnNumber = await promptForTurnNumber(`What turn was ${player.name} eliminated on?`);
  if (turnNumber === null) {
    refreshLiveTrackerUi();
    return;
  }

  saveUndoSnapshot();
  player.cannotLoseTheGame = false;
  const reason = player.life <= 0 ? 'life-total' : 'commander-damage';
  eliminateLivePlayer(player, sourcePlayerId, turnNumber, reason);
  recordLiveEvent({
    type: 'elimination',
    actorPlayerId: sourcePlayerId,
    targetPlayerId: player.id,
    amount: 0,
    turnNumber,
    notes: 'Resolved after cannot lose was disabled.',
  });

  const alivePlayers = getActiveAlivePlayers(activeGameState);
  if (alivePlayers.length === 1) {
    alivePlayers[0].place = 1;
  }

  persistActiveGameState(activeGameState);
  refreshLiveTrackerUi();

  if (alivePlayers.length === 1 && await promptLiveConfirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`, {
    title: 'Finish game?',
    confirmLabel: 'Finish and save',
  })) {
    await completeActiveGame();
  }
}

async function markPlayerAutomaticWinner(playerId) {
  if (!activeGameState) {
    return;
  }

  const winner = activeGameState.players.find((entry) => entry.id === playerId);
  if (!winner || winner.eliminatedAt) {
    return;
  }

  if (!await promptLiveConfirm(`Mark ${winner.name} as the winner and finish the game?`, {
    title: 'Confirm winner',
    confirmLabel: 'Mark winner',
  })) {
    return;
  }

  const turnNumber = await promptForTurnNumber(`What turn did ${winner.name} win on?`);
  if (turnNumber === null) {
    return;
  }

  const alternateWinCondition = await promptLiveText(`Enter ${winner.name}'s alternate win condition for the notes. Leave blank for a normal win.`, {
    title: 'Alternate win condition',
    confirmLabel: 'Save',
    defaultValue: activeGameState.alternateWinCondition || '',
    placeholder: 'Optional details',
  });
  if (alternateWinCondition === null) {
    return;
  }

  saveUndoSnapshot();
  activeGameState.turnNumber = turnNumber;
  activeGameState.alternateWinCondition = alternateWinCondition.trim();
  activeGameState.players.forEach((player) => {
    if (player.id === winner.id) {
      player.place = 1;
      return;
    }

    if (!player.eliminatedAt) {
      player.eliminatedAt = new Date().toISOString();
      player.eliminatedTurnNumber = activeGameState.turnNumber;
      player.eliminatedByPlayerId = '';
      player.eliminationReason = 'automatic-win';
    }
  });

  recordLiveEvent({
    type: 'automatic-win',
    actorPlayerId: winner.id,
    targetPlayerId: winner.id,
    amount: 0,
    turnNumber,
    notes: activeGameState.alternateWinCondition ? `Alternate win: ${activeGameState.alternateWinCondition}` : '',
  });
  await completeActiveGame();
}

function undoLastLiveAction() {
  if (!activeGameUndoState.length) {
    return;
  }

  const previousState = activeGameUndoState[activeGameUndoState.length - 1] || null;
  persistActiveGameState(cloneActiveGameState(previousState));
  persistActiveGameUndoState(activeGameUndoState.slice(0, -1));
  refresh();
}

async function completeActiveGame() {
  if (!activeGameState) {
    return;
  }

  const alivePlayers = getActiveAlivePlayers(activeGameState);
  if (alivePlayers.length > 1 && !await promptLiveConfirm('More than one player is still alive. Finish and score remaining players by life total?', {
    title: 'Finish active game?',
    confirmLabel: 'Finish game',
  })) {
    return;
  }

  const completedPlayers = activeGameState.players
    .map((player) => ({ ...player }))
    .sort((a, b) => {
      if (a.eliminatedTurnNumber && b.eliminatedTurnNumber && a.eliminatedTurnNumber !== b.eliminatedTurnNumber) {
        return a.eliminatedTurnNumber - b.eliminatedTurnNumber;
      }
      if (a.eliminatedAt && b.eliminatedAt) {
        return new Date(a.eliminatedAt).getTime() - new Date(b.eliminatedAt).getTime();
      }
      if (a.eliminatedAt) {
        return -1;
      }
      if (b.eliminatedAt) {
        return 1;
      }
      return a.life - b.life;
    });

  let nextPlace = completedPlayers.length;
  completedPlayers.forEach((player) => {
    if (player.place) {
      nextPlace = Math.min(nextPlace, player.place - 1);
      return;
    }
    player.place = nextPlace;
    nextPlace -= 1;
  });

  const finalPlayers = completedPlayers.slice().sort((a, b) => a.place - b.place);
  const durationMinutes = Math.max(1, Math.round((Date.now() - new Date(activeGameState.startedAt).getTime()) / 60000));
  const recordStats = buildLiveGameRecordStats(activeGameState, finalPlayers, durationMinutes);
  const playerRows = finalPlayers.map((player) => ({
    player: player.name,
    commander: player.commander,
    place: player.place,
    kills: player.kills,
    killed: player.killedPlayers || [],
  }));
  const finishOrder = finalPlayers.map((player) => player.name);

  const games = loadGames();
  games.push({
    id: activeGameState.id,
    date: activeGameState.date,
    playerRows,
    players: playerRows.map((row) => row.player),
    playerCommanders: playerRows.map((row) => ({ player: row.player, commander: row.commander })),
    finishOrder,
    notes: buildActiveGameSummary({ ...activeGameState, players: finalPlayers }),
    liveSummary: {
      startingPlayer: getPlayerNameById(activeGameState.startingPlayerId, activeGameState),
      alternateWinCondition: activeGameState.alternateWinCondition || '',
      alternateLoseConditions: finalPlayers
        .filter((player) => player.eliminationDetails)
        .map((player) => ({ player: player.name, details: player.eliminationDetails })),
      durationMinutes,
      firstBlood: activeGameState.firstBlood
        ? {
          actorPlayer: getPlayerNameById(activeGameState.firstBlood.actorPlayerId, activeGameState),
          actorCommander: finalPlayers.find((player) => player.id === activeGameState.firstBlood.actorPlayerId)?.commander || '',
          targetPlayer: getPlayerNameById(activeGameState.firstBlood.targetPlayerId, activeGameState),
          turnNumber: activeGameState.firstBlood.turnNumber,
        }
        : null,
      recordStats,
      turnNumber: activeGameState.turnNumber,
      eventCount: activeGameState.events.length,
    },
  });
  saveGames(games);
  persistActiveGameState(null);
  persistActiveGameUndoState(null);
  refresh();
  refreshLiveTrackerUi();
  await promptLiveAlert('Live game saved to history.', 'Game saved');
}

async function abandonActiveGame() {
  if (!activeGameState) {
    return;
  }

  if (!await promptLiveConfirm('Abandon the current live game? This removes the in-progress tracker without saving to history.', {
    title: 'Abandon live game?',
    confirmLabel: 'Abandon game',
  })) {
    return;
  }

  persistActiveGameState(null);
  persistActiveGameUndoState(null);
  refreshLiveTrackerUi();
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

  populateRowSelectors(row);
  attachLookupWrapperHandlers(row);
  return row;
}

function addPlayerRow(data = {}) {
  playerTableBody.appendChild(createPlayerRow(data));
}

function removePlayerRow() {
  const rows = playerTableBody.querySelectorAll('tr');
  if (rows.length > 2) {
    rows[rows.length - 1].remove();
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
    const commander = canonicalizeIdentityValue(row.querySelector('[name="commander"]')?.value || '', commanderMap);
    const placeRaw = row.querySelector('input[name="place"]').value.trim();
    const killsRaw = row.querySelector('input[name="kills"]').value.trim();
    const turnKilledRaw = row.querySelector('input[name="turnKilled"]').value.trim();
    const killed = normalizeList(row.querySelector('[name="killed"]').value);

    return {
      player,
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
  }));
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
  queueCloudSync();
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
  const existingRows = game.playerRows && game.playerRows.length ? game.playerRows : getGameRows(game);

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

function getHistoryFilterOptions(games) {
  const winners = new Set();
  const commanders = new Set();
  const players = new Set();

  games.forEach((game) => {
    const winner = getGameWinner(game);
    if (winner) {
      winners.add(winner);
    }

    getGameRows(game).forEach((row) => {
      if (row.player) {
        players.add(row.player);
      }

      if (row.commander) {
        commanders.add(row.commander);
      }
    });
  });

  return {
    winners: [...winners].sort((a, b) => a.localeCompare(b)),
    commanders: [...commanders].sort((a, b) => a.localeCompare(b)),
    players: [...players].sort((a, b) => a.localeCompare(b)),
  };
}

function buildFilterOptions(selectElement, values, label) {
  if (!selectElement) {
    return;
  }

  const currentValue = selectElement.value || 'all';
  const options = ['all', ...values];
  selectElement.innerHTML = options
    .map((value) => {
      const display = value === 'all' ? `All ${label}` : escapeHtml(value);
      const selected = currentValue === value ? ' selected' : '';
      return `<option value="${escapeHtml(value)}"${selected}>${display}</option>`;
    })
    .join('');
}

function updateHistoryFilters(games) {
  if (!historyFilterWinner && !historyFilterCommander && !historyFilterPlayer) {
    return;
  }

  const filters = getHistoryFilterOptions(games);
  buildFilterOptions(historyFilterWinner, filters.winners, 'winners');
  buildFilterOptions(historyFilterCommander, filters.commanders, 'commanders');
  buildFilterOptions(historyFilterPlayer, filters.players, 'players');
}

function applyHistoryQueryFilters() {
  if (historyQueryFiltersApplied) {
    return;
  }

  const winner = getQueryParam('winner');
  const commander = getQueryParam('commander');
  const player = getQueryParam('player');
  const fromDate = getQueryParam('from');
  const toDate = getQueryParam('to');

  if (historyFilterWinner && winner) {
    historyFilterWinner.value = winner;
  }
  if (historyFilterCommander && commander) {
    historyFilterCommander.value = commander;
  }
  if (historyFilterPlayer && player) {
    historyFilterPlayer.value = player;
  }
  if (historyFilterDateFrom && fromDate) {
    historyFilterDateFrom.value = fromDate;
  }
  if (historyFilterDateTo && toDate) {
    historyFilterDateTo.value = toDate;
  }

  historyQueryFiltersApplied = true;
}

function getCurrentHistoryFilters() {
  return {
    winner: historyFilterWinner?.value && historyFilterWinner.value !== 'all' ? historyFilterWinner.value : '',
    commander: historyFilterCommander?.value && historyFilterCommander.value !== 'all' ? historyFilterCommander.value : '',
    player: historyFilterPlayer?.value && historyFilterPlayer.value !== 'all' ? historyFilterPlayer.value : '',
    from: historyFilterDateFrom?.value || '',
    to: historyFilterDateTo?.value || '',
  };
}

function syncHistoryFilterQuery() {
  if (!historyList) {
    return;
  }

  const href = buildHistoryFilterHref(getCurrentHistoryFilters());
  const currentPath = `${window.location.pathname.split('/').pop() || 'history.html'}${window.location.search}`;
  if (currentPath !== href) {
    window.history.replaceState({}, '', href);
  }
}

function renderHistoryActiveFilters(totalCount, filteredCount) {
  if (!historyActiveFilters) {
    return;
  }

  const filters = getCurrentHistoryFilters();
  const activeFilters = [
    filters.winner ? `Winner: ${filters.winner}` : '',
    filters.commander ? `Commander: ${filters.commander}` : '',
    filters.player ? `Player: ${filters.player}` : '',
    filters.from ? `From: ${filters.from}` : '',
    filters.to ? `To: ${filters.to}` : '',
  ].filter(Boolean);

  const summaryLabel = filteredCount === totalCount
    ? `Showing all ${totalCount} game${totalCount === 1 ? '' : 's'}.`
    : `Showing ${filteredCount} of ${totalCount} game${totalCount === 1 ? '' : 's'}.`;

  if (!activeFilters.length) {
    historyActiveFilters.innerHTML = `<p class="history-active-filter-summary">${summaryLabel}</p>`;
    return;
  }

  historyActiveFilters.innerHTML = `
    <p class="history-active-filter-summary">${summaryLabel}</p>
    <div class="history-active-filter-list" aria-label="Active history filters">
      ${activeFilters.map((label) => `<span class="history-active-filter-chip">${escapeHtml(label)}</span>`).join('')}
    </div>`;
}

function resetHistoryFilters() {
  if (historyFilterWinner) {
    historyFilterWinner.value = 'all';
  }
  if (historyFilterCommander) {
    historyFilterCommander.value = 'all';
  }
  if (historyFilterPlayer) {
    historyFilterPlayer.value = 'all';
  }
  if (historyFilterDateFrom) {
    historyFilterDateFrom.value = '';
  }
  if (historyFilterDateTo) {
    historyFilterDateTo.value = '';
  }

  if (window.location.pathname.endsWith('/history.html') || window.location.pathname.endsWith('history.html')) {
    window.history.replaceState({}, '', 'history.html');
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

function buildDatalistOptions(element, values) {
  if (!element) {
    return;
  }

  element.innerHTML = values
    .map((value) => `<option value="${escapeHtml(value)}"></option>`)
    .join('');
}

function buildSelectOptions(element, values, selectedValue, placeholder) {
  if (!element) {
    return;
  }

  const normalized = getUniqueValues(values);
  element.innerHTML = [
    `<option value="">${escapeHtml(placeholder)}</option>`,
    ...normalized.map((value) => {
      const selected = selectedValue && value === selectedValue ? ' selected' : '';
      return `<option value="${escapeHtml(value)}"${selected}>${escapeHtml(value)}</option>`;
    }),
  ].join('');
}

function populateRowSelectors(row) {
  if (!row) {
    return;
  }

  const playerMenu = row.querySelector('.player-dropdown-menu');
  const commanderMenu = row.querySelector('.commander-dropdown-menu');

  if (playerMenu) {
    buildDropdownMenu(playerMenu, knownPlayers);
  }

  if (commanderMenu) {
    buildDropdownMenu(commanderMenu, knownCommanders);
  }
}

function buildDropdownMenu(menuElement, values) {
  if (!menuElement) {
    return;
  }
  
  const normalized = getUniqueValues(values);
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

  const optionValues = getUniqueValues(
    Array.from(menu.querySelectorAll('.dropdown-item')).map((item) => item.dataset.value || '').filter(Boolean),
  );

  mobileSelect.innerHTML = [
    `<option value="">${optionValues.length ? 'Choose saved option' : 'No saved options yet'}</option>`,
    ...optionValues.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`),
  ].join('');
  mobileSelect.disabled = optionValues.length === 0;
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

  knownPlayers = getUniqueValues(players);
  knownCommanders = getUniqueValues(commanders);

  if (playerDatalist) {
    buildDatalistOptions(playerDatalist, knownPlayers);
  }
  if (commanderDatalist) {
    buildDatalistOptions(commanderDatalist, knownCommanders);
  }
  if (firstBloodPlayerMenu) {
    buildDropdownMenu(firstBloodPlayerMenu, knownPlayers);
  }

  refreshRowSelectors();
  attachLookupWrapperHandlers(form || document);
  populateDeckCommanderSelector();
  populateRecordLookupMenus();
}

function normalizeDeckListEntry(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const playerMap = buildCanonicalIdentityMapFromValues(knownPlayers);
  const commanderMap = buildCanonicalIdentityMapFromValues(knownCommanders.concat(Object.keys(loadCommanderPowerLevels())));
  const commander = canonicalizeIdentityValue(entry.commander, commanderMap);
  const owner = canonicalizeIdentityValue(entry.owner, playerMap);
  const url = String(entry.url || '').trim();
  if (!commander || !url) {
    return null;
  }

  return {
    id: entry.id || generateId(),
    commander,
    owner,
    url,
  };
}

function getSortedDeckLists() {
  return loadDeckLists()
    .map(normalizeDeckListEntry)
    .filter(Boolean)
    .sort((a, b) => a.commander.localeCompare(b.commander));
}

function populateDeckCommanderSelector() {
  if (!deckCommanderMenu) {
    if (deckOwnerMenu) {
      buildDropdownMenu(deckOwnerMenu, knownPlayers);
      attachLookupWrapperHandlers(deckListForm || document);
    }
    return;
  }

  buildDropdownMenu(deckCommanderMenu, knownCommanders);
  if (deckOwnerMenu) {
    buildDropdownMenu(deckOwnerMenu, knownPlayers);
  }
  attachLookupWrapperHandlers(deckListForm || document);
}

function populateRecordLookupMenus() {
  if (recordsTableBody) {
    recordsTableBody.querySelectorAll('.record-holder-menu').forEach((menu) => {
      buildDropdownMenu(menu, knownPlayers);
    });

    recordsTableBody.querySelectorAll('.record-commander-menu').forEach((menu) => {
      buildDropdownMenu(menu, knownCommanders);
    });
  }

  if (customRecordHolderMenu) {
    buildDropdownMenu(customRecordHolderMenu, knownPlayers);
  }

  if (customRecordCommanderMenu) {
    buildDropdownMenu(customRecordCommanderMenu, knownCommanders);
  }

  attachLookupWrapperHandlers(recordsForm || document);
  attachLookupWrapperHandlers(customRecordForm || document);
}

function renderDeckLookup() {
  if (!deckLookupSelect || !deckLookupResult) {
    return;
  }

  const deckLists = getSortedDeckLists();
  const selectedCommander = deckLookupSelect.value;

  buildSelectOptions(
    deckLookupSelect,
    deckLists.map((entry) => entry.commander),
    selectedCommander,
    'Select a commander',
  );

  const activeCommander = deckLookupSelect.value;
  if (!deckLists.length) {
    deckLookupResult.innerHTML = '<p>No deck lists saved yet.</p>';
    return;
  }

  if (!activeCommander) {
    deckLookupResult.innerHTML = '<p>Select a commander to view its saved deck URL.</p>';
    return;
  }

  const selectedDeck = deckLists.find((entry) => entry.commander === activeCommander);
  if (!selectedDeck) {
    deckLookupResult.innerHTML = '<p>No saved URL found for that commander.</p>';
    return;
  }

  const safeCommander = escapeHtml(selectedDeck.commander);
  const safeUrl = escapeHtml(selectedDeck.url);
  const safeOwner = escapeHtml(selectedDeck.owner || 'Unassigned');
  deckLookupResult.innerHTML = `
    <p class="deck-lookup-label">${safeCommander}</p>
    <p>Owner: ${safeOwner}</p>
    <a href="${safeUrl}" target="_blank" rel="noopener noreferrer" title="${safeUrl}">${safeUrl}</a>
  `;
}

function getDeckOwnerGroups() {
  return getSortedDeckLists().reduce((groups, entry) => {
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

function getSelectedDeckSelectorOwners() {
  if (!deckSelectorOwnerList) {
    return [];
  }

  return Array.from(deckSelectorOwnerList.querySelectorAll('input[name="deck-selector-owner"]:checked'))
    .map((input) => input.value)
    .filter(Boolean);
}

function chooseRandomDeck(deckOptions) {
  if (!Array.isArray(deckOptions) || !deckOptions.length) {
    return null;
  }

  const index = Math.floor(Math.random() * deckOptions.length);
  return deckOptions[index] || null;
}

function shuffleList(items) {
  const shuffled = [...items];
  for (let index = shuffled.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    [shuffled[index], shuffled[swapIndex]] = [shuffled[swapIndex], shuffled[index]];
  }
  return shuffled;
}

function getDeckSelectorPool(selectedOwners) {
  const ownerGroups = getDeckOwnerGroups();
  return shuffleList(
    selectedOwners.flatMap((owner) => ownerGroups[owner] || []),
  );
}

function getDeckWheelPalette(count) {
  const palette = [
    '#6aa9ff',
    '#7fd4b8',
    '#ffd36d',
    '#ff9f7a',
    '#c8a8ff',
    '#8fd3ff',
    '#ffb6c7',
    '#b7df76',
    '#ffc280',
    '#8ec5a4',
  ];

  return Array.from({ length: count }, (_, index) => palette[index % palette.length]);
}

function polarToCartesian(centerX, centerY, radius, angleInDegrees) {
  const angleInRadians = ((angleInDegrees - 90) * Math.PI) / 180;
  return {
    x: centerX + (radius * Math.cos(angleInRadians)),
    y: centerY + (radius * Math.sin(angleInRadians)),
  };
}

function describeWheelSegment(startAngle, endAngle) {
  const start = polarToCartesian(50, 50, 48, endAngle);
  const end = polarToCartesian(50, 50, 48, startAngle);
  const largeArcFlag = endAngle - startAngle <= 180 ? '0' : '1';
  return [
    'M 50 50',
    `L ${start.x} ${start.y}`,
    `A 48 48 0 ${largeArcFlag} 0 ${end.x} ${end.y}`,
    'Z',
  ].join(' ');
}

function getDeckWheelLabelFontSize(labelLength, segmentSize, radialPathLength = 26) {
  const minFontSize = 1.55;
  const maxFontSize = 2;
  const averageRadius = 31;
  const segmentRadians = (segmentSize * Math.PI) / 180;
  const wedgeWidth = averageRadius * segmentRadians;
  const estimatedLengthUnits = (labelLength * 0.62) + 1.2;
  const lengthConstrainedSize = radialPathLength / Math.max(estimatedLengthUnits, 1);
  const wedgeConstrainedSize = wedgeWidth * 0.72;
  return Math.max(minFontSize, Math.min(maxFontSize, lengthConstrainedSize, wedgeConstrainedSize));
}

function truncateDeckWheelLabel(labelText, maxCharacters) {
  if (labelText.length <= maxCharacters) {
    return labelText;
  }

  if (maxCharacters <= 4) {
    return `${labelText.slice(0, Math.max(1, maxCharacters - 1))}…`;
  }

  const hardTrimmed = labelText.slice(0, maxCharacters - 1).trimEnd();
  const lastSpaceIndex = hardTrimmed.lastIndexOf(' ');
  const charactersDroppedForWordBoundary = lastSpaceIndex >= 0
    ? hardTrimmed.length - lastSpaceIndex
    : Number.POSITIVE_INFINITY;
  const trimmed = charactersDroppedForWordBoundary <= 2
    ? hardTrimmed.slice(0, lastSpaceIndex)
    : hardTrimmed;

  return `${trimmed.trim()}…`;
}

function fitDeckWheelLabelText(labelText, segmentSize, radialPathLength = 26) {
  return {
    fontSize: getDeckWheelLabelFontSize(labelText.length, segmentSize, radialPathLength),
    text: labelText,
  };
}

function getDeckWheelLabelPath(midAngle, innerRadius, outerRadius) {
  const innerPoint = polarToCartesian(50, 50, innerRadius, midAngle);
  const outerPoint = polarToCartesian(50, 50, outerRadius, midAngle);
  const isLeftSide = midAngle > 90 && midAngle < 270;
  return {
    start: isLeftSide ? outerPoint : innerPoint,
    end: isLeftSide ? innerPoint : outerPoint,
  };
}

function getDeckWheelSvgMarkup(deckPool) {
  const count = deckPool.length;
  const colors = getDeckWheelPalette(count);
  const segmentSize = 360 / count;

  const segments = deckPool.map((deck, index) => {
    const startAngle = index * segmentSize;
    const endAngle = startAngle + segmentSize;
    const midAngle = startAngle + (segmentSize / 2);
    const pathId = `deck-wheel-label-path-${index}`;
    const path = getDeckWheelLabelPath(midAngle, 18, 44);
    const fittedLabel = fitDeckWheelLabelText(deck.commander, segmentSize);

    return `
      <path d="${describeWheelSegment(startAngle, endAngle)}" fill="${colors[index]}" class="deck-wheel-segment" />
      <path id="${pathId}" d="M ${path.start.x} ${path.start.y} L ${path.end.x} ${path.end.y}" class="deck-wheel-label-guide" />
      <text class="deck-wheel-segment-text" data-full-label="${escapeHtml(deck.commander)}" data-segment-size="${segmentSize}" style="font-size: ${fittedLabel.fontSize.toFixed(2)}px;">
        <textPath href="#${pathId}" startOffset="8%">${escapeHtml(fittedLabel.text)}</textPath>
      </text>`;
  }).join('');

  return `<svg viewBox="0 0 100 100" class="deck-wheel-svg" aria-hidden="true">${segments}</svg>`;
}

function fitDeckWheelSvgLabels(rootElement) {
  const minFontSize = 1.55;
  const guaranteedLabelLength = 31;
  const guaranteedMinFontSize = 0.95;
  const labelElements = rootElement?.querySelectorAll?.('.deck-wheel-segment-text') || [];

  labelElements.forEach((labelElement) => {
    const textPath = labelElement.querySelector('textPath');
    const fullLabel = labelElement.dataset.fullLabel || textPath?.textContent || '';
    const segmentSize = Number(labelElement.dataset.segmentSize || '0');
    const pathReference = textPath?.getAttribute('href');

    if (!textPath || !fullLabel || !pathReference) {
      return;
    }

    const guidePath = rootElement.querySelector(pathReference);

    if (!guidePath || typeof guidePath.getTotalLength !== 'function' || typeof labelElement.getComputedTextLength !== 'function') {
      textPath.textContent = fullLabel;
      return;
    }

    const availableLength = guidePath.getTotalLength() * 0.9;
    const preferredFontSize = getDeckWheelLabelFontSize(fullLabel.length, segmentSize);
    textPath.textContent = fullLabel;
    labelElement.style.fontSize = `${preferredFontSize.toFixed(2)}px`;

    let renderedLength = labelElement.getComputedTextLength();
    if (renderedLength <= availableLength) {
      return;
    }

    if (fullLabel.length <= guaranteedLabelLength) {
      const guaranteedSize = Math.max(guaranteedMinFontSize, preferredFontSize * (availableLength / renderedLength));
      labelElement.style.fontSize = `${guaranteedSize.toFixed(2)}px`;
      return;
    }

    const scaledFontSize = Math.max(minFontSize, preferredFontSize * (availableLength / renderedLength));
    labelElement.style.fontSize = `${scaledFontSize.toFixed(2)}px`;
    renderedLength = labelElement.getComputedTextLength();

    if (renderedLength <= availableLength || scaledFontSize > minFontSize) {
      return;
    }

    labelElement.style.fontSize = `${minFontSize.toFixed(2)}px`;

    let bestFit = truncateDeckWheelLabel(fullLabel, 3);
    let low = 3;
    let high = fullLabel.length;

    while (low <= high) {
      const middle = Math.floor((low + high) / 2);
      const candidate = truncateDeckWheelLabel(fullLabel, middle);
      textPath.textContent = candidate;

      if (labelElement.getComputedTextLength() <= availableLength) {
        bestFit = candidate;
        low = middle + 1;
      } else {
        high = middle - 1;
      }
    }

    textPath.textContent = bestFit;
  });
}

function renderDeckSelectorWheel(deckPool, centerLabel = 'Ready to Spin') {
  if (!deckSelectorWheel || !deckSelectorWheelDisc) {
    return;
  }

  if (!deckPool.length) {
    deckSelectorWheel.classList.add('is-empty');
    deckSelectorWheelDisc.innerHTML = `<div class="deck-wheel-center-label">${escapeHtml(centerLabel)}</div>`;
    deckSelectorWheelDisc.style.transform = 'rotate(0deg)';
    return;
  }

  deckSelectorWheel.classList.remove('is-empty');
  deckSelectorWheelDisc.innerHTML = `
    ${getDeckWheelSvgMarkup(deckPool)}
    <div class="deck-wheel-center-label">${escapeHtml(centerLabel)}</div>
  `;
  fitDeckWheelSvgLabels(deckSelectorWheelDisc);
  deckSelectorWheelDisc.style.transform = `rotate(${deckSelectorRotation}deg)`;
}

function renderDeckSelectorResult(selectedOwners, deck) {
  if (!deckSelectorResults) {
    return;
  }

  const safeCommander = escapeHtml(deck.commander);
  const safeUrl = escapeHtml(deck.url);
  const safeOwner = escapeHtml(deck.owner || 'Unassigned');
  const safePool = escapeHtml(selectedOwners.join(', '));

  deckSelectorResults.innerHTML = `
    <article class="deck-selector-card">
      <p class="deck-selector-owner">From pool: ${safePool}</p>
      <h3>${safeCommander}</h3>
      <p>Owned by ${safeOwner}</p>
      <a href="${safeUrl}" target="_blank" rel="noopener noreferrer">Open deck list</a>
    </article>`;
}

function renderDeckSelectorAssignments(selectedOwners) {
  if (!deckSelectorResults || !deckSelectorWheelDisc) {
    return;
  }

  if (deckSelectorSpinTimer) {
    clearTimeout(deckSelectorSpinTimer);
    deckSelectorSpinTimer = null;
  }

  if (!selectedOwners.length) {
    deckSelectorResults.innerHTML = '<p>Select at least one player to randomize decks.</p>';
    if (deckSelectorWheelStatus) {
      deckSelectorWheelStatus.textContent = 'Select players and click Randomize decks.';
    }
    renderDeckSelectorWheel([], 'Select Players');
    return;
  }

  const pooledDecks = getDeckSelectorPool(selectedOwners);

  if (!pooledDecks.length) {
    deckSelectorResults.innerHTML = '<p>No owned decks were found for the selected players.</p>';
    if (deckSelectorWheelStatus) {
      deckSelectorWheelStatus.textContent = 'No eligible decks found in the selected pool.';
    }
    renderDeckSelectorWheel([], 'No Decks');
    return;
  }

  const winningIndex = Math.floor(Math.random() * pooledDecks.length);
  const deck = pooledDecks[winningIndex];
  if (!deck) {
    deckSelectorResults.innerHTML = '<p>No owned decks were found for the selected players.</p>';
    if (deckSelectorWheelStatus) {
      deckSelectorWheelStatus.textContent = 'No eligible decks found in the selected pool.';
    }
    return;
  }

  const segmentSize = 360 / pooledDecks.length;
  const winningCenterAngle = (winningIndex * segmentSize) + (segmentSize / 2);
  const extraTurns = 5 + Math.floor(Math.random() * 2);
  const normalizedRotation = ((deckSelectorRotation % 360) + 360) % 360;
  const neededOffset = ((360 - winningCenterAngle - normalizedRotation) + 360) % 360;
  const targetRotation = deckSelectorRotation + (extraTurns * 360) + neededOffset;
  const spinDuration = 5200;

  renderDeckSelectorWheel(pooledDecks, 'Spinning');
  deckSelectorResults.innerHTML = '<p>The wheel is spinning...</p>';
  if (deckSelectorWheelStatus) {
    deckSelectorWheelStatus.textContent = `Spinning through ${pooledDecks.length} eligible decks...`;
  }
  if (deckSelectorSubmitButton) {
    deckSelectorSubmitButton.disabled = true;
  }

  deckSelectorWheelDisc.style.transition = 'none';
  deckSelectorWheelDisc.style.transform = `rotate(${deckSelectorRotation}deg)`;
  deckSelectorWheelDisc.getBoundingClientRect();
  deckSelectorWheelDisc.style.transition = `transform ${spinDuration}ms cubic-bezier(0.16, 1, 0.3, 1)`;
  deckSelectorWheelDisc.style.transform = `rotate(${targetRotation}deg)`;
  deckSelectorRotation = targetRotation;

  deckSelectorSpinTimer = window.setTimeout(() => {
    deckSelectorSpinTimer = null;
    renderDeckSelectorWheel(pooledDecks, 'Winner');
    deckSelectorWheelDisc.style.transition = 'none';
    deckSelectorWheelDisc.style.transform = `rotate(${deckSelectorRotation}deg)`;
    renderDeckSelectorResult(selectedOwners, deck);
    if (deckSelectorWheelStatus) {
      deckSelectorWheelStatus.textContent = `${deck.commander} selected.`;
    }
    if (deckSelectorSubmitButton) {
      deckSelectorSubmitButton.disabled = false;
    }
  }, spinDuration + 80);

}

function renderDeckSelector() {
  if (!deckSelectorOwnerList || !deckSelectorResults) {
    return;
  }

  const selectedOwners = new Set(getSelectedDeckSelectorOwners());
  const ownerGroups = getDeckOwnerGroups();
  const owners = Object.keys(ownerGroups).sort((a, b) => a.localeCompare(b));

  if (!owners.length) {
    deckSelectorOwnerList.innerHTML = '<p>No owned decks available yet. Add deck owners in Deck Lists first.</p>';
    deckSelectorResults.innerHTML = '<p>Add deck owners in Deck Lists to start randomizing decks.</p>';
    if (deckSelectorWheelStatus) {
      deckSelectorWheelStatus.textContent = 'Add deck owners in Deck Lists to enable the wheel.';
    }
    renderDeckSelectorWheel([], 'No Decks');
    return;
  }

  deckSelectorOwnerList.innerHTML = owners
    .map((owner) => `
      <label class="deck-selector-option">
        <input type="checkbox" name="deck-selector-owner" value="${escapeHtml(owner)}"${selectedOwners.has(owner) ? ' checked' : ''} />
        <span>${escapeHtml(owner)}</span>
      </label>`)
    .join('');

  if (!deckSelectorResults.dataset.initialized) {
    deckSelectorResults.innerHTML = '<p>Select players and click Randomize decks.</p>';
    deckSelectorResults.dataset.initialized = 'true';
  }

  if (deckSelectorWheelStatus && !deckSelectorWheelStatus.dataset.initialized) {
    deckSelectorWheelStatus.textContent = 'Select players and click Randomize decks.';
    deckSelectorWheelStatus.dataset.initialized = 'true';
  }

  if (deckSelectorWheel && deckSelectorWheel.classList.contains('is-empty')) {
    renderDeckSelectorWheel([], 'Ready to Spin');
  }
}

function resetDeckListForm() {
  if (!deckListForm) {
    return;
  }

  editingDeckListId = null;
  deckListForm.reset();
  if (deckListSubmitButton) {
    deckListSubmitButton.textContent = 'Add deck list';
  }
  if (deckListCancelButton) {
    deckListCancelButton.hidden = true;
  }
}

function startDeckListEdit(deckId) {
  if (!deckListForm) {
    return;
  }

  const entry = loadDeckLists().find((deck) => deck.id === deckId);
  if (!entry) {
    return;
  }

  editingDeckListId = entry.id;
  deckCommanderInput.value = entry.commander;
  if (deckOwnerInput) {
    deckOwnerInput.value = entry.owner || '';
  }
  deckUrlInput.value = entry.url;

  if (deckListSubmitButton) {
    deckListSubmitButton.textContent = 'Save deck list';
  }
  if (deckListCancelButton) {
    deckListCancelButton.hidden = false;
  }
}

function deleteDeckList(deckId) {
  const remaining = loadDeckLists().filter((deck) => deck.id !== deckId);
  saveDeckLists(remaining);

  if (editingDeckListId === deckId) {
    resetDeckListForm();
  }

  refresh();
}

function renderDeckLists() {
  if (!deckListTableBody) {
    return;
  }

  const sortState = getTableSort('deckLists', 'commander', false);
  const deckLists = getSortedDeckLists()
    .slice()
    .sort((a, b) => {
      let result = 0;
      switch (sortState.column) {
        case 'commander':
          result = compareTextValues(a.commander, b.commander);
          break;
        case 'owner':
          result = compareTextValues(a.owner, b.owner);
          break;
        case 'url':
          result = compareTextValues(a.url, b.url);
          break;
        default:
          result = compareTextValues(a.commander, b.commander);
          break;
      }

      if (result === 0) {
        result = compareTextValues(a.commander, b.commander);
      }

      return finalizeSortResult(result, sortState.descending);
    });

  if (!deckLists.length) {
    deckListTableBody.innerHTML = '<tr><td colspan="4">No deck lists saved yet.</td></tr>';
    return;
  }

  deckListTableBody.innerHTML = deckLists
    .map((entry) => {
      const safeCommander = escapeHtml(entry.commander);
      const safeOwner = escapeHtml(entry.owner || '—');
      const safeUrl = escapeHtml(entry.url);
      return `
        <tr>
          <td>${safeCommander}</td>
          <td>${safeOwner}</td>
          <td><a href="${safeUrl}" target="_blank" rel="noopener noreferrer" aria-label="Open saved deck list for ${safeCommander}">${safeUrl}</a></td>
          <td>
            <button type="button" class="secondary-button deck-list-edit" data-id="${escapeHtml(entry.id)}" aria-label="Edit deck list for ${safeCommander}">Edit</button>
            <button type="button" class="history-delete-button deck-list-delete" data-id="${escapeHtml(entry.id)}" aria-label="Delete deck list for ${safeCommander}">Delete</button>
          </td>
        </tr>`;
    })
    .join('');

  updateSortableTableIndicators('deckLists');
}

function collectRecordsFromTable() {
  if (!recordsTableBody) {
    return loadRecords();
  }

  return Array.from(recordsTableBody.querySelectorAll('tr[data-record-id]'))
    .map((row, index) => {
      const isCustom = row.dataset.custom === 'true';
      const valueInput = row.querySelector('[name="value"]');
      const holderInput = row.querySelector('[name="holder"]');
      const commanderInput = row.querySelector('[name="commander"]');
      const dateField = row.querySelector('[name="date"]');
      const notesInput = row.querySelector('[name="notes"]');

      return normalizeRecordEntry({
        id: row.dataset.recordId,
        key: row.dataset.key || '',
        title: row.dataset.title || '',
        unit: row.dataset.unit || '',
        value: valueInput?.value || '',
        holder: holderInput?.value || '',
        commander: commanderInput?.value || '',
        date: dateField?.value || '',
        notes: notesInput?.value || '',
        isCustom,
      }, index);
    })
    .filter(Boolean);
}

function renderRecords() {
  if (!recordsTableBody) {
    return;
  }

  const records = loadRecords();
  recordsTableBody.innerHTML = records
    .map((record) => `
        <tr
          data-record-id="${escapeHtml(record.id)}"
          data-key="${escapeHtml(record.key || '')}"
          data-title="${escapeHtml(record.title)}"
          data-unit="${escapeHtml(record.unit || '')}"
          data-custom="${record.isCustom ? 'true' : 'false'}"
        >
          <td class="record-title-cell">
            <strong>${escapeHtml(record.title)}</strong>
          </td>
          <td class="record-value-cell"><input type="text" name="value" value="${escapeHtml(record.value)}" placeholder="Record" /></td>
          <td class="record-unit-cell">
            <span class="record-unit-badge">${escapeHtml(record.unit || 'open')}</span>
          </td>
          <td>
            <div class="combined-input-wrapper record-lookup-wrapper">
              <input class="lookup-input" type="text" name="holder" list="player-list" value="${escapeHtml(record.holder)}" placeholder="Player" autocomplete="off" autocapitalize="none" autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" />
              <button type="button" class="dropdown-button" title="Show players">▼</button>
              <div class="dropdown-menu record-holder-menu"></div>
            </div>
          </td>
          <td class="record-commander-cell">
            <div class="combined-input-wrapper record-lookup-wrapper">
              <input class="lookup-input" type="text" name="commander" list="commander-list" value="${escapeHtml(record.commander)}" placeholder="Commander or deck" autocomplete="off" autocapitalize="none" autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" />
              <button type="button" class="dropdown-button" title="Show commanders">▼</button>
              <div class="dropdown-menu record-commander-menu"></div>
            </div>
          </td>
          <td><input type="date" name="date" value="${escapeHtml(record.date)}" /></td>
          <td><textarea name="notes" rows="2" placeholder="How it happened, matchup, table notes...">${escapeHtml(record.notes)}</textarea></td>
        </tr>`)
    .join('');

  populateRecordLookupMenus();
}

async function handleDeckListTableAction(event) {
  const button = event.target.closest('button');
  if (!button || !deckListTableBody.contains(button)) {
    return;
  }

  const deckId = button.dataset.id;
  if (!deckId) {
    return;
  }

  if (button.classList.contains('deck-list-edit')) {
    startDeckListEdit(deckId);
    return;
  }

  if (button.classList.contains('deck-list-delete')) {
    if (await promptLiveConfirm('Delete this deck list? This removes the saved deck link immediately.', {
      title: 'Delete deck list?',
      confirmLabel: 'Delete deck list',
    })) {
      deleteDeckList(deckId);
    }
  }
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

function getSortedHistoryGames(games) {
  const sorted = [...games];

  sorted.sort((a, b) => {
    if (historySortKey === 'date') {
      const result = (a.date || '').localeCompare(b.date || '');
      return historySortDescending ? -result : result;
    }

    if (historySortKey === 'winner') {
      const aWinner = getGameWinner(a) || '';
      const bWinner = getGameWinner(b) || '';
      const result = aWinner.localeCompare(bWinner);
      return historySortDescending ? -result : result;
    }

    if (historySortKey === 'players') {
      const result = getGamePlayerCount(a) - getGamePlayerCount(b);
      return historySortDescending ? -result : result;
    }

    if (historySortKey === 'totalKills') {
      const result = getGameTotalKills(a) - getGameTotalKills(b);
      return historySortDescending ? -result : result;
    }

    return 0;
  });

  return sorted;
}

function renderHistoryGame(game) {
  const rows = getGameRows(game);
  const winner = getGameWinner(game) || '—';
  const totalKills = getGameTotalKills(game);
  const playerCount = rows.length;
  const firstBloodLabel = getGameFirstBloodLabel(game);
  const notes = game.notes ? escapeHtml(game.notes) : 'No notes recorded.';
  const gameDate = escapeHtml(game.date || 'Unknown date');

  const playerRows = rows
    .map((row) => {
      const killed = Array.isArray(row.killed) ? row.killed.join(', ') : row.killed || '';
      return `
        <tr>
          <td>${escapeHtml(row.player)}</td>
          <td>${escapeHtml(row.commander)}</td>
          <td>${row.place || '—'}</td>
          <td>${typeof row.kills === 'number' ? row.kills : 0}</td>
          <td>${escapeHtml(killed)}</td>
        </tr>`;
    })
    .join('');

  return `
    <article class="history-item">
      <div class="row history-item-meta">
        <h3>${gameDate}</h3>
        <div class="history-item-facts" aria-label="Game summary">
          <span class="history-item-fact">Winner: ${escapeHtml(winner)}</span>
          <span class="history-item-fact">${playerCount} players</span>
          <span class="history-item-fact">${totalKills} kills</span>
          <span class="history-item-fact">First blood: ${escapeHtml(firstBloodLabel)}</span>
        </div>
      </div>
      <div class="history-item-actions">
        <button type="button" class="secondary-button history-edit-button" data-id="${escapeHtml(game.id)}" aria-label="Edit saved game from ${gameDate}">Edit</button>
        <button type="button" class="history-delete-button" data-id="${escapeHtml(game.id)}" aria-label="Delete saved game from ${gameDate}">Delete</button>
      </div>
      <p>${notes}</p>
      <table class="player-table history-game-table" aria-label="Players for game on ${gameDate}">
        <thead>
          <tr>
            <th scope="col">Player</th>
            <th scope="col">Commander</th>
            <th scope="col">Place</th>
            <th scope="col">Kills</th>
            <th scope="col">Killed</th>
          </tr>
        </thead>
        <tbody>
          ${playerRows}
        </tbody>
      </table>
    </article>`;
}

function renderHistory(games) {
  if (!historyList) {
    return;
  }

  const sortedGames = getSortedHistoryGames(games);
  const winnerFilter = historyFilterWinner?.value || 'all';
  const commanderFilter = historyFilterCommander?.value || 'all';
  const playerFilter = historyFilterPlayer?.value || 'all';
  const dateFromFilter = historyFilterDateFrom?.value || '';
  const dateToFilter = historyFilterDateTo?.value || '';

  const filteredGames = sortedGames.filter((game) => {
    if (dateFromFilter && (game.date || '') < dateFromFilter) {
      return false;
    }

    if (dateToFilter && (game.date || '') > dateToFilter) {
      return false;
    }

    if (winnerFilter !== 'all' && getGameWinner(game) !== winnerFilter) {
      return false;
    }

    if (playerFilter !== 'all') {
      const foundPlayer = getGameRows(game).some((row) => row.player === playerFilter);
      if (!foundPlayer) {
        return false;
      }
    }

    if (commanderFilter !== 'all') {
      const foundCommander = getGameRows(game).some((row) => row.commander === commanderFilter);
      if (!foundCommander) {
        return false;
      }
    }

    return true;
  });

  syncHistoryFilterQuery();
  renderHistoryActiveFilters(sortedGames.length, filteredGames.length);

  if (!filteredGames.length) {
    historyList.innerHTML = '<p class="history-empty-state">No games match the current filters.</p>';
    return;
  }

  historyList.innerHTML = filteredGames.map(renderHistoryGame).join('');
}

async function handleHistoryAction(event) {
  const button = event.target.closest('button');
  if (!button || !historyList.contains(button)) {
    return;
  }

  const gameId = button.dataset.id;
  if (!gameId) {
    return;
  }

  if (button.classList.contains('history-delete-button')) {
    if (await promptLiveConfirm('Delete this game? This removes it from history, rankings, and all derived stats.', {
      title: 'Delete saved game?',
      confirmLabel: 'Delete game',
    })) {
      deleteGame(gameId);
    }
    return;
  }

  if (button.classList.contains('history-edit-button')) {
    const game = getGameById(gameId);
    if (game) {
      window.location.href = `index.html?editId=${encodeURIComponent(gameId)}`;
    }
    return;
  }
}

function updateHistorySortOrderLabel() {
  if (!historySortOrderButton) {
    return;
  }

  if (historySortKey === 'date') {
    historySortOrderButton.textContent = historySortDescending ? 'Newest first' : 'Oldest first';
  } else {
    historySortOrderButton.textContent = historySortDescending ? 'Descending' : 'Ascending';
  }
}

function getGameRows(game) {
  if (Array.isArray(game.playerRows) && game.playerRows.length) {
    return game.playerRows;
  }

  if (Array.isArray(game.playerCommanders) && game.playerCommanders.length) {
    return game.playerCommanders.map((entry) => ({
      player: entry.player || '',
      commander: entry.commander || '',
      place: typeof entry.place !== 'undefined'
        ? entry.place
        : (Array.isArray(game.finishOrder) ? game.finishOrder.indexOf(entry.player) + 1 : null),
      kills: typeof entry.kills === 'number' ? entry.kills : 0,
      killed: entry.killed ?? [],
    }));
  }

  if (Array.isArray(game.players) && game.players.length) {
    return game.players.map((player) => ({
      player,
      commander: '',
      place: Array.isArray(game.finishOrder) ? game.finishOrder.indexOf(player) + 1 : null,
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
  return [...games]
    .map((game, index) => ({ game, index }))
    .sort((a, b) => {
      const dateResult = (a.game.date || '').localeCompare(b.game.date || '');
      if (dateResult) {
        return dateResult;
      }
      return a.index - b.index;
    })
    .map(({ game }) => game);
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
  const stats = {};
  const eloKFactor = 28;
  const eloKillBonus = 2;
  const streakBonusStep = 3;
  const maxStreakBonus = 18;
  const playerGameWindowLookup = getPlayerGameWindowLookup(games, 10);

  getGamesSortedByDateAscending(games).forEach((game) => {
    const rows = getGameRows(game);
    const firstBlood = getGameFirstBloodInfo(game);
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
          commanders: {},
        };
      }

      const entry = stats[player];
      const place = getGameRowPlace(row, game);
      const kills = getRowKills(row);
      const commander = String(row.commander || '').trim();
      const firstBloodBonus = firstBlood?.actorPlayer === player ? 1 : 0;
      const isInPlayerWindow = playerGameWindowLookup.get(player)?.has(game.id) ?? true;

      if (isInPlayerWindow) {
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

      stats[player].rating += baseRatingChange + streakBonus + (kills * eloKillBonus);
    });
  });

  return Object.values(stats)
    .filter((entry) => entry.games > 0)
    .map((entry) => ({
      ...entry,
      rating: Math.round(entry.rating),
      avgPlace: entry.placementGames ? entry.placementTotal / entry.placementGames : null,
      pointsPerGame: entry.games ? entry.points / entry.games : 0,
      winRate: entry.games ? (entry.wins / entry.games) * 100 : 0,
      favoriteCommander: getMaxCountKey(entry.commanders),
    }))
    .sort(compareRankingsEntries);
}

function buildCommanderRankingEntries(games) {
  const stats = {};

  games.forEach((game) => {
    const rows = getGameRows(game);
    const firstBlood = getGameFirstBloodInfo(game);

    rows.forEach((row) => {
      const commander = String(row.commander || '').trim();
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

  return Object.values(stats)
    .map((entry) => ({
      ...entry,
      avgPlace: entry.placementGames ? entry.placementTotal / entry.placementGames : null,
      pointsPerGame: entry.games ? entry.points / entry.games : 0,
      winRate: entry.games ? (entry.wins / entry.games) * 100 : 0,
    }))
    .sort(compareAggregateEntries);
}

function buildStreakEntries(games, keySelector) {
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

  return Object.entries(appearances)
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

function renderPodRankings(games) {
  if (!rankingsSummary || !rankingsTableBody) {
    return;
  }

  const sortState = getTableSort('rankingsMain', 'rank', false);
  const entries = buildPlayerRankingEntries(games);
  const recentEntries = buildPlayerRankingEntries(getGamesSortedByDateAscending(games).slice(-10));
  const commanderEntries = buildCommanderRankingEntries(games);
  const playerStreaks = buildStreakEntries(games, (row) => row.player);
  const playerStreaksByName = new Map(playerStreaks.map((entry) => [entry.name, entry]));

  if (!entries.length) {
    renderStatCardGroup(rankingsSummary, [
      { title: 'No rankings yet', body: 'Save a few games to generate pod standings.' },
    ]);
    rankingsTableBody.innerHTML = '<tr><td colspan="12">No rankings available yet.</td></tr>';
    return;
  }

  const topPlayer = entries[0];
  const bestEfficiency = entries
    .filter((entry) => entry.games >= 3)
    .sort((a, b) => b.pointsPerGame - a.pointsPerGame)[0] || topPlayer;
  const hottestRecent = recentEntries[0] || topPlayer;
  const activeStreakLeader = playerStreaks.find((entry) => entry.currentWinStreak > 0) || null;
  const topCommander = commanderEntries
    .filter((entry) => entry.games >= 3)
    .sort((a, b) => b.pointsPerGame - a.pointsPerGame || compareAggregateEntries(a, b))[0]
    || commanderEntries[0]
    || null;

  renderStatCardGroup(rankingsSummary, [
    {
      title: '#1 Player',
      body: `${topPlayer.name} leads at ${formatRating(topPlayer.rating)} ELO across ${topPlayer.games} games.`,
    },
    {
      title: 'Best Efficiency',
      body: `${bestEfficiency.name} averages ${formatPointsPerGame(bestEfficiency.pointsPerGame)} placement points per game.`,
    },
    {
      title: 'Hottest Last 10',
      body: `${hottestRecent.name} tops the last 10 at ${formatRating(hottestRecent.rating)} ELO.`,
    },
    {
      title: 'Best Commander Form',
      body: topCommander ? `${topCommander.name} averages ${formatPointsPerGame(topCommander.pointsPerGame)} placement points per game.` : 'No commander data yet.',
    },
    {
      title: 'Active Win Streak',
      body: activeStreakLeader ? `${activeStreakLeader.name} is on ${activeStreakLeader.currentWinStreak} straight wins.` : 'No active player win streak right now.',
    },
  ]);

  const rankingRows = entries.map((entry, index) => {
    const streakEntry = playerStreaksByName.get(entry.name) || { currentWinStreak: 0, bestWinStreak: 0 };
    return {
      rank: index + 1,
      player: entry.name,
      rating: entry.rating,
      perfGame: entry.pointsPerGame,
      games: entry.games,
      wins: entry.wins,
      currentStreak: streakEntry.currentWinStreak,
      longestStreak: streakEntry.bestWinStreak,
      kills: entry.kills,
      firstBlood: entry.firstBloods,
      avgPlace: entry.avgPlace,
      favoriteCommander: entry.favoriteCommander || '',
      entry,
      streakEntry,
    };
  });

  rankingRows.sort((a, b) => {
    let result = 0;
    switch (sortState.column) {
      case 'rank':
        result = compareNumberValues(a.rank, b.rank);
        break;
      case 'player':
        result = compareTextValues(a.player, b.player);
        break;
      case 'rating':
        result = compareNumberValues(a.rating, b.rating);
        break;
      case 'perfGame':
        result = compareNumberValues(a.perfGame, b.perfGame);
        break;
      case 'games':
        result = compareNumberValues(a.games, b.games);
        break;
      case 'wins':
        result = compareNumberValues(a.wins, b.wins);
        break;
      case 'currentStreak':
        result = compareNumberValues(a.currentStreak, b.currentStreak);
        break;
      case 'longestStreak':
        result = compareNumberValues(a.longestStreak, b.longestStreak);
        break;
      case 'kills':
        result = compareNumberValues(a.kills, b.kills);
        break;
      case 'firstBlood':
        result = compareNumberValues(a.firstBlood, b.firstBlood);
        break;
      case 'avgPlace':
        result = compareNumberValues(a.avgPlace, b.avgPlace);
        break;
      case 'favoriteCommander':
        result = compareTextValues(a.favoriteCommander, b.favoriteCommander);
        break;
      default:
        result = compareNumberValues(a.rank, b.rank);
        break;
    }

    if (result === 0) {
      result = compareNumberValues(a.rank, b.rank);
    }

    return finalizeSortResult(result, sortState.descending);
  });

  rankingsTableBody.innerHTML = rankingRows
    .map((row) => {
      const { entry, streakEntry } = row;
      return `
      <tr>
        <td>${row.rank}</td>
        <td>${buildHistoryFilterLink(entry.name, { player: entry.name })}</td>
        <td>${formatRating(entry.rating)}</td>
        <td>${formatPointsPerGame(entry.pointsPerGame)}</td>
        <td class="rankings-games-cell">${entry.games}</td>
        <td>${entry.wins}</td>
        <td>${streakEntry.currentWinStreak}</td>
        <td>${streakEntry.bestWinStreak}</td>
        <td>${entry.kills}</td>
        <td>${entry.firstBloods}</td>
        <td>${formatAveragePlace(entry.avgPlace)}</td>
        <td>${entry.favoriteCommander ? buildHistoryFilterLink(entry.favoriteCommander, { commander: entry.favoriteCommander }) : '—'}</td>
      </tr>`;
    })
    .join('');

  updateSortableTableIndicators('rankingsMain');
}

function renderRecentTrends(games) {
  if (!recentTrendsSummary || !recentPlayerTrendsBody || !recentCommanderTrendsBody) {
    return;
  }

  const playerSortState = getTableSort('recentPlayerTrends', 'points', true);
  const commanderSortState = getTableSort('recentCommanderTrends', 'points', true);
  const recentGames = getGamesSortedByDateAscending(games).slice(-10);
  const playerEntries = buildPlayerRankingEntries(recentGames);
  const commanderEntries = buildCommanderRankingEntries(recentGames);

  if (!recentGames.length) {
    renderStatCardGroup(recentTrendsSummary, [
      { title: 'No recent trend data', body: 'Play a few games and the last-10-game view will appear here.' },
    ]);
    recentPlayerTrendsBody.innerHTML = '<tr><td colspan="7">No recent player trend data yet.</td></tr>';
    recentCommanderTrendsBody.innerHTML = '<tr><td colspan="7">No recent commander trend data yet.</td></tr>';
    return;
  }

  const hottestPlayer = playerEntries
    .filter((entry) => entry.games >= 2)
    .sort((a, b) => b.winRate - a.winRate || compareAggregateEntries(a, b))[0] || playerEntries[0] || null;
  const hottestCommander = commanderEntries
    .filter((entry) => entry.games >= 2)
    .sort((a, b) => b.winRate - a.winRate || compareAggregateEntries(a, b))[0] || commanderEntries[0] || null;
  const killLeader = playerEntries.slice().sort((a, b) => b.kills - a.kills || compareAggregateEntries(a, b))[0] || null;
  const firstBloodLeader = playerEntries.slice().sort((a, b) => b.firstBloods - a.firstBloods || compareAggregateEntries(a, b))[0] || null;

  renderStatCardGroup(recentTrendsSummary, [
    {
      title: 'Hottest Player',
      body: hottestPlayer ? `${hottestPlayer.name} is at ${formatPercent(hottestPlayer.winRate)} over the last 10 games.` : 'No player trend data yet.',
    },
    {
      title: 'Hottest Commander',
      body: hottestCommander ? `${hottestCommander.name} is at ${formatPercent(hottestCommander.winRate)} lately.` : 'No commander trend data yet.',
    },
    {
      title: 'Recent Kill Leader',
      body: killLeader ? `${killLeader.name} logged ${killLeader.kills} kills in the last 10 games.` : 'No recent kills logged yet.',
    },
    {
      title: 'First Blood Leader',
      body: firstBloodLeader ? `${firstBloodLeader.name} opened ${firstBloodLeader.firstBloods} games with first blood.` : 'No recent first blood leader yet.',
    },
  ]);

  const sortedPlayerEntries = playerEntries.slice().sort((a, b) => {
    let result = 0;
    switch (playerSortState.column) {
      case 'player':
        result = compareTextValues(a.name, b.name);
        break;
      case 'games':
        result = compareNumberValues(a.games, b.games);
        break;
      case 'points':
        result = compareNumberValues(a.points, b.points);
        break;
      case 'winRate':
        result = compareNumberValues(a.winRate, b.winRate);
        break;
      case 'kills':
        result = compareNumberValues(a.kills, b.kills);
        break;
      case 'firstBlood':
        result = compareNumberValues(a.firstBloods, b.firstBloods);
        break;
      case 'avgPlace':
        result = compareNumberValues(a.avgPlace, b.avgPlace);
        break;
      default:
        result = compareAggregateEntries(a, b);
        return playerSortState.descending ? result : -result;
    }
    if (result === 0) {
      result = compareAggregateEntries(a, b);
      return playerSortState.descending ? result : -result;
    }
    return finalizeSortResult(result, playerSortState.descending);
  });

  recentPlayerTrendsBody.innerHTML = (sortedPlayerEntries.length ? sortedPlayerEntries : [])
    .map((entry) => `
      <tr>
        <td>${buildHistoryFilterLink(entry.name, { player: entry.name })}</td>
        <td>${entry.games}</td>
        <td>${entry.points}</td>
        <td>${formatPercent(entry.winRate)}</td>
        <td>${entry.kills}</td>
        <td>${entry.firstBloods}</td>
        <td>${formatAveragePlace(entry.avgPlace)}</td>
      </tr>`)
    .join('') || '<tr><td colspan="7">No recent player trend data yet.</td></tr>';

  const sortedCommanderEntries = commanderEntries.slice().sort((a, b) => {
    let result = 0;
    switch (commanderSortState.column) {
      case 'commander':
        result = compareTextValues(a.name, b.name);
        break;
      case 'games':
        result = compareNumberValues(a.games, b.games);
        break;
      case 'points':
        result = compareNumberValues(a.points, b.points);
        break;
      case 'winRate':
        result = compareNumberValues(a.winRate, b.winRate);
        break;
      case 'kills':
        result = compareNumberValues(a.kills, b.kills);
        break;
      case 'firstBlood':
        result = compareNumberValues(a.firstBloods, b.firstBloods);
        break;
      case 'avgPlace':
        result = compareNumberValues(a.avgPlace, b.avgPlace);
        break;
      default:
        result = compareAggregateEntries(a, b);
        return commanderSortState.descending ? result : -result;
    }
    if (result === 0) {
      result = compareAggregateEntries(a, b);
      return commanderSortState.descending ? result : -result;
    }
    return finalizeSortResult(result, commanderSortState.descending);
  });

  recentCommanderTrendsBody.innerHTML = (sortedCommanderEntries.length ? sortedCommanderEntries : [])
    .map((entry) => `
      <tr>
        <td>${buildHistoryFilterLink(entry.name, { commander: entry.name })}</td>
        <td>${entry.games}</td>
        <td>${entry.points}</td>
        <td>${formatPercent(entry.winRate)}</td>
        <td>${entry.kills}</td>
        <td>${entry.firstBloods}</td>
        <td>${formatAveragePlace(entry.avgPlace)}</td>
      </tr>`)
    .join('') || '<tr><td colspan="7">No recent commander trend data yet.</td></tr>';

  updateSortableTableIndicators('recentPlayerTrends');
  updateSortableTableIndicators('recentCommanderTrends');
}

function renderStreaks(games) {
  if (!streaksSummary || !playerStreaksBody || !commanderStreaksBody) {
    return;
  }

  const playerSortState = getTableSort('playerStreaks', 'currentWins', true);
  const commanderSortState = getTableSort('commanderStreaks', 'currentWins', true);
  const playerEntries = buildStreakEntries(games, (row) => row.player);
  const commanderEntries = buildStreakEntries(games, (row) => row.commander);

  if (!playerEntries.length && !commanderEntries.length) {
    renderStatCardGroup(streaksSummary, [
      { title: 'No streak data yet', body: 'Streaks appear once you have saved a few games.' },
    ]);
    playerStreaksBody.innerHTML = '<tr><td colspan="6">No player streaks available yet.</td></tr>';
    commanderStreaksBody.innerHTML = '<tr><td colspan="6">No commander streaks available yet.</td></tr>';
    return;
  }

  const activePlayerStreak = playerEntries.find((entry) => entry.currentWinStreak > 0) || null;
  const bestPlayerStreak = playerEntries.slice().sort((a, b) => b.bestWinStreak - a.bestWinStreak || a.name.localeCompare(b.name))[0] || null;
  const activeCommanderStreak = commanderEntries.find((entry) => entry.currentWinStreak > 0) || null;
  const longestDrought = playerEntries.slice().sort((a, b) => b.drought - a.drought || a.name.localeCompare(b.name))[0] || null;

  renderStatCardGroup(streaksSummary, [
    {
      title: 'Active Player Streak',
      body: activePlayerStreak ? `${activePlayerStreak.name} has ${activePlayerStreak.currentWinStreak} consecutive wins.` : 'No active player win streak right now.',
    },
    {
      title: 'Best Player Run',
      body: bestPlayerStreak ? `${bestPlayerStreak.name} peaked at ${bestPlayerStreak.bestWinStreak} straight wins.` : 'No player win streaks yet.',
    },
    {
      title: 'Active Commander Streak',
      body: activeCommanderStreak ? `${activeCommanderStreak.name} has ${activeCommanderStreak.currentWinStreak} consecutive wins.` : 'No active commander streak right now.',
    },
    {
      title: 'Longest Drought',
      body: longestDrought ? `${longestDrought.name} has gone ${longestDrought.drought} appearances without a win.` : 'No drought data yet.',
    },
  ]);

  const sortedPlayerEntries = playerEntries.slice().sort((a, b) => {
    let result = 0;
    switch (playerSortState.column) {
      case 'player':
        result = compareTextValues(a.name, b.name);
        break;
      case 'games':
        result = compareNumberValues(a.games, b.games);
        break;
      case 'currentWins':
        result = compareNumberValues(a.currentWinStreak, b.currentWinStreak);
        break;
      case 'bestWinStreak':
        result = compareNumberValues(a.bestWinStreak, b.bestWinStreak);
        break;
      case 'drought':
        result = compareNumberValues(a.drought, b.drought);
        break;
      case 'lastWin':
        result = compareDateValues(a.lastWinDate, b.lastWinDate);
        break;
      default:
        result = compareNumberValues(a.currentWinStreak, b.currentWinStreak);
        break;
    }
    if (result === 0) {
      result = compareTextValues(a.name, b.name);
    }
    return finalizeSortResult(result, playerSortState.descending);
  });

  playerStreaksBody.innerHTML = sortedPlayerEntries.length
    ? sortedPlayerEntries.map((entry) => `
        <tr>
          <td>${buildHistoryFilterLink(entry.name, { player: entry.name })}</td>
          <td>${entry.games}</td>
          <td>${entry.currentWinStreak}</td>
          <td>${entry.bestWinStreak}</td>
          <td>${entry.drought}</td>
          <td>${escapeHtml(entry.lastWinDate || '—')}</td>
        </tr>`).join('')
    : '<tr><td colspan="6">No player streaks available yet.</td></tr>';

  const sortedCommanderEntries = commanderEntries.slice().sort((a, b) => {
    let result = 0;
    switch (commanderSortState.column) {
      case 'commander':
        result = compareTextValues(a.name, b.name);
        break;
      case 'games':
        result = compareNumberValues(a.games, b.games);
        break;
      case 'currentWins':
        result = compareNumberValues(a.currentWinStreak, b.currentWinStreak);
        break;
      case 'bestWinStreak':
        result = compareNumberValues(a.bestWinStreak, b.bestWinStreak);
        break;
      case 'drought':
        result = compareNumberValues(a.drought, b.drought);
        break;
      case 'lastWin':
        result = compareDateValues(a.lastWinDate, b.lastWinDate);
        break;
      default:
        result = compareNumberValues(a.currentWinStreak, b.currentWinStreak);
        break;
    }
    if (result === 0) {
      result = compareTextValues(a.name, b.name);
    }
    return finalizeSortResult(result, commanderSortState.descending);
  });

  commanderStreaksBody.innerHTML = sortedCommanderEntries.length
    ? sortedCommanderEntries.map((entry) => `
        <tr>
          <td>${buildHistoryFilterLink(entry.name, { commander: entry.name })}</td>
          <td>${entry.games}</td>
          <td>${entry.currentWinStreak}</td>
          <td>${entry.bestWinStreak}</td>
          <td>${entry.drought}</td>
          <td>${escapeHtml(entry.lastWinDate || '—')}</td>
        </tr>`).join('')
    : '<tr><td colspan="6">No commander streaks available yet.</td></tr>';

  updateSortableTableIndicators('playerStreaks');
  updateSortableTableIndicators('commanderStreaks');
}

function renderRankingsPage(games) {
  renderPodRankings(games);
  renderRecentTrends(games);
  renderStreaks(games);
}

function getPlayerStatsData(games) {
  const stats = {};

  games.forEach((game) => {
    const rows = getGameRows(game);
    const winner = Array.isArray(game.finishOrder) && game.finishOrder.length ? game.finishOrder[0] : null;
    const firstBlood = getGameFirstBloodInfo(game);

    if (firstBlood?.actorPlayer) {
      const firstBloodStat = ensurePlayerStats(stats, firstBlood.actorPlayer);
      firstBloodStat.firstBloods = (firstBloodStat.firstBloods || 0) + 1;
    }

    rows.forEach((row) => {
      const player = (row.player || '').trim();
      if (!player) {
        return;
      }

      const playerStat = ensurePlayerStats(stats, player);
      playerStat.games += 1;
      if (winner === player) {
        playerStat.wins += 1;
      }

      const commander = (row.commander || '').trim();
      if (commander) {
        playerStat.commanders[commander] = (playerStat.commanders[commander] || 0) + 1;
        if (!playerStat.commanderStats[commander]) {
          playerStat.commanderStats[commander] = { played: 0, wins: 0 };
        }
        playerStat.commanderStats[commander].played += 1;
        if (winner === player) {
          playerStat.commanderStats[commander].wins += 1;
        }
      }

      const killedList = getCleanKilledList(row.killed);
      const killsCount = typeof row.kills === 'number' && !Number.isNaN(row.kills)
        ? row.kills
        : killedList.length;
      playerStat.kills += killsCount;

      killedList.forEach((target) => {
        if (!target) {
          return;
        }

        playerStat.victimCounts[target] = (playerStat.victimCounts[target] || 0) + 1;
        const targetStat = ensurePlayerStats(stats, target);
        targetStat.killerCounts[player] = (targetStat.killerCounts[player] || 0) + 1;
      });
    });
  });

  return stats;
}

function renderPlayerStats(games) {
  if (!playerStatsTableBody) {
    return;
  }

  const sortState = getTableSort('playerStats', 'games', true);
  const stats = getPlayerStatsData(games);
  const players = Object.keys(stats);

  if (!players.length) {
    playerStatsTableBody.innerHTML = '<tr><td colspan="11">No player stats available.</td></tr>';
    return;
  }

  const rows = players
    .map((player) => {
      const stat = stats[player];
      const winRateValue = stat.games ? (stat.wins / stat.games) * 100 : 0;
      const favoriteCommander = getMaxCountKey(stat.commanders);
      const nemesis = getMaxCountKey(stat.killerCounts);
      const victim = getMaxCountKey(stat.victimCounts);
      const killAverage = stat.games ? (stat.kills / stat.games).toFixed(1) : '0.0';

      let bestDeck = '—';
      Object.entries(stat.commanderStats).forEach(([commander, data]) => {
        if (!data.played) {
          return;
        }
        const successRate = data.wins / data.played;
        const currentBest = bestDeck === '—' ? null : bestDeck;
        if (currentBest === null) {
          bestDeck = `${commander} (${formatPercent(successRate * 100)} from ${data.played})`;
          return;
        }

        const [, bestRatePart] = currentBest.match(/\(([-\d.]+)%/) || [null, '0'];
        const bestRate = Number(bestRatePart);
        if (successRate * 100 > bestRate || (successRate * 100 === bestRate && data.played > Number(currentBest.match(/from (\d+)/)?.[1] || 0))) {
          bestDeck = `${commander} (${formatPercent(successRate * 100)} from ${data.played})`;
        }
      });

      return {
        player,
        games: stat.games,
        wins: stat.wins,
        winRate: winRateValue,
        firstBloods: stat.firstBloods || 0,
        favoriteCommander: favoriteCommander || '',
        nemesis: nemesis || '',
        victim: victim || '',
        kills: stat.kills,
        kd: Number.parseFloat(killAverage),
        bestDeck,
      };
    });

  rows.sort((a, b) => {
    let result = 0;
    switch (sortState.column) {
      case 'player':
        result = compareTextValues(a.player, b.player);
        break;
      case 'games':
        result = compareNumberValues(a.games, b.games);
        break;
      case 'wins':
        result = compareNumberValues(a.wins, b.wins);
        break;
      case 'winRate':
        result = compareNumberValues(a.winRate, b.winRate);
        break;
      case 'firstBlood':
        result = compareNumberValues(a.firstBloods, b.firstBloods);
        break;
      case 'favoriteCommander':
        result = compareTextValues(a.favoriteCommander, b.favoriteCommander);
        break;
      case 'nemesis':
        result = compareTextValues(a.nemesis, b.nemesis);
        break;
      case 'victim':
        result = compareTextValues(a.victim, b.victim);
        break;
      case 'kills':
        result = compareNumberValues(a.kills, b.kills);
        break;
      case 'kd':
        result = compareNumberValues(a.kd, b.kd);
        break;
      case 'bestDeck':
        result = compareTextValues(a.bestDeck, b.bestDeck);
        break;
      default:
        result = compareNumberValues(a.games, b.games);
        break;
    }

    if (result === 0) {
      result = compareTextValues(a.player, b.player);
    }

    return finalizeSortResult(result, sortState.descending);
  });

  const html = rows.map((row) => {
      return `
        <tr>
          <td>${buildHistoryFilterLink(row.player, { player: row.player })}</td>
          <td>${row.games}</td>
          <td>${row.wins}</td>
          <td>${formatPercent(row.winRate)}</td>
          <td>${row.firstBloods}</td>
          <td>${row.favoriteCommander ? buildHistoryFilterLink(row.favoriteCommander, { commander: row.favoriteCommander }) : '—'}</td>
          <td>${row.nemesis ? buildHistoryFilterLink(row.nemesis, { player: row.nemesis }) : '—'}</td>
          <td>${row.victim ? buildHistoryFilterLink(row.victim, { player: row.victim }) : '—'}</td>
          <td>${row.kills}</td>
          <td>${row.kd.toFixed(1)}</td>
          <td>${row.bestDeck}</td>
        </tr>`;
    })
    .join('');

  playerStatsTableBody.innerHTML = html;
  updateSortableTableIndicators('playerStats');
}

function getCommanderStatsData(games) {
  const stats = {};

  games.forEach((game) => {
    const firstBlood = getGameFirstBloodInfo(game);
    getGameRows(game).forEach((row) => {
      const commander = (row.commander || '').trim();
      if (!commander) {
        return;
      }

      if (!stats[commander]) {
        stats[commander] = { games: 0, wins: 0, kills: 0, firstBloods: 0, placementTotal: 0, placementScoreTotal: 0, placementGames: 0 };
      }

      stats[commander].games += 1;
      if (Array.isArray(game.finishOrder) && game.finishOrder[0] === row.player) {
        stats[commander].wins += 1;
      }

      if (Array.isArray(game.finishOrder) && game.finishOrder.length) {
        const place = game.finishOrder.indexOf(row.player) + 1;
        if (place > 0) {
          stats[commander].placementTotal += place;
          stats[commander].placementGames += 1;

          const playerCount = game.finishOrder.length;
          const placementScore = playerCount > 1 ? 1 - ((place - 1) / (playerCount - 1)) : 1;
          stats[commander].placementScoreTotal += placementScore;
        }
      }

      const kills = typeof row.kills === 'number' && !Number.isNaN(row.kills) ? row.kills : 0;
      stats[commander].kills += kills;

      if (firstBlood?.actorPlayer === row.player) {
        stats[commander].firstBloods += 1;
      }
    });
  });

  return stats;
}

function getCommanderActualPower(commanderStats) {
  const entries = Object.entries(commanderStats);
  if (!entries.length) {
    return {};
  }

  const maxKillsPerGame = Math.max(...entries.map(([, stat]) => (stat.games ? stat.kills / stat.games : 0)));
  const actual = {};

  entries.forEach(([commander, stat]) => {
    const winRate = stat.games ? stat.wins / stat.games : 0;
    const killsPerGame = stat.games ? stat.kills / stat.games : 0;
    const killScore = maxKillsPerGame ? killsPerGame / maxKillsPerGame : 0;
    const placementScore = stat.placementGames ? stat.placementScoreTotal / stat.placementGames : 0;
    const rawScore = 0.5 * winRate + 0.25 * killScore + 0.25 * placementScore;
    actual[commander] = Math.round(rawScore * 100) / 10;
  });

  return actual;
}

function getCalibratedCommanderActualPower(relativeActualPowers, commanderStats) {
  const commanders = Object.keys(relativeActualPowers);
  const expectedByCommander = {};

  commanders.forEach((commander) => {
    expectedByCommander[commander] = getCommanderExpectedPower(commander);
  });

  const overlap = commanders.filter((commander) => typeof expectedByCommander[commander] === 'number');
  const calibrated = {};

  if (overlap.length >= 2) {
    const relativeValues = overlap.map((commander) => relativeActualPowers[commander]);
    const expectedValues = overlap.map((commander) => expectedByCommander[commander]);
    const relativeMean = getMean(relativeValues);
    const expectedMean = getMean(expectedValues);
    const relativeStd = getStandardDeviation(relativeValues, relativeMean);
    const expectedStd = getStandardDeviation(expectedValues, expectedMean);

    commanders.forEach((commander) => {
      const relative = relativeActualPowers[commander];
      const normalized = relativeStd > 0 ? (relative - relativeMean) / relativeStd : 0;
      const targetStd = expectedStd > 0 ? expectedStd : 1;
      const value = expectedMean + (normalized * targetStd);
      calibrated[commander] = Math.round(clamp(value, 0, 10) * 10) / 10;
    });

    return calibrated;
  }

  if (overlap.length === 1) {
    const singleExpected = expectedByCommander[overlap[0]];
    const singleRelative = relativeActualPowers[overlap[0]];
    const shift = singleExpected - singleRelative;

    commanders.forEach((commander) => {
      const value = relativeActualPowers[commander] + shift;
      calibrated[commander] = Math.round(clamp(value, 0, 10) * 10) / 10;
    });

    return calibrated;
  }

  commanders.forEach((commander) => {
    calibrated[commander] = relativeActualPowers[commander];
  });

  return calibrated;
}

function renderCommanderStats(games) {
  if (!commanderStatsTableBody) {
    return;
  }

  const sortState = getTableSort('commanderStats', 'games', true);
  const searchTerm = commanderSearch?.value.trim().toLowerCase() || '';
  const commanderStats = getCommanderStatsData(games);
  const actualPowers = getCommanderActualPower(commanderStats);
  const calibratedPowers = getCalibratedCommanderActualPower(actualPowers, commanderStats);
  
  let entries = Object.entries(commanderStats)
    .filter(([commander]) => commander.toLowerCase().includes(searchTerm))
    .map(([commander, stat]) => {
      const winRateValue = stat.games ? (stat.wins / stat.games) * 100 : 0;
      const killsPerGame = stat.games ? stat.kills / stat.games : 0;
      const averagePlacement = stat.placementGames ? stat.placementTotal / stat.placementGames : 0;
      return {
        commander,
        games: stat.games,
        wins: stat.wins,
        winRate: winRateValue,
        kills: stat.kills,
        firstBloods: stat.firstBloods,
        kd: killsPerGame,
        averagePlacement,
        expected: getCommanderExpectedPower(commander),
        actual: typeof actualPowers[commander] === 'number' ? actualPowers[commander] : 0,
        actualCal: typeof calibratedPowers[commander] === 'number' ? calibratedPowers[commander] : 0,
        delta: 0,
        confidence: getCommanderConfidence(stat.games),
        stat,
        actualPowers
      };
    })
    .map((entry) => ({
      ...entry,
      delta: typeof entry.expected === 'number' ? entry.actualCal - entry.expected : 0,
    }));

  entries.sort((a, b) => {
    let aVal, bVal;
    switch (sortState.column) {
      case 'commander':
        aVal = a.commander.toLowerCase();
        bVal = b.commander.toLowerCase();
        return sortState.descending ? bVal.localeCompare(aVal) : aVal.localeCompare(bVal);
      case 'games':
        aVal = a.games;
        bVal = b.games;
        break;
      case 'wins':
        aVal = a.wins;
        bVal = b.wins;
        break;
      case 'winRate':
        aVal = a.winRate;
        bVal = b.winRate;
        break;
      case 'kills':
        aVal = a.kills;
        bVal = b.kills;
        break;
      case 'firstBloods':
        aVal = a.firstBloods;
        bVal = b.firstBloods;
        break;
      case 'kd':
        aVal = a.kd;
        bVal = b.kd;
        break;
      case 'avgPlace':
        aVal = a.averagePlacement;
        bVal = b.averagePlacement;
        break;
      case 'expected':
        aVal = typeof a.expected === 'number' ? a.expected : 0;
        bVal = typeof b.expected === 'number' ? b.expected : 0;
        break;
      case 'actual':
        aVal = a.actual;
        bVal = b.actual;
        break;
      case 'actualCal':
        aVal = a.actualCal;
        bVal = b.actualCal;
        break;
      case 'delta':
        aVal = a.delta;
        bVal = b.delta;
        break;
      case 'confidence':
        aVal = a.confidence.score;
        bVal = b.confidence.score;
        break;
      default:
        return 0;
    }
    const diff = aVal - bVal;
    return sortState.descending ? -diff : diff;
  });

  const rows = entries
    .map(({ commander, stat, actual, actualCal, expected, delta, confidence }) => {
      const winRateValue = stat.games ? (stat.wins / stat.games) * 100 : 0;
      const killsPerGame = stat.games ? stat.kills / stat.games : 0;
      const averagePlacement = stat.placementGames ? stat.placementTotal / stat.placementGames : 0;
      const roundedExpected = typeof expected === 'number' ? expected.toFixed(1) : '';
      const actualPower = typeof actual === 'number' ? actual.toFixed(1) : '0.0';
      const actualCalibrated = typeof actualCal === 'number' ? actualCal.toFixed(1) : '0.0';

      let deltaClass = 'delta-neutral';
      if (delta > 0.15) {
        deltaClass = 'delta-positive';
      } else if (delta < -0.15) {
        deltaClass = 'delta-negative';
      }

      const deltaText = typeof expected === 'number' ? `${delta > 0 ? '+' : ''}${delta.toFixed(1)}` : '—';

      return `
        <tr>
          <td>${buildHistoryFilterLink(commander, { commander })}</td>
          <td>${stat.games}</td>
          <td>${stat.wins}</td>
          <td>${formatPercent(winRateValue)}</td>
          <td>${stat.kills}</td>
          <td>${stat.firstBloods || 0}</td>
          <td>${killsPerGame.toFixed(1)}</td>
          <td>${averagePlacement ? averagePlacement.toFixed(2) : '—'}</td>
          <td>
            <input
              type="number"
              class="commander-expected-input"
              data-commander="${escapeHtml(commander)}"
              min="0"
              max="10"
              step="0.1"
              value="${roundedExpected}"
              placeholder="0.0"
            />
          </td>
          <td>${actualPower}</td>
          <td>${actualCalibrated}</td>
          <td><span class="delta-pill ${deltaClass}">${deltaText}</span></td>
          <td>
            <div class="confidence-cell">
              <span class="confidence-label confidence-${confidence.levelClass}">${confidence.label}</span>
              <div class="confidence-bar"><span style="width: ${confidence.percent}%;"></span></div>
            </div>
          </td>
        </tr>`;
    })
    .join('');

  commanderStatsTableBody.innerHTML = rows || '<tr><td colspan="13">No commanders match your search.</td></tr>';
  updateSortableTableIndicators('commanderStats');
}

function handleCommanderExpectedInput(event) {
  const input = event.target.closest('.commander-expected-input');
  if (!input) {
    return;
  }

  const commander = input.dataset.commander;
  const value = parseFloat(input.value);
  if (Number.isNaN(value)) {
    setCommanderExpectedPower(commander, null);
  } else {
    setCommanderExpectedPower(commander, value);
    input.value = Math.round(value * 10) / 10;
  }
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

    const playerCommandersList = game.playerCommanders || (game.playerRows || []).map(({ player, commander }) => ({ player, commander }));
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
    .map(([commander, count]) => `${commander} (${count})`)
    .join(', ');

  summaryEl.appendChild(createStatCard('Total games', total));
  summaryEl.appendChild(createStatCard('Most recent game', mostRecent));
  summaryEl.appendChild(createStatCard('Top winners', bestPlayers || 'No winners yet')); 
  summaryEl.appendChild(createStatCard('Popular commanders', bestCommanders || 'No commander data yet'));
}

function refresh() {
  const games = loadGames();
  updateFormDatalists(games);
  renderIdentityRenameOptions();
  renderSummary(games);
  renderRankingsPage(games);
  renderPlayerStats(games);
  updateHistoryFilters(games);
  applyHistoryQueryFilters();
  renderHistory(games);
  renderCommanderStats(games);
  renderDeckLookup();
  renderDeckLists();
  renderDeckSelector();
  renderRecords();
  refreshLiveTrackerUi();
  initializeMobileSortControls();
  applyResponsiveTableLabels();
}

function resetForm() {
  form.reset();
  resetPlayerTable();
  dateInput.valueAsDate = new Date();
}

if (addPlayerRowButton) {
  addPlayerRowButton.addEventListener('click', () => {
    addPlayerRow();
  });
}

if (removePlayerRowButton) {
  removePlayerRowButton.addEventListener('click', () => {
    removePlayerRow();
  });
}

if (playerTableBody) {
  playerTableBody.addEventListener('input', (event) => {
    const input = event.target.closest('input[name="player"], input[name="commander"]');
    if (!input) {
      return;
    }

    const row = input.closest('tr');
    if (!row) {
      return;
    }

    // Update dropdown menus as user types
    populateRowSelectors(row);
  });
}

if (liveGamePlayerBody) {
  liveGamePlayerBody.addEventListener('input', () => {
    updateLiveSetupSeatLabels();
    renderLiveOrderPreview();
  });
}

if (form) {
  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const games = loadGames();
    const existingGame = editingGameId ? games.find((game) => game.id === editingGameId) || null : null;

    const rows = getPlayerRows();
    if (!rows.length) {
      await promptLiveAlert('Please add at least one player row with a player name.', 'Unable to save game');
      return;
    }

    const firstBloodPlayer = firstBloodPlayerInput?.value.trim() || '';
    const firstBloodTurn = parseOptionalPositiveInteger(firstBloodTurnInput?.value || '');
    const winningTurn = parseOptionalPositiveInteger(winningTurnInput?.value || '');

    const validationError = validateManualGameEntry(rows, {
      firstBloodPlayer,
      firstBloodTurn,
      winningTurn,
    });

    if (validationError) {
      await promptLiveAlert(validationError, 'Unable to save game');
      return;
    }

    const sanitizedRows = sanitizeManualGameRows(rows);

    const players = sanitizedRows.map(({ player }) => player);
    const playerCommanders = sanitizedRows.map(({ player, commander }) => ({ player, commander }));
    const finishOrder = sanitizedRows
      .slice()
      .filter((row) => row.place !== null)
      .sort((a, b) => (a.place || 999) - (b.place || 999))
      .map((row) => row.player);
    const manualLiveSummary = buildManualGameLiveSummary(
      sanitizedRows,
      dateInput.value || new Date().toISOString().slice(0, 10),
      firstBloodPlayer,
      firstBloodTurn,
      winningTurn,
    );

    const newGame = {
      id: editingGameId || generateId(),
      date: dateInput.value || new Date().toISOString().slice(0, 10),
      playerRows: sanitizedRows,
      players,
      playerCommanders,
      finishOrder,
      liveSummary: mergeEditedGameLiveSummary(existingGame?.liveSummary, manualLiveSummary),
      notes: notesInput.value.trim(),
    };

    if (editingGameId) {
      const index = games.findIndex((game) => game.id === editingGameId);
      if (index >= 0) {
        games[index] = newGame;
      } else {
        games.push(newGame);
      }
    } else {
      games.push(newGame);
    }

    saveGames(games);
    resetEditMode();
    refresh();
    resetForm();
  });
}

if (liveAddPlayerButton) {
  liveAddPlayerButton.addEventListener('click', () => {
    addLiveSetupRow();
    renderLiveOrderPreview();
  });
}

if (liveRemovePlayerButton) {
  liveRemovePlayerButton.addEventListener('click', () => {
    removeLiveSetupRow();
    renderLiveOrderPreview();
  });
}

if (liveRandomizeFirstButton) {
  liveRandomizeFirstButton.addEventListener('click', () => {
    randomizeLiveFirstPlayer();
  });
}

if (liveGameForm) {
  liveGameForm.addEventListener('submit', (event) => {
    event.preventDefault();
    startLiveGame();
  });
}

if (liveSourceOptions) {
  liveSourceOptions.addEventListener('click', (event) => {
    const button = event.target.closest('[data-source-id]');
    if (!button) {
      return;
    }

    hideLiveSourcePrompt(button.dataset.sourceId || '');
  });
}

if (liveSourceCancelButton) {
  liveSourceCancelButton.addEventListener('click', () => {
    hideLiveSourcePrompt(null);
  });
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

if (livePlayerGrid) {
  livePlayerGrid.addEventListener('click', (event) => {
    const toggle = event.target.closest('[data-action="toggle-cannot-lose"]');
    if (toggle) {
      setPlayerCannotLoseState(toggle.dataset.playerId || '', Boolean(toggle.checked));
      return;
    }

    const commanderDamageButton = event.target.closest('[data-action="manual-commander-damage"]');
    if (commanderDamageButton) {
      applyCommanderDamageToPlayer(commanderDamageButton.dataset.playerId || '');
      return;
    }

    const manualLifeEntryButton = event.target.closest('[data-action="manual-life-entry"]');
    if (manualLifeEntryButton) {
      applyManualLifeEntry(manualLifeEntryButton.dataset.playerId || '');
      return;
    }

    const eliminateButton = event.target.closest('[data-action="manual-eliminate"]');
    if (eliminateButton) {
      manuallyEliminatePlayer(eliminateButton.dataset.playerId || '');
      return;
    }

    const autoWinButton = event.target.closest('[data-action="auto-win"]');
    if (autoWinButton) {
      markPlayerAutomaticWinner(autoWinButton.dataset.playerId || '');
      return;
    }

    const button = event.target.closest('[data-action="adjust-life"]');
    if (!button) {
      return;
    }

    if (liveHoldRepeated) {
      liveHoldRepeated = false;
      return;
    }

    const playerId = button.dataset.playerId || '';
    const delta = parseInt(button.dataset.delta || '0', 10);
    if (!playerId || Number.isNaN(delta) || delta === 0) {
      return;
    }

    applyQuickLifeChange(playerId, delta);
  });

  livePlayerGrid.addEventListener('pointerdown', (event) => {
    const button = event.target.closest('[data-action="adjust-life"]');
    if (!button) {
      return;
    }

    startLiveHoldRepeat(button);
  });

  ['pointerup', 'pointerleave', 'pointercancel'].forEach((eventName) => {
    livePlayerGrid.addEventListener(eventName, () => {
      stopLiveHoldRepeat();
    });
  });

  ['pointerup', 'pointercancel'].forEach((eventName) => {
    document.addEventListener(eventName, () => {
      stopLiveHoldRepeat();
    });
  });
}

if (liveActionsToggleButton) {
  liveActionsToggleButton.addEventListener('click', () => {
    toggleLiveActionsMenu();
  });
}

if (liveBackButton) {
  liveBackButton.addEventListener('click', () => {
    closeLiveActionsMenu();
    if (window.history.length > 1) {
      window.history.back();
      return;
    }
    window.location.href = 'index.html';
  });
}

if (liveUndoButton) {
  liveUndoButton.addEventListener('click', () => {
    closeLiveActionsMenu();
    undoLastLiveAction();
  });
}

if (liveSwapLifeButton) {
  liveSwapLifeButton.addEventListener('click', () => {
    closeLiveActionsMenu();
    swapLivePlayerLifeTotals();
  });
}

if (liveFinishGameButton) {
  liveFinishGameButton.addEventListener('click', () => {
    closeLiveActionsMenu();
    completeActiveGame();
  });
}

if (liveAbandonGameButton) {
  liveAbandonGameButton.addEventListener('click', () => {
    closeLiveActionsMenu();
    abandonActiveGame();
  });
}

document.addEventListener('click', (event) => {
  if (!liveActiveActions || !liveActiveActions.classList.contains('is-open')) {
    return;
  }

  if (!event.target.closest('#live-active-actions')) {
    closeLiveActionsMenu();
  }
});

document.addEventListener('keydown', (event) => {
  if (event.key !== 'Escape') {
    return;
  }

  if (liveModalPrompt && !liveModalPrompt.hidden) {
    hideLiveModal(null);
    return;
  }

  if (liveSourcePrompt && !liveSourcePrompt.hidden) {
    hideLiveSourcePrompt(null);
  }
});

if (historySortSelect) {
  historySortSelect.addEventListener('change', (event) => {
    historySortKey = event.target.value;
    updateHistorySortOrderLabel();
    renderHistory(loadGames());
  });
}

if (historyList) {
  historyList.addEventListener('click', handleHistoryAction);
}

if (historySortOrderButton) {
  historySortOrderButton.addEventListener('click', () => {
    historySortDescending = !historySortDescending;
    updateHistorySortOrderLabel();
    renderHistory(loadGames());
  });
}

if (historyFilterWinner) {
  historyFilterWinner.addEventListener('change', () => {
    renderHistory(loadGames());
  });
}

if (historyFilterCommander) {
  historyFilterCommander.addEventListener('change', () => {
    renderHistory(loadGames());
  });
}

if (historyFilterPlayer) {
  historyFilterPlayer.addEventListener('change', () => {
    renderHistory(loadGames());
  });
}

if (historyFilterDateFrom) {
  historyFilterDateFrom.addEventListener('change', () => {
    renderHistory(loadGames());
  });
}

if (historyFilterDateTo) {
  historyFilterDateTo.addEventListener('change', () => {
    renderHistory(loadGames());
  });
}

if (historyResetFiltersButton) {
  historyResetFiltersButton.addEventListener('click', () => {
    resetHistoryFilters();
    renderHistory(loadGames());
  });
}

if (commanderSearch) {
  commanderSearch.addEventListener('input', () => {
    renderCommanderStats(loadGames());
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

if (commanderStatsTableBody) {
  commanderStatsTableBody.addEventListener('change', handleCommanderExpectedInput);
}

if (deckListTableBody) {
  deckListTableBody.addEventListener('click', handleDeckListTableAction);
}

if (deckLookupSelect) {
  deckLookupSelect.addEventListener('change', () => {
    renderDeckLookup();
  });
}

if (deckListCancelButton) {
  deckListCancelButton.addEventListener('click', () => {
    resetDeckListForm();
  });
}

if (deckListForm && deckCommanderInput && deckUrlInput) {
  deckListForm.addEventListener('submit', async (event) => {
    event.preventDefault();

    const commander = deckCommanderInput.value.trim();
    const owner = deckOwnerInput?.value.trim() || '';
    const url = deckUrlInput.value.trim();

    if (!commander) {
      await promptLiveAlert('Please choose or enter a commander.', 'Unable to save deck list');
      return;
    }

    if (!url) {
      await promptLiveAlert('Please enter a deck URL.', 'Unable to save deck list');
      return;
    }

    let parsedUrl;
    try {
      parsedUrl = new URL(url);
    } catch (error) {
      await promptLiveAlert('Please enter a valid URL (include https://).', 'Unable to save deck list');
      return;
    }

    if (!['http:', 'https:'].includes(parsedUrl.protocol)) {
      await promptLiveAlert('Deck URL must use http or https.', 'Unable to save deck list');
      return;
    }

    const deckLists = loadDeckLists()
      .map(normalizeDeckListEntry)
      .filter(Boolean);

    const normalizedUrl = parsedUrl.toString();
    const normalizedCommander = commander.toLowerCase();
    const duplicateCommanderIndex = deckLists.findIndex((entry) => {
      if (editingDeckListId && entry.id === editingDeckListId) {
        return false;
      }
      return entry.commander.toLowerCase() === normalizedCommander;
    });

    if (editingDeckListId) {
      const index = deckLists.findIndex((entry) => entry.id === editingDeckListId);
      if (index >= 0) {
        deckLists[index] = { id: editingDeckListId, commander, owner, url: normalizedUrl };
      } else {
        deckLists.push({ id: generateId(), commander, owner, url: normalizedUrl });
      }

      if (duplicateCommanderIndex >= 0) {
        deckLists.splice(duplicateCommanderIndex, 1);
      }
    } else {
      if (duplicateCommanderIndex >= 0) {
        deckLists[duplicateCommanderIndex] = {
          ...deckLists[duplicateCommanderIndex],
          commander,
          owner,
          url: normalizedUrl,
        };
      } else {
        deckLists.push({ id: generateId(), commander, owner, url: normalizedUrl });
      }
    }

    saveDeckLists(deckLists);
    resetDeckListForm();
    refresh();
  });
}

if (deckSelectorForm) {
  deckSelectorForm.addEventListener('submit', (event) => {
    event.preventDefault();
    renderDeckSelectorAssignments(getSelectedDeckSelectorOwners());
  });
}

if (recordsForm) {
  recordsForm.addEventListener('submit', (event) => {
    event.preventDefault();

    const records = collectRecordsFromTable();
    saveRecords(records);
    renderRecords();
  });
}

if (customRecordForm) {
  customRecordForm.addEventListener('submit', async (event) => {
    event.preventDefault();

    const title = customRecordTitleInput?.value.trim() || '';
    if (!title) {
      await promptLiveAlert('Please enter a title for the custom record.', 'Unable to add custom record');
      return;
    }

    const records = collectRecordsFromTable();
    records.push({
      id: generateId(),
      key: '',
      title,
      unit: customRecordUnitInput?.value.trim() || '',
      value: customRecordValueInput?.value.trim() || '',
      holder: customRecordHolderInput?.value.trim() || '',
      commander: customRecordCommanderInput?.value.trim() || '',
      date: customRecordDateInput?.value || '',
      notes: customRecordNotesInput?.value.trim() || '',
      isCustom: true,
    });

    saveRecords(records);
    customRecordForm.reset();
    renderRecords();
  });
}

document.querySelectorAll('[data-sort-table]').forEach((table) => {
  table.addEventListener('click', handleSortableHeaderClick);
  table.addEventListener('keydown', handleSortableHeaderKeydown);
});

updateHistorySortOrderLabel();

window.addEventListener('storage', (event) => {
  if (
    event.key === STORAGE_KEY
    || event.key === EXPECTED_POWER_STORAGE_KEY
    || event.key === DECK_LIST_STORAGE_KEY
    || event.key === RECORDS_STORAGE_KEY
    || event.key === ACTIVE_GAME_STORAGE_KEY
    || event.key === ACTIVE_GAME_UNDO_STORAGE_KEY
  ) {
    appState = loadLocalState();
    activeGameState = loadActiveGameState();
    activeGameUndoState = loadActiveGameUndoState();
    refresh();
  }
});

window.addEventListener('resize', () => {
  closeLiveActionsMenu();
  refreshLiveTrackerUi();
});

window.addEventListener('orientationchange', () => {
  closeLiveActionsMenu();
  window.setTimeout(() => {
    refreshLiveTrackerUi();
  }, 100);
});

window.addEventListener('load', () => {
  refreshLiveTrackerUi();
});

window.visualViewport?.addEventListener('resize', () => {
  refreshLiveTrackerUi();
});

window.addEventListener('online', async () => {
  refreshSyncStatus();
  await checkCloudStateFreshness({ autoPull: !syncPendingChanges, force: true });
  if (!syncConflictInfo && hasSyncCredentials() && (syncPendingChanges || syncLastErrorMessage)) {
    queueCloudSync(250);
  }
});

window.addEventListener('offline', () => {
  clearSyncRetryTimer();
  refreshSyncStatus();
});

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible') {
    checkCloudStateFreshness({ autoPull: !syncPendingChanges && !syncConflictInfo });
  }
});

window.addEventListener('focus', () => {
  checkCloudStateFreshness({ autoPull: !syncPendingChanges && !syncConflictInfo });
});

function setupSyncUi() {
  if (!syncUserInput || !syncTokenInput) {
    return;
  }

  const credentials = getSyncCredentials();
  syncUserInput.value = credentials.user;
  syncTokenInput.value = credentials.token;
  syncConnectionState = credentials.user && credentials.token ? 'configured' : 'local';
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
    syncPendingChanges = false;
    syncLastSuccessAt = null;
    syncLastErrorMessage = '';
    updateSyncMetadata();
    clearSyncConflict();
    syncConnectionState = 'connecting';
    localStorage.setItem(SYNC_USER_STORAGE_KEY, user);
    localStorage.setItem(SYNC_TOKEN_STORAGE_KEY, token);
    updateSyncControls();
    refreshSyncStatus();

    try {
      await pullCloudState();
      refreshSyncStatus();
      setSyncUiCollapsed(true);
    } catch (error) {
      syncConnectionState = 'configured';
      syncLastErrorMessage = error.message;
      refreshSyncStatus();
    } finally {
      updateSyncControls();
    }
  };

  if (syncConnectButton) {
    syncConnectButton.addEventListener('click', handleSyncConnect);
  }

  [syncUserInput, syncTokenInput].forEach((input) => {
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
      syncPendingChanges = false;
      syncRetryCount = 0;
      syncLastSuccessAt = null;
      syncLastErrorMessage = '';
      updateSyncMetadata();
      clearSyncConflict();
      syncConnectionState = 'local';
      localStorage.removeItem(SYNC_USER_STORAGE_KEY);
      localStorage.removeItem(SYNC_TOKEN_STORAGE_KEY);
      syncUserInput.value = '';
      syncTokenInput.value = '';
      updateSyncControls();
      refreshSyncStatus();
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
      syncPendingChanges = true;
      syncLastErrorMessage = '';
      refreshSyncStatus();
      await pushCloudState();
    });
  }
}

async function initializeApp() {
  appState = loadLocalState();
  activeGameState = loadActiveGameState();
  activeGameUndoState = loadActiveGameUndoState();
  hideLiveSourcePrompt();
  setupSyncUi();

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

  if (liveGameForm) {
    liveGameDateInput.value = new Date().toISOString().slice(0, 10);
    if (!liveGamePlayerBody.children.length) {
      for (let index = 0; index < 4; index += 1) {
        addLiveSetupRow();
      }
    }
    renderLiveOrderPreview();
  }

  refresh();

  if (hasSyncCredentials()) {
    try {
      await pullCloudState();
      setSyncUiCollapsed(true);
      refreshSyncStatus();
    } catch (error) {
      syncConnectionState = 'configured';
      setSyncUiCollapsed(false);
      syncLastErrorMessage = `${error.message}. Using local cache.`;
      refreshSyncStatus();
    }
  }
}

initializeApp();
