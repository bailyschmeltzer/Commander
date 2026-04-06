const STORAGE_KEY = 'commanderTrackerGames';
const EXPECTED_POWER_STORAGE_KEY = 'commanderExpectedPowerLevels';
const DECK_LIST_STORAGE_KEY = 'commanderDeckLists';
const RECORDS_STORAGE_KEY = 'commanderTrackerRecords';
const ACTIVE_GAME_STORAGE_KEY = 'commanderTrackerActiveGame';
const ACTIVE_GAME_UNDO_STORAGE_KEY = 'commanderTrackerActiveGameUndo';
const SYNC_USER_STORAGE_KEY = 'commanderTrackerSyncUser';
const SYNC_TOKEN_STORAGE_KEY = 'commanderTrackerSyncToken';
const CLOUD_SYNC_ENDPOINT = '/api/state';
const form = document.getElementById('game-form');
const dateInput = document.getElementById('game-date');
const playerTableBody = document.getElementById('player-table-body');
const addPlayerRowButton = document.getElementById('add-player-row');
const notesInput = document.getElementById('game-notes');
const summaryEl = document.getElementById('summary');
const playerDatalist = document.getElementById('player-list');
const commanderDatalist = document.getElementById('commander-list');
const playerStatsTableBody = document.getElementById('player-stats-body');
const historyList = document.getElementById('history-list');
const historySortSelect = document.getElementById('history-sort');
const historySortOrderButton = document.getElementById('history-sort-order');
const historyFilterWinner = document.getElementById('history-filter-winner');
const historyFilterCommander = document.getElementById('history-filter-commander');
const commanderSearch = document.getElementById('commander-search');
const commanderStatsTableBody = document.getElementById('commander-stats-body');
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
const liveFirstPlayer = document.getElementById('live-first-player');
const liveTurnNumberInput = document.getElementById('live-turn-number');
const liveFirstBlood = document.getElementById('live-first-blood');
const livePlayerGrid = document.getElementById('live-player-grid');
const liveEventLog = document.getElementById('live-event-log');
const liveUndoButton = document.getElementById('live-undo');
const liveFinishGameButton = document.getElementById('live-finish-game');
const liveAbandonGameButton = document.getElementById('live-abandon-game');
const liveSourcePrompt = document.getElementById('live-source-prompt');
const liveSourceTitle = document.getElementById('live-source-title');
const liveSourceCopy = document.getElementById('live-source-copy');
const liveSourceOptions = document.getElementById('live-source-options');
const liveSourceCancelButton = document.getElementById('live-source-cancel');

const syncUserInput = document.getElementById('sync-user');
const syncTokenInput = document.getElementById('sync-token');
const syncConnectButton = document.getElementById('sync-connect');
const syncDisconnectButton = document.getElementById('sync-disconnect');
const syncNowButton = document.getElementById('sync-now');
const syncStatus = document.getElementById('sync-status');
const syncPanel = document.querySelector('.sync-panel');

let historySortKey = 'date';
let historySortDescending = true;
let editingGameId = null;
let knownPlayers = [];
let knownCommanders = [];
let commanderSortColumn = 'games';
let commanderSortDescending = true;
let appState = { games: [], powerLevels: {}, deckLists: [], records: [] };
let editingDeckListId = null;
let syncQueueTimer = null;
let syncInFlight = false;
let deckSelectorSpinTimer = null;
let deckSelectorRotation = 0;
let activeGameState = null;
let activeGameUndoState = null;
let liveSetupFirstPlayerId = null;
let liveSourcePromptResolver = null;
let liveHoldTimerId = null;
let liveHoldIntervalId = null;
let liveHoldRepeated = false;

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

function parseJsonSafe(value, fallback) {
  try {
    return JSON.parse(value);
  } catch (error) {
    return fallback;
  }
}

function loadLocalState() {
  const games = parseJsonSafe(localStorage.getItem(STORAGE_KEY) || '[]', []);
  const powerLevels = parseJsonSafe(localStorage.getItem(EXPECTED_POWER_STORAGE_KEY) || '{}', {});
  const deckLists = parseJsonSafe(localStorage.getItem(DECK_LIST_STORAGE_KEY) || '[]', []);
  const records = parseJsonSafe(localStorage.getItem(RECORDS_STORAGE_KEY) || '[]', []);
  return {
    games: Array.isArray(games) ? games : [],
    powerLevels: powerLevels && typeof powerLevels === 'object' ? powerLevels : {},
    deckLists: Array.isArray(deckLists) ? deckLists : [],
    records: Array.isArray(records) ? records : [],
  };
}

function persistLocalState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.games || []));
  localStorage.setItem(EXPECTED_POWER_STORAGE_KEY, JSON.stringify(state.powerLevels || {}));
  localStorage.setItem(DECK_LIST_STORAGE_KEY, JSON.stringify(state.deckLists || []));
  localStorage.setItem(RECORDS_STORAGE_KEY, JSON.stringify(state.records || []));
}

function loadActiveGameState() {
  return parseJsonSafe(localStorage.getItem(ACTIVE_GAME_STORAGE_KEY) || 'null', null);
}

function loadActiveGameUndoState() {
  return parseJsonSafe(localStorage.getItem(ACTIVE_GAME_UNDO_STORAGE_KEY) || 'null', null);
}

function persistActiveGameUndoState(state) {
  activeGameUndoState = state || null;
  if (!state) {
    localStorage.removeItem(ACTIVE_GAME_UNDO_STORAGE_KEY);
    return;
  }

  localStorage.setItem(ACTIVE_GAME_UNDO_STORAGE_KEY, JSON.stringify(state));
}

function cloneActiveGameState(state) {
  return state ? parseJsonSafe(JSON.stringify(state), null) : null;
}

function saveUndoSnapshot() {
  if (!activeGameState) {
    persistActiveGameUndoState(null);
    return;
  }

  persistActiveGameUndoState(cloneActiveGameState(activeGameState));
}

function persistActiveGameState(state) {
  activeGameState = state || null;
  if (!state) {
    localStorage.removeItem(ACTIVE_GAME_STORAGE_KEY);
    return;
  }

  localStorage.setItem(ACTIVE_GAME_STORAGE_KEY, JSON.stringify(state));
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
  if (syncDisconnectButton) {
    syncDisconnectButton.disabled = !hasCredentials;
  }
  if (syncNowButton) {
    syncNowButton.disabled = !hasCredentials;
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
    try {
      const body = await response.json();
      if (body && body.error) {
        message = body.error;
      }
    } catch (error) {
      // Keep default message.
    }
    throw new Error(message);
  }

  return response.json();
}

async function pullCloudState() {
  const payload = await cloudRequest(CLOUD_SYNC_ENDPOINT, { method: 'GET' });
  const games = Array.isArray(payload.games) ? payload.games : [];
  const powerLevels = payload.powerLevels && typeof payload.powerLevels === 'object' ? payload.powerLevels : {};
  const deckLists = Array.isArray(payload.deckLists) ? payload.deckLists : [];
  const records = Array.isArray(payload.records) ? payload.records : [];
  appState = { games, powerLevels, deckLists, records };
  persistLocalState(appState);
  refresh();
}

async function pushCloudState() {
  if (!hasSyncCredentials() || syncInFlight) {
    return;
  }

  syncInFlight = true;
  try {
    await cloudRequest(CLOUD_SYNC_ENDPOINT, {
      method: 'PUT',
      body: JSON.stringify({
        games: appState.games,
        powerLevels: appState.powerLevels,
        deckLists: appState.deckLists,
        records: appState.records,
      }),
    });
    setSyncStatus('Synced to cloud.', 'success');
  } catch (error) {
    setSyncStatus(`Sync failed: ${error.message}`, 'error');
  } finally {
    syncInFlight = false;
  }
}

function queueCloudSync(delay = 500) {
  if (!hasSyncCredentials()) {
    return;
  }

  if (syncQueueTimer) {
    clearTimeout(syncQueueTimer);
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
  appState.games = Array.isArray(games) ? games : [];
  persistLocalState(appState);
  queueCloudSync();
}

function loadDeckLists() {
  return Array.isArray(appState.deckLists) ? appState.deckLists : [];
}

function saveDeckLists(deckLists) {
  appState.deckLists = Array.isArray(deckLists) ? deckLists : [];
  persistLocalState(appState);
  queueCloudSync();
}

function normalizeList(value) {
  return value
    .split(',')
    .map((item) => item.trim())
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
    .filter((player) => player.id !== targetPlayer.id)
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

function shouldPromptForSource(targetPlayer, projectedLife, eventType, sourcePlayerId = '') {
  if (!activeGameState || !targetPlayer) {
    return false;
  }

  if (eventType === 'elimination' || eventType === 'commander-damage') {
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

function eliminateLivePlayer(targetPlayer, sourcePlayerId, turnNumber, reason) {
  if (!targetPlayer || targetPlayer.eliminatedAt || targetPlayer.cannotLoseTheGame) {
    return false;
  }

  const aliveCount = getActiveAlivePlayers(activeGameState).length;
  targetPlayer.eliminatedAt = new Date().toISOString();
  targetPlayer.eliminatedTurnNumber = turnNumber;
  targetPlayer.eliminatedByPlayerId = sourcePlayerId || '';
  targetPlayer.eliminationReason = reason;
  targetPlayer.place = aliveCount;

  if (sourcePlayerId && sourcePlayerId !== targetPlayer.id) {
    const sourcePlayer = activeGameState.players.find((player) => player.id === sourcePlayerId);
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
      ? `Who eliminated ${targetPlayer.name}?`
      : `Who caused ${targetPlayer.name} to lose ${amount} life?`;
  const requireOpponent = eventType === 'commander-damage' || eventType === 'elimination';
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
    activeGameState.shouldPromptForSource = projectedLife <= 0 && !targetPlayer.cannotLoseTheGame;
  }

  return selectedSourceId || '';
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
  const rawValue = parseInt(liveTurnNumberInput?.value || `${activeGameState?.turnNumber || 1}`, 10);
  return Math.max(1, Number.isNaN(rawValue) ? 1 : rawValue);
}

function syncActiveGameTurnFromInput() {
  const turnNumber = getLiveTrackedTurnNumber();
  if (liveTurnNumberInput) {
    liveTurnNumberInput.value = String(turnNumber);
  }
  if (!activeGameState) {
    return turnNumber;
  }

  activeGameState.turnNumber = turnNumber;
  persistActiveGameState(activeGameState);
  return turnNumber;
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

  return Array.from(liveGamePlayerBody.querySelectorAll('tr'))
    .map((row, index) => ({
      id: row.dataset.playerId || generateId(),
      player: row.querySelector('[name="player"]')?.value.trim() || '',
      commander: row.querySelector('[name="commander"]')?.value.trim() || '',
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

  livePlayerGrid.innerHTML = activeGameState.players
    .slice()
    .sort((a, b) => a.seat - b.seat)
    .map((player) => {
      const damageEntries = Object.entries(player.commanderDamageTaken || {}).filter(([, amount]) => amount > 0);
      const damageMarkup = damageEntries.length
        ? `<ul class="live-card-damage-list">${damageEntries.map(([sourceId, amount]) => `<li>${escapeHtml(getPlayerNameById(sourceId, activeGameState))}: <strong>${amount}</strong></li>`).join('')}</ul>`
        : '<p>No commander damage tracked.</p>';

      return `
        <article class="live-player-card${player.eliminatedAt ? ' is-eliminated' : ''}">
          <div class="live-player-card-header">
            <div>
              <h3>${escapeHtml(player.name)}</h3>
              <p>${escapeHtml(player.commander || 'No commander')}</p>
            </div>
            <span class="live-seat-badge">Seat ${player.seat}</span>
          </div>
          <div class="live-player-life">${player.life}</div>
          <div class="live-quick-actions">
            <button type="button" class="live-quick-action is-negative" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="-1">-1</button>
            <button type="button" class="live-quick-action is-negative" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="-5">-5</button>
            <button type="button" class="live-quick-action is-positive" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="1">+1</button>
            <button type="button" class="live-quick-action is-positive" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="5">+5</button>
            <button type="button" class="live-quick-action" data-action="manual-commander-damage" data-player-id="${escapeHtml(player.id)}">Cmdr</button>
            <button type="button" class="live-quick-action" data-action="manual-eliminate" data-player-id="${escapeHtml(player.id)}">Out</button>
            <button type="button" class="live-quick-action" data-action="auto-win" data-player-id="${escapeHtml(player.id)}">Win</button>
          </div>
          <div class="live-player-meta">
            <div>Status: <strong>${escapeHtml(player.eliminatedAt ? `Out in place ${player.place || '—'}` : 'Still alive')}</strong></div>
            <div>Kills: <strong>${player.kills}</strong></div>
            <div>Killed: <strong>${escapeHtml((player.killedPlayers || []).join(', ') || 'None')}</strong></div>
            <label class="live-player-toggle">
              <input type="checkbox" data-action="toggle-cannot-lose" data-player-id="${escapeHtml(player.id)}"${player.cannotLoseTheGame ? ' checked' : ''} />
              <span>Cannot lose the game</span>
            </label>
            <div>Commander damage received:</div>
            ${damageMarkup}
          </div>
        </article>`;
    })
    .join('');
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

  if (!activeGameState) {
    if (liveGameStatus) {
      liveGameStatus.textContent = 'No game in progress.';
    }
    if (liveFirstPlayer) {
      liveFirstPlayer.textContent = '—';
    }
    if (liveTurnNumberInput) {
      liveTurnNumberInput.value = '1';
    }
    if (liveFirstBlood) {
      liveFirstBlood.textContent = 'None yet';
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
  if (liveFirstPlayer) {
    liveFirstPlayer.textContent = getPlayerNameById(activeGameState.startingPlayerId, activeGameState);
  }
  if (liveTurnNumberInput) {
    liveTurnNumberInput.value = String(activeGameState.turnNumber || 1);
  }
  if (liveFirstBlood) {
    if (activeGameState.firstBlood) {
      const actor = activeGameState.firstBlood.actorPlayerId
        ? getPlayerNameById(activeGameState.firstBlood.actorPlayerId, activeGameState)
        : 'Environment';
      liveFirstBlood.textContent = `${actor} drew first blood on ${getPlayerNameById(activeGameState.firstBlood.targetPlayerId, activeGameState)} during turn ${activeGameState.firstBlood.turnNumber}.`;
    } else {
      liveFirstBlood.textContent = 'None yet';
    }
  }
  if (liveUndoButton) {
    liveUndoButton.disabled = !activeGameUndoState;
  }
  renderLivePlayerGrid();
  renderLiveEventLog();
}

function refreshLiveTrackerUi() {
  if (!liveSourcePromptResolver) {
    hideLiveSourcePrompt();
  }
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

  return [
    `First player: ${getPlayerNameById(gameState.startingPlayerId, gameState)}`,
    firstBloodSummary,
    `Kill leader: ${killLeader ? `${killLeader.name} (${killLeader.kills})` : 'None'}`,
    `Finish order: ${finishOrderSummary}`,
  ].join(' · ');
}

function startLiveGame() {
  const players = getLiveSetupRows();
  if (players.length < 2) {
    alert('Add at least two players to start a live game.');
    return;
  }

  const startingLife = parseInt(liveStartingLifeInput?.value || '40', 10);
  if (!startingLife || startingLife < 1) {
    alert('Enter a valid starting life total.');
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
      place: null,
    })),
  });
  refreshLiveTrackerUi();
}

function promptForPositiveNumber(message, defaultValue = 1) {
  const rawValue = window.prompt(message, String(defaultValue));
  if (rawValue === null) {
    return null;
  }

  const parsed = parseInt(rawValue, 10);
  if (Number.isNaN(parsed) || parsed < 1) {
    alert('Enter a whole number greater than 0.');
    return null;
  }

  return parsed;
}

async function applyCommanderDamageToPlayer(targetPlayerId) {
  if (!activeGameState) {
    return;
  }

  const targetPlayer = activeGameState.players.find((player) => player.id === targetPlayerId);
  if (!targetPlayer || targetPlayer.eliminatedAt) {
    return;
  }

  const amount = promptForPositiveNumber(`How much commander damage did ${targetPlayer.name} take?`, 1);
  if (amount === null) {
    return;
  }

  const eventTurnNumber = syncActiveGameTurnFromInput();
  const sourcePlayerId = await resolveLiveSourceSelection({
    targetPlayerId,
    eventType: 'commander-damage',
    amount,
    projectedLife: targetPlayer.life - amount,
  });
  if (sourcePlayerId === null) {
    return;
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

  if (alivePlayers.length === 1 && confirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`)) {
    completeActiveGame();
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

  const eventTurnNumber = syncActiveGameTurnFromInput();
  const sourcePlayerId = await resolveLiveSourceSelection({
    targetPlayerId,
    eventType: 'elimination',
    amount: 0,
    projectedLife: targetPlayer.life,
  });
  if (sourcePlayerId === null) {
    return;
  }

  saveUndoSnapshot();
  eliminateLivePlayer(targetPlayer, sourcePlayerId, eventTurnNumber, 'manual');
  recordLiveEvent({
    type: 'elimination',
    actorPlayerId: sourcePlayerId,
    targetPlayerId,
    amount: 0,
    turnNumber: eventTurnNumber,
  });

  const alivePlayers = getActiveAlivePlayers(activeGameState);
  if (alivePlayers.length === 1) {
    alivePlayers[0].place = 1;
  }

  persistActiveGameState(activeGameState);
  refreshLiveTrackerUi();

  if (alivePlayers.length === 1 && confirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`)) {
    completeActiveGame();
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

  const turnNumber = syncActiveGameTurnFromInput();
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

  if (alivePlayers.length === 1 && confirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`)) {
    completeActiveGame();
  }
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

  const turnNumber = syncActiveGameTurnFromInput();
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

  if (alivePlayers.length === 1 && confirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`)) {
    completeActiveGame();
  }
}

function markPlayerAutomaticWinner(playerId) {
  if (!activeGameState) {
    return;
  }

  const winner = activeGameState.players.find((entry) => entry.id === playerId);
  if (!winner || winner.eliminatedAt) {
    return;
  }

  if (!confirm(`Mark ${winner.name} as the winner and finish the game?`)) {
    return;
  }

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
    turnNumber: activeGameState.turnNumber,
  });
  completeActiveGame();
}

function undoLastLiveAction() {
  if (!activeGameUndoState) {
    return;
  }

  persistActiveGameState(cloneActiveGameState(activeGameUndoState));
  persistActiveGameUndoState(null);
  refresh();
}

function completeActiveGame() {
  if (!activeGameState) {
    return;
  }

  const alivePlayers = getActiveAlivePlayers(activeGameState);
  if (alivePlayers.length > 1 && !confirm('More than one player is still alive. Finish and score remaining players by life total?')) {
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
      firstBlood: activeGameState.firstBlood,
      turnNumber: activeGameState.turnNumber,
      eventCount: activeGameState.events.length,
    },
  });
  saveGames(games);
  persistActiveGameState(null);
  persistActiveGameUndoState(null);
  refresh();
  refreshLiveTrackerUi();
  alert('Live game saved to history.');
}

function abandonActiveGame() {
  if (!activeGameState) {
    return;
  }

  if (!confirm('Abandon the current live game? This removes the in-progress tracker without saving to history.')) {
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
    <td><input type="number" name="place" min="1" value="${escapeHtml(data.place || '')}" placeholder="Place" /></td>
    <td><input type="number" name="kills" min="0" value="${escapeHtml(data.kills || 0)}" placeholder="Kills" /></td>
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
  return Array.from(playerTableBody.querySelectorAll('tr')).map((row) => {
    const player = row.querySelector('[name="player"]').value.trim();
    const commander = row.querySelector('[name="commander"]').value.trim();
    const placeRaw = row.querySelector('input[name="place"]').value;
    const killsRaw = row.querySelector('input[name="kills"]').value;
    const killed = normalizeList(row.querySelector('[name="killed"]').value);

    return {
      player,
      commander,
      place: placeRaw ? parseInt(placeRaw, 10) : null,
      kills: killsRaw ? parseInt(killsRaw, 10) : 0,
      killed,
    };
  }).filter((row) => row.player);
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
  appState.powerLevels = levels && typeof levels === 'object' ? levels : {};
  persistLocalState(appState);
  queueCloudSync();
}

function setCommanderExpectedPower(commander, value) {
  const levels = loadCommanderPowerLevels();
  if (typeof value !== 'number' || Number.isNaN(value)) {
    delete levels[commander];
  } else {
    levels[commander] = Math.min(10, Math.max(0, Math.round(value * 10) / 10));
  }
  saveCommanderPowerLevels(levels);
}

function getCommanderExpectedPower(commander) {
  const levels = loadCommanderPowerLevels();
  return typeof levels[commander] === 'number' ? levels[commander] : '';
}

function normalizeRecordEntry(entry, index = 0) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const isCustom = Boolean(entry.isCustom);
  const key = isCustom ? '' : String(entry.key || '').trim();
  const title = String(entry.title || '').trim();
  if (!title) {
    return null;
  }

  const fallbackId = key || `${title.toLowerCase().replace(/[^a-z0-9]+/g, '-') || 'record'}-${index}`;
  return {
    id: String(entry.id || fallbackId),
    key,
    title,
    unit: String(entry.unit || '').trim(),
    value: String(entry.value || '').trim(),
    holder: String(entry.holder || '').trim(),
    commander: String(entry.commander || '').trim(),
    date: String(entry.date || '').trim(),
    notes: String(entry.notes || '').trim(),
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
    isCustom: false,
  }, index));
}

function mergeRecordsWithDefaults(records) {
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
  }));

  const customRecords = normalized
    .filter((entry) => entry.isCustom)
    .sort((a, b) => a.title.localeCompare(b.title));

  return [...mergedDefaults, ...customRecords];
}

function loadRecords() {
  return mergeRecordsWithDefaults(appState.records);
}

function saveRecords(records) {
  appState.records = mergeRecordsWithDefaults(records);
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

  resetPlayerTable();
  const rows = getPlayerRows();
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

  games.forEach((game) => {
    const winner = getGameWinner(game);
    if (winner) {
      winners.add(winner);
    }

    getGameRows(game).forEach((row) => {
      if (row.commander) {
        commanders.add(row.commander);
      }
    });
  });

  return {
    winners: [...winners].sort((a, b) => a.localeCompare(b)),
    commanders: [...commanders].sort((a, b) => a.localeCompare(b)),
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
  if (!historyFilterWinner && !historyFilterCommander) {
    return;
  }

  const filters = getHistoryFilterOptions(games);
  buildFilterOptions(historyFilterWinner, filters.winners, 'winners');
  buildFilterOptions(historyFilterCommander, filters.commanders, 'commanders');
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

  refreshRowSelectors();
  populateDeckCommanderSelector();
  populateRecordLookupMenus();
}

function normalizeDeckListEntry(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const commander = String(entry.commander || '').trim();
  const owner = String(entry.owner || '').trim();
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

  const deckLists = getSortedDeckLists();

  if (!deckLists.length) {
    deckListTableBody.innerHTML = '<tr><td colspan="4">No deck lists saved yet.</td></tr>';
    return;
  }

  deckListTableBody.innerHTML = deckLists
    .map((entry) => {
      const safeUrl = escapeHtml(entry.url);
      return `
        <tr>
          <td>${escapeHtml(entry.commander)}</td>
          <td>${escapeHtml(entry.owner || '—')}</td>
          <td><a href="${safeUrl}" target="_blank" rel="noopener noreferrer">${safeUrl}</a></td>
          <td>
            <button type="button" class="secondary-button deck-list-edit" data-id="${escapeHtml(entry.id)}">Edit</button>
            <button type="button" class="history-delete-button deck-list-delete" data-id="${escapeHtml(entry.id)}">Delete</button>
          </td>
        </tr>`;
    })
    .join('');
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

function handleDeckListTableAction(event) {
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
    if (confirm('Delete this deck list?')) {
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
  const notes = game.notes ? escapeHtml(game.notes) : 'No notes recorded.';

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
        <h3>${escapeHtml(game.date || 'Unknown date')}</h3>
        <small>Winner: ${escapeHtml(winner)} · ${playerCount} players · ${totalKills} kills</small>
      </div>
      <div class="history-item-actions">
        <button type="button" class="secondary-button history-edit-button" data-id="${escapeHtml(game.id)}">Edit</button>
        <button type="button" class="history-delete-button" data-id="${escapeHtml(game.id)}">Delete</button>
      </div>
      <p>${notes}</p>
      <table class="player-table history-game-table">
        <thead>
          <tr>
            <th>Player</th>
            <th>Commander</th>
            <th>Place</th>
            <th>Kills</th>
            <th>Killed</th>
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

  const filteredGames = sortedGames.filter((game) => {
    if (winnerFilter !== 'all' && getGameWinner(game) !== winnerFilter) {
      return false;
    }

    if (commanderFilter !== 'all') {
      const foundCommander = getGameRows(game).some((row) => row.commander === commanderFilter);
      if (!foundCommander) {
        return false;
      }
    }

    return true;
  });

  if (!filteredGames.length) {
    historyList.innerHTML = '<p>No games match the current filters.</p>';
    return;
  }

  historyList.innerHTML = filteredGames.map(renderHistoryGame).join('');
}

function handleHistoryAction(event) {
  const button = event.target.closest('button');
  if (!button || !historyList.contains(button)) {
    return;
  }

  const gameId = button.dataset.id;
  if (!gameId) {
    return;
  }

  if (button.classList.contains('history-delete-button')) {
    if (confirm('Delete this game? This cannot be undone.')) {
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

function getPlayerStatsData(games) {
  const stats = {};

  games.forEach((game) => {
    const rows = getGameRows(game);
    const winner = Array.isArray(game.finishOrder) && game.finishOrder.length ? game.finishOrder[0] : null;

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

  const stats = getPlayerStatsData(games);
  const players = Object.keys(stats);

  if (!players.length) {
    playerStatsTableBody.innerHTML = '<tr><td colspan="10">No player stats available.</td></tr>';
    return;
  }

  const html = players
    .sort((a, b) => {
      const diff = stats[b].games - stats[a].games;
      if (diff) return diff;
      return stats[b].wins - stats[a].wins;
    })
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

      return `
        <tr>
          <td>${player}</td>
          <td>${stat.games}</td>
          <td>${stat.wins}</td>
          <td>${formatPercent(winRateValue)}</td>
          <td>${favoriteCommander}</td>
          <td>${nemesis}</td>
          <td>${victim}</td>
          <td>${stat.kills}</td>
          <td>${killAverage}</td>
          <td>${bestDeck}</td>
        </tr>`;
    })
    .join('');

  playerStatsTableBody.innerHTML = html;
}

function getCommanderStatsData(games) {
  const stats = {};

  games.forEach((game) => {
    getGameRows(game).forEach((row) => {
      const commander = (row.commander || '').trim();
      if (!commander) {
        return;
      }

      if (!stats[commander]) {
        stats[commander] = { games: 0, wins: 0, kills: 0, placementTotal: 0, placementScoreTotal: 0, placementGames: 0 };
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

  // Sort by selected column
  entries.sort((a, b) => {
    let aVal, bVal;
    switch (commanderSortColumn) {
      case 'commander':
        aVal = a.commander.toLowerCase();
        bVal = b.commander.toLowerCase();
        return commanderSortDescending ? bVal.localeCompare(aVal) : aVal.localeCompare(bVal);
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
    return commanderSortDescending ? -diff : diff;
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
          <td>${escapeHtml(commander)}</td>
          <td>${stat.games}</td>
          <td>${stat.wins}</td>
          <td>${formatPercent(winRateValue)}</td>
          <td>${stat.kills}</td>
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

  commanderStatsTableBody.innerHTML = rows || '<tr><td colspan="12">No commanders match your search.</td></tr>';
  updateCommanderSortIndicators();
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

function updateCommanderSortIndicators() {
  const headers = document.querySelectorAll('.commander-stats-table .sortable-header');
  headers.forEach((header) => {
    const column = header.dataset.column;
    const indicator = header.querySelector('.sort-indicator');
    if (!indicator) return;

    if (column === commanderSortColumn) {
      indicator.textContent = commanderSortDescending ? ' ▼' : ' ▲';
      indicator.style.opacity = '1';
    } else {
      indicator.textContent = '';
      indicator.style.opacity = '0.3';
    }
  });
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

function handleCommanderHeaderClick(event) {
  const header = event.target.closest('.sortable-header');
  if (!header) {
    return;
  }

  const column = header.dataset.column;
  if (column === commanderSortColumn) {
    commanderSortDescending = !commanderSortDescending;
  } else {
    commanderSortColumn = column;
    commanderSortDescending = true;
  }

  renderCommanderStats(loadGames());
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
  renderSummary(games);
  renderPlayerStats(games);
  updateHistoryFilters(games);
  renderHistory(games);
  renderCommanderStats(games);
  renderDeckLookup();
  renderDeckLists();
  renderDeckSelector();
  renderRecords();
  refreshLiveTrackerUi();
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
  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const games = loadGames();

    const rows = getPlayerRows();
    if (!rows.length) {
      alert('Please add at least one player row with a player name.');
      return;
    }

    const players = rows.map(({ player }) => player);
    const playerCommanders = rows.map(({ player, commander }) => ({ player, commander }));
    const finishOrder = rows
      .slice()
      .filter((row) => row.place !== null)
      .sort((a, b) => (a.place || 999) - (b.place || 999))
      .map((row) => row.player);

    const newGame = {
      id: editingGameId || generateId(),
      date: dateInput.value || new Date().toISOString().slice(0, 10),
      playerRows: rows,
      players: rows.map(({ player }) => player),
      playerCommanders,
      finishOrder,
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

if (liveTurnNumberInput) {
  liveTurnNumberInput.addEventListener('change', () => {
    syncActiveGameTurnFromInput();
    refreshLiveTrackerUi();
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

if (liveUndoButton) {
  liveUndoButton.addEventListener('click', () => {
    undoLastLiveAction();
  });
}

if (liveFinishGameButton) {
  liveFinishGameButton.addEventListener('click', () => {
    completeActiveGame();
  });
}

if (liveAbandonGameButton) {
  liveAbandonGameButton.addEventListener('click', () => {
    abandonActiveGame();
  });
}

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

if (commanderSearch) {
  commanderSearch.addEventListener('input', () => {
    renderCommanderStats(loadGames());
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
  deckListForm.addEventListener('submit', (event) => {
    event.preventDefault();

    const commander = deckCommanderInput.value.trim();
    const owner = deckOwnerInput?.value.trim() || '';
    const url = deckUrlInput.value.trim();

    if (!commander) {
      alert('Please choose or enter a commander.');
      return;
    }

    if (!url) {
      alert('Please enter a deck URL.');
      return;
    }

    let parsedUrl;
    try {
      parsedUrl = new URL(url);
    } catch (error) {
      alert('Please enter a valid URL (include https://).');
      return;
    }

    if (!['http:', 'https:'].includes(parsedUrl.protocol)) {
      alert('Deck URL must use http or https.');
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
  customRecordForm.addEventListener('submit', (event) => {
    event.preventDefault();

    const title = customRecordTitleInput?.value.trim() || '';
    if (!title) {
      alert('Please enter a title for the custom record.');
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

const commanderStatsTable = document.querySelector('.commander-stats-table');
if (commanderStatsTable) {
  commanderStatsTable.addEventListener('click', handleCommanderHeaderClick);
}

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

function setupSyncUi() {
  if (!syncUserInput || !syncTokenInput) {
    return;
  }

  const credentials = getSyncCredentials();
  syncUserInput.value = credentials.user;
  syncTokenInput.value = credentials.token;
  updateSyncControls();

  if (!credentials.user || !credentials.token) {
    setSyncStatus('Cloud sync not connected. Data remains local until you connect.', 'muted');
  } else {
    setSyncStatus('Cloud sync configured. Pulling latest state...', 'neutral');
  }

  if (syncConnectButton) {
    syncConnectButton.addEventListener('click', async () => {
      const user = syncUserInput.value.trim();
      const token = syncTokenInput.value.trim();

      if (!user || !token) {
        setSyncStatus('Enter both display name and pod access code.', 'error');
        return;
      }

      localStorage.setItem(SYNC_USER_STORAGE_KEY, user);
      localStorage.setItem(SYNC_TOKEN_STORAGE_KEY, token);
      updateSyncControls();
      setSyncStatus('Connecting to cloud...', 'neutral');

      try {
        await pullCloudState();
        setSyncStatus(`Connected as ${user}.`, 'success');
      } catch (error) {
        setSyncStatus(`Connection failed: ${error.message}`, 'error');
      }
    });
  }

  if (syncDisconnectButton) {
    syncDisconnectButton.addEventListener('click', () => {
      localStorage.removeItem(SYNC_USER_STORAGE_KEY);
      localStorage.removeItem(SYNC_TOKEN_STORAGE_KEY);
      syncUserInput.value = '';
      syncTokenInput.value = '';
      updateSyncControls();
      setSyncStatus('Cloud sync disconnected. Local mode active.', 'muted');
    });
  }

  if (syncNowButton) {
    syncNowButton.addEventListener('click', async () => {
      setSyncStatus('Syncing now...', 'neutral');
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

  if (hasSyncCredentials()) {
    try {
      await pullCloudState();
      setSyncStatus('Cloud data loaded.', 'success');
      setSyncUiCollapsed(true, getSyncCredentials().user);
    } catch (error) {
      setSyncUiCollapsed(false);
      setSyncStatus(`Using local cache: ${error.message}`, 'error');
    }
  }

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
    if (liveTurnNumberInput) {
      liveTurnNumberInput.value = '1';
    }
    if (!liveGamePlayerBody.children.length) {
      for (let index = 0; index < 4; index += 1) {
        addLiveSetupRow();
      }
    }
    renderLiveOrderPreview();
  }

  refresh();
}

initializeApp();
