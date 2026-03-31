const STORAGE_KEY = 'commanderTrackerGames';
const EXPECTED_POWER_STORAGE_KEY = 'commanderExpectedPowerLevels';
const DECK_LIST_STORAGE_KEY = 'commanderDeckLists';
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
const deckUrlInput = document.getElementById('deck-url');
const deckListTableBody = document.getElementById('deck-list-body');
const deckListCancelButton = document.getElementById('deck-list-cancel');
const deckListSubmitButton = document.querySelector('#deck-list-form button[type="submit"]');
const deckLookupSelect = document.getElementById('deck-lookup-commander');
const deckLookupResult = document.getElementById('deck-lookup-result');

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
let appState = { games: [], powerLevels: {}, deckLists: [] };
let editingDeckListId = null;
let syncQueueTimer = null;
let syncInFlight = false;

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
  return {
    games: Array.isArray(games) ? games : [],
    powerLevels: powerLevels && typeof powerLevels === 'object' ? powerLevels : {},
    deckLists: Array.isArray(deckLists) ? deckLists : [],
  };
}

function persistLocalState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.games || []));
  localStorage.setItem(EXPECTED_POWER_STORAGE_KEY, JSON.stringify(state.powerLevels || {}));
  localStorage.setItem(DECK_LIST_STORAGE_KEY, JSON.stringify(state.deckLists || []));
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
  appState = { games, powerLevels, deckLists };
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

function createPlayerRow(data = {}) {
  const row = document.createElement('tr');
  const killedValue = Array.isArray(data.killed) ? data.killed.join(', ') : data.killed || '';

  row.innerHTML = `
    <td class="lookup-field-cell">
      <div class="combined-input-wrapper">
        <input class="lookup-input" type="text" name="player" list="player-list" placeholder="Player" required value="${escapeHtml(data.player || '')}" />
        <button type="button" class="dropdown-button player-dropdown-button" title="Show options">▼</button>
        <div class="dropdown-menu player-dropdown-menu"></div>
      </div>
    </td>
    <td class="lookup-field-cell">
      <div class="combined-input-wrapper">
        <input class="lookup-input" type="text" name="commander" list="commander-list" placeholder="Commander" value="${escapeHtml(data.commander || '')}" />
        <button type="button" class="dropdown-button commander-dropdown-button" title="Show options">▼</button>
        <div class="dropdown-menu commander-dropdown-menu"></div>
      </div>
    </td>
    <td><input type="number" name="place" min="1" value="${escapeHtml(data.place || '')}" placeholder="Place" /></td>
    <td><input type="number" name="kills" min="0" value="${escapeHtml(data.kills || 0)}" placeholder="Kills" /></td>
    <td><textarea name="killed" placeholder="Killed">${escapeHtml(killedValue)}</textarea></td>
  `;

  populateRowSelectors(row);
  attachDropdownHandlers(row);
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
  menuElement.innerHTML = normalized
    .map((value) => `<div class="dropdown-item" data-value="${escapeHtml(value)}">${escapeHtml(value)}</div>`)
    .join('');
}

function updateDropdownLayeringState() {
  const hasActiveDropdown = document.querySelector('.dropdown-menu.active');
  document.querySelectorAll('.player-table-wrapper').forEach((wrapper) => {
    wrapper.classList.toggle('dropdown-open', Boolean(hasActiveDropdown));
  });
}

function attachDropdownHandlers(row) {
  if (!row || row.dataset.dropdownHandlersAttached) {
    return;
  }
  row.dataset.dropdownHandlersAttached = 'true';

  row.querySelectorAll('.dropdown-button').forEach((button) => {
    button.addEventListener('click', (e) => {
      e.preventDefault();
      const wrapper = button.closest('.combined-input-wrapper');
      const menu = wrapper.querySelector('.dropdown-menu');
      
      // Close all other menus
      document.querySelectorAll('.dropdown-menu.active').forEach((m) => {
        if (m !== menu) m.classList.remove('active');
      });
      
      // Toggle this menu
      if (menu.classList.contains('active')) {
        menu.classList.remove('active');
      } else {
        menu.classList.add('active');
      }

      updateDropdownLayeringState();
    });
  });

  // Use event delegation on dropdown menus
  row.querySelectorAll('.dropdown-menu').forEach((menu) => {
    menu.addEventListener('click', (e) => {
      const item = e.target.closest('.dropdown-item');
      if (!item) return;
      
      const value = item.dataset.value;
      const wrapper = menu.closest('.combined-input-wrapper');
      const input = wrapper.querySelector('input[name]');
      
      input.value = value;
      menu.classList.remove('active');
      updateDropdownLayeringState();
    });
  });

}

// Global outside-click handler — registered once
document.addEventListener('click', (e) => {
  if (!e.target.closest('.combined-input-wrapper')) {
    document.querySelectorAll('.dropdown-menu').forEach((m) => m.classList.remove('active'));
    updateDropdownLayeringState();
  }
});

function refreshRowSelectors() {
  Array.from(document.querySelectorAll('tr')).forEach((row) => {
    populateRowSelectors(row);
    attachDropdownHandlers(row);
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
}

function normalizeDeckListEntry(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const commander = String(entry.commander || '').trim();
  const url = String(entry.url || '').trim();
  if (!commander || !url) {
    return null;
  }

  return {
    id: entry.id || generateId(),
    commander,
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
    return;
  }
  buildDropdownMenu(deckCommanderMenu, knownCommanders);
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
  deckLookupResult.innerHTML = `
    <p class="deck-lookup-label">${safeCommander}</p>
    <a href="${safeUrl}" target="_blank" rel="noopener noreferrer" title="${safeUrl}">${safeUrl}</a>
  `;
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
    deckListTableBody.innerHTML = '<tr><td colspan="3">No deck lists saved yet.</td></tr>';
    return;
  }

  deckListTableBody.innerHTML = deckLists
    .map((entry) => {
      const safeUrl = escapeHtml(entry.url);
      return `
        <tr>
          <td>${escapeHtml(entry.commander)}</td>
          <td><a href="${safeUrl}" target="_blank" rel="noopener noreferrer">${safeUrl}</a></td>
          <td>
            <button type="button" class="secondary-button deck-list-edit" data-id="${escapeHtml(entry.id)}">Edit</button>
            <button type="button" class="history-delete-button deck-list-delete" data-id="${escapeHtml(entry.id)}">Delete</button>
          </td>
        </tr>`;
    })
    .join('');
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

if (deckCommanderDropdownButton && deckCommanderMenu && deckCommanderInput) {
  deckCommanderDropdownButton.addEventListener('click', (event) => {
    event.preventDefault();
    document.querySelectorAll('.dropdown-menu.active').forEach((menu) => {
      if (menu !== deckCommanderMenu) {
        menu.classList.remove('active');
      }
    });
    deckCommanderMenu.classList.toggle('active');
    updateDropdownLayeringState();
  });

  deckCommanderMenu.addEventListener('click', (event) => {
    const item = event.target.closest('.dropdown-item');
    if (!item) {
      return;
    }
    deckCommanderInput.value = item.dataset.value || '';
    deckCommanderMenu.classList.remove('active');
    updateDropdownLayeringState();
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
        deckLists[index] = { id: editingDeckListId, commander, url: normalizedUrl };
      } else {
        deckLists.push({ id: generateId(), commander, url: normalizedUrl });
      }

      if (duplicateCommanderIndex >= 0) {
        deckLists.splice(duplicateCommanderIndex, 1);
      }
    } else {
      if (duplicateCommanderIndex >= 0) {
        deckLists[duplicateCommanderIndex] = {
          ...deckLists[duplicateCommanderIndex],
          commander,
          url: normalizedUrl,
        };
      } else {
        deckLists.push({ id: generateId(), commander, url: normalizedUrl });
      }
    }

    saveDeckLists(deckLists);
    resetDeckListForm();
    refresh();
  });
}

const commanderStatsTable = document.querySelector('.commander-stats-table');
if (commanderStatsTable) {
  commanderStatsTable.addEventListener('click', handleCommanderHeaderClick);
}

updateHistorySortOrderLabel();

window.addEventListener('storage', (event) => {
  if (event.key === STORAGE_KEY || event.key === EXPECTED_POWER_STORAGE_KEY || event.key === DECK_LIST_STORAGE_KEY) {
    appState = loadLocalState();
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

  refresh();
}

initializeApp();
