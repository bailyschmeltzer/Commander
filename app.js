const STORAGE_KEY = 'commanderTrackerGames';
const form = document.getElementById('game-form');
const dateInput = document.getElementById('game-date');
const playerTableBody = document.getElementById('player-table-body');
const addPlayerRowButton = document.getElementById('add-player-row');
const notesInput = document.getElementById('game-notes');
const summaryEl = document.getElementById('summary');
const playerStatsTableBody = document.getElementById('player-stats-body');
const historyList = document.getElementById('history-list');
const historySortSelect = document.getElementById('history-sort');
const historySortOrderButton = document.getElementById('history-sort-order');
const historyFilterWinner = document.getElementById('history-filter-winner');
const historyFilterCommander = document.getElementById('history-filter-commander');
const commanderSearch = document.getElementById('commander-search');
const commanderStatsTableBody = document.getElementById('commander-stats-body');
const clearAllButton = document.getElementById('clear-all');
let historySortKey = 'date';
let historySortDescending = true;
let editingGameId = null;

function generateId() {
  return crypto.randomUUID?.() || `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

function getQueryParam(name) {
  return new URLSearchParams(window.location.search).get(name);
}

function loadGames() {
  try {
    const games = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
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
  } catch (error) {
    return [];
  }
}

function saveGames(games) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(games));
}

function normalizeList(value) {
  return value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function normalizePlayerName(name) {
  return String(name || '').trim().toLowerCase();
}

function formatPlayerCommanders(list) {
  return list.map(({ player, commander }) => `${player}: ${commander || '—'}`).join(', ');
}

function createPlayerRow(data = {}) {
  const row = document.createElement('tr');
  const killedValue = Array.isArray(data.killed) ? data.killed.join(', ') : data.killed || '';

  row.innerHTML = `
    <td><textarea name="player" placeholder="Players" required>${data.player || ''}</textarea></td>
    <td><textarea name="commander" placeholder="Commander">${data.commander || ''}</textarea></td>
    <td><input type="number" name="place" min="1" value="${data.place || ''}" placeholder="Place" /></td>
    <td><input type="number" name="kills" min="0" value="${data.kills || 0}" placeholder="Kills" /></td>
    <td><textarea name="killed" placeholder="Killed">${killedValue}</textarea></td>
  `;

  return row;
}

function addPlayerRow(data = {}) {
  playerTableBody.appendChild(createPlayerRow(data));
}

function resetPlayerTable() {
  playerTableBody.innerHTML = '';
  for (let i = 0; i < 2; i += 1) {
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

function ensurePlayerStats(stats, player, displayName) {
  if (!stats[player]) {
    stats[player] = {
      displayName: displayName || player,
      games: 0,
      wins: 0,
      kills: 0,
      commanders: {},
      commanderStats: {},
      killerCounts: {},
      victimCounts: {},
    };
  } else if (displayName && !stats[player].displayName) {
    stats[player].displayName = displayName;
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

function loadCommanderPowerLevels() {
  try {
    return JSON.parse(localStorage.getItem('commanderExpectedPowerLevels') || '{}');
  } catch (error) {
    return {};
  }
}

function saveCommanderPowerLevels(levels) {
  localStorage.setItem('commanderExpectedPowerLevels', JSON.stringify(levels));
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
      const canonicalPlayer = normalizePlayerName(player);
      if (!canonicalPlayer) {
        return;
      }

      const playerStat = ensurePlayerStats(stats, canonicalPlayer, player);
      playerStat.games += 1;
      if (normalizePlayerName(winner) === canonicalPlayer) {
        playerStat.wins += 1;
      }

      const commander = (row.commander || '').trim();
      if (commander) {
        playerStat.commanders[commander] = (playerStat.commanders[commander] || 0) + 1;
        if (!playerStat.commanderStats[commander]) {
          playerStat.commanderStats[commander] = { played: 0, wins: 0 };
        }
        playerStat.commanderStats[commander].played += 1;
        if (normalizePlayerName(winner) === canonicalPlayer) {
          playerStat.commanderStats[commander].wins += 1;
        }
      }

      const killedList = getCleanKilledList(row.killed);
      const killsCount = typeof row.kills === 'number' && !Number.isNaN(row.kills)
        ? row.kills
        : killedList.length;
      playerStat.kills += killsCount;

      killedList.forEach((targetRaw) => {
        const target = normalizePlayerName(targetRaw);
        if (!target) {
          return;
        }

        playerStat.victimCounts[target] = (playerStat.victimCounts[target] || 0) + 1;
        const targetStat = ensurePlayerStats(stats, target, targetRaw.trim());
        targetStat.killerCounts[canonicalPlayer] = (targetStat.killerCounts[canonicalPlayer] || 0) + 1;
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
      const nemesisKey = getMaxCountKey(stat.killerCounts);
      const victimKey = getMaxCountKey(stat.victimCounts);
      const nemesis = stats[nemesisKey]?.displayName || nemesisKey;
      const victim = stats[victimKey]?.displayName || victimKey;
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

      const displayName = stat.displayName || player;
      return `
        <tr>
          <td>${displayName}</td>
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
        stats[commander] = { games: 0, wins: 0, kills: 0 };
      }

      stats[commander].games += 1;
      if (Array.isArray(game.finishOrder) && game.finishOrder[0] === row.player) {
        stats[commander].wins += 1;
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
    const rawScore = 0.7 * winRate + 0.3 * killScore;
    actual[commander] = Math.round(rawScore * 100) / 10;
  });

  return actual;
}

function renderCommanderStats(games) {
  if (!commanderStatsTableBody) {
    return;
  }

  const searchTerm = commanderSearch?.value.trim().toLowerCase() || '';
  const commanderStats = getCommanderStatsData(games);
  const actualPowers = getCommanderActualPower(commanderStats);
  const rows = Object.entries(commanderStats)
    .filter(([commander]) => commander.toLowerCase().includes(searchTerm))
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([commander, stat]) => {
      const winRateValue = stat.games ? (stat.wins / stat.games) * 100 : 0;
      const killsPerGame = stat.games ? stat.kills / stat.games : 0;
      const expected = getCommanderExpectedPower(commander);
      const roundedExpected = typeof expected === 'number' ? expected.toFixed(1) : '';
      const actualPower = typeof actualPowers[commander] === 'number' ? actualPowers[commander].toFixed(1) : '0.0';

      return `
        <tr>
          <td>${escapeHtml(commander)}</td>
          <td>${stat.games}</td>
          <td>${stat.wins}</td>
          <td>${formatPercent(winRateValue)}</td>
          <td>${stat.kills}</td>
          <td>${killsPerGame.toFixed(1)}</td>
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
        </tr>`;
    })
    .join('');

  commanderStatsTableBody.innerHTML = rows || '<tr><td colspan="8">No commanders match your search.</td></tr>';
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
  renderSummary(games);
  renderPlayerStats(games);
  updateHistoryFilters(games);
  renderHistory(games);
  renderCommanderStats(games);
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

if (clearAllButton) {
  clearAllButton.addEventListener('click', () => {
    if (confirm('Remove all saved games? This cannot be undone.')) {
      localStorage.removeItem(STORAGE_KEY);
      refresh();
    }
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

updateHistorySortOrderLabel();

window.addEventListener('storage', (event) => {
  if (event.key === STORAGE_KEY) {
    refresh();
  }
});

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
