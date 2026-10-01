(function () {
  const list = document.getElementById('history-list');
  const sortSelect = document.getElementById('history-sort');
  const sortOrderButton = document.getElementById('history-sort-order');
  const winnerFilter = document.getElementById('history-filter-winner');
  const commanderFilter = document.getElementById('history-filter-commander');
  const playerFilter = document.getElementById('history-filter-player');
  const dateFromFilter = document.getElementById('history-filter-date-from');
  const dateToFilter = document.getElementById('history-filter-date-to');
  const resetFiltersButton = document.getElementById('history-reset-filters');
  const activeFilters = document.getElementById('history-active-filters');
  let queryFiltersApplied = false;
  let playerFilterDefaulted = false;

  function getFilterOptions(games) {
    const winners = [];
    const commanders = [];
    const players = [];
    const playerMap = buildCanonicalIdentityMapFromValues(knownPlayers);
    const commanderMap = buildCanonicalIdentityMapFromValues(knownCommanders);

    games.forEach((game) => {
      const winner = getGameWinner(game);
      if (winner) winners.push(canonicalizeIdentityValue(winner, playerMap));

      getGameRows(game).forEach((row) => {
        if (row.player) players.push(canonicalizeIdentityValue(row.player, playerMap));
        if (row.commander) commanders.push(canonicalizeIdentityValue(row.commander, commanderMap));
      });
    });

    return {
      winners: getUniqueValuesBySimilarity(winners.filter(Boolean)),
      commanders: getUniqueValuesBySimilarity(commanders.filter(Boolean)),
      players: getUniqueValuesBySimilarity(players.filter(Boolean)),
    };
  }

  function buildFilterOptions(selectElement, values, label) {
    if (!selectElement) return;
    const currentValue = selectElement.value || 'all';
    const options = ['all', ...values];
    selectElement.innerHTML = options.map((value) => {
      const display = value === 'all' ? `All ${label}` : escapeHtml(value);
      const selected = currentValue === value ? ' selected' : '';
      return `<option value="${escapeHtml(value)}"${selected}>${display}</option>`;
    }).join('');
  }

  function updateFilters(games) {
    if (!winnerFilter && !commanderFilter && !playerFilter) return;
    const filters = getFilterOptions(games);
    buildFilterOptions(winnerFilter, filters.winners, 'winners');
    buildFilterOptions(commanderFilter, filters.commanders, 'commanders');
    buildFilterOptions(playerFilter, filters.players, 'players');
  }

  function applyQueryFilters(force = false) {
    if (queryFiltersApplied && !force) return;
    const winner = getQueryParam('winner');
    const commander = getQueryParam('commander');
    const player = getQueryParam('player');
    const fromDate = getQueryParam('from');
    const toDate = getQueryParam('to');
    const sortKey = getQueryParam('sort');
    const order = getQueryParam('order');

    if (winnerFilter) winnerFilter.value = winner || 'all';
    if (commanderFilter) commanderFilter.value = commander || 'all';
    if (playerFilter) {
      playerFilter.value = player || 'all';
      if (player) playerFilterDefaulted = true;
    }
    if (dateFromFilter) dateFromFilter.value = fromDate || '';
    if (dateToFilter) dateToFilter.value = toDate || '';
    if (sortSelect && sortKey) sortSelect.value = sortKey;

    historySortKey = sortKey || historySortKey || 'date';
    if (order === 'asc') historySortDescending = false;
    else if (order === 'desc') historySortDescending = true;
    updateSortOrderLabel();
    queryFiltersApplied = true;
  }

  function getCurrentFilters() {
    return {
      winner: winnerFilter?.value && winnerFilter.value !== 'all' ? winnerFilter.value : '',
      commander: commanderFilter?.value && commanderFilter.value !== 'all' ? commanderFilter.value : '',
      player: playerFilter?.value && playerFilter.value !== 'all' ? playerFilter.value : '',
      from: dateFromFilter?.value || '',
      to: dateToFilter?.value || '',
    };
  }

  function syncFilterQuery() {
    if (!list) return;
    const href = buildHistoryFilterHref(getCurrentFilters());
    const currentPath = `${window.location.pathname.split('/').pop() || 'history.html'}${window.location.search}`;
    if (currentPath !== href) window.history.replaceState({}, '', href);
  }

  function renderActiveFilters(totalCount, filteredCount) {
    if (!activeFilters) return;
    const filters = getCurrentFilters();
    const active = [
      filters.winner ? `Winner: ${filters.winner}` : '',
      filters.commander ? `Commander: ${filters.commander}` : '',
      filters.player ? `Player: ${filters.player}` : '',
      filters.from ? `From: ${filters.from}` : '',
      filters.to ? `To: ${filters.to}` : '',
    ].filter(Boolean);
    const summary = filteredCount === totalCount
      ? `Showing all ${totalCount} game${totalCount === 1 ? '' : 's'}.`
      : `Showing ${filteredCount} of ${totalCount} game${totalCount === 1 ? '' : 's'}.`;

    if (!active.length) {
      activeFilters.innerHTML = `<p class="history-active-filter-summary">${summary}</p>`;
      return;
    }

    activeFilters.innerHTML = `
    <p class="history-active-filter-summary">${summary}</p>
    <div class="history-active-filter-list" aria-label="Active history filters">
      ${active.map((label) => `<span class="history-active-filter-chip">${escapeHtml(label)}</span>`).join('')}
    </div>`;
  }

  function resetFilters() {
    if (winnerFilter) winnerFilter.value = 'all';
    if (commanderFilter) commanderFilter.value = 'all';
    if (playerFilter) playerFilter.value = 'all';
    if (dateFromFilter) dateFromFilter.value = '';
    if (dateToFilter) dateToFilter.value = '';
    historySortKey = 'date';
    historySortDescending = true;
    if (sortSelect) sortSelect.value = historySortKey;
    updateSortOrderLabel();
    if (window.location.pathname.endsWith('/history.html') || window.location.pathname.endsWith('history.html')) {
      window.history.replaceState({}, '', 'history.html');
    }
  }

  function getSortedHistoryGames(games) {
    const sorted = [...games];
    sorted.sort((a, b) => {
      if (historySortKey === 'date') {
        const result = (a.date || '').localeCompare(b.date || '');
        return historySortDescending ? -result : result;
      }
      if (historySortKey === 'winner') {
        const result = (getGameWinner(a) || '').localeCompare(getGameWinner(b) || '');
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

  function renderHistoryGame(game, commanderMap) {
    const rows = getGameRows(game);
    const winner = getGameWinner(game) || '—';
    const totalKills = getGameTotalKills(game);
    const playerCount = rows.length;
    const endingTurn = parseOptionalPositiveInteger(game?.liveSummary?.turnNumber);
    const firstBloodLabel = getGameFirstBloodLabel(game);
    const notes = game.notes ? escapeHtml(game.notes) : 'No notes recorded.';
    const gameDate = escapeHtml(game.date || 'Unknown date');
    const playerRows = rows.map((row) => {
      const killed = Array.isArray(row.killed) ? row.killed.join(', ') : row.killed || '';
      const commander = canonicalizeIdentityValue(row.commander, commanderMap);
      return `
        <tr>
          <td>${escapeHtml(row.player)}</td>
          <td>${buildCommanderTextHtml(commander)}</td>
          <td>${row.place || '—'}</td>
          <td>${typeof row.kills === 'number' ? row.kills : 0}</td>
          <td>${escapeHtml(killed)}</td>
        </tr>`;
    }).join('');

    return `
    <article class="history-item">
      <div class="row history-item-meta">
        <h3>${gameDate}</h3>
        <div class="history-item-facts" aria-label="Game summary">
          <span class="history-item-fact">Winner: ${escapeHtml(winner)}</span>
          <span class="history-item-fact">${playerCount} players</span>
          <span class="history-item-fact">${totalKills} kills</span>
          ${endingTurn ? `<span class="history-item-fact">Ended on turn ${endingTurn}</span>` : ''}
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
        <tbody>${playerRows}</tbody>
      </table>
    </article>`;
  }

  function render(games) {
    if (!list) return;
    const sortedGames = getSortedHistoryGames(games);
    if (playerFilter && !playerFilterDefaulted) {
      const displayName = normalizeIdentityLabel(getCurrentSyncDisplayName());
      const currentUserId = normalizeIdentityLabel(getCurrentSyncUserId());
      const availablePlayers = Array.from(playerFilter.options || [])
        .map((option) => String(option.value || '').trim())
        .filter((value) => value && value !== 'all');
      const defaultPlayer = availablePlayers.find((value) => {
        const key = getIdentityKey(value);
        return Boolean(key) && (key === getIdentityKey(displayName) || key === getIdentityKey(currentUserId));
      });
      if (defaultPlayer) {
        playerFilter.value = defaultPlayer;
        playerFilterDefaulted = true;
      }
    }

    const winner = winnerFilter?.value || 'all';
    const commander = commanderFilter?.value || 'all';
    const player = playerFilter?.value || 'all';
    const from = dateFromFilter?.value || '';
    const to = dateToFilter?.value || '';
    const allCommanders = [];
    sortedGames.forEach((game) => {
      getGameRows(game).forEach((row) => {
        if (row.commander) allCommanders.push(row.commander);
      });
    });
    const commanderMap = buildCanonicalIdentityMapFromValues([...getKnownCommanderOptions(), ...allCommanders]);

    const filteredGames = sortedGames.filter((game) => {
      if (from && (game.date || '') < from) return false;
      if (to && (game.date || '') > to) return false;
      if (winner !== 'all' && getGameWinner(game) !== winner) return false;
      if (player !== 'all' && !getGameRows(game).some((row) => row.player === player)) return false;
      if (commander !== 'all' && !getGameRows(game).some((row) => canonicalizeIdentityValue(row.commander, commanderMap) === commander)) return false;
      return true;
    });

    syncFilterQuery();
    renderActiveFilters(sortedGames.length, filteredGames.length);
    if (!filteredGames.length) {
      list.innerHTML = '<p class="history-empty-state">No games match the current filters.</p>';
      return;
    }
    list.innerHTML = filteredGames.map((game) => renderHistoryGame(game, commanderMap)).join('');
  }

  async function handleHistoryAction(event) {
    const button = event.target.closest('button');
    if (!button || !list.contains(button)) return;
    const gameId = button.dataset.id;
    if (!gameId) return;

    if (button.classList.contains('history-delete-button')) {
      const changedRecordTitles = getChangedRecordTitlesAfterGameRemoval(gameId);
      const impactSuffix = changedRecordTitles.length
        ? ` This will also recalculate records such as ${changedRecordTitles.slice(0, 3).join(', ')}${changedRecordTitles.length > 3 ? ', and others' : ''}.`
        : '';
      if (await promptLiveConfirm(`Delete this game? This removes it from history, rankings, and all derived stats.${impactSuffix}`, {
        title: 'Delete saved game?',
        confirmLabel: 'Delete game',
      })) deleteGame(gameId);
      return;
    }

    if (button.classList.contains('history-edit-button') && getGameById(gameId)) {
      window.location.href = `index.html?editId=${encodeURIComponent(gameId)}`;
    }
  }

  function updateSortOrderLabel() {
    if (!sortOrderButton) return;
    sortOrderButton.textContent = historySortKey === 'date'
      ? (historySortDescending ? 'Newest first' : 'Oldest first')
      : (historySortDescending ? 'Descending' : 'Ascending');
  }

  sortSelect?.addEventListener('change', (event) => {
    historySortKey = event.target.value;
    updateSortOrderLabel();
    render(loadGames());
  });
  list?.addEventListener('click', handleHistoryAction);
  sortOrderButton?.addEventListener('click', () => {
    historySortDescending = !historySortDescending;
    updateSortOrderLabel();
    render(loadGames());
  });
  [winnerFilter, commanderFilter, playerFilter, dateFromFilter, dateToFilter].forEach((filter) => {
    filter?.addEventListener('change', () => render(loadGames()));
  });
  resetFiltersButton?.addEventListener('click', () => {
    resetFilters();
    render(loadGames());
  });

  window.CommanderHistory = {
    hasView: () => Boolean(list),
    updateFilters,
    applyQueryFilters,
    getCurrentFilters,
    render,
    updateSortOrderLabel,
    resetFilters,
    resetPlayerFilterForUserChange: () => { playerFilterDefaulted = false; },
  };
})();