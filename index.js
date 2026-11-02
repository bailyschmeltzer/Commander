(function () {
  const gameForm = document.getElementById('game-form');
  const dateInput = document.getElementById('game-date');
  const notesInput = document.getElementById('game-notes');
  const firstBloodPlayerInput = document.getElementById('game-first-blood-player');
  const firstBloodTurnInput = document.getElementById('game-first-blood-turn');
  const winningTurnInput = document.getElementById('game-winning-turn');
  const firstBloodMenu = document.getElementById('game-first-blood-menu');
  const summaryElement = document.getElementById('summary');
  const addRowButton = document.getElementById('add-player-row');
  const removeRowButton = document.getElementById('remove-player-row');
  const playerRows = document.getElementById('player-table-body');
  let editingGameId = null;

  function getPlayerRows() {
    const playerMap = buildCanonicalIdentityMapFromValues(knownPlayers);
    const commanderMap = buildCanonicalIdentityMapFromValues(knownCommanders);
    return Array.from(playerRows.querySelectorAll('tr')).map((row) => {
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
    playerRows.appendChild(row);
    populateRowSelectors(row);
  }

  function removePlayerRow() {
    const rows = playerRows.querySelectorAll('tr');
    if (rows.length > 2) {
      rows[rows.length - 1].remove();
      playerRows.querySelectorAll('tr').forEach((row) => populateRowSelectors(row));
    }
  }

  function resetPlayerTable() {
    playerRows.innerHTML = '';
    for (let rowIndex = 0; rowIndex < 4; rowIndex += 1) {
      addPlayerRow();
    }
  }

  function sanitizeManualGameRows(rows) {
    const playerBuckets = new Map();
    const commanderBuckets = new Map();
    rows.forEach((row) => {
      recordIdentityVariant(playerBuckets, row.player);
      recordIdentityVariant(commanderBuckets, row.commander);
      row.killed.forEach((player) => recordIdentityVariant(playerBuckets, player));
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
    const rowsByPlayerName = new Map(rows.map((row) => [row.player.toLocaleLowerCase(), row]));
    const winnerRow = rows.find((row) => row.place === 1) || null;
    const eliminatedRows = rows.filter((row) => row.place !== 1).slice().sort((a, b) => b.place - a.place);
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
    const maximumCreditedKills = rows.length - 1;
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

  function resetForm() {
    gameForm.reset();
    resetPlayerTable();
    dateInput.valueAsDate = new Date();
  }

  function setRows(rows) {
    playerRows.innerHTML = '';
    (Array.isArray(rows) ? rows : []).forEach((row) => addPlayerRow(row));
  }

  function setEditMode(game) {
    if (!gameForm || !game) {
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
    setRows(getGameRows(game));
    const submitButton = gameForm.querySelector('button[type="submit"]');
    if (submitButton) {
      submitButton.textContent = 'Save changes';
    }
  }

  function resetEditMode() {
    editingGameId = null;
    const submitButton = gameForm?.querySelector('button[type="submit"]');
    if (submitButton) {
      submitButton.textContent = 'Save game';
    }
  }

  function updatePlayerOptions(players) {
    if (firstBloodMenu) {
      buildDropdownMenu(firstBloodMenu, Array.isArray(players) ? players : []);
    }
  }

  function renderSummary(games) {
    if (!summaryElement) {
      return;
    }
    summaryElement.innerHTML = '';
    if (!games.length) {
      summaryElement.appendChild(createStatCard('No games yet', 'Add a game to start tracking commander results.'));
      return;
    }
    const total = games.length;
    const mostRecent = games.slice().sort((a, b) => b.date.localeCompare(a.date))[0].date;
    const playerWins = {};
    const commanderCount = {};
    games.forEach((game) => {
      if (game.finishOrder && game.finishOrder.length) {
        const winner = game.finishOrder[0];
        playerWins[winner] = (playerWins[winner] || 0) + 1;
      }
      getGameRows(game).forEach(({ commander }) => {
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
      .map(([commander, commanderGames]) => `${buildCommanderTextHtml(commander)} (${commanderGames})`)
      .join('<br />');
    summaryElement.appendChild(createStatCard('Total games', total));
    summaryElement.appendChild(createStatCard('Most recent game', mostRecent));
    summaryElement.appendChild(createStatCard('Top winners', bestPlayers || 'No winners yet'));
    summaryElement.appendChild(createStatCard('Popular commanders', bestCommanders || 'No commander data yet'));
  }

  window.CommanderGameTracker = {
    hasView: () => Boolean(gameForm),
    resetForm,
    setEditMode,
    resetEditMode,
    addPlayerRow,
    removePlayerRow,
    resetPlayerTable,
    getPlayerRows,
    updatePlayerOptions,
    renderSummary,
  };

  addRowButton?.addEventListener('click', () => {
    addPlayerRow();
  });

  removeRowButton?.addEventListener('click', () => {
    removePlayerRow();
  });

  playerRows?.addEventListener('input', (event) => {
    const input = event.target.closest('input[name="player"], input[name="commander"]');
    if (!input) return;
    const row = input.closest('tr');
    if (row) populateRowSelectors(row);
  });

  gameForm?.addEventListener('submit', async (event) => {
    event.preventDefault();
    const games = loadGames();
    const existingGame = editingGameId ? games.find((game) => game.id === editingGameId) || null : null;
    const rows = getPlayerRows();

    if (!rows.length) {
      await promptLiveAlert('Please add at least one player row with a player name.', 'Unable to save game');
      return;
    }

    const commanderValidationError = await validateCommanderEntries(rows, { allowExactCardLookup: true });
    if (commanderValidationError) {
      await promptLiveAlert(commanderValidationError, 'Unable to save game');
      return;
    }

    const firstBloodPlayer = firstBloodPlayerInput?.value.trim() || '';
    const firstBloodTurn = parseOptionalPositiveInteger(firstBloodTurnInput?.value || '');
    const winningTurn = parseOptionalPositiveInteger(winningTurnInput?.value || '');
    const validationError = validateManualGameEntry(rows, { firstBloodPlayer, firstBloodTurn, winningTurn });
    if (validationError) {
      await promptLiveAlert(validationError, 'Unable to save game');
      return;
    }

    const advisories = getManualGameAdvisories(rows, { firstBloodPlayer, firstBloodTurn, winningTurn });
    if (advisories.length && !await promptLiveConfirm(`This game is valid, but a few details are unusual: ${advisories.join(' ')} Save it anyway?`, {
      title: 'Save unusual game?',
      confirmLabel: 'Save game',
    })) return;

    const playerRows = sanitizeManualGameRows(rows);
    const players = playerRows.map(({ player }) => player);
    const playerCommanders = playerRows.map(({ player, commander }) => ({ player, commander }));
    const finishOrder = playerRows
      .slice()
      .filter((row) => row.place !== null)
      .sort((first, second) => (first.place || 999) - (second.place || 999))
      .map((row) => row.player);
    const manualLiveSummary = buildManualGameLiveSummary(
      playerRows,
      dateInput.value || new Date().toISOString().slice(0, 10),
      firstBloodPlayer,
      firstBloodTurn,
      winningTurn,
    );
    const game = buildCompactSavedGameEntry({
      id: editingGameId || generateId(),
      date: dateInput.value || new Date().toISOString().slice(0, 10),
      playerRows,
      players,
      playerCommanders,
      finishOrder,
      liveSummary: mergeEditedGameLiveSummary(existingGame?.liveSummary, manualLiveSummary),
      notes: notesInput.value.trim(),
    });

    if (editingGameId) {
      const index = games.findIndex((entry) => entry.id === editingGameId);
      if (index >= 0) games[index] = game;
      else games.push(game);
    } else {
      games.push(game);
    }

    saveGames(games);
    resetEditMode();
    refresh();
    resetForm();
  });
})();