(function () {
  const gameForm = document.getElementById('game-form');
  const addRowButton = document.getElementById('add-player-row');
  const removeRowButton = document.getElementById('remove-player-row');
  const playerRows = document.getElementById('player-table-body');

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