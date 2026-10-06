(function () {
  const createButton = document.getElementById('deck-library-create');
  const playerFilter = document.getElementById('deck-library-player-filter');
  const rotationFilter = document.getElementById('deck-library-rotation-filter');
  const clearFiltersButton = document.getElementById('deck-library-player-filter-clear');
  const tableBody = document.getElementById('deck-library-body');
  let playerFilterDefaulted = false;
  let cloudWaitStartAt = 0;
  let cloudWaitTimer = null;

  function render() {
    if (!tableBody) return;

    // While connected and the cloud state hasn't resolved yet, show a loading row
    // instead of painting local cache. This prevents the deck list from briefly (or
    // misleadingly) showing stale/empty local data before the authoritative cloud
    // pull completes and re-renders with fresh data. Bounded: if the fetch hangs or
    // fails, fall back to local data after a few seconds so it never hangs forever.
    const stillResolving = typeof hasLoadedCloudState === 'function'
      && !hasLoadedCloudState()
      && typeof hasSyncCredentials === 'function'
      && hasSyncCredentials()
      && navigator.onLine;

    if (stillResolving) {
      if (!cloudWaitStartAt) {
        cloudWaitStartAt = Date.now();
      }
      const elapsed = Date.now() - cloudWaitStartAt;
      if (elapsed < 4000) {
        tableBody.innerHTML = '<tr><td colspan="7">Loading decks…</td></tr>';
        updateSortableTableIndicators('decks');
        // Re-render when the wait window expires so we fall back to local data even
        // if the cloud fetch never resolves (otherwise this could hang on "Loading…").
        if (!cloudWaitTimer) {
          cloudWaitTimer = setTimeout(() => {
            cloudWaitTimer = null;
            render();
          }, 4000 - elapsed + 50);
        }
        return;
      }
      // Timed out waiting — fall through and render local data rather than hang.
    } else {
      cloudWaitStartAt = 0;
      if (cloudWaitTimer) {
        clearTimeout(cloudWaitTimer);
        cloudWaitTimer = null;
      }
    }

    const currentUserId = getCurrentSyncUserId();
    const sortState = getTableSort('decks', 'updatedAt', true);
    const deckUsageLookup = buildDeckUsageLookup();
    const visibleDecks = loadDecks().filter((deck) => {
      const ownerUserId = String(deck?.ownerUserId || '').trim().toLowerCase();
      const isOwnedByCurrentUser = Boolean(ownerUserId) && Boolean(currentUserId) && ownerUserId === currentUserId;
      const hasGameRecord = isDeckUsedInGameFromLookup(deck, deckUsageLookup);
      const isInRotation = deck?.inRotation !== false;
      return isOwnedByCurrentUser || isInRotation || hasGameRecord;
    });

    const sortedDecks = visibleDecks.slice().sort((first, second) => {
      let result = 0;
      switch (sortState.column) {
        case 'name': result = compareTextValues(first.name, second.name); break;
        case 'owner': result = compareTextValues(first.owner, second.owner); break;
        case 'commander': result = compareTextValues(first.commander?.name, second.commander?.name); break;
        case 'powerLevel': result = compareNumberValues(first.powerLevel, second.powerLevel); break;
        case 'updatedAt':
        default: result = compareDateValues(first.updatedAt, second.updatedAt); break;
      }

      if (result === 0) result = compareTextValues(first.name, second.name);
      return finalizeSortResult(result, sortState.descending);
    });

    const ownerFilterOptions = getUniqueValues(sortedDecks.map((deck) => normalizeIdentityLabel(deck.owner || '')).filter(Boolean));
    const requestedOwnerFilter = normalizeIdentityLabel(playerFilter?.value || '');
    let activeOwnerFilter = ownerFilterOptions.includes(requestedOwnerFilter) ? requestedOwnerFilter : '';

    if (!activeOwnerFilter && !playerFilterDefaulted) {
      const displayName = normalizeIdentityLabel(getCurrentSyncDisplayName());
      const currentUser = getCurrentSyncUserId();
      const ownedDeckByUserId = sortedDecks.find((deck) => String(deck?.ownerUserId || '').trim().toLowerCase() === currentUser);

      if (ownedDeckByUserId) {
        activeOwnerFilter = normalizeIdentityLabel(ownedDeckByUserId.owner || displayName || currentUser);
        playerFilterDefaulted = true;
      }

      if (!activeOwnerFilter && displayName && ownerFilterOptions.includes(displayName)) {
        activeOwnerFilter = displayName;
        playerFilterDefaulted = true;
      } else if (!activeOwnerFilter && ownedDeckByUserId) {
        activeOwnerFilter = normalizeIdentityLabel(ownedDeckByUserId.owner || displayName || currentUser);
        playerFilterDefaulted = true;
      }
    }

    if (playerFilter) {
      buildSelectOptions(playerFilter, ownerFilterOptions, activeOwnerFilter, 'All players');
    }

    const activeOwnerFilterKey = getIdentityKey(activeOwnerFilter);
    const activeRotationFilter = String(rotationFilter?.value || 'all').trim().toLowerCase();
    let decks = activeOwnerFilter
      ? sortedDecks.filter((deck) => {
        const ownerLabel = normalizeIdentityLabel(deck.owner || '');
        const ownerUserIdKey = getIdentityKey(deck.ownerUserId || '');
        return ownerLabel === activeOwnerFilter || (activeOwnerFilterKey && ownerUserIdKey === activeOwnerFilterKey);
      })
      : sortedDecks;

    if (activeRotationFilter === 'active') {
      decks = decks.filter((deck) => deck.inRotation !== false);
    } else if (activeRotationFilter === 'inactive') {
      decks = decks.filter((deck) => deck.inRotation === false);
    }

    if (activeOwnerFilter && !decks.length) {
      activeOwnerFilter = '';
      if (playerFilter) buildSelectOptions(playerFilter, ownerFilterOptions, '', 'All players');
      decks = activeRotationFilter === 'active'
        ? sortedDecks.filter((deck) => deck.inRotation !== false)
        : activeRotationFilter === 'inactive'
          ? sortedDecks.filter((deck) => deck.inRotation === false)
          : sortedDecks;
    }

    const commanderPointsPerGameByIdentity = new Map(
      buildCommanderRankingEntries(loadGames()).map((entry) => [getIdentityKey(entry.name), entry.pointsPerGame])
    );

    if (!sortedDecks.length) {
      tableBody.innerHTML = '<tr><td colspan="7">No built decks yet. Click Add New Deck to start one.</td></tr>';
      updateSortableTableIndicators('decks');
      return;
    }

    if (!decks.length) {
      const emptyMessage = activeOwnerFilter
        ? 'No built decks found for that player and deck status.'
        : activeRotationFilter === 'active'
          ? 'No active decks found.'
          : activeRotationFilter === 'inactive'
            ? 'No inactive decks found.'
            : 'No built decks found for that filter.';
      tableBody.innerHTML = `<tr><td colspan="7">${escapeHtml(emptyMessage)}</td></tr>`;
      updateSortableTableIndicators('decks');
      return;
    }

    tableBody.innerHTML = decks.map((deck) => {
      const summary = getDeckValidationSummary(deck);
      const powerLevel = Number.isFinite(deck.powerLevel) ? deck.powerLevel.toFixed(1).replace(/\.0$/, '') : '—';
      const canEditDeck = canCurrentUserEditDeck(deck);
      const hasGameRecord = isDeckUsedInGameFromLookup(deck, deckUsageLookup);
      const isInRotation = deck.inRotation !== false;
      const rotationLabel = isInRotation ? 'In rotation' : 'Out of rotation';
      const warnings = [
        deck.ownerUserId && !canEditDeck ? 'locked' : '',
        summary.bannedCards.length ? `${summary.bannedCards.length} banned` : '',
      ].filter(Boolean).join(', ') || '—';
      const commanderPointsPerGame = commanderPointsPerGameByIdentity.get(getIdentityKey(deck.commander?.name || ''));

      return `
      <tr>
        <td data-label="Deck">
          <span class="deck-library-name-cell">
            ${escapeHtml(deck.name)}
            <span class="deck-rotation-indicator ${isInRotation ? 'is-in' : 'is-out'}" role="img" aria-label="${escapeHtml(rotationLabel)}" title="${escapeHtml(rotationLabel)}">${isInRotation ? '✓' : 'X'}</span>
          </span>
        </td>
        <td data-label="Owner">${escapeHtml(deck.owner || '—')}</td>
        <td data-label="Commander">${deck.commander?.name ? buildCommanderTextHtml(deck.commander.name, deck.commander) : '—'}</td>
        <td data-label="Power Level">${escapeHtml(String(powerLevel))}</td>
        <td data-label="Avg Points/Game">${deck.commander?.name ? formatPointsPerGame(commanderPointsPerGame ?? 0) : '—'}</td>
        <td data-label="Status">${escapeHtml(getDeckSummaryLabel(deck, summary))}${warnings !== '—' ? `<div class="deck-library-warning-text">${escapeHtml(warnings)}</div>` : ''}</td>
        <td data-label="Actions">
          <button type="button" class="secondary-button deck-library-open" data-id="${escapeHtml(deck.id)}">Open</button>
          ${hasGameRecord ? '' : `<button type="button" class="history-delete-button deck-library-delete" data-id="${escapeHtml(deck.id)}"${canEditDeck ? '' : ' disabled'}>Delete</button>`}
        </td>
      </tr>`;
    }).join('');

    updateSortableTableIndicators('decks');
  }

  async function deleteDeck(deckId) {
    const deck = loadDecks().find((entry) => entry.id === deckId);
    if (!deck) return;

    if (!canCurrentUserEditDeck(deck)) {
      setDeckBuilderSaveStatus(getDeckReadOnlyMessage(deck), 'error');
      return;
    }

    const confirmed = await promptLiveConfirm(`Delete ${deck.name}? This removes the saved deck from this app.`, {
      title: 'Delete deck',
      confirmLabel: 'Delete deck',
    });
    if (!confirmed) return;

    saveDecks(loadDecks().filter((entry) => entry.id !== deckId));
    render();
  }

  function resetFilterForUserChange({ clearSelection = false } = {}) {
    playerFilterDefaulted = false;
    if (clearSelection && playerFilter) playerFilter.value = '';
  }

  createButton?.addEventListener('click', () => {
    window.location.href = 'deckbuilder.html?new=1';
  });

  playerFilter?.addEventListener('change', () => {
    render();
    applyResponsiveTableLabels();
  });

  rotationFilter?.addEventListener('change', () => {
    render();
    applyResponsiveTableLabels();
  });

  clearFiltersButton?.addEventListener('click', () => {
    if (playerFilter) playerFilter.value = '';
    if (rotationFilter) rotationFilter.value = 'all';
    render();
    applyResponsiveTableLabels();
  });

  tableBody?.addEventListener('click', async (event) => {
    const openButton = event.target.closest('.deck-library-open');
    if (openButton) {
      const deckId = openButton.dataset.id || '';
      if (deckId) window.location.href = getDeckBuilderHref(deckId);
      return;
    }

    const deleteButton = event.target.closest('.deck-library-delete');
    if (deleteButton) {
      const deckId = deleteButton.dataset.id || '';
      if (deckId) await deleteDeck(deckId);
    }
  });

  window.CommanderDeckLibrary = {
    hasView: () => Boolean(tableBody),
    render,
    resetFilterForUserChange,
  };
})();