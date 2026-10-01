(function () {
  const nameInput = document.getElementById('deck-builder-name');
  const ownerInput = document.getElementById('deck-builder-owner');
  const powerInput = document.getElementById('deck-builder-power-level');
  const inRotationInput = document.getElementById('deck-builder-in-rotation');
  const searchInput = document.getElementById('deck-builder-search');
  const tokenSearchInput = document.getElementById('deck-builder-token-search');
  const searchResults = document.getElementById('deck-builder-search-results');
  const tokenSearchResults = document.getElementById('deck-builder-token-search-results');
  const selection = document.getElementById('deck-builder-selection');
  const cards = document.getElementById('deck-builder-cards');
  const exportButton = document.getElementById('deck-builder-export');
  const importButton = document.getElementById('deck-builder-import');
  const textList = document.getElementById('deck-builder-text-list');
  const importStatus = document.getElementById('deck-builder-import-status');
  const edhplButton = document.getElementById('deck-builder-edhpl');
  const preconSelect = document.getElementById('deck-builder-precon');
  const preconLoadButton = document.getElementById('deck-builder-precon-load');
  const discardButton = document.getElementById('deck-builder-discard');
  const saveButton = document.getElementById('deck-builder-save');
  const undoButton = document.getElementById('deck-builder-undo');
  const manaCurve = document.getElementById('deck-builder-mana-curve');
  const backToDecksLink = document.querySelector('.page-deckbuilder .page-action-link[href="decklists.html"], .page-deck-builder .page-action-link[href="decklists.html"]');

  nameInput?.addEventListener('input', () => queueDeckBuilderMetaSave());
  ownerInput?.addEventListener('input', () => queueDeckBuilderMetaSave());
  powerInput?.addEventListener('input', () => queueDeckBuilderMetaSave());
  inRotationInput?.addEventListener('change', () => queueDeckBuilderMetaSave());

  searchInput?.addEventListener('input', () => queueDeckBuilderSearch(searchInput.value));
  searchInput?.addEventListener('keydown', async (event) => {
    if (event.key !== 'Enter' || deckBuilderSearchResultsState.length !== 1) return;
    event.preventDefault();
    await selectAndAddDeckSearchResult(deckBuilderSearchResultsState[0]);
  });

  tokenSearchInput?.addEventListener('input', () => queueDeckBuilderTokenSearch(tokenSearchInput.value));
  tokenSearchInput?.addEventListener('keydown', async (event) => {
    if (event.key !== 'Enter' || deckBuilderTokenSearchResultsState.length !== 1) return;
    event.preventDefault();
    await selectAndAddDeckTokenSearchResult(deckBuilderTokenSearchResultsState[0]);
  });

  searchResults?.addEventListener('click', async (event) => {
    const option = event.target.closest('.deck-search-result');
    if (option) await selectDeckBuilderSearchResult(option.dataset.name || '');
  });
  searchResults?.addEventListener('dblclick', async (event) => {
    const option = event.target.closest('.deck-search-result');
    if (option) await selectAndAddDeckSearchResult(option.dataset.name || '');
  });

  tokenSearchResults?.addEventListener('click', async (event) => {
    const option = event.target.closest('[data-token-name]');
    if (option) await selectAndAddDeckTokenSearchResult(option.dataset.tokenName || '');
  });
  tokenSearchResults?.addEventListener('dblclick', async (event) => {
    const option = event.target.closest('[data-token-name]');
    if (option) await selectAndAddDeckTokenSearchResult(option.dataset.tokenName || '');
  });

  selection?.addEventListener('click', async (event) => {
    if (event.target.closest('#deck-builder-add-card')) {
      event.stopPropagation();
      await addSelectedCardToDeck();
      return;
    }
    if (event.target.closest('#deck-builder-add-token')) {
      event.stopPropagation();
      await addSelectedCardToTokens();
      return;
    }
    if (event.target.closest('#deck-builder-add-maybeboard')) {
      event.stopPropagation();
      await addSelectedCardToMaybeboard();
      return;
    }
    if (event.target.closest('#deck-builder-set-commander')) {
      event.stopPropagation();
      await setSelectedCardAsCommander();
    }
  });

  document.addEventListener('click', async (event) => {
    if (event.target.closest('#deck-builder-add-card')) {
      await addSelectedCardToDeck();
      return;
    }
    if (event.target.closest('#deck-builder-add-token')) {
      await addSelectedCardToTokens();
      return;
    }
    if (event.target.closest('#deck-builder-add-maybeboard')) {
      await addSelectedCardToMaybeboard();
      return;
    }
    if (event.target.closest('#deck-builder-set-commander')) await setSelectedCardAsCommander();
  });

  cards?.addEventListener('pointerdown', (event) => {
    const holdButton = event.target.closest('[data-add-basic], [data-remove-basic]');
    if (holdButton) startDeckBuilderHoldRepeat(holdButton, { applyInitialChange: true });
  });

  cards?.addEventListener('click', async (event) => {
    const cardImage = event.target.closest('.deck-card-row-image');
    if (cardImage) {
      event.preventDefault();
      event.stopPropagation();
      openCardImageModal(cardImage.currentSrc || cardImage.src, cardImage.alt || 'Card image');
      return;
    }

    const changeArtButton = event.target.closest('[data-change-art-id]');
    if (changeArtButton) {
      const cardId = changeArtButton.dataset.changeArtId || '';
      if (!cardId) return;
      if (deckBuilderArtPickerCardId === cardId) {
        deckBuilderArtPickerCardId = '';
        deckBuilderArtPickerState = { status: 'idle', cardId: '', options: [], message: '' };
        const deck = ensureActiveDeckBuilderRecord();
        if (deck) renderDeckBuilderCards(deck);
        return;
      }
      await openDeckBuilderArtPicker(cardId);
      return;
    }

    const applyArtButton = event.target.closest('[data-apply-art-print-id][data-apply-art-card-id]');
    if (applyArtButton) {
      const cardId = applyArtButton.dataset.applyArtCardId || '';
      const printId = applyArtButton.dataset.applyArtPrintId || '';
      const selectedPrint = (deckBuilderArtPickerState.options || []).find((option) => String(option.id || '') === String(printId || ''));
      if (cardId && selectedPrint) applyDeckBuilderCardArt(cardId, selectedPrint);
      return;
    }

    const removeCommanderButton = event.target.closest('[data-remove-commander="true"]');
    if (removeCommanderButton) {
      const isSecond = removeCommanderButton.dataset.removeSecondCommander === 'true';
      removeDeckBuilderCard('', { isCommander: !isSecond, isSecondCommander: isSecond });
      return;
    }

    const setCommanderButton = event.target.closest('[data-set-commander-id]');
    if (setCommanderButton) {
      await setDeckBuilderCardAsCommander(setCommanderButton.dataset.setCommanderId || '');
      return;
    }

    const addFromMaybeboardButton = event.target.closest('[data-add-from-maybeboard-id]');
    if (addFromMaybeboardButton) {
      await addMaybeboardCardToDeck(addFromMaybeboardButton.dataset.addFromMaybeboardId || '');
      return;
    }

    const moveToMaybeboardButton = event.target.closest('[data-move-to-maybeboard-id]');
    if (moveToMaybeboardButton) {
      await moveDeckCardToMaybeboard(moveToMaybeboardButton.dataset.moveToMaybeboardId || '');
      return;
    }

    const removeTokenButton = event.target.closest('[data-remove-token-id]');
    if (removeTokenButton) {
      updateDeckBuilderTokenQuantity(removeTokenButton.dataset.removeTokenId || '', -1);
      return;
    }

    const addTokenButton = event.target.closest('[data-add-token-id]');
    if (addTokenButton) {
      updateDeckBuilderTokenQuantity(addTokenButton.dataset.addTokenId || '', 1);
      return;
    }

    const removeBasicButton = event.target.closest('[data-remove-basic]');
    if (removeBasicButton) {
      if (event.detail !== 0) return;
      removeBasicLandFromDeck(removeBasicButton.dataset.removeBasic || '');
      return;
    }

    const removeCardButton = event.target.closest('.deck-builder-remove-card[data-card-id]');
    if (removeCardButton) {
      removeDeckBuilderCard(removeCardButton.dataset.cardId || '');
      return;
    }

    const addBasicButton = event.target.closest('[data-add-basic]');
    if (addBasicButton) {
      if (event.detail !== 0) return;
      await addBasicLandToDeck(addBasicButton.dataset.addBasic || '');
      return;
    }

    const addUnlimitedButton = event.target.closest('[data-add-unlimited]');
    if (addUnlimitedButton) {
      await addUnlimitedCopyCardToDeck(addUnlimitedButton.dataset.addUnlimited || '');
      return;
    }

    const deckCardRow = event.target.closest('.deck-card-row[data-card-name]');
    if (deckCardRow && !event.target.closest('.deck-builder-remove-card')) {
      const cardId = deckCardRow.dataset.cardId || '';
      if (deckBuilderSelectedDeckCardId === cardId) {
        deckBuilderSelectedDeckCardId = null;
        if (deckBuilderArtPickerCardId === cardId) {
          deckBuilderArtPickerCardId = '';
          deckBuilderArtPickerState = { status: 'idle', cardId: '', options: [], message: '' };
        }
      } else {
        deckBuilderSelectedDeckCardId = cardId;
        deckBuilderSelectedFaceIndex = 0;
        void selectDeckCardByName(deckCardRow.dataset.cardName || '', cardId);
      }
      const deck = ensureActiveDeckBuilderRecord();
      if (deck) renderDeckBuilderCards(deck);
    }
  });

  ['pointerup', 'pointerleave', 'pointercancel'].forEach((eventName) => {
    cards?.addEventListener(eventName, () => stopDeckBuilderHoldRepeat());
  });
  ['pointerup', 'pointercancel'].forEach((eventName) => {
    document.addEventListener(eventName, () => stopDeckBuilderHoldRepeat());
  });

  exportButton?.addEventListener('click', () => {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) {
      setDeckBuilderImportStatus('No active deck to export.', 'error');
      return;
    }
    if (textList) textList.value = exportDeckAsText(deck);
    setDeckBuilderImportStatus('Deck exported to text above.', 'success');
  });

  importButton?.addEventListener('click', async () => {
    if (!textList?.value?.trim()) {
      setDeckBuilderImportStatus('Paste a deck list above first.', 'error');
      return;
    }
    await importDeckFromText(textList.value);
  });

  if (preconSelect) {
    let preconListLoaded = false;
    preconSelect.addEventListener('focus', async () => {
      if (preconListLoaded) return;
      const defaultOption = preconSelect.options[0];
      defaultOption.textContent = 'Loading precons...';
      defaultOption.disabled = true;
      try {
        const decks = await fetchPreconList();
        preconListLoaded = true;
        defaultOption.textContent = 'Select a precon...';
        defaultOption.disabled = false;
        decks.forEach((deck) => {
          const option = document.createElement('option');
          option.value = deck.fileName;
          option.textContent = `${deck.name} (${deck.code})`;
          preconSelect.appendChild(option);
        });
      } catch (_) {
        defaultOption.textContent = 'Failed to load precons';
        defaultOption.disabled = false;
      }
    });
    preconSelect.addEventListener('change', () => {
      if (preconLoadButton) preconLoadButton.disabled = !preconSelect.value;
    });
  }

  preconLoadButton?.addEventListener('click', async () => {
    const fileName = preconSelect?.value;
    if (!fileName) return;
    preconLoadButton.disabled = true;
    setDeckBuilderImportStatus('Loading precon...', 'muted');
    try {
      const text = await loadPreconDeckText(fileName);
      if (textList) textList.value = text;
      await importDeckFromText(text);
    } catch (error) {
      setDeckBuilderImportStatus(error instanceof Error ? error.message : 'Failed to load precon.', 'error');
    } finally {
      preconLoadButton.disabled = false;
    }
  });

  edhplButton?.addEventListener('click', () => {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) {
      setDeckBuilderImportStatus('No active deck to analyze.', 'error');
      return;
    }
    const deckText = buildEdhPowerLevelDeckText(deck);
    if (!deckText.trim()) {
      setDeckBuilderImportStatus('Add cards before analyzing deck strength.', 'error');
      return;
    }
    window.open(`https://edhpowerlevel.com?d=${encodeEdhPowerLevelDeck(deckText)}`, '_blank', 'noopener,noreferrer');
    setDeckBuilderImportStatus('Opened EDHPowerLevel analysis in a new tab.', 'success');
  });

  discardButton?.addEventListener('click', async () => {
    const deck = activeDeckBuilderRecord;
    if (!deck?.id) return;
    const confirmed = await promptLiveConfirm(`Delete "${deck.name}"? This cannot be undone.`, {
      title: 'Delete deck',
      confirmLabel: 'Delete deck',
    });
    if (!confirmed) return;
    saveDecks(loadDecks().filter((entry) => entry.id !== deck.id));
    activeDeckBuilderId = '';
    activeDeckBuilderRecord = null;
    if (syncQueueTimer) {
      clearTimeout(syncQueueTimer);
      syncQueueTimer = null;
    }
    try {
      await pushCloudState();
    } catch (_) {
      // Local deletion is already saved; cloud sync can retry later.
    }
    window.location.href = 'decklists.html';
  });

  saveButton?.addEventListener('click', async () => await saveActiveDeckBuilderDeck());

  undoButton?.addEventListener('click', async () => await undoDeckBuilderChange());

  manaCurve?.addEventListener('click', async (event) => {
    const bar = event.target.closest('.deck-builder-curve-bar[data-mana-value]');
    if (!(bar instanceof HTMLButtonElement) || bar.disabled) return;
    const manaValue = Number.parseInt(bar.dataset.manaValue || '', 10);
    if (Number.isFinite(manaValue)) await showDeckBuilderManaValueCards(manaValue);
  });

  window.addEventListener('beforeunload', (event) => {
    if (typeof hasUnsavedDeckBuilderChanges === 'function' && hasUnsavedDeckBuilderChanges()) {
      event.preventDefault();
      event.returnValue = '';
    }
  });

  backToDecksLink?.addEventListener('click', (event) => {
    // Navigating away without pressing Save Deck discards any pending changes.
    if (hasUnsavedDeckBuilderChanges()) {
      setDeckBuilderSaveStatus('Unsaved changes discarded.', 'muted');
    }
    window.location.href = backToDecksLink.getAttribute('href') || 'decklists.html';
  });
})();