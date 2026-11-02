(function () {
  const DECK_BUILDER_SELECTED_CARD_STORAGE_KEY = 'deckBuilderSelectedCardDraft';
  const DECK_BUILDER_META_SAVE_DEBOUNCE_MS = 120;
  const DECK_CARD_ARTS_ENDPOINT = '/api/deck-card-arts';
  const BASIC_LAND_NAMES = ['Plains', 'Island', 'Swamp', 'Mountain', 'Forest', 'Wastes'];
  const DECK_BUILDER_HOLD_REPEAT_START_DELAY_MS = 1000;
  const DECK_BUILDER_HOLD_REPEAT_INTERVAL_MS = 90;
  const deckBuilderPage = document.querySelector('.page-deckbuilder, .page-deck-builder');
  const deckBuilderTitle = document.getElementById('deck-builder-title');
  const deckBuilderAccessBadge = document.getElementById('deck-builder-access-badge');
  const deckBuilderNameInput = document.getElementById('deck-builder-name');
  const deckBuilderOwnerInput = document.getElementById('deck-builder-owner');
  const deckBuilderPowerInput = document.getElementById('deck-builder-power-level');
  const deckBuilderInRotationInput = document.getElementById('deck-builder-in-rotation');
  const deckBuilderOwnerMenu = document.getElementById('deck-builder-owner-menu');
  const deckBuilderCardCount = document.getElementById('deck-builder-card-count');
  const deckBuilderSaveStatus = document.getElementById('deck-builder-save-status');
  const deckBuilderValidation = document.getElementById('deck-builder-validation');
  const deckBuilderManaCurve = document.getElementById('deck-builder-mana-curve');
  const deckBuilderSearchInput = document.getElementById('deck-builder-search');
  const deckBuilderSearchResults = document.getElementById('deck-builder-search-results');
  const deckBuilderSearchStatus = document.getElementById('deck-builder-search-status');
  const deckBuilderTokenSearchInput = document.getElementById('deck-builder-token-search');
  const deckBuilderTokenSearchResults = document.getElementById('deck-builder-token-search-results');
  const deckBuilderTokenSearchStatus = document.getElementById('deck-builder-token-search-status');
  const deckBuilderSelection = document.getElementById('deck-builder-selection');
  const deckBuilderCards = document.getElementById('deck-builder-cards');
  const deckBuilderBreakdown = document.getElementById('deck-builder-breakdown');
  const deckBuilderTextList = document.getElementById('deck-builder-text-list');
  const deckBuilderImportButton = document.getElementById('deck-builder-import');
  const deckBuilderImportStatus = document.getElementById('deck-builder-import-status');
  const deckBuilderPreconSelect = document.getElementById('deck-builder-precon');
  const deckBuilderPreconLoadButton = document.getElementById('deck-builder-precon-load');
  const deckBuilderUndoButton = document.getElementById('deck-builder-undo');
  const deckBuilderDiscardButton = document.getElementById('deck-builder-discard');
  const deckBuilderSaveButton = document.getElementById('deck-builder-save');

  if (deckBuilderSaveStatus) {
    deckBuilderSaveStatus.setAttribute('role', 'status');
    deckBuilderSaveStatus.setAttribute('aria-live', 'polite');
  }
  if (deckBuilderSearchStatus) {
    deckBuilderSearchStatus.setAttribute('role', 'status');
    deckBuilderSearchStatus.setAttribute('aria-live', 'polite');
  }
  if (deckBuilderTokenSearchStatus) {
    deckBuilderTokenSearchStatus.setAttribute('role', 'status');
    deckBuilderTokenSearchStatus.setAttribute('aria-live', 'polite');
  }

  let activeDeckBuilderId = '';
  let activeDeckBuilderRecord = null;
  let deckBuilderHasUnsavedChanges = false;
  let deckBuilderSaveConfirmTimer = null;
  let deckBuilderSearchRequestId = 0;
  let deckBuilderSearchTimer = null;
  let deckBuilderSearchLoading = false;
  let deckBuilderSearchResultsState = [];
  let deckBuilderTokenSearchRequestId = 0;
  let deckBuilderTokenSearchTimer = null;
  let deckBuilderTokenSearchLoading = false;
  let deckBuilderTokenSearchResultsState = [];
  let deckBuilderSelectedCard = null;
  let deckBuilderSelectedDeckCardId = null;
  let deckBuilderSelectedFaceIndex = 0;
  let deckBuilderSaveTimer = null;
  let deckBuilderArtPickerCardId = '';
  let deckBuilderArtPickerState = { status: 'idle', cardId: '', options: [], message: '' };
  let deckBuilderCommanderPrefill = '';
  const deckBuilderArtOptionsCache = new Map();
  const deckBuilderUndoStacks = {};
  const hydratedDeckIds = new Set();
  let deckBuilderBasicLandWarmPromise = null;
  let deckBuilderHoldTimerId = null;
  let deckBuilderHoldIntervalId = null;
  let deckBuilderMutationQueue = Promise.resolve();

  function getDeckBuilderGroupedCards(deck) {
    const typeOrder = ['Creature', 'Artifact', 'Enchantment', 'Planeswalker', 'Battle', 'Instant', 'Sorcery', 'Land', 'Other'];
    const groups = new Map(typeOrder.map((type) => [type, []]));
    (deck?.cards || []).forEach((card) => {
      const type = typeOrder.includes(card.cardType) ? card.cardType : 'Other';
      groups.get(type).push(card);
    });
    typeOrder.forEach((type) => groups.get(type).sort((first, second) => compareTextValues(first.name, second.name)));
    return typeOrder.map((type) => ({ type, cards: groups.get(type) })).filter((group) => group.cards.length);
  }

  function renderDeckBuilderValidation(deck) {
    if (!deckBuilderValidation || !deckBuilderCardCount) return;
    const summary = getDeckValidationSummary(deck);
    deckBuilderCardCount.textContent = `${summary.totalCards} / 100 cards`;
    const lines = [
      (summary.commanderCount === 1 || summary.commanderCount === 2) ? null : { label: 'Deck needs at least one commander.', tone: 'error' },
      summary.totalCards === 100 ? null : { label: `Deck currently has ${summary.totalCards} cards.`, tone: 'error' },
      summary.duplicates.length ? { label: `Duplicate card names found: ${summary.duplicates.join(', ')}.`, tone: 'error' } : null,
      summary.bannedCards.length ? { label: `Banned cards: ${summary.bannedCards.join(', ')}.`, tone: 'error' } : null,
      summary.colorViolations.length ? { label: `Color identity violations: ${summary.colorViolations.join(', ')}.`, tone: 'error' } : null,
    ].filter(Boolean);
    deckBuilderValidation.innerHTML = lines.map((line) => `<li class="deck-validation-item deck-validation-${line.tone}">${escapeHtml(line.label)}</li>`).join('');
  }

  function getDeckManaSymbolValue(rawSymbol) {
    const symbol = String(rawSymbol || '').trim().toUpperCase();
    if (!symbol || /^\d+$/.test(symbol) || ['X', 'Y', 'Z'].includes(symbol)) {
      return /^\d+$/.test(symbol) ? Number.parseInt(symbol, 10) : 0;
    }
    if (symbol.includes('/')) {
      const numericPart = symbol.split('/').map((part) => part.trim()).find((part) => /^\d+$/.test(part));
      return numericPart ? Number.parseInt(numericPart, 10) : 1;
    }
    return 1;
  }

  function getDeckCardManaValue(card) {
    const normalizedCard = normalizeDeckCardEntry(card);
    if (!normalizedCard || normalizedCard.cardType === 'Land') return null;
    const manaCost = normalizedCard.manaCost || normalizedCard.cardFaces[0]?.manaCost || '';
    const symbols = manaCost.match(/\{[^}]+\}/g);
    if (!symbols?.length) return 0;
    return symbols.reduce((total, symbol) => total + getDeckManaSymbolValue(symbol.slice(1, -1)), 0);
  }

  function getDeckBuilderManaCurveGroups(deck) {
    const groups = new Map();
    const entries = [];
    if (deck?.commander) entries.push({ card: deck.commander, label: `${deck.commander.name} (Commander)` });
    (deck?.cards || []).forEach((card) => entries.push({ card, label: card.name }));
    entries.forEach(({ card, label }) => {
      const manaValue = getDeckCardManaValue(card);
      if (!Number.isInteger(manaValue) || manaValue < 0) return;
      if (!groups.has(manaValue)) groups.set(manaValue, []);
      groups.get(manaValue).push(label);
    });
    groups.forEach((cards) => cards.sort((first, second) => compareTextValues(first, second)));
    return groups;
  }

  async function showDeckBuilderManaValueCards(manaValue) {
    const deck = ensureActiveDeckBuilderRecord();
    const cards = getDeckBuilderManaCurveGroups(deck).get(manaValue) || [];
    if (!cards.length) return;
    await showLiveModal({
      title: `Mana value ${manaValue}`,
      description: `${cards.length} card${cards.length === 1 ? '' : 's'}\n${cards.join('\n')}`,
      confirmLabel: 'Close',
      showCancel: false,
      showInput: false,
    });
  }

  function renderDeckBuilderManaCurve(deck) {
    if (!deckBuilderManaCurve) return;
    const groups = getDeckBuilderManaCurveGroups(deck);
    const manaValues = [...groups.keys()].sort((first, second) => first - second);
    if (!manaValues.length) {
      deckBuilderManaCurve.innerHTML = '<div class="deck-breakdown-section"><h3>Mana Curve</h3><p class="status-muted">Add nonland cards to see the curve.</p></div>';
      return;
    }
    const maxManaValue = manaValues[manaValues.length - 1];
    const maxCount = Math.max(...manaValues.map((manaValue) => groups.get(manaValue)?.length || 0), 1);
    const bars = Array.from({ length: maxManaValue + 1 }, (_, manaValue) => {
      const count = groups.get(manaValue)?.length || 0;
      const heightPercent = count ? Math.max(Math.round((count / maxCount) * 100), 14) : 8;
      return `<button type="button" class="deck-builder-curve-bar${count ? '' : ' is-empty'}" data-mana-value="${manaValue}" aria-label="Mana value ${manaValue}: ${count} card${count === 1 ? '' : 's'}"${count ? '' : ' disabled'}><span class="deck-builder-curve-count">${count}</span><span class="deck-builder-curve-column" style="height:${heightPercent}%"></span><span class="deck-builder-curve-label">${manaValue}</span></button>`;
    }).join('');
    deckBuilderManaCurve.innerHTML = `<div class="deck-breakdown-section"><h3>Mana Curve</h3><div class="deck-builder-mana-curve-chart" role="group" aria-label="Deck mana curve">${bars}</div></div>`;
  }

  function renderDeckBuilderBreakdown(deck) {
    if (!deckBuilderBreakdown) return;
    const typeOrder = ['Creature', 'Artifact', 'Enchantment', 'Planeswalker', 'Battle', 'Instant', 'Sorcery', 'Land', 'Other'];
    const typeCounts = new Map(typeOrder.map((type) => [type, 0]));
    (deck?.cards || []).forEach((card) => {
      const type = typeOrder.includes(card.cardType) ? card.cardType : 'Other';
      typeCounts.set(type, typeCounts.get(type) + 1);
    });
    const totalDeckCards = (deck?.cards || []).length + (deck?.commander ? 1 : 0) + (deck?.secondCommander ? 1 : 0);
    const gameChangerCards = [deck?.commander, deck?.secondCommander, ...(deck?.cards || [])].filter((card) => card?.isGameChanger);
    const commanderRows = [deck?.commander, deck?.secondCommander].filter(Boolean).map((card) => `<tr><td>Commander</td><td class="deck-breakdown-count">1</td></tr>`).join('');
    const typeRows = typeOrder.filter((type) => typeCounts.get(type) > 0)
      .map((type) => `<tr><td>${escapeHtml(type)}</td><td class="deck-breakdown-count">${typeCounts.get(type)}</td></tr>`).join('');
    const typeMarkup = `<div class="deck-breakdown-section"><h3>Card Types</h3><table class="deck-breakdown-table"><tbody>${commanderRows}${typeRows}<tr class="deck-breakdown-total"><td>Total</td><td class="deck-breakdown-count">${totalDeckCards}</td></tr>${gameChangerCards.length ? `<tr><td>&#9889; Game Changers</td><td class="deck-breakdown-count">${gameChangerCards.length}</td></tr>` : ''}</tbody></table></div>`;

    let matchedRows = [];
    if (deck?.commander?.name) {
      const allGames = loadGames();
      const rawCommanders = [];
      allGames.forEach((game) => getGameRows(game).forEach((row) => {
        const commander = String(row.commander || '').trim();
        if (commander) rawCommanders.push(commander);
      }));
      const commanderMap = buildCanonicalIdentityMapFromValues([
        ...rawCommanders,
        deck.commander.name,
        deck.secondCommander?.name || '',
      ]);
      const commanderKeys = [deck.commander, deck.secondCommander]
        .filter(Boolean)
        .map((card) => getIdentityKey(canonicalizeIdentityValue(card.name, commanderMap)));
      allGames.forEach((game) => getGameRows(game).forEach((row) => {
        const canonicalCommander = canonicalizeIdentityValue(String(row.commander || '').trim(), commanderMap);
        if (commanderKeys.includes(getIdentityKey(canonicalCommander))) matchedRows.push(row);
      }));
    }
    let performanceMarkup = '<div class="deck-breakdown-section"><h3>Performance</h3><p class="status-muted">No games registered</p></div>';
    if (matchedRows.length) {
      const games = matchedRows.length;
      const wins = matchedRows.filter((row) => row.place === 1).length;
      const winRate = Math.round((wins / games) * 100);
      const totalKills = matchedRows.reduce((sum, row) => sum + (typeof row.kills === 'number' ? row.kills : 0), 0);
      const avgKills = (totalKills / games).toFixed(1);
      const validPlaces = matchedRows.filter((row) => typeof row.place === 'number' && row.place > 0);
      const avgPlace = validPlaces.length ? (validPlaces.reduce((sum, row) => sum + row.place, 0) / validPlaces.length).toFixed(1) : '—';
      performanceMarkup = `<div class="deck-breakdown-section"><h3>Performance</h3><table class="deck-breakdown-table"><tbody><tr><td>Games played</td><td class="deck-breakdown-count">${games}</td></tr><tr><td>Wins</td><td class="deck-breakdown-count">${wins}</td></tr><tr><td>Win rate</td><td class="deck-breakdown-count">${winRate}%</td></tr><tr><td>Avg kills / game</td><td class="deck-breakdown-count">${avgKills}</td></tr><tr><td>Avg finish position</td><td class="deck-breakdown-count">${avgPlace}</td></tr></tbody></table></div>`;
    }
    deckBuilderBreakdown.innerHTML = `<div class="deck-breakdown-grid">${typeMarkup}<div class="deck-builder-mana-curve-slot"></div>${performanceMarkup}</div>`;
    const curveSlot = deckBuilderBreakdown.querySelector('.deck-builder-mana-curve-slot');
    if (curveSlot && deckBuilderManaCurve) curveSlot.appendChild(deckBuilderManaCurve);
  }

  async function hydrateDeckCardStats(deck) {
    if (!deck?.id || hydratedDeckIds.has(deck.id)) return;
    hydratedDeckIds.add(deck.id);
    const mightHaveStats = (card) => /creature|planeswalker|battle/i.test(card?.typeLine || '');
    const isMissingStats = (card) => !card?.power && !card?.toughness && !card?.loyalty && !card?.defense;
    const toRefresh = [];
    if (deck.commander && mightHaveStats(deck.commander) && isMissingStats(deck.commander)) {
      toRefresh.push({ type: 'commander', card: deck.commander, index: -1 });
    }
    (deck.cards || []).forEach((card, index) => {
      if (mightHaveStats(card) && isMissingStats(card)) toRefresh.push({ type: 'card', card, index });
    });
    if (!toRefresh.length) return;
    let bulkMap;
    try {
      bulkMap = await fetchDeckCardsByNamesBulk(toRefresh.map(({ card }) => card.name));
    } catch (error) {
      return;
    }
    const latestDeck = ensureActiveDeckBuilderRecord();
    if (!latestDeck || latestDeck.id !== deck.id) return;
    const nextCommander = latestDeck.commander ? { ...latestDeck.commander } : null;
    const nextCards = (latestDeck.cards || []).map((card) => ({ ...card }));
    let updated = false;
    toRefresh.forEach(({ type, card, index }) => {
      const fresh = bulkMap.get(getIdentityKey(card.name));
      if (!fresh || (!fresh.power && !fresh.toughness && !fresh.loyalty && !fresh.defense)) return;
      const target = type === 'commander' ? nextCommander : nextCards.find((entry) => entry.id === card.id);
      if (!target) return;
      Object.assign(target, {
        power: fresh.power,
        toughness: fresh.toughness,
        loyalty: fresh.loyalty,
        defense: fresh.defense,
        cardFaces: fresh.cardFaces,
      });
      updated = true;
    });
    if (updated) persistDeckBuilderRecord({ ...latestDeck, commander: nextCommander, cards: nextCards }, 'Card stats refreshed.', 'muted', { skipUndo: true });
  }

  function queueDeckBuilderMetaSave() {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) return;
    const capturedName = String(deckBuilderNameInput?.value || '').trim() || 'Untitled Deck';
    const capturedOwner = normalizeIdentityLabel(deckBuilderOwnerInput?.value || '');
    const capturedPower = normalizeDeckPowerLevel(deckBuilderPowerInput?.value);
    const capturedInRotation = deckBuilderInRotationInput ? deckBuilderInRotationInput.checked : (deck.inRotation !== false);
    const capturedDeck = { ...deck, name: capturedName, owner: capturedOwner, powerLevel: capturedPower, inRotation: capturedInRotation };
    renderDeckBuilderBreakdown(capturedDeck);
    if (deckBuilderSaveTimer) clearTimeout(deckBuilderSaveTimer);
    deckBuilderSaveTimer = setTimeout(() => {
      deckBuilderSaveTimer = null;
      persistDeckBuilderRecord({ ...capturedDeck, ownerUserId: deck.ownerUserId || '' }, 'Unsaved changes', 'muted');
    }, DECK_BUILDER_META_SAVE_DEBOUNCE_MS);
  }

  function populateDeckBuilderLookupMenus() {
    if (deckBuilderOwnerMenu) buildDropdownMenu(deckBuilderOwnerMenu, knownPlayers);
  }

  window.CommanderDeckBuilder = {
    hasView: () => Boolean(deckBuilderPage),
    getActiveDeckId: () => activeDeckBuilderId,
    getSyncState: () => ({ activeDeckId: activeDeckBuilderId, hasUnsavedChanges: deckBuilderHasUnsavedChanges }),
    setCommanderPrefill: (value) => { deckBuilderCommanderPrefill = String(value || ''); },
    render: renderDeckBuilderPage,
    renderSelection: renderDeckBuilderSelection,
    renderCards: renderDeckBuilderCards,
    renderCardRow: renderDeckCardRow,
    renderDeckGroupCardRows,
    renderBreakdown: renderDeckBuilderBreakdown,
    applyDeckBuilderDraftMeta,
    populateDeckBuilderLookupMenus,
    setDeckBuilderSaveStatus,
    flashDeckBuilderSaveConfirmation,
    persistDeckBuilderSelectedCard,
    loadDeckBuilderSelectedCardFromStorage,
    getCurrentDeckBuilderSelectedCard,
    addSelectedCardToDeck,
    addSelectedCardToMaybeboard,
    addSelectedCardToTokens,
    ensureActiveDeckBuilderRecord,
    hasUnsavedDeckBuilderChanges,
  };

  function createEmptyDeckRecord() {
    const timestamp = new Date().toISOString();
    return normalizeDeckRecord({
      id: generateId(),
      name: 'Untitled Deck',
      owner: getCurrentSyncDisplayName(),
      ownerUserId: getCurrentSyncUserId(),
      commander: null,
      secondCommander: null,
      cards: [],
      maybeboard: [],
      tokens: [],
      inRotation: true,
      powerLevel: null,
      createdAt: timestamp,
      updatedAt: timestamp,
    });
  }

  function applyDeckBuilderAccessState(deck) {
    const isReadOnly = Boolean(deck) && !canCurrentUserEditDeck(deck);
    if (deckBuilderAccessBadge) {
      deckBuilderAccessBadge.hidden = !isReadOnly;
      deckBuilderAccessBadge.textContent = isReadOnly ? 'Read-only' : '';
    }
    if (deckBuilderNameInput) deckBuilderNameInput.disabled = isReadOnly;
    if (deckBuilderOwnerInput) deckBuilderOwnerInput.disabled = isReadOnly || !isCurrentSyncUserAdmin();
    if (deckBuilderPowerInput) deckBuilderPowerInput.disabled = isReadOnly;
    if (deckBuilderInRotationInput) deckBuilderInRotationInput.disabled = isReadOnly;
    if (deckBuilderSearchInput) deckBuilderSearchInput.disabled = isReadOnly;
    if (deckBuilderTokenSearchInput) deckBuilderTokenSearchInput.disabled = isReadOnly;
    if (deckBuilderImportButton) deckBuilderImportButton.disabled = isReadOnly;
    if (deckBuilderTextList) deckBuilderTextList.disabled = isReadOnly;
    if (deckBuilderUndoButton) {
      deckBuilderUndoButton.disabled = isReadOnly || ((deckBuilderUndoStacks[deck?.id || ''] || []).length === 0);
    }
    if (deckBuilderDiscardButton) {
      deckBuilderDiscardButton.hidden = isReadOnly || !deck?.id || isDeckUsedInGame(deck);
    }
    updateDeckBuilderSaveButton();
  }

  function setDeckBuilderSaveStatus(message, tone = 'muted') {
    if (!deckBuilderSaveStatus) return;
    deckBuilderSaveStatus.textContent = message;
    deckBuilderSaveStatus.classList.remove('status-muted', 'status-success', 'status-error', 'status-neutral');
    deckBuilderSaveStatus.classList.add(`status-${tone}`);
  }

  function flashDeckBuilderSaveConfirmation(message) {
    setDeckBuilderSaveStatus(message, 'success');
    if (deckBuilderSaveConfirmTimer) {
      clearTimeout(deckBuilderSaveConfirmTimer);
      deckBuilderSaveConfirmTimer = null;
    }
    if (deckBuilderSaveButton) {
      deckBuilderSaveButton.classList.remove('save-confirm-flash');
      void deckBuilderSaveButton.offsetWidth;
      deckBuilderSaveButton.classList.add('save-confirm-flash');
    }
    deckBuilderSaveConfirmTimer = setTimeout(() => {
      deckBuilderSaveConfirmTimer = null;
      if (deckBuilderSaveButton) deckBuilderSaveButton.classList.remove('save-confirm-flash');
      if (!deckBuilderHasUnsavedChanges) setDeckBuilderSaveStatus('Changes saved locally.', 'muted');
    }, 2600);
  }

  function persistDeckBuilderSelectedCard(card) {
    const normalizedCard = normalizeDeckCardEntry(card);
    if (!normalizedCard) {
      removeLocalStorageValue(DECK_BUILDER_SELECTED_CARD_STORAGE_KEY);
      return;
    }
    writeLocalStorageValue(DECK_BUILDER_SELECTED_CARD_STORAGE_KEY, JSON.stringify(normalizedCard));
  }

  function loadDeckBuilderSelectedCardFromStorage() {
    const rawValue = readLocalStorageValue(DECK_BUILDER_SELECTED_CARD_STORAGE_KEY) || '';
    if (!rawValue) return null;
    try {
      return normalizeDeckCardEntry(JSON.parse(rawValue));
    } catch (error) {
      return null;
    }
  }

  function getCurrentDeckBuilderSelectedCard() {
    const normalizedSelectedCard = normalizeDeckCardEntry(deckBuilderSelectedCard);
    if (normalizedSelectedCard) return normalizedSelectedCard;
    const serializedCard = deckBuilderSelection?.dataset.selectedCard || '';
    if (!serializedCard) return null;
    try {
      return normalizeDeckCardEntry(JSON.parse(serializedCard));
    } catch (error) {
      return loadDeckBuilderSelectedCardFromStorage();
    }
  }

  function applyDeckBuilderDraftMeta(deck) {
    if (!deck || !deckBuilderPage) return deck;
    const nextName = String(deckBuilderNameInput?.value || deck.name || '').trim() || 'Untitled Deck';
    const nextOwner = normalizeIdentityLabel(deckBuilderOwnerInput?.value || deck.owner || '');
    const powerInputValue = String(deckBuilderPowerInput?.value || '').trim();
    const nextPowerLevel = powerInputValue ? normalizeDeckPowerLevel(powerInputValue) : normalizeDeckPowerLevel(deck.powerLevel);
    const nextInRotation = deckBuilderInRotationInput ? deckBuilderInRotationInput.checked : (deck.inRotation !== false);
    return { ...deck, name: nextName, owner: nextOwner, powerLevel: nextPowerLevel, inRotation: nextInRotation };
  }

  function deckCardShowPartnerCommanderButton(deck, card) {
    if (!isCommanderEligibleCard(card)) return false;
    if (!deck?.commander) return false;
    if (!cardHasPairingMechanic(deck.commander)) return false;
    if (!cardHasPairingMechanic(card)) return false;
    return true;
  }

  function cardHasGenericPartner(card) {
    return getCardOracleTexts(card).some((text) => /\bpartner\b(?!\s+with\b)/i.test(text));
  }

  function getCardPartnerWithName(card) {
    for (const text of getCardOracleTexts(card)) {
      const match = text.match(/\bpartner with ([^\n(]+)/i);
      if (match) return match[1].trim();
    }
    return null;
  }

  function cardHasPartnerWith(card) {
    return Boolean(getCardPartnerWithName(card));
  }

  function cardHasFriendsForever(card) {
    return getCardOracleTexts(card).some((text) => /\bfriends forever\b/i.test(text));
  }

  function cardHasChooseABackground(card) {
    return getCardOracleTexts(card).some((text) => /\bchoose a background\b/i.test(text));
  }

  function cardIsBackground(card) {
    const texts = [
      String(card?.typeLine || ''),
      ...(Array.isArray(card?.cardFaces) ? card.cardFaces.map((face) => String(face?.typeLine || '')) : []),
    ];
    return texts.some((text) => /\bbackground\b/i.test(text));
  }

  function cardHasDoctorsCompanion(card) {
    return getCardOracleTexts(card).some((text) => /\bdoctor's companion\b/i.test(text));
  }

  function cardIsTheDoctor(card) {
    return /^the doctor$/i.test(String(card?.name || '').trim());
  }

  function getCardPairingMechanicLabel(card) {
    if (cardHasGenericPartner(card)) return 'Partner';
    if (cardHasPartnerWith(card)) return `Partner with ${getCardPartnerWithName(card)}`;
    if (cardHasFriendsForever(card)) return 'Friends forever';
    if (cardHasChooseABackground(card)) return 'Choose a Background';
    if (cardIsBackground(card)) return 'Background';
    if (cardHasDoctorsCompanion(card)) return "Doctor's companion";
    if (cardIsTheDoctor(card)) return 'The Doctor';
    return null;
  }

  function isCompatiblePartnerPair(card1, card2) {
    if (cardHasGenericPartner(card1) && cardHasGenericPartner(card2)) return true;
    const target1 = getCardPartnerWithName(card1);
    const target2 = getCardPartnerWithName(card2);
    if (target1 && target2 && getIdentityKey(target1) === getIdentityKey(card2.name) && getIdentityKey(target2) === getIdentityKey(card1.name)) {
      return true;
    }
    if (cardHasFriendsForever(card1) && cardHasFriendsForever(card2)) return true;
    if (cardHasChooseABackground(card1) && cardIsBackground(card2)) return true;
    if (cardHasChooseABackground(card2) && cardIsBackground(card1)) return true;
    if (cardHasDoctorsCompanion(card1) && cardIsTheDoctor(card2)) return true;
    if (cardHasDoctorsCompanion(card2) && cardIsTheDoctor(card1)) return true;
    return false;
  }

  function cardHasPairingMechanic(card) {
    return cardHasGenericPartner(card)
      || cardHasPartnerWith(card)
      || cardHasFriendsForever(card)
      || cardHasChooseABackground(card)
      || cardIsBackground(card)
      || cardHasDoctorsCompanion(card)
      || cardIsTheDoctor(card);
  }

  function isCommanderEligibleCard(card) {
    if (!card || card.isBanned || !card.isCommanderLegal) return false;
    const typeText = [
      String(card.typeLine || ''),
      ...(Array.isArray(card.cardFaces) ? card.cardFaces.map((face) => String(face?.typeLine || '')) : []),
    ].join(' ').toLowerCase();
    const rulesText = [
      String(card.oracleText || ''),
      ...(Array.isArray(card.cardFaces) ? card.cardFaces.map((face) => String(face?.oracleText || '')) : []),
    ].join(' ').toLowerCase();
    return (/\blegendary\b/.test(typeText) && /\bcreature\b/.test(typeText)) || /can be your commander/.test(rulesText);
  }

  function canSetCardAsCommanderForDeck(deck, card) {
    if (!isCommanderEligibleCard(card)) return false;
    if (!deck?.commander || deck.secondCommander) return true;
    return true;
  }

  function getDeckBuilderCardNameSet(deck) {
    const nameSet = new Set();
    if (deck?.commander?.name) nameSet.add(getIdentityKey(deck.commander.name));
    if (deck?.secondCommander?.name) nameSet.add(getIdentityKey(deck.secondCommander.name));
    (deck?.cards || []).forEach((card) => {
      if (card?.name) nameSet.add(getIdentityKey(card.name));
    });
    return nameSet;
  }

  async function resolveCommanderLegalCard(card) {
    if (!card || !card.name) return null;
    if (isCommanderEligibleCard(card)) return card;
    try {
      const refreshed = await fetchDeckCardByName(card.name);
      if (isCommanderEligibleCard(refreshed)) return refreshed;
    } catch (error) {
    }
    return null;
  }

  function warmDeckBuilderBasicLandCache() {
    const missingNames = BASIC_LAND_NAMES.filter((name) => !deckBuilderCardCache.has(name.toLowerCase()));
    if (!missingNames.length) {
      return Promise.resolve(new Map());
    }

    if (deckBuilderBasicLandWarmPromise) {
      return deckBuilderBasicLandWarmPromise;
    }

    deckBuilderBasicLandWarmPromise = fetchDeckCardsByNamesBulk(missingNames)
      .catch(() => new Map())
      .finally(() => {
        deckBuilderBasicLandWarmPromise = null;
      });

    return deckBuilderBasicLandWarmPromise;
  }

  function queueDeckBuilderMutation(task) {
    const runTask = async () => task();
    const queuedTask = deckBuilderMutationQueue.then(runTask, runTask);
    deckBuilderMutationQueue = queuedTask.catch(() => {});
    return queuedTask;
  }

  function stopDeckBuilderHoldRepeat() {
    if (deckBuilderHoldTimerId) {
      clearTimeout(deckBuilderHoldTimerId);
      deckBuilderHoldTimerId = null;
    }
    if (deckBuilderHoldIntervalId) {
      clearInterval(deckBuilderHoldIntervalId);
      deckBuilderHoldIntervalId = null;
    }
  }

  function getDeckBuilderBasicLandHoldAction(target) {
    const button = target?.closest?.('[data-add-basic], [data-remove-basic]');
    if (!button || button.disabled) {
      return null;
    }

    const addBasicName = String(button.dataset.addBasic || '').trim();
    if (addBasicName) {
      return () => {
        void addBasicLandToDeck(addBasicName);
      };
    }

    const removeBasicName = String(button.dataset.removeBasic || '').trim();
    if (removeBasicName) {
      return () => {
        removeBasicLandFromDeck(removeBasicName);
      };
    }

    return null;
  }

  function startDeckBuilderHoldRepeat(target, { applyInitialChange = false } = {}) {
    stopDeckBuilderHoldRepeat();

    const action = getDeckBuilderBasicLandHoldAction(target);
    if (!action) {
      return;
    }

    if (applyInitialChange) {
      action();
    }

    deckBuilderHoldTimerId = setTimeout(() => {
      deckBuilderHoldIntervalId = setInterval(() => {
        action();
      }, DECK_BUILDER_HOLD_REPEAT_INTERVAL_MS);
    }, DECK_BUILDER_HOLD_REPEAT_START_DELAY_MS);
  }

  async function addUnlimitedCopyCardToDeck(cardName) {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) return;

    let card = deckBuilderCardCache.get(cardName.toLowerCase());
    if (!card) {
      const existing = (deck.cards || []).find((entry) => getIdentityKey(entry.name) === getIdentityKey(cardName));
      card = existing || null;
    }
    if (!card) {
      try {
        card = await fetchDeckCardByName(cardName);
      } catch (error) {
        setDeckBuilderSaveStatus(`Could not find ${cardName}.`, 'error');
        return;
      }
    }
    if (!card) {
      setDeckBuilderSaveStatus(`Could not find ${cardName}.`, 'error');
      return;
    }

    const latestDeck = ensureActiveDeckBuilderRecord();
    if (!latestDeck) return;

    persistDeckBuilderRecord({ ...latestDeck, cards: [...latestDeck.cards, { ...card, id: generateId() }] }, `${cardName} added.`, 'success', {
      skipFullRefresh: true,
      persistDelayMs: DECKS_RAPID_ACTION_PERSIST_DEBOUNCE_MS,
    });
  }

  async function addBasicLandToDeck(landName) {
    return queueDeckBuilderMutation(async () => {
      const deck = ensureActiveDeckBuilderRecord({ createIfMissing: true });
      if (!deck) return;
      let card = deckBuilderCardCache.get(landName.toLowerCase());
      if (!card) {
        await warmDeckBuilderBasicLandCache();
        card = deckBuilderCardCache.get(landName.toLowerCase());
      }
      if (!card) {
        try {
          card = await fetchDeckCardByName(landName);
        } catch (error) {
          setDeckBuilderSaveStatus(`Could not find ${landName}.`, 'error');
          return;
        }
      }
      if (!card) {
        setDeckBuilderSaveStatus(`Could not find ${landName}.`, 'error');
        return;
      }

      const latestDeck = ensureActiveDeckBuilderRecord({ createIfMissing: true });
      if (!latestDeck) return;

      persistDeckBuilderRecord({ ...latestDeck, cards: [...latestDeck.cards, { ...card, id: generateId() }] }, `${landName} added.`, 'success', {
        skipFullRefresh: true,
        persistDelayMs: DECKS_RAPID_ACTION_PERSIST_DEBOUNCE_MS,
      });
    });
  }

  function removeBasicLandFromDeck(landName) {
    return queueDeckBuilderMutation(() => {
      const deck = ensureActiveDeckBuilderRecord();
      if (!deck) return;

      const match = (deck.cards || []).find((card) => getIdentityKey(card?.name) === getIdentityKey(landName));
      if (!match?.id) return;

      removeDeckBuilderCard(match.id);
    });
  }

  async function fetchDeckCardArtOptions(card) {
    const oracleId = String(card?.oracleId || '').trim();
    const name = String(card?.name || '').trim();
    if (!oracleId && !name) return [];

    const query = new URLSearchParams();
    if (oracleId) query.set('oracleId', oracleId);
    if (name) query.set('name', name);
    const response = await fetch(`${DECK_CARD_ARTS_ENDPOINT}?${query.toString()}`, { cache: 'no-store' });
    if (!response.ok) {
      let message = `Request failed (${response.status})`;
      try {
        const payload = await response.json();
        if (payload?.detail) message = `${payload.error || message} ${payload.detail}`.trim();
        else if (payload?.error) message = payload.error;
      } catch (error) {
      }
      throw new Error(message);
    }
    const payload = await response.json();
    return Array.isArray(payload?.prints) ? payload.prints : [];
  }

  function getDeckBuilderArtCacheKey(card) {
    const oracleId = String(card?.oracleId || '').trim().toLowerCase();
    if (oracleId) return `oracle:${oracleId}`;
    const nameKey = getIdentityKey(card?.name || '');
    return nameKey ? `name:${nameKey}` : '';
  }

  function findDeckBuilderCardById(deck, cardId) {
    if (!deck || !cardId) return null;
    if (deck.commander?.id === cardId) return { zone: 'commander', card: deck.commander };

    const mainCard = (deck.cards || []).find((card) => card.id === cardId);
    if (mainCard) return { zone: 'cards', card: mainCard };
    const tokenCard = (deck.tokens || []).find((card) => card.id === cardId);
    if (tokenCard) return { zone: 'tokens', card: tokenCard };
    const maybeCard = (deck.maybeboard || []).find((card) => card.id === cardId);
    if (maybeCard) return { zone: 'maybeboard', card: maybeCard };
    return null;
  }

  async function openDeckBuilderArtPicker(cardId) {
    const deck = ensureActiveDeckBuilderRecord();
    const located = findDeckBuilderCardById(deck, cardId);
    if (!deck || !located?.card) return;

    const card = located.card;
    const cacheKey = getDeckBuilderArtCacheKey(card);
    deckBuilderArtPickerCardId = cardId;
    if (cacheKey && deckBuilderArtOptionsCache.has(cacheKey)) {
      deckBuilderArtPickerState = {
        status: 'ready', cardId, options: deckBuilderArtOptionsCache.get(cacheKey), message: '',
      };
      renderDeckBuilderCards(deck);
      return;
    }

    deckBuilderArtPickerState = { status: 'loading', cardId, options: [], message: 'Loading art options...' };
    renderDeckBuilderCards(deck);
    try {
      const options = await fetchDeckCardArtOptions(card);
      if (cacheKey) deckBuilderArtOptionsCache.set(cacheKey, options);
      deckBuilderArtPickerState = {
        status: 'ready',
        cardId,
        options,
        message: options.length ? '' : 'No alternate arts found for this card.',
      };
    } catch (error) {
      deckBuilderArtPickerState = {
        status: 'error',
        cardId,
        options: [],
        message: error instanceof Error ? error.message : 'Unable to load art options right now.',
      };
    }
    renderDeckBuilderCards(deck);
  }

  function applyDeckBuilderCardArt(cardId, print) {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck || !print || !cardId) return;

    const located = findDeckBuilderCardById(deck, cardId);
    const targetNameKey = getIdentityKey(located?.card?.name || '');
    const shouldApplyToSameNameCopies = located?.zone === 'cards' && Boolean(targetNameKey) && allowsMultipleCopies(located.card);
    const applyPrint = (card) => ({
      ...card,
      imageUri: String(print.imageUri || card.imageUri || '').trim(),
      imageLargeUri: String(print.imageLargeUri || card.imageLargeUri || '').trim(),
      cardFaces: Array.isArray(print.cardFaces) && print.cardFaces.length ? print.cardFaces : (card.cardFaces || []),
      scryfallUri: String(print.scryfallUri || card.scryfallUri || '').trim(),
    });

    let changed = false;
    const nextCommander = deck.commander?.id === cardId ? applyPrint(deck.commander) : deck.commander;
    changed = changed || Boolean(deck.commander?.id === cardId);
    const updateList = (list) => (list || []).map((card) => {
      const matchesById = card.id === cardId;
      const matchesSameNameCopy = shouldApplyToSameNameCopies && getIdentityKey(card.name) === targetNameKey;
      if (!matchesById && !matchesSameNameCopy) return card;
      changed = true;
      return applyPrint(card);
    });

    const nextDeck = {
      ...deck,
      commander: nextCommander,
      cards: updateList(deck.cards),
      tokens: updateList(deck.tokens),
      maybeboard: updateList(deck.maybeboard),
    };
    if (changed) persistDeckBuilderRecord(nextDeck, 'Card art updated.');
  }

  async function selectDeckCardByName(cardName, deckCardId = '') {
    try {
      const card = await fetchDeckCardByName(cardName);
      if (!card) return;
      deckBuilderSelectedCard = card;
      deckBuilderSelectedFaceIndex = 0;
      if (deckCardId && card.cardFaces?.length > 1) applyDeckBuilderCardArt(deckCardId, card);
      persistDeckBuilderSelectedCard(card);
      if (deckBuilderSearchResults) {
        deckBuilderSearchResults.innerHTML = '';
        deckBuilderSearchResults.hidden = true;
      }
      renderDeckBuilderSelection();
    } catch (error) {
      console.warn('Failed to select deck card:', error);
    }
  }

  function getDeckImportSectionType(value) {
    const normalized = String(value || '').toLowerCase();
    if (!normalized) {
      return '';
    }

    if (normalized.includes('command zone') || normalized.includes('commander')) {
      return 'commander';
    }

    if (normalized.includes('maybeboard') || normalized.includes('sideboard') || normalized.includes('considering') || normalized.includes('wishlist')) {
      return 'maybeboard';
    }

    if (normalized.includes('token')) {
      return 'tokens';
    }

    if (normalized.includes('mainboard') || normalized.includes('main board') || normalized.includes('deck')) {
      return 'deck';
    }

    return '';
  }

  function getDeckImportLineEntry(value) {
    const line = String(value || '').trim().replace(/^[-*]\s+/, '');
    if (!line) {
      return null;
    }

    let section = '';
    let working = line;

    const sideboardPrefix = working.match(/^(sb|sideboard)\s*:\s*(.+)$/i);
    if (sideboardPrefix) {
      section = 'maybeboard';
      working = sideboardPrefix[2].trim();
    }

    const commanderPrefix = working.match(/^(cmdr|commander|command\s*zone)\s*:\s*(.+)$/i);
    if (commanderPrefix) {
      section = 'commander';
      working = commanderPrefix[2].trim();
    }

    let count = 1;
    let name = '';

    const csvLeadingCount = working.match(/^(\d+)\s*,\s*(.+)$/);
    const leadingCount = working.match(/^(\d+)\s*x?\s+(.+)$/i);
    const compactLeadingCount = working.match(/^(\d+)x\s*(.+)$/i);
    const trailingCount = working.match(/^(.+?)\s+x\s*(\d+)$/i);

    if (csvLeadingCount) {
      count = Math.max(1, parseInt(csvLeadingCount[1], 10));
      name = csvLeadingCount[2].trim();
    } else if (leadingCount) {
      count = Math.max(1, parseInt(leadingCount[1], 10));
      name = leadingCount[2].trim();
    } else if (compactLeadingCount) {
      count = Math.max(1, parseInt(compactLeadingCount[1], 10));
      name = compactLeadingCount[2].trim();
    } else if (trailingCount) {
      count = Math.max(1, parseInt(trailingCount[2], 10));
      name = trailingCount[1].trim();
    } else {
      name = working.trim();
    }

    if (!name) {
      return null;
    }

    name = name
      .replace(/\s*\[[^\]]*\]\s*/g, ' ')
      .replace(/\s*\*[^*]+\*\s*/g, ' ')
      .replace(/\s*\([A-Za-z0-9]{2,}\)\s*\d+[A-Za-z]*$/i, '')
      .replace(/\s+\d+[A-Za-z]*$/i, '')
      .replace(/\s+#.*$/, '')
      .replace(/^"|"$/g, '')
      .replace(/\s+/g, ' ')
      .trim();

    if (!name) {
      return null;
    }

    if (/\s*\(Commander\)\s*$/i.test(name)) {
      section = 'commander';
      name = name.replace(/\s*\(Commander\)\s*$/i, '').trim();
    }

    return { name, count, section };
  }

  function parseDeckTextList(text) {
    let commanderName = null;
    let currentSection = 'deck';
    let hasMaybeboardSection = false;
    let hasTokenSection = false;
    const deckCounts = new Map();
    const deckNames = new Map();
    const maybeboardCounts = new Map();
    const maybeboardNames = new Map();
    const tokenCounts = new Map();
    const tokenNames = new Map();

    const addEntry = (countsMap, namesMap, name, count) => {
      const key = getIdentityKey(name);
      if (!key) {
        return;
      }
      countsMap.set(key, (countsMap.get(key) || 0) + Math.max(1, count));
      if (!namesMap.has(key)) {
        namesMap.set(key, name);
      }
    };

    String(text || '').split(/\r?\n/).forEach((line) => {
      const trimmed = String(line || '').trim();
      if (!trimmed) {
        return;
      }

      const headerLine = trimmed.replace(/^\/\/\s*|^#\s*/, '').trim();
      const explicitHeader = getDeckImportSectionType(headerLine.replace(/:$/, ''));
      if (explicitHeader && /^([a-z\s]+:?|\[[^\]]+\])$/i.test(headerLine)) {
        currentSection = explicitHeader;
        if (explicitHeader === 'maybeboard') {
          hasMaybeboardSection = true;
        }
        if (explicitHeader === 'tokens') {
          hasTokenSection = true;
        }
        return;
      }

      if (trimmed.startsWith('//') || trimmed.startsWith('#')) {
        const sectionFromComment = getDeckImportSectionType(headerLine);
        if (sectionFromComment) {
          currentSection = sectionFromComment;
          if (sectionFromComment === 'maybeboard') {
            hasMaybeboardSection = true;
          }
          if (sectionFromComment === 'tokens') {
            hasTokenSection = true;
          }
        }
        return;
      }

      const parsed = getDeckImportLineEntry(trimmed);
      if (!parsed) {
        return;
      }

      const entrySection = parsed.section || currentSection;
      if (entrySection === 'commander' && !commanderName) {
        commanderName = parsed.name;
        return;
      }

      if (entrySection === 'maybeboard') {
        hasMaybeboardSection = true;
        addEntry(maybeboardCounts, maybeboardNames, parsed.name, parsed.count);
        return;
      }

      if (entrySection === 'tokens') {
        hasTokenSection = true;
        addEntry(tokenCounts, tokenNames, parsed.name, parsed.count);
        return;
      }

      addEntry(deckCounts, deckNames, parsed.name, parsed.count);
    });

    const toEntries = (countsMap, namesMap) => [...countsMap.entries()].map(([key, count]) => ({
      name: namesMap.get(key) || key,
      count,
    }));

    return {
      commanderName,
      entries: toEntries(deckCounts, deckNames),
      maybeboardEntries: toEntries(maybeboardCounts, maybeboardNames),
      tokenEntries: toEntries(tokenCounts, tokenNames),
      hasMaybeboardSection,
      hasTokenSection,
    };
  }

  function getDeckImportNameVariants(name) {
    const raw = String(name || '').trim();
    if (!raw) {
      return [];
    }

    const variants = [];
    const pushVariant = (value) => {
      const normalized = String(value || '').replace(/\s+/g, ' ').trim();
      if (!normalized) {
        return;
      }
      const key = normalized.toLowerCase();
      if (!variants.some((entry) => entry.toLowerCase() === key)) {
        variants.push(normalized);
      }
    };

    pushVariant(raw);
    pushVariant(raw.replace(/[’‘`´]/g, "'"));
    pushVariant(raw.split(/\s*\/\/\s*/)[0]);
    pushVariant(raw.replace(/\s+\([A-Za-z0-9]+\)\s+\d+[A-Za-z]*$/i, ''));
    pushVariant(raw.replace(/\s+\d+[A-Za-z]*$/i, ''));

    return variants;
  }

  function getPrimaryDeckImportLookupName(name) {
    const variants = getDeckImportNameVariants(name);
    return variants.find((value) => !value.includes('//')) || variants[0] || String(name || '').trim();
  }

  function getBulkResolvedImportCard(entryName, bulkCardsByQuery) {
    const variants = getDeckImportNameVariants(entryName);
    for (let i = 0; i < variants.length; i += 1) {
      const key = getIdentityKey(variants[i]);
      if (!key) {
        continue;
      }

      const card = bulkCardsByQuery.get(key);
      if (card) {
        return card;
      }
    }

    return null;
  }

  async function fetchDeckCardForImport(name) {
    const candidates = getDeckImportNameVariants(name);
    const wait = (ms) => new Promise((resolve) => window.setTimeout(resolve, ms));

    for (let i = 0; i < candidates.length; i += 1) {
      const candidate = candidates[i];
      for (let attempt = 0; attempt < 2; attempt += 1) {
        try {
          const card = await fetchDeckCardByName(candidate);
          if (card) {
            return card;
          }
          break;
        } catch (error) {
          const message = String(error instanceof Error ? error.message : error || '').toLowerCase();
          const mayBeTransient = message.includes('rate-limit') || message.includes('temporarily') || message.includes('request failed (5');
          if (!mayBeTransient || attempt === 1) {
            break;
          }

          await wait(350);
        }
      }
    }

    return null;
  }

  async function fetchPreconList() {
    const res = await fetch('https://mtgjson.com/api/v5/DeckList.json');
    if (!res.ok) throw new Error(`Failed to load precon list (${res.status})`);
    const json = await res.json();
    return (json.data || [])
      .filter((deck) => deck.type === 'Commander Deck')
      .sort((first, second) => second.releaseDate.localeCompare(first.releaseDate));
  }

  async function loadPreconDeckText(fileName) {
    const res = await fetch(`https://mtgjson.com/api/v5/decks/${encodeURIComponent(fileName)}.json`);
    if (!res.ok) throw new Error(`Failed to load deck (${res.status})`);
    const json = await res.json();
    const data = json.data || {};
    const lines = [];
    const commanders = data.commander || [];
    if (commanders.length) {
      lines.push('Commander:');
      commanders.forEach((card) => lines.push(`${card.count || 1} ${card.name}`));
      lines.push('');
    }
    lines.push('Deck:');
    (data.mainBoard || []).forEach((card) => lines.push(`${card.count || 1} ${card.name}`));
    return lines.join('\n');
  }

  function setDeckBuilderSearchStatus(message, tone = 'muted') {
    if (!deckBuilderSearchStatus) {
      return;
    }

    deckBuilderSearchStatus.textContent = message;
    deckBuilderSearchStatus.classList.remove('status-muted', 'status-success', 'status-error', 'status-neutral');
    deckBuilderSearchStatus.classList.add(`status-${tone}`);
  }

  function renderDeckBuilderSearchResults() {
    if (!deckBuilderSearchResults) {
      return;
    }

    if (!deckBuilderSearchResultsState.length) {
      deckBuilderSearchResults.innerHTML = '';
      deckBuilderSearchResults.hidden = true;
      return;
    }

    deckBuilderSearchResults.innerHTML = deckBuilderSearchResultsState
      .map((name) => `<button type="button" class="deck-search-result" data-name="${escapeHtml(name)}">${escapeHtml(name)}</button>`)
      .join('');
    deckBuilderSearchResults.hidden = false;
  }

  async function runDeckBuilderSearch(query) {
    const normalizedQuery = String(query || '').trim();
    const requestId = deckBuilderSearchRequestId + 1;
    deckBuilderSearchRequestId = requestId;

    if (normalizedQuery.length < 3) {
      deckBuilderSearchLoading = false;
      deckBuilderSearchResultsState = [];
      renderDeckBuilderSearchResults();
      setDeckBuilderSearchStatus('Type at least 3 characters to search for a card.', 'muted');
      return;
    }

    deckBuilderSearchLoading = true;
    setDeckBuilderSearchStatus(`Searching for ${normalizedQuery}...`, 'neutral');

    try {
      const results = await fetchDeckSearchResultsList(normalizedQuery);
      if (requestId !== deckBuilderSearchRequestId) {
        return;
      }

      deckBuilderSearchLoading = false;
      deckBuilderSearchResultsState = results;
      renderDeckBuilderSearchResults();
      setDeckBuilderSearchStatus(results.length ? `Found ${results.length} matching cards.` : 'No cards matched that search.', results.length ? 'success' : 'muted');
    } catch (error) {
      if (requestId !== deckBuilderSearchRequestId) {
        return;
      }

      deckBuilderSearchLoading = false;
      deckBuilderSearchResultsState = [];
      renderDeckBuilderSearchResults();
      setDeckBuilderSearchStatus(error instanceof Error ? error.message : 'Unable to search cards right now.', 'error');
    }
  }

  function queueDeckBuilderSearch(query) {
    if (deckBuilderSearchTimer) {
      clearTimeout(deckBuilderSearchTimer);
    }

    deckBuilderSearchTimer = setTimeout(() => {
      deckBuilderSearchTimer = null;
      runDeckBuilderSearch(query);
    }, 320);
  }

  async function selectDeckBuilderSearchResult(name) {
    const normalizedName = String(name || '').trim();
    if (!normalizedName) {
      return;
    }

    setDeckBuilderSearchStatus(`Loading ${normalizedName}...`, 'neutral');
    deckBuilderSelectedCard = null;
    persistDeckBuilderSelectedCard(null);
    try {
      const card = await fetchDeckCardByName(normalizedName);
      if (!card) {
        setDeckBuilderSearchStatus(`${normalizedName} was not found.`, 'error');
        renderDeckBuilderSelection();
        return;
      }
      if (card.isToken) {
        setDeckBuilderSearchStatus('That is a token card. Use Token Search to add it to the token pool.', 'error');
        return;
      }
      deckBuilderSelectedFaceIndex = 0;
      deckBuilderSelectedCard = card;
      persistDeckBuilderSelectedCard(deckBuilderSelectedCard);
      renderDeckBuilderSelection();
      setDeckBuilderSearchStatus(`${normalizedName} loaded. Choose Add to Deck or Add to Maybeboard.`, 'success');
    } catch (error) {
      setDeckBuilderSearchStatus(error instanceof Error ? error.message : 'Unable to load that card right now.', 'error');
    }
  }

  async function selectAndAddDeckSearchResult(name) {
    const normalizedName = String(name || '').trim();
    if (!normalizedName) {
      return;
    }

    setDeckBuilderSearchStatus(`Adding ${normalizedName}...`, 'neutral');
    deckBuilderSelectedCard = null;
    persistDeckBuilderSelectedCard(null);
    try {
      const card = await fetchDeckCardByName(normalizedName);
      if (!card) {
        setDeckBuilderSearchStatus(`${normalizedName} was not found.`, 'error');
        renderDeckBuilderSelection();
        return;
      }
      if (card.isToken) {
        setDeckBuilderSearchStatus('That is a token card. Use Token Search to add it to the token pool.', 'error');
        return;
      }
      deckBuilderSelectedCard = card;
      persistDeckBuilderSelectedCard(deckBuilderSelectedCard);
      renderDeckBuilderSelection();
      await addSelectedCardToDeck();
      if (deckBuilderSearchInput) {
        deckBuilderSearchInput.value = '';
      }
      deckBuilderSearchResultsState = [];
      renderDeckBuilderSearchResults();
      setDeckBuilderSearchStatus('Search for a card to preview it before adding it to the deck.', 'muted');
    } catch (error) {
      setDeckBuilderSearchStatus(error instanceof Error ? error.message : 'Unable to load that card right now.', 'error');
    }
  }

  function setDeckBuilderTokenSearchStatus(message, tone = 'muted') {
    if (!deckBuilderTokenSearchStatus) {
      return;
    }

    deckBuilderTokenSearchStatus.textContent = message;
    deckBuilderTokenSearchStatus.classList.remove('status-muted', 'status-success', 'status-error', 'status-neutral');
    deckBuilderTokenSearchStatus.classList.add(`status-${tone}`);
  }

  function renderDeckBuilderTokenSearchResults() {
    if (!deckBuilderTokenSearchResults) {
      return;
    }

    if (!deckBuilderTokenSearchResultsState.length) {
      deckBuilderTokenSearchResults.innerHTML = '';
      deckBuilderTokenSearchResults.hidden = true;
      return;
    }

    deckBuilderTokenSearchResults.innerHTML = deckBuilderTokenSearchResultsState
      .map((name) => `<button type="button" class="deck-search-result" data-token-name="${escapeHtml(name)}">${escapeHtml(name)}</button>`)
      .join('');
    deckBuilderTokenSearchResults.hidden = false;
  }

  async function fetchTokenSearchResultsList(query) {
    const normalizedQuery = String(query || '').trim().toLowerCase();
    if (!normalizedQuery) {
      return [];
    }

    const response = await fetch(`${TOKEN_SEARCH_ENDPOINT}?q=${encodeURIComponent(query)}`, {
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
      }
      throw new Error(message);
    }

    const payload = await response.json();
    return Array.isArray(payload?.results) ? payload.results.map((value) => String(value || '').trim()).filter(Boolean) : [];
  }

  async function runDeckBuilderTokenSearch(query) {
    const normalizedQuery = String(query || '').trim();
    const requestId = deckBuilderTokenSearchRequestId + 1;
    deckBuilderTokenSearchRequestId = requestId;

    if (normalizedQuery.length < 3) {
      deckBuilderTokenSearchLoading = false;
      deckBuilderTokenSearchResultsState = [];
      renderDeckBuilderTokenSearchResults();
      setDeckBuilderTokenSearchStatus('Type at least 3 characters to search tokens.', 'muted');
      return;
    }

    deckBuilderTokenSearchLoading = true;
    setDeckBuilderTokenSearchStatus(`Searching token cards for ${normalizedQuery}...`, 'neutral');

    try {
      const results = await fetchTokenSearchResultsList(normalizedQuery);
      if (requestId !== deckBuilderTokenSearchRequestId) {
        return;
      }

      deckBuilderTokenSearchLoading = false;
      deckBuilderTokenSearchResultsState = results;
      renderDeckBuilderTokenSearchResults();
      setDeckBuilderTokenSearchStatus(results.length ? `Found ${results.length} matching token cards.` : 'No token cards matched that search.', results.length ? 'success' : 'muted');
    } catch (error) {
      if (requestId !== deckBuilderTokenSearchRequestId) {
        return;
      }

      deckBuilderTokenSearchLoading = false;
      deckBuilderTokenSearchResultsState = [];
      renderDeckBuilderTokenSearchResults();
      setDeckBuilderTokenSearchStatus(error instanceof Error ? error.message : 'Unable to search token cards right now.', 'error');
    }
  }

  function queueDeckBuilderTokenSearch(query) {
    if (deckBuilderTokenSearchTimer) {
      clearTimeout(deckBuilderTokenSearchTimer);
    }

    deckBuilderTokenSearchTimer = setTimeout(() => {
      deckBuilderTokenSearchTimer = null;
      runDeckBuilderTokenSearch(query);
    }, 320);
  }

  async function selectAndAddDeckTokenSearchResult(name) {
    const normalizedName = String(name || '').trim();
    if (!normalizedName) {
      return;
    }

    setDeckBuilderTokenSearchStatus(`Adding ${normalizedName} to tokens...`, 'neutral');
    try {
      deckBuilderSelectedCard = await fetchDeckCardByName(normalizedName);
      if (!deckBuilderSelectedCard?.isToken) {
        setDeckBuilderTokenSearchStatus(`${normalizedName} is not a token card.`, 'error');
        return;
      }

      persistDeckBuilderSelectedCard(deckBuilderSelectedCard);
      await addSelectedCardToTokens();
      if (deckBuilderTokenSearchInput) {
        deckBuilderTokenSearchInput.value = '';
      }
      deckBuilderTokenSearchResultsState = [];
      renderDeckBuilderTokenSearchResults();
    } catch (error) {
      setDeckBuilderTokenSearchStatus(error instanceof Error ? error.message : 'Unable to load that token right now.', 'error');
    }
  }

  async function addSelectedCardToDeck() {
    const deck = ensureActiveDeckBuilderRecord({ createIfMissing: true });
    const card = getCurrentDeckBuilderSelectedCard();
    if (!deck) {
      setDeckBuilderSaveStatus('Unable to open or create a deck right now.', 'error');
      return;
    }

    if (!card) {
      setDeckBuilderSaveStatus('Select a card first, then try Add to Deck again.', 'error');
      return;
    }

    if (card.isToken) {
      await promptLiveAlert('Token cards can only be added to the Tokens section.', 'Token card');
      return;
    }

    const isMultiAllowed = allowsMultipleCopies(card);
    const cardNameKey = getIdentityKey(card.name);
    const deckCardNameSet = getDeckBuilderCardNameSet(deck);

    if (!isMultiAllowed && deckCardNameSet.has(cardNameKey)) {
      const isCommanderMatch = (deck.commander && getIdentityKey(deck.commander.name) === cardNameKey)
        || (deck.secondCommander && getIdentityKey(deck.secondCommander.name) === cardNameKey);
      const message = isCommanderMatch
        ? `${card.name} is already set as commander and cannot also be in the main deck.`
        : `${card.name} is already in this deck. Only basic lands and cards that explicitly allow multiple copies can be added more than once.`;
      await promptLiveAlert(message, 'Duplicate card');
      return;
    }

    persistDeckBuilderRecord({
      ...deck,
      cards: [...deck.cards, card],
    }, `${card.name} added to ${deck.name}.`);
  }

  async function addSelectedCardToMaybeboard() {
    const deck = ensureActiveDeckBuilderRecord({ createIfMissing: true });
    const card = getCurrentDeckBuilderSelectedCard();
    if (!deck) {
      setDeckBuilderSaveStatus('Unable to open or create a deck right now.', 'error');
      return;
    }

    if (!card) {
      setDeckBuilderSaveStatus('Select a card first, then try Add to Maybeboard again.', 'error');
      return;
    }

    if (card.isToken) {
      await promptLiveAlert('Token cards can only be added to the Tokens section.', 'Token card');
      return;
    }

    const cardNameKey = getIdentityKey(card.name);
    const existsInMaybeboard = (deck.maybeboard || []).some((entry) => getIdentityKey(entry.name) === cardNameKey);
    if (existsInMaybeboard) {
      setDeckBuilderSaveStatus(`${card.name} is already in the maybeboard.`, 'neutral');
      return;
    }

    persistDeckBuilderRecord({
      ...deck,
      maybeboard: [...(deck.maybeboard || []), card],
    }, `${card.name} added to maybeboard.`);
    deckBuilderSelectedCard = null;
    persistDeckBuilderSelectedCard(null);
    renderDeckBuilderSelection();
  }

  async function addSelectedCardToTokens() {
    const deck = ensureActiveDeckBuilderRecord({ createIfMissing: true });
    const card = getCurrentDeckBuilderSelectedCard();
    if (!deck) {
      setDeckBuilderSaveStatus('Unable to open or create a deck right now.', 'error');
      return;
    }

    if (!card) {
      setDeckBuilderSaveStatus('Select a card first, then try Add to Tokens again.', 'error');
      return;
    }

    if (!card.isToken) {
      await promptLiveAlert('Only token cards can be added to the Tokens section.', 'Token card');
      return;
    }

    const existingTokens = Array.isArray(deck.tokens) ? [...deck.tokens] : [];
    const cardKey = getIdentityKey(card.name);
    const existingToken = existingTokens.find((entry) => getIdentityKey(entry.name) === cardKey);
    if (existingToken) {
      const updatedTokens = existingTokens.map((entry) => {
        if (getIdentityKey(entry.name) !== cardKey) {
          return entry;
        }
        return {
          ...entry,
          count: (Number.isFinite(Number(entry.count)) ? Math.max(1, Number(entry.count)) : 1) + 1,
        };
      });
      persistDeckBuilderRecord({
        ...deck,
        tokens: updatedTokens,
      }, `${card.name} quantity increased in tokens.`);
      return;
    }

    persistDeckBuilderRecord({
      ...deck,
      tokens: [...existingTokens, { ...card, count: 1 }],
    }, `${card.name} added to tokens.`);
  }

  async function setSelectedCardAsCommander() {
    const deck = ensureActiveDeckBuilderRecord({ createIfMissing: true });
    const selectedCard = getCurrentDeckBuilderSelectedCard();
    if (!deck) {
      setDeckBuilderSaveStatus('Unable to open or create a deck right now.', 'error');
      return;
    }

    if (!selectedCard) {
      setDeckBuilderSaveStatus('Select a card first, then try Set as Commander again.', 'error');
      return;
    }

    const card = await resolveCommanderLegalCard(selectedCard);
    if (!card) {
      await promptLiveAlert(`${selectedCard.name} is not legal as a commander.`, 'Commander legality');
      return;
    }

    const duplicateInDeck = deck.cards.some((entry) => getIdentityKey(entry.name) === getIdentityKey(card.name));
    if (duplicateInDeck) {
      await promptLiveAlert(`${card.name} is already in the main deck. Remove it there before setting it as commander.`, 'Duplicate card');
      return;
    }

    if (!deck.commander) {
      const nextDeck = applyDeckBuilderDraftMeta(deck);
      persistDeckBuilderRecord({ ...nextDeck, commander: card }, `${card.name} set as commander.`);
      return;
    }

    if (getIdentityKey(deck.commander.name) === getIdentityKey(card.name)) {
      return;
    }

    if (deck.secondCommander && getIdentityKey(deck.secondCommander.name) === getIdentityKey(card.name)) {
      return;
    }

    if (!deck.secondCommander && isCompatiblePartnerPair(deck.commander, card)) {
      const nextDeck = applyDeckBuilderDraftMeta(deck);
      persistDeckBuilderRecord({ ...nextDeck, secondCommander: card }, `${card.name} added as partner commander.`);
      return;
    }

    if (!deck.secondCommander && (cardHasPairingMechanic(card) || cardHasPairingMechanic(deck.commander))) {
      const c1Label = getCardPairingMechanicLabel(deck.commander);
      const c2Label = getCardPairingMechanicLabel(card);
      const detail = c1Label && c2Label
        ? `${deck.commander.name} has "${c1Label}" and ${card.name} has "${c2Label}" — these mechanics are not compatible with each other.`
        : c1Label
          ? `${deck.commander.name} has "${c1Label}" but ${card.name} is not a valid pairing for it.`
          : `${card.name} has "${c2Label}" but ${deck.commander.name} is not a valid pairing for it.`;
      await promptLiveAlert(detail, 'Incompatible pairing');
      return;
    }

    const replacingPair = deck.secondCommander
      ? `This will also remove ${deck.secondCommander.name} from the command zone.`
      : '';
    const confirmed = await promptLiveConfirm(
      `Replace ${deck.commander.name} as the commander for ${deck.name}?${replacingPair ? ' ' + replacingPair : ''}`,
      { title: 'Replace commander', confirmLabel: 'Set commander' }
    );
    if (!confirmed) return;

    const nextDeck = applyDeckBuilderDraftMeta(deck);
    persistDeckBuilderRecord({
      ...nextDeck,
      commander: card,
      secondCommander: null,
    }, `${card.name} set as commander.`);
  }

  async function setDeckBuilderCardAsCommander(cardId) {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) {
      setDeckBuilderSaveStatus('Unable to open the deck right now.', 'error');
      return;
    }

    const selectedCard = deck.cards.find((entry) => entry.id === cardId) || null;
    if (!selectedCard) {
      setDeckBuilderSaveStatus('Could not find that card in your deck.', 'error');
      return;
    }

    const nextCommander = await resolveCommanderLegalCard(selectedCard);
    if (!nextCommander) {
      await promptLiveAlert(`${selectedCard.name} is not legal as a commander.`, 'Commander legality');
      return;
    }

    if (!deck.commander) {
      const nextCards = deck.cards.filter((entry) => entry.id !== cardId);
      deckBuilderSelectedDeckCardId = null;
      const nextDeck = applyDeckBuilderDraftMeta(deck);
      persistDeckBuilderRecord({ ...nextDeck, commander: nextCommander, cards: nextCards }, `${nextCommander.name} set as commander.`);
      return;
    }

    if (getIdentityKey(deck.commander.name) === getIdentityKey(nextCommander.name)) {
      return;
    }

    if (!deck.secondCommander && isCompatiblePartnerPair(deck.commander, nextCommander)) {
      const nextCards = deck.cards.filter((entry) => entry.id !== cardId);
      deckBuilderSelectedDeckCardId = null;
      const nextDeck = applyDeckBuilderDraftMeta(deck);
      persistDeckBuilderRecord({ ...nextDeck, secondCommander: nextCommander, cards: nextCards }, `${nextCommander.name} added as partner commander.`);
      return;
    }

    if (!deck.secondCommander && (cardHasPairingMechanic(nextCommander) || cardHasPairingMechanic(deck.commander))) {
      const c1Label = getCardPairingMechanicLabel(deck.commander);
      const c2Label = getCardPairingMechanicLabel(nextCommander);
      const detail = c1Label && c2Label
        ? `${deck.commander.name} has "${c1Label}" and ${nextCommander.name} has "${c2Label}" — these mechanics are not compatible with each other.`
        : c1Label
          ? `${deck.commander.name} has "${c1Label}" but ${nextCommander.name} is not a valid pairing for it.`
          : `${nextCommander.name} has "${c2Label}" but ${deck.commander.name} is not a valid pairing for it.`;
      await promptLiveAlert(detail, 'Incompatible pairing');
      return;
    }

    const replacingPair = deck.secondCommander
      ? ` This will also remove ${deck.secondCommander.name} from the command zone.`
      : '';
    const confirmed = await promptLiveConfirm(
      `Replace ${deck.commander.name} as the commander for ${deck.name}?${replacingPair}`,
      { title: 'Replace commander', confirmLabel: 'Set commander' }
    );
    if (!confirmed) return;

    const nextCards = deck.cards.filter((entry) => entry.id !== cardId);
    deckBuilderSelectedDeckCardId = null;
    const nextDeck = applyDeckBuilderDraftMeta(deck);
    persistDeckBuilderRecord({
      ...nextDeck,
      commander: nextCommander,
      secondCommander: null,
      cards: nextCards,
    }, `${nextCommander.name} set as commander.`);
  }

  async function addMaybeboardCardToDeck(cardId) {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) {
      setDeckBuilderSaveStatus('Unable to open the deck right now.', 'error');
      return;
    }

    const card = (deck.maybeboard || []).find((entry) => entry.id === cardId) || null;
    if (!card) {
      setDeckBuilderSaveStatus('Could not find that maybeboard card.', 'error');
      return;
    }

    if (card.isToken) {
      await promptLiveAlert('Token cards cannot be added to the main deck.', 'Token card');
      return;
    }

    const isMultiAllowed = allowsMultipleCopies(card);
    const cardNameKey = getIdentityKey(card.name);
    const deckCardNameSet = getDeckBuilderCardNameSet(deck);
    if (!isMultiAllowed && deckCardNameSet.has(cardNameKey)) {
      await promptLiveAlert(`${card.name} is already in this deck. Only basic lands and cards that explicitly allow multiple copies can be added more than once.`, 'Duplicate card');
      return;
    }

    persistDeckBuilderRecord({
      ...deck,
      cards: [...deck.cards, { ...card, id: generateId() }],
      maybeboard: (deck.maybeboard || []).filter((entry) => entry.id !== cardId),
    }, `${card.name} moved from maybeboard to deck.`);
  }

  async function moveDeckCardToMaybeboard(cardId) {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) {
      setDeckBuilderSaveStatus('Unable to open the deck right now.', 'error');
      return;
    }

    const card = (deck.cards || []).find((entry) => entry.id === cardId) || null;
    if (!card) {
      setDeckBuilderSaveStatus('Could not find that card in your deck.', 'error');
      return;
    }

    const existsInMaybeboard = (deck.maybeboard || []).some((entry) => getIdentityKey(entry.name) === getIdentityKey(card.name));
    if (existsInMaybeboard) {
      await promptLiveAlert(`${card.name} is already in the maybeboard.`, 'Duplicate card');
      return;
    }

    persistDeckBuilderRecord({
      ...deck,
      cards: (deck.cards || []).filter((entry) => entry.id !== cardId),
      maybeboard: [...(deck.maybeboard || []), { ...card, id: generateId() }],
    }, `${card.name} moved to maybeboard.`);
  }

  function updateDeckBuilderTokenQuantity(cardId, delta) {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) {
      return;
    }

    const existingTokens = Array.isArray(deck.tokens) ? [...deck.tokens] : [];
    let updated = false;
    const nextTokens = existingTokens.reduce((result, entry) => {
      if (entry.id !== cardId) {
        result.push(entry);
        return result;
      }

      const currentCount = Number.isFinite(Number(entry.count)) ? Math.max(1, Number(entry.count)) : 1;
      const nextCount = currentCount + delta;
      if (nextCount <= 0) {
        updated = true;
        return result;
      }

      updated = true;
      result.push({ ...entry, count: nextCount });
      return result;
    }, []);

    if (!updated) {
      return;
    }

    persistDeckBuilderRecord({
      ...deck,
      tokens: nextTokens,
    }, delta < 0 ? 'Token quantity decreased.' : 'Token quantity increased.', 'success', {
      persistDelayMs: DECKS_RAPID_ACTION_PERSIST_DEBOUNCE_MS,
    });
  }

  function removeDeckBuilderCard(cardId, { isCommander = false, isSecondCommander = false } = {}) {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) {
      return;
    }

    if (isSecondCommander) {
      persistDeckBuilderRecord({
        ...deck,
        secondCommander: null,
      }, 'Partner commander removed.', 'neutral');
      return;
    }

    if (isCommander) {
      persistDeckBuilderRecord({
        ...deck,
        commander: null,
        secondCommander: null,
      }, 'Commander removed.', 'neutral');
      return;
    }

    const removedMainDeckCard = (deck.cards || []).find((card) => card.id === cardId) || null;
    const useRapidDelay = Boolean(removedMainDeckCard && allowsMultipleCopies(removedMainDeckCard));

    persistDeckBuilderRecord({
      ...deck,
      cards: deck.cards.filter((card) => card.id !== cardId),
      maybeboard: (deck.maybeboard || []).filter((card) => card.id !== cardId),
      tokens: (deck.tokens || []).filter((card) => card.id !== cardId),
    }, 'Card removed.', 'neutral', useRapidDelay
      ? { persistDelayMs: DECKS_RAPID_ACTION_PERSIST_DEBOUNCE_MS }
      : undefined);
  }

  function ensureActiveDeckBuilderRecord({ createIfMissing = false } = {}) {
    if (!deckBuilderPage) {
      return null;
    }

    const requestedDeckId = getQueryParam('deckId');
    const shouldCreateNew = getQueryParam('new') === '1';

    if (requestedDeckId) {
      if (
        activeDeckBuilderRecord
        && activeDeckBuilderRecord.id === requestedDeckId
        && deckBuilderHasUnsavedChanges
      ) {
        return activeDeckBuilderRecord;
      }

      const requestedDeck = loadDecks().find((deck) => deck.id === requestedDeckId) || null;
      if (requestedDeck) {
        activeDeckBuilderId = requestedDeck.id;
        activeDeckBuilderRecord = requestedDeck;
        deckBuilderHasUnsavedChanges = false;
        updateDeckBuilderSaveButton();
        return requestedDeck;
      }
    }

    if (shouldCreateNew) {
      const newDeck = createEmptyDeckRecord();
      activeDeckBuilderId = newDeck.id;
      activeDeckBuilderRecord = newDeck;
      deckBuilderHasUnsavedChanges = true;
      updateDeckBuilderSaveButton();
      window.history.replaceState({}, '', getDeckBuilderHref(newDeck.id));
      return newDeck;
    }

    if (createIfMissing) {
      const newDeck = createEmptyDeckRecord();
      activeDeckBuilderId = newDeck.id;
      activeDeckBuilderRecord = newDeck;
      deckBuilderHasUnsavedChanges = true;
      updateDeckBuilderSaveButton();
      window.history.replaceState({}, '', getDeckBuilderHref(newDeck.id));
      setDeckBuilderSaveStatus('Started a new deck. Press Save Deck to keep it.', 'neutral');
      return newDeck;
    }

    activeDeckBuilderId = '';
    activeDeckBuilderRecord = null;
    return null;
  }

  function updateDeckBuilderUndoButton() {
    if (!deckBuilderUndoButton) return;
    if (activeDeckBuilderRecord && !canCurrentUserEditDeck(activeDeckBuilderRecord)) {
      deckBuilderUndoButton.disabled = true;
      return;
    }
    const stackLen = activeDeckBuilderRecord
      ? (deckBuilderUndoStacks[activeDeckBuilderRecord.id] || []).length
      : 0;
    deckBuilderUndoButton.disabled = stackLen === 0;
  }

  async function undoDeckBuilderChange() {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) return;
    const stack = deckBuilderUndoStacks[deck.id] || [];
    if (stack.length === 0) {
      setDeckBuilderSaveStatus('Nothing to undo.', 'muted');
      return;
    }
    const previousState = stack.pop();
    persistDeckBuilderRecord(previousState, 'Change undone.', 'muted', { skipUndo: true });
  }

  function persistDeckBuilderRecord(nextDeck, statusMessage = 'Saved locally.', tone = 'success', {
    skipUndo = false,
    skipFullRefresh = false,
    persistDelayMs = DECKS_RAPID_ACTION_PERSIST_DEBOUNCE_MS,
  } = {}) {
    const ownedDeck = applyDeckOwnership(nextDeck);
    if (!ownedDeck) {
      return;
    }

    if (!canCurrentUserEditDeck(ownedDeck)) {
      setDeckBuilderSaveStatus(getDeckReadOnlyMessage(ownedDeck), 'error');
      return;
    }

    if (!skipUndo && activeDeckBuilderRecord) {
      if (!deckBuilderUndoStacks[activeDeckBuilderRecord.id]) {
        deckBuilderUndoStacks[activeDeckBuilderRecord.id] = [];
      }
      deckBuilderUndoStacks[activeDeckBuilderRecord.id].push(JSON.parse(JSON.stringify(activeDeckBuilderRecord)));
    }
    const normalizedDeck = normalizeDeckRecord({
      ...ownedDeck,
      updatedAt: new Date().toISOString(),
    });

    activeDeckBuilderId = normalizedDeck.id;
    activeDeckBuilderRecord = normalizedDeck;
    appState = normalizeAppStateData({
      ...appState,
      decks: (Array.isArray(appState.decks) && appState.decks.some((deck) => deck.id === normalizedDeck.id))
        ? appState.decks.map((deck) => (deck.id === normalizedDeck.id ? normalizedDeck : deck))
        : [...(Array.isArray(appState.decks) ? appState.decks : []), normalizedDeck],
    });
    deckBuilderHasUnsavedChanges = true;

    setDeckBuilderSaveStatus(statusMessage, tone);
    updateDeckBuilderUndoButton();
    updateDeckBuilderSaveButton();

    if (deckBuilderPage) {
      renderDeckBuilderPage();
    }
  }

  function hasUnsavedDeckBuilderChanges() {
    return deckBuilderHasUnsavedChanges;
  }

  function updateDeckBuilderSaveButton() {
    if (!deckBuilderSaveButton) {
      return;
    }
    const deck = activeDeckBuilderRecord;
    const readOnly = deck ? !canCurrentUserEditDeck(deck) : true;
    deckBuilderSaveButton.disabled = readOnly || !deckBuilderHasUnsavedChanges;
    deckBuilderSaveButton.classList.toggle('has-unsaved-changes', deckBuilderHasUnsavedChanges && !readOnly);
  }

  async function saveActiveDeckBuilderDeck() {
    const deck = ensureActiveDeckBuilderRecord();
    if (!deck) {
      return;
    }

    if (!canCurrentUserEditDeck(deck)) {
      setDeckBuilderSaveStatus(getDeckReadOnlyMessage(deck), 'error');
      return;
    }

    const stagedDeck = applyDeckBuilderDraftMeta(deck);
    const normalizedDeck = normalizeDeckRecord({
      ...stagedDeck,
      updatedAt: new Date().toISOString(),
    });

    const existingDecks = loadDecks();
    const nextDecks = existingDecks.some((entry) => entry.id === normalizedDeck.id)
      ? existingDecks.map((entry) => (entry.id === normalizedDeck.id ? normalizedDeck : entry))
      : [...existingDecks, normalizedDeck];

    saveDecks(nextDecks, { deferPersist: false, delay: 0 });
    flushQueuedDeckPersist({ force: true });

    activeDeckBuilderId = normalizedDeck.id;
    activeDeckBuilderRecord = normalizedDeck;
    deckBuilderHasUnsavedChanges = false;

    if (typeof normalizedDeck.powerLevel === 'number' && normalizedDeck.commander?.name) {
      setCommanderExpectedPower(normalizedDeck.commander.name, normalizedDeck.powerLevel);
    }
    linkDeckListToDeck(normalizedDeck);
    flashDeckBuilderSaveConfirmation('✓ Deck saved.');
    updateDeckBuilderSaveButton();

    if (syncQueueTimer) {
      clearTimeout(syncQueueTimer);
      syncQueueTimer = null;
    }
    if (hasSyncCredentials() && !syncConflictInfo) {
      try {
        await pushCloudState();
        flashDeckBuilderSaveConfirmation('✓ Deck saved and synced.');
      } catch (_) {
        setDeckBuilderSaveStatus('Deck saved locally. Cloud sync will retry.', 'muted');
      }
    }
  }

  function renderDeckBuilderSelection() {
    if (!deckBuilderSelection) {
      return;
    }

    const deck = ensureActiveDeckBuilderRecord();
    const card = normalizeDeckCardEntry(deckBuilderSelectedCard);
    const isReadOnly = deck ? !canCurrentUserEditDeck(deck) : false;

    if (!card) {
      persistDeckBuilderSelectedCard(null);
      delete deckBuilderSelection.dataset.selectedCard;

      if (deck?.commander) {
        const commander = deck.commander;
        const commanderFaces = Array.isArray(commander.cardFaces) ? commander.cardFaces : [];
        const activeFaceIndex = commanderFaces.length
          ? Math.min(Math.max(0, deckBuilderSelectedFaceIndex), commanderFaces.length - 1)
          : 0;
        const activeFace = commanderFaces[activeFaceIndex] || null;
        const activeFaceName = activeFace?.name || commander.name;
        const activeFaceImage = activeFace?.imageLargeUri || activeFace?.imageUri || commander.imageLargeUri || commander.imageUri || '';
        const activeFaceTypeLine = activeFace?.typeLine || commander.typeLine || 'No type line available';
        const activeFaceManaCost = activeFace?.manaCost || commander.manaCost || '—';
        const badges = [
          commander.isBanned ? '<span class="deck-card-badge deck-card-badge-banned">Banned</span>' : '',
          commander.isGameChanger ? '<span class="deck-card-badge deck-card-badge-gamechanger" title="Game Changer" aria-label="Game Changer">&#9889;</span>' : '',
        ].filter(Boolean).join('');
        const rulesText = activeFace?.oracleText || getDeckCardRulesText(commander);
        const statLine = getDeckCardStatLine(activeFace || commander);

        deckBuilderSelection.innerHTML = `
        <article class="deck-card-preview">
          ${activeFaceImage ? `<img class="deck-card-preview-image" src="${escapeHtml(activeFaceImage)}" alt="${escapeHtml(activeFaceName)}" loading="lazy" />` : ''}
          <div class="deck-card-preview-copy">
            <p class="deck-card-preview-kicker">Current Commander</p>
            <h3>${buildCommanderDisplayHtml(activeFaceName, escapeHtml(activeFaceName), commander)}</h3>
            <p class="deck-card-preview-meta">${escapeHtml(activeFaceTypeLine)}</p>
            <p class="deck-card-preview-meta">Mana cost: ${escapeHtml(activeFaceManaCost)}</p>
            ${statLine ? `<p class="deck-card-preview-meta">${escapeHtml(statLine)}</p>` : ''}
            ${rulesText ? `<div class="deck-card-preview-rules-text">${formatCommanderBuilderRichText(rulesText)}</div>` : ''}
            ${badges ? `<div class="deck-card-badge-row">${badges}</div>` : ''}
          </div>
          <div class="actions deck-card-preview-actions">
            ${commanderFaces.length > 1 ? `<button type="button" id="deck-builder-flip-card" class="secondary-button">Flip</button>` : ''}
            <button type="button" class="secondary-button" onclick="window.__deckBuilderRemoveCommander && window.__deckBuilderRemoveCommander(event)"${isReadOnly ? ' disabled' : ''}>Remove Commander</button>
          </div>
        </article>`;

        deckBuilderSelection.querySelector('#deck-builder-flip-card')?.addEventListener('click', () => {
          deckBuilderSelectedFaceIndex = (activeFaceIndex + 1) % commanderFaces.length;
          renderDeckBuilderSelection();
        });

        const removeButton = deckBuilderSelection.querySelector('[onclick*="RemoveCommander"]');
        if (removeButton) {
          removeButton.addEventListener('click', async (event) => {
            event.preventDefault();
            event.stopPropagation();
            removeDeckBuilderCard('', { isCommander: true });
          });
        }
      } else {
        deckBuilderSelection.innerHTML = '<p class="status-muted">Search for a card or click a card in your deck to see its details here.</p>';
      }
      return;
    }

    persistDeckBuilderSelectedCard(card);
    deckBuilderSelection.dataset.selectedCard = JSON.stringify(card);

    const cardFaces = Array.isArray(card.cardFaces) ? card.cardFaces : [];
    const activeFaceIndex = cardFaces.length
      ? Math.min(Math.max(0, deckBuilderSelectedFaceIndex), cardFaces.length - 1)
      : 0;
    const activeFace = cardFaces[activeFaceIndex] || null;
    const activeFaceName = activeFace?.name || card.name;
    const activeFaceImage = activeFace?.imageLargeUri || activeFace?.imageUri || card.imageLargeUri || card.imageUri || '';
    const activeFaceTypeLine = activeFace?.typeLine || card.typeLine || 'No type line available';
    const activeFaceManaCost = activeFace?.manaCost || card.manaCost || '—';
    const activeFaceRulesText = activeFace?.oracleText || getDeckCardRulesText(card);
    const activeFaceStatLine = getDeckCardStatLine(activeFace || card);
    const canFlipFaces = cardFaces.length > 1;
    const badges = [
      card.isBanned ? '<span class="deck-card-badge deck-card-badge-banned">Banned</span>' : '',
      card.isGameChanger ? '<span class="deck-card-badge deck-card-badge-gamechanger" title="Game Changer" aria-label="Game Changer">&#9889;</span>' : '',
    ].filter(Boolean).join('');
    const canSetCommander = canSetCardAsCommanderForDeck(deck, card);

    deckBuilderSelection.innerHTML = `
    <article class="deck-card-preview">
      ${activeFaceImage ? `<img class="deck-card-preview-image" src="${escapeHtml(activeFaceImage)}" alt="${escapeHtml(activeFaceName)}" loading="lazy" />` : ''}
      <div class="deck-card-preview-copy">
        <p class="deck-card-preview-kicker">Selected Card</p>
        <h3>${escapeHtml(activeFaceName)}</h3>
        <p class="deck-card-preview-meta">${escapeHtml(activeFaceTypeLine)}</p>
        <p class="deck-card-preview-meta">Mana cost: ${escapeHtml(activeFaceManaCost)}</p>
        ${activeFaceStatLine ? `<p class="deck-card-preview-meta">${escapeHtml(activeFaceStatLine)}</p>` : ''}
        ${activeFaceRulesText ? `<div class="deck-card-preview-rules-text">${formatCommanderBuilderRichText(activeFaceRulesText)}</div>` : ''}
        ${badges ? `<div class="deck-card-badge-row">${badges}</div>` : ''}
      </div>
      <div class="actions deck-card-preview-actions">
        ${canFlipFaces ? `<button type="button" id="deck-builder-flip-card" class="secondary-button">Flip</button>` : ''}
        <button type="button" id="deck-builder-add-card" onclick="window.__deckBuilderAddSelectedCard && window.__deckBuilderAddSelectedCard(event)"${isReadOnly ? ' disabled' : ''}>Add to Deck</button>
        ${card.isToken ? `<button type="button" id="deck-builder-add-token" class="secondary-button" onclick="window.__deckBuilderAddSelectedToken && window.__deckBuilderAddSelectedToken(event)"${isReadOnly ? ' disabled' : ''}>Add to Tokens</button>` : ''}
        <button type="button" id="deck-builder-add-maybeboard" class="secondary-button" onclick="window.__deckBuilderAddSelectedMaybeboard && window.__deckBuilderAddSelectedMaybeboard(event)"${isReadOnly ? ' disabled' : ''}>Add to Maybeboard</button>
        ${canSetCommander ? `<button type="button" id="deck-builder-set-commander" class="secondary-button" onclick="window.__deckBuilderSetCommander && window.__deckBuilderSetCommander(event)"${isReadOnly ? ' disabled' : ''}>Set as Commander</button>` : ''}
      </div>
    </article>`;

    deckBuilderSelection.querySelector('#deck-builder-flip-card')?.addEventListener('click', () => {
      deckBuilderSelectedFaceIndex = (activeFaceIndex + 1) % cardFaces.length;
      renderDeckBuilderSelection();
    });
  }

  function renderDeckCardRow(card, options = {}) {
    const isReadOnly = Boolean(options.readOnly);
    const cardNameMarkup = options.isCommander ? buildCommanderDisplayHtml(card.name, escapeHtml(card.name), card) : escapeHtml(card.name);
    const badges = [
      card.isBanned ? '<span class="deck-card-badge deck-card-badge-banned">Banned</span>' : '',
      card.isGameChanger ? '<span class="deck-card-badge deck-card-badge-gamechanger" title="Game Changer" aria-label="Game Changer">&#9889;</span>' : '',
    ].filter(Boolean).join('');
    const rulesText = getDeckCardRulesText(card);
    const statLine = getDeckCardStatLine(card);
    const isExpanded = options.isCommander || options.isSelected || Boolean(options.showArtPicker);

    const imageMarkup = isExpanded && (card.imageLargeUri || card.imageUri)
      ? `
      <div class="deck-card-row-media">
        <img
          class="deck-card-row-image"
          src="${escapeHtml(card.imageLargeUri || card.imageUri)}"
          alt="${escapeHtml(card.name)}"
          loading="lazy"
          decoding="async"
        />
      </div>`
      : '';
    const removeAction = isExpanded
      ? (options.isSecondCommander
        ? `<button type="button" class="history-delete-button deck-builder-remove-card" data-remove-commander="true" data-remove-second-commander="true"${isReadOnly ? ' disabled' : ''}>Remove</button>`
        : options.isCommander
          ? `<button type="button" class="history-delete-button deck-builder-remove-card" data-remove-commander="true"${isReadOnly ? ' disabled' : ''}>Remove</button>`
          : `<button type="button" class="history-delete-button deck-builder-remove-card" data-card-id="${escapeHtml(card.id)}"${isReadOnly ? ' disabled' : ''}>Remove</button>`)
      : '';
    const commanderAction = isExpanded && !options.isCommander && !options.fromMaybeboard && !options.fromTokens && Boolean(options.canSetCommander)
      ? `<button type="button" class="secondary-button deck-builder-set-row-commander" data-set-commander-id="${escapeHtml(card.id)}"${isReadOnly ? ' disabled' : ''}>Set as Commander</button>`
      : '';
    const addToDeckAction = isExpanded && options.fromMaybeboard
      ? `<button type="button" class="secondary-button deck-builder-maybe-to-deck" data-add-from-maybeboard-id="${escapeHtml(card.id)}"${isReadOnly ? ' disabled' : ''}>Add to Deck</button>`
      : '';
    const moveToMaybeboardAction = isExpanded && !options.isCommander && !options.fromMaybeboard && !options.fromTokens
      ? `<button type="button" class="secondary-button deck-builder-move-to-maybeboard" data-move-to-maybeboard-id="${escapeHtml(card.id)}"${isReadOnly ? ' disabled' : ''}>To Maybeboard</button>`
      : '';
    const tokenQuantityAction = options.fromTokens
      ? `
      <button type="button" class="deck-token-minus" data-remove-token-id="${escapeHtml(card.id)}" aria-label="Remove one ${escapeHtml(card.name)}"${isReadOnly ? ' disabled' : ''}>−</button>
      <span class="deck-token-quantity">×${escapeHtml(String(Number.isFinite(Number(card.count)) ? Math.max(1, Number(card.count)) : 1))}</span>
      <button type="button" class="deck-token-plus" data-add-token-id="${escapeHtml(card.id)}" aria-label="Add one ${escapeHtml(card.name)}"${isReadOnly ? ' disabled' : ''}>+</button>
    `
      : '';
    const artAction = isExpanded && !options.isBasicLand
      ? `<button type="button" class="secondary-button deck-builder-change-art" data-change-art-id="${escapeHtml(card.id)}"${isReadOnly ? ' disabled' : ''}>Change Art</button>`
      : '';
    const isArtPickerOpen = Boolean(options.showArtPicker);
    const artState = isArtPickerOpen ? deckBuilderArtPickerState : null;
    const artBody = isArtPickerOpen
      ? `
      <div class="deck-card-art-picker">
        ${artState?.status === 'loading' ? `<p class="deck-card-art-status">${escapeHtml(artState.message || 'Loading art options...')}</p>` : ''}
        ${artState?.status === 'error' ? `<p class="deck-card-art-status status-error">${escapeHtml(artState.message || 'Unable to load art options.')}</p>` : ''}
        ${artState?.status === 'ready' && !(artState.options || []).length ? `<p class="deck-card-art-status status-muted">${escapeHtml(artState.message || 'No alternate arts found for this card.')}</p>` : ''}
        ${artState?.status === 'ready' && (artState.options || []).length ? `
          <div class="deck-card-art-grid">
            ${(artState.options || []).map((print) => `
              <button
                type="button"
                class="deck-card-art-option"
                data-apply-art-card-id="${escapeHtml(card.id)}"
                data-apply-art-print-id="${escapeHtml(print.id || '')}"
                title="${escapeHtml(`${print.setName || print.set || 'Print'} ${print.collectorNumber || ''}`.trim())}"
              >
                ${print.imageUri ? `<img src="${escapeHtml(print.imageUri)}" alt="${escapeHtml(print.name || card.name)}" loading="lazy" decoding="async" />` : ''}
                <span>${escapeHtml(`${(print.set || '').toUpperCase()} ${print.collectorNumber || ''}`.trim() || (print.lang || 'Print'))}</span>
              </button>
            `).join('')}
          </div>` : ''}
      </div>`
      : '';

    const cardDataAttr = options.isCommander ? '' : ` data-card-id="${escapeHtml(card.id)}" data-card-name="${escapeHtml(card.name)}"`;

    return `
    <div class="deck-card-row${card.isBanned ? ' is-banned' : ''}${options.isCommander ? ' is-commander' : ''}${options.isSelected ? ' is-selected' : ''}${options.isIllegal ? ' is-illegal' : ''}"${cardDataAttr}>
      ${imageMarkup}
      <div class="deck-card-row-main">
        <div class="deck-card-row-copy">
          <p class="deck-card-name">${cardNameMarkup}${Number.isFinite(Number(card.count)) && Number(card.count) > 1 ? `<span class="deck-card-count">×${escapeHtml(String(card.count))}</span>` : ''}</p>
          <p class="deck-card-meta">${escapeHtml(card.typeLine || card.cardType || 'Unknown')}</p>
          ${isExpanded && card.manaCost ? `<p class="deck-card-meta">Mana cost: ${escapeHtml(card.manaCost)}</p>` : ''}
          ${isExpanded && statLine ? `<p class="deck-card-meta">${escapeHtml(statLine)}</p>` : ''}
          ${isExpanded && rulesText ? `<div class="deck-card-rules-text">${formatCommanderBuilderRichText(rulesText)}</div>` : ''}
          ${badges ? `<div class="deck-card-badge-row">${badges}</div>` : ''}
        </div>
        <div class="deck-card-row-actions">
          ${options.fromTokens ? tokenQuantityAction : ''}
          ${addToDeckAction}
          ${moveToMaybeboardAction}
          ${commanderAction}
          ${artAction}
          ${removeAction}
        </div>
      </div>
      ${artBody}
    </div>`;
  }

  function renderDeckGroupCardRows(cards, { deck, isIllegalCard, isReadOnly }) {
    const unlimitedGroups = new Map();
    const regularCards = [];
    cards.forEach((card) => {
      if (isUnlimitedCopiesCard(card)) {
        const key = getIdentityKey(card.name);
        if (!unlimitedGroups.has(key)) unlimitedGroups.set(key, { name: card.name, typeLine: card.typeLine || '', cards: [] });
        unlimitedGroups.get(key).cards.push(card);
      } else {
        regularCards.push(card);
      }
    });
    const unlimitedRows = [...unlimitedGroups.values()]
      .map(({ name, cards: groupedCards }) => renderUnlimitedCopyCardRow(name, groupedCards))
      .join('');
    const regularRows = regularCards.map((card) => renderDeckCardRow(card, {
      isSelected: card.id === deckBuilderSelectedDeckCardId,
      isIllegal: isIllegalCard(card),
      canSetCommander: deckCardShowPartnerCommanderButton(deck, card),
      showArtPicker: card.id === deckBuilderArtPickerCardId,
      readOnly: isReadOnly,
    })).join('');
    return unlimitedRows + regularRows;
  }

  function renderUnlimitedCopyCardRow(name, cards) {
    const count = cards.length;
    const sampleId = cards[0]?.id || '';
    const typeLine = cards[0]?.typeLine || '';
    const deck = ensureActiveDeckBuilderRecord();
    const isReadOnly = deck ? !canCurrentUserEditDeck(deck) : false;
    const isExpanded = sampleId && (deckBuilderSelectedDeckCardId === sampleId || deckBuilderArtPickerCardId === sampleId);
    const artButton = sampleId && isExpanded
      ? `<button type="button" class="secondary-button deck-builder-change-art" data-change-art-id="${escapeHtml(sampleId)}" aria-label="Change art for ${escapeHtml(name)}"${isReadOnly ? ' disabled' : ''}>Art</button>`
      : '';
    return `
    <div class="deck-card-row deck-card-row-basic-land${isExpanded ? ' is-selected' : ''}" data-card-id="${escapeHtml(sampleId)}" data-card-name="${escapeHtml(name)}">
      <div class="deck-card-row-copy">
        <p class="deck-card-name">${escapeHtml(name)}</p>
        ${typeLine ? `<p class="deck-card-meta">${escapeHtml(typeLine)}</p>` : ''}
      </div>
      <div class="deck-card-row-actions deck-card-row-actions-land">
        <button type="button" class="deck-land-minus deck-builder-remove-card" data-card-id="${escapeHtml(sampleId)}" aria-label="Remove one ${escapeHtml(name)}"${isReadOnly ? ' disabled' : ''}>−</button>
        <span class="deck-land-count">×${escapeHtml(String(count))}</span>
        <button type="button" class="deck-land-plus" data-add-unlimited="${escapeHtml(name)}" aria-label="Add one ${escapeHtml(name)}"${isReadOnly ? ' disabled' : ''}>+</button>
        ${artButton}
      </div>
    </div>`;
  }

  function renderBasicLandRow(name, cards) {
    const count = cards.length;
    const sampleId = cards[0]?.id || '';
    const deck = ensureActiveDeckBuilderRecord();
    const isReadOnly = deck ? !canCurrentUserEditDeck(deck) : false;
    const isExpanded = sampleId && (deckBuilderSelectedDeckCardId === sampleId || deckBuilderArtPickerCardId === sampleId);
    const artButton = sampleId && isExpanded
      ? `<button type="button" class="secondary-button deck-builder-change-art" data-change-art-id="${escapeHtml(sampleId)}" aria-label="Change art for ${escapeHtml(name)}"${isReadOnly ? ' disabled' : ''}>Art</button>`
      : '';
    return `
    <div class="deck-card-row deck-card-row-basic-land${isExpanded ? ' is-selected' : ''}" data-basic-land-name="${escapeHtml(name)}" data-card-id="${escapeHtml(sampleId)}" data-card-name="${escapeHtml(name)}">
      <div class="deck-card-row-copy">
        <p class="deck-card-name">${escapeHtml(name)}</p>
        <p class="deck-card-meta">Basic Land</p>
      </div>
      <div class="deck-card-row-actions deck-card-row-actions-land">
        <button type="button" class="deck-land-minus deck-builder-remove-card" data-card-id="${escapeHtml(sampleId)}" data-remove-basic="${escapeHtml(name)}" aria-label="Remove one ${escapeHtml(name)}"${isReadOnly ? ' disabled' : ''}>−</button>
        <span class="deck-land-count">×${escapeHtml(String(count))}</span>
        <button type="button" class="deck-land-plus" data-add-basic="${escapeHtml(name)}" aria-label="Add one ${escapeHtml(name)}"${isReadOnly ? ' disabled' : ''}>+</button>
        ${artButton}
      </div>
    </div>`;
  }

  function renderDeckBuilderCards(deck) {
    if (!deckBuilderCards) {
      return;
    }

    const groups = getDeckBuilderGroupedCards(deck);
    const summary = getDeckValidationSummary(deck);
    const colorViolationNames = new Set(
      summary.colorViolations.map((violation) => getIdentityKey(violation.replace(/\s*\([^)]*\)$/, '').trim())
    );
    const duplicateCardNames = new Set(
      summary.duplicates.map((name) => getIdentityKey(name))
    );
    const isIllegalCard = (card) => {
      const key = getIdentityKey(card?.name);
      return Boolean(key) && (colorViolationNames.has(key) || duplicateCardNames.has(key));
    };
    const isReadOnly = !canCurrentUserEditDeck(deck);

    const commanderCards = [deck?.commander, deck?.secondCommander].filter(Boolean);
    const commanderMarkup = commanderCards.length
      ? `
      <section class="deck-builder-group">
        <div class="deck-builder-group-header">
          <h3>Commander</h3>
          <p>${commanderCards.length === 2 ? '2 cards — Partner' : '1 card'}</p>
        </div>
        <div class="deck-builder-group-cards deck-builder-group-cards--single">
          ${renderDeckCardRow(deck.commander, { isCommander: true, isIllegal: isIllegalCard(deck.commander), showArtPicker: deck.commander?.id === deckBuilderArtPickerCardId, readOnly: isReadOnly })}
          ${deck.secondCommander ? renderDeckCardRow(deck.secondCommander, { isCommander: true, isSecondCommander: true, isIllegal: isIllegalCard(deck.secondCommander), showArtPicker: deck.secondCommander?.id === deckBuilderArtPickerCardId, readOnly: isReadOnly }) : ''}
        </div>
      </section>`
      : `
      <section class="deck-builder-group">
        <div class="deck-builder-group-header">
          <h3>Commander</h3>
          <p>Not set</p>
        </div>
        <p class="deck-builder-empty-copy">Choose a card from search and use Set as Commander.</p>
      </section>`;

    const basicLandKeySet = new Set(BASIC_LAND_NAMES.map((name) => getIdentityKey(name)));

    const groupMarkup = groups.map((group) => {
      if (group.type === 'Land') {
        const basicGroups = new Map();
        const nonBasics = [];
        group.cards.forEach((card) => {
          if (basicLandKeySet.has(getIdentityKey(card.name))) {
            const key = getIdentityKey(card.name);
            if (!basicGroups.has(key)) basicGroups.set(key, { name: card.name, cards: [] });
            basicGroups.get(key).cards.push(card);
          } else {
            nonBasics.push(card);
          }
        });
        const totalCards = group.cards.length;
        const basicRows = [...basicGroups.values()].map(({ name, cards }) => renderBasicLandRow(name, cards)).join('');
        const nonBasicRows = renderDeckGroupCardRows(nonBasics, { deck, isIllegalCard, isReadOnly });
        const quickAddButtons = BASIC_LAND_NAMES.map((name) =>
          `<button type="button" class="secondary-button deck-land-quick-add" data-add-basic="${escapeHtml(name)}">${escapeHtml(name)}</button>`
        ).join('');
        return `
        <section class="deck-builder-group">
          <div class="deck-builder-group-header">
            <h3>Land</h3>
            <p>${escapeHtml(String(totalCards))} card${totalCards === 1 ? '' : 's'}</p>
          </div>
          <div class="deck-builder-group-cards">${basicRows}${nonBasicRows}</div>
          <div class="deck-land-quick-add-row">
            <p class="deck-land-quick-add-label">Add basic land:</p>
            ${quickAddButtons}
          </div>
        </section>`;
      }
      return `
      <section class="deck-builder-group">
        <div class="deck-builder-group-header">
          <h3>${escapeHtml(group.type)}</h3>
          <p>${escapeHtml(String(group.cards.length))} card${group.cards.length === 1 ? '' : 's'}</p>
        </div>
        <div class="deck-builder-group-cards">${renderDeckGroupCardRows(group.cards, { deck, isIllegalCard, isReadOnly })}</div>
      </section>`;
    }).join('');

    const landGroupExists = groups.some((group) => group.type === 'Land');
    const quickAddButtons = BASIC_LAND_NAMES.map((name) =>
      `<button type="button" class="secondary-button deck-land-quick-add" data-add-basic="${escapeHtml(name)}">${escapeHtml(name)}</button>`
    ).join('');
    const landQuickAddSection = landGroupExists ? '' : `
    <section class="deck-builder-group">
      <div class="deck-builder-group-header">
        <h3>Land</h3>
        <p>0 cards</p>
      </div>
      <div class="deck-land-quick-add-row">
        <p class="deck-land-quick-add-label">Add basic land:</p>
        ${quickAddButtons}
      </div>
    </section>`;

    const tokenCards = (deck?.tokens || []).slice().sort((first, second) => compareTextValues(first.name, second.name));
    const tokenCount = tokenCards.reduce((sum, card) => sum + (Number.isFinite(Number(card.count)) ? Math.max(1, Number(card.count)) : 1), 0);
    const tokensMarkup = `
    <section class="deck-builder-group">
      <div class="deck-builder-group-header">
        <h3>Tokens</h3>
        <p>${escapeHtml(String(tokenCount))} token${tokenCount === 1 ? '' : 's'}</p>
      </div>
      ${tokenCards.length
        ? `<div class="deck-builder-group-cards">${tokenCards.map((card) => renderDeckCardRow(card, {
          isSelected: card.id === deckBuilderSelectedDeckCardId,
          fromTokens: true,
          canSetCommander: false,
          showArtPicker: card.id === deckBuilderArtPickerCardId,
          readOnly: isReadOnly,
        })).join('')}</div>`
        : '<p class="deck-builder-empty-copy">Track token cards here. Tokens do not count toward deck legality.</p>'}
    </section>`;

    const maybeboardCards = (deck?.maybeboard || []).slice().sort((first, second) => compareTextValues(first.name, second.name));
    const maybeboardMarkup = `
    <section class="deck-builder-group">
      <div class="deck-builder-group-header">
        <h3>Maybeboard</h3>
        <p>${escapeHtml(String(maybeboardCards.length))} card${maybeboardCards.length === 1 ? '' : 's'}</p>
      </div>
      ${maybeboardCards.length
        ? `<div class="deck-builder-group-cards">${maybeboardCards.map((card) => renderDeckCardRow(card, {
          isSelected: card.id === deckBuilderSelectedDeckCardId,
          fromMaybeboard: true,
          canSetCommander: deckCardShowPartnerCommanderButton(deck, card),
          showArtPicker: card.id === deckBuilderArtPickerCardId,
          readOnly: isReadOnly,
        })).join('')}</div>`
        : '<p class="deck-builder-empty-copy">Add cards here as possible includes. Maybeboard cards do not affect deck legality.</p>'}
    </section>`;

    deckBuilderCards.innerHTML = `${commanderMarkup}${groupMarkup || ''}${landQuickAddSection}${tokensMarkup}${maybeboardMarkup}`;
  }

  function renderDeckBuilderPage() {
    if (!deckBuilderPage) {
      return;
    }

    void warmDeckBuilderBasicLandCache();

    const prefilledCommanderName = deckBuilderCommanderPrefill;
    deckBuilderCommanderPrefill = '';
    const deck = ensureActiveDeckBuilderRecord({ createIfMissing: !!prefilledCommanderName });
    if (!deck) {
      if (deckBuilderCards) {
        deckBuilderCards.innerHTML = '<p>No deck selected. Go back to the deck list page and create or open a deck.</p>';
      }
      if (deckBuilderValidation) {
        deckBuilderValidation.innerHTML = '';
      }
      if (deckBuilderBreakdown) {
        deckBuilderBreakdown.innerHTML = '';
      }
      return;
    }

    if (deckBuilderTitle) {
      deckBuilderTitle.textContent = deck.name;
    }
    if (deckBuilderNameInput && deckBuilderNameInput.value !== deck.name && !deckBuilderSaveTimer && document.activeElement !== deckBuilderNameInput) {
      deckBuilderNameInput.value = deck.name;
    }
    if (deckBuilderOwnerInput && deckBuilderOwnerInput.value !== (deck.owner || '') && !deckBuilderSaveTimer && document.activeElement !== deckBuilderOwnerInput) {
      deckBuilderOwnerInput.value = deck.owner || '';
    }
    if (deckBuilderPowerInput && document.activeElement !== deckBuilderPowerInput) {
      const fallbackExpectedPower = deck.commander?.name ? getCommanderExpectedPower(deck.commander.name) : '';
      const displayPower = typeof deck.powerLevel === 'number' ? deck.powerLevel : fallbackExpectedPower;
      const nextValue = typeof displayPower === 'number' ? String(displayPower) : '';
      if (deckBuilderPowerInput.value !== nextValue) {
        deckBuilderPowerInput.value = nextValue;
      }
    }
    if (deckBuilderInRotationInput && document.activeElement !== deckBuilderInRotationInput) {
      const nextInRotationValue = deck.inRotation !== false;
      if (deckBuilderInRotationInput.checked !== nextInRotationValue) {
        deckBuilderInRotationInput.checked = nextInRotationValue;
      }
    }

    applyDeckBuilderAccessState(deck);
    if (!canCurrentUserEditDeck(deck)) {
      setDeckBuilderSaveStatus(getDeckReadOnlyMessage(deck), 'muted');
    }
    renderDeckBuilderValidation(deck);
    renderDeckBuilderManaCurve(deck);
    renderDeckBuilderSelection();
    renderDeckBuilderCards(deck);
    renderDeckBuilderBreakdown(deck);
    renderDeckBuilderSearchResults();
    renderDeckBuilderTokenSearchResults();
    hydrateDeckCardStats(deck);

    if (prefilledCommanderName && !deck.commander && canCurrentUserEditDeck(deck)) {
      fetchDeckCardByName(prefilledCommanderName).then((card) => {
        if (!card) return;
        const currentDeck = ensureActiveDeckBuilderRecord();
        if (!currentDeck || currentDeck.commander) return;
        const eligibleCard = isCommanderEligibleCard(card) ? card : null;
        if (!eligibleCard) {
          deckBuilderSelectedCard = card;
          persistDeckBuilderSelectedCard(card);
          renderDeckBuilderSelection();
          setDeckBuilderSaveStatus(`${card.name} pre-selected. Use Set as Commander if eligible.`, 'neutral');
          return;
        }
        persistDeckBuilderRecord({
          ...currentDeck,
          name: eligibleCard.name,
          commander: eligibleCard,
        }, `${eligibleCard.name} set as commander.`);
      }).catch(() => {});
    }
  }

  function exportDeckAsText(deck) {
    const lines = [];
    if (deck.commander) {
      lines.push('// Commander');
      lines.push(`1 ${deck.commander.name}`);
      lines.push('');
      lines.push('// Deck');
    }
    const counts = new Map();
    deck.cards.forEach((card) => {
      const key = card.name;
      counts.set(key, (counts.get(key) || 0) + 1);
    });
    [...counts.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .forEach(([name, count]) => lines.push(`${count} ${name}`));

    const tokens = Array.isArray(deck.tokens) ? deck.tokens : [];
    if (tokens.length > 0) {
      lines.push('');
      lines.push('// Tokens');
      [...tokens]
        .sort((a, b) => String(a.name).localeCompare(String(b.name)))
        .forEach((token) => lines.push(`${token.count || 1} ${token.name}`));
    }

    const maybeboard = Array.isArray(deck.maybeboard) ? deck.maybeboard : [];
    if (maybeboard.length > 0) {
      lines.push('');
      lines.push('// Maybeboard');
      [...maybeboard]
        .sort((a, b) => String(a.name).localeCompare(String(b.name)))
        .forEach((card) => lines.push(`1 ${card.name}`));
    }

    return lines.join('\n');
  }

  function buildEdhPowerLevelDeckText(deck) {
    const lines = [];
    if (deck?.commander?.name) {
      lines.push(`1 ${deck.commander.name}`);
    }

    const counts = new Map();
    (deck?.cards || []).forEach((card) => {
      if (!card?.name) {
        return;
      }
      counts.set(card.name, (counts.get(card.name) || 0) + 1);
    });

    [...counts.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .forEach(([name, count]) => lines.push(`${count} ${name}`));

    return lines.join('\n');
  }

  function encodeEdhPowerLevelDeck(deckText) {
    const rows = String(deckText || '').split(/[\r\n~]/);
    const cleaned = rows
      .map((row) => row
        .split('(')[0]
        .replace(/ *\[[^\]]*\] */g, ' ')
        .replace(/ *<[^>]*> */g, ' ')
        .trim())
      .filter(Boolean);

    return `${encodeURIComponent(cleaned.join('~')).replace(/%20/g, '+')}~Z~`;
  }

  function setDeckBuilderImportStatus(message, tone = 'muted') {
    if (!importStatus) return;
    importStatus.textContent = message;
    importStatus.className = `deck-builder-import-status status-${tone}`;
  }

  async function importDeckFromText(text) {
    const deck = ensureActiveDeckBuilderRecord({ createIfMissing: true });
    if (!deck) { setDeckBuilderImportStatus('No active deck.', 'error'); return; }

    const {
      entries,
      maybeboardEntries,
      tokenEntries,
      commanderName,
      hasMaybeboardSection,
      hasTokenSection,
    } = parseDeckTextList(text);
    const allEntries = [
      ...entries.map((entry) => ({ ...entry, target: 'deck' })),
      ...maybeboardEntries.map((entry) => ({ ...entry, target: 'maybeboard' })),
      ...tokenEntries.map((entry) => ({ ...entry, target: 'tokens' })),
    ];
    if (!allEntries.length && !commanderName) { setDeckBuilderImportStatus('No cards found in the import text.', 'error'); return; }

    const totalUnique = allEntries.length + (commanderName ? 1 : 0);
    setDeckBuilderImportStatus(`Importing 0 / ${totalUnique}...`, 'neutral');

    if (importButton) importButton.disabled = true;

    const bulkLookupNames = [];
    const seenBulkNameKeys = new Set();
    const lookupEntries = commanderName
      ? [{ name: commanderName, count: 1 }, ...allEntries]
      : allEntries;

    lookupEntries.forEach((entry) => {
      const variants = getDeckImportNameVariants(getPrimaryDeckImportLookupName(entry.name));
      variants.forEach((variant) => {
        const key = getIdentityKey(variant);
        if (!key || seenBulkNameKeys.has(key)) {
          return;
        }
        seenBulkNameKeys.add(key);
        bulkLookupNames.push(variant);
      });
    });

    let bulkCardsByQuery = new Map();
    let bulkLookupFailed = false;
    try {
      setDeckBuilderImportStatus(`Bulk loading ${totalUnique} cards...`, 'neutral');
      bulkCardsByQuery = await fetchDeckCardsByNamesBulk(bulkLookupNames);
    } catch (error) {
      bulkCardsByQuery = new Map();
      bulkLookupFailed = true;
    }

    const newCards = [];
    const importedMaybeboard = [];
    const importedTokens = [];
    let newCommander = null;
    const failed = [];

    if (commanderName) {
      try {
        const bulkCard = getBulkResolvedImportCard(commanderName, bulkCardsByQuery);
        const commanderCard = bulkCard || await fetchDeckCardForImport(commanderName);
        if (commanderCard) {
          newCommander = commanderCard;
        } else {
          failed.push(commanderName);
        }
      } catch (error) {
        failed.push(commanderName);
      }
    }

    for (let i = 0; i < allEntries.length; i++) {
      const { count, name, target } = allEntries[i];
      setDeckBuilderImportStatus(`Importing ${i + 1} / ${totalUnique} — ${name}`, 'neutral');
      try {
        const bulkCard = getBulkResolvedImportCard(name, bulkCardsByQuery);
        const card = bulkCard || await fetchDeckCardForImport(name);
        if (!card) { failed.push(name); continue; }

        if (card.isToken || target === 'tokens') {
          importedTokens.push({ ...card, id: generateId(), count });
        } else if (target === 'maybeboard') {
          importedMaybeboard.push({ ...card, id: generateId() });
        } else {
          for (let c = 0; c < count; c++) {
            newCards.push({ ...card, id: generateId() });
          }
        }
      } catch (error) {
        failed.push(name);
      }

      if (i < allEntries.length - 1 && bulkLookupFailed) {
        await new Promise((resolve) => window.setTimeout(resolve, 30));
      }
    }

    if (importButton) importButton.disabled = false;

    const importedTokenCount = importedTokens.reduce((sum, token) => sum + (Number.isFinite(Number(token.count)) ? Math.max(1, Number(token.count)) : 1), 0);
    const nextDeck = applyDeckBuilderDraftMeta(deck);
    const importOnlyAcceptedCards = [];
    const importOnlyNameCounts = new Map();
    const skippedImportDuplicateNames = new Set();

    if (newCommander?.name) {
      const commanderKey = getIdentityKey(newCommander.name);
      if (commanderKey) {
        importOnlyNameCounts.set(commanderKey, 1);
      }
    }

    newCards.forEach((card) => {
      const key = getIdentityKey(card?.name);
      if (!key) {
        importOnlyAcceptedCards.push(card);
        return;
      }

      if (!allowsMultipleCopies(card) && (importOnlyNameCounts.get(key) || 0) > 0) {
        skippedImportDuplicateNames.add(card.name);
        return;
      }

      importOnlyAcceptedCards.push(card);
      importOnlyNameCounts.set(key, (importOnlyNameCounts.get(key) || 0) + 1);
    });

    const importedDeckCardTotal = importOnlyAcceptedCards.length + (newCommander ? 1 : 0);
    const shouldReplaceDeckContents = importedDeckCardTotal === 100;
    const baseCommander = shouldReplaceDeckContents ? (newCommander || null) : (nextDeck.commander ?? newCommander);
    const baseSecondCommander = shouldReplaceDeckContents ? null : (nextDeck.secondCommander || null);
    const existingCards = shouldReplaceDeckContents ? [] : (Array.isArray(nextDeck.cards) ? nextDeck.cards : []);
    const acceptedNewCards = [];
    const overflowDeckCards = [];
    const skippedDuplicateNames = new Set(skippedImportDuplicateNames);
    const mainDeckNameCounts = new Map();
    const commanderCount = (baseCommander ? 1 : 0) + (baseSecondCommander ? 1 : 0);
    let remainingDeckSlots = shouldReplaceDeckContents
      ? Math.max(0, 100 - commanderCount)
      : Math.max(0, 100 - (existingCards.length + commanderCount));

    if (baseCommander?.name) {
      const commanderKey = getIdentityKey(baseCommander.name);
      if (commanderKey) {
        mainDeckNameCounts.set(commanderKey, 1);
      }
    }

    if (baseSecondCommander?.name) {
      const secondCommanderKey = getIdentityKey(baseSecondCommander.name);
      if (secondCommanderKey) {
        mainDeckNameCounts.set(secondCommanderKey, 1);
      }
    }

    existingCards.forEach((card) => {
      const key = getIdentityKey(card?.name);
      if (!key) {
        return;
      }
      mainDeckNameCounts.set(key, (mainDeckNameCounts.get(key) || 0) + 1);
    });

    importOnlyAcceptedCards.forEach((card) => {
      const key = getIdentityKey(card?.name);
      if (!key) {
        if (remainingDeckSlots > 0) {
          acceptedNewCards.push(card);
          remainingDeckSlots -= 1;
        } else {
          overflowDeckCards.push(card);
        }
        return;
      }

      if (!allowsMultipleCopies(card) && (mainDeckNameCounts.get(key) || 0) > 0) {
        skippedDuplicateNames.add(card.name);
        return;
      }

      if (remainingDeckSlots <= 0) {
        overflowDeckCards.push(card);
        return;
      }

      acceptedNewCards.push(card);
      mainDeckNameCounts.set(key, (mainDeckNameCounts.get(key) || 0) + 1);
      remainingDeckSlots -= 1;
    });

    const existingMaybeboard = Array.isArray(nextDeck.maybeboard) ? nextDeck.maybeboard : [];
    const maybeboardNameSet = new Set(
      existingMaybeboard
        .map((card) => getIdentityKey(card?.name))
        .filter(Boolean)
    );
    const acceptedMaybeboardCards = [];
    const skippedMaybeboardNames = new Set();
    const partialOverflowNames = new Set();

    overflowDeckCards.forEach((card) => {
      const key = getIdentityKey(card?.name);
      if (!key) {
        acceptedMaybeboardCards.push(card);
        return;
      }

      if (maybeboardNameSet.has(key) || mainDeckNameCounts.has(key)) {
        skippedMaybeboardNames.add(card.name);
        return;
      }

      acceptedMaybeboardCards.push(card);
      maybeboardNameSet.add(key);
      partialOverflowNames.add(card.name);
    });

    importedMaybeboard.forEach((card) => {
      const key = getIdentityKey(card?.name);
      if (!key) {
        acceptedMaybeboardCards.push(card);
        return;
      }

      if (maybeboardNameSet.has(key) || mainDeckNameCounts.has(key)) {
        skippedMaybeboardNames.add(card.name);
        return;
      }

      acceptedMaybeboardCards.push(card);
      maybeboardNameSet.add(key);
    });

    const mergedMaybeboard = [
      ...existingMaybeboard,
      ...acceptedMaybeboardCards,
    ];

    const existingTokens = Array.isArray(nextDeck.tokens) ? nextDeck.tokens : [];
    const mergedTokens = mergeDeckBuilderTokenCards([...existingTokens, ...importedTokens]);

    const importedCardCount = acceptedNewCards.length + acceptedMaybeboardCards.length + importedTokenCount + (newCommander ? 1 : 0);
    const statusParts = [];
    if (importedCardCount) {
      statusParts.push(`Imported ${importedCardCount} card(s).`);
    }
    if (shouldReplaceDeckContents) {
      statusParts.push('Detected a 100-card import and replaced the current deck contents.');
    } else if (partialOverflowNames.size) {
      const overflowNames = [...partialOverflowNames];
      statusParts.push(`Partial import filled the deck to 100 cards and moved ${overflowNames.length} extra card(s) to the maybeboard: ${overflowNames.slice(0, 3).join(', ')}${overflowNames.length > 3 ? '…' : ''}.`);
    }
    if (failed.length) {
      statusParts.push(`Could not find ${failed.length} card(s): ${failed.slice(0, 3).join(', ')}${failed.length > 3 ? '…' : ''}.`);
    }
    if (skippedDuplicateNames.size) {
      const skippedNames = [...skippedDuplicateNames];
      statusParts.push(`Skipped ${skippedNames.length} duplicate singleton card(s): ${skippedNames.slice(0, 3).join(', ')}${skippedNames.length > 3 ? '…' : ''}.`);
    }
    if (skippedMaybeboardNames.size) {
      const skippedNames = [...skippedMaybeboardNames];
      statusParts.push(`Skipped ${skippedNames.length} maybeboard duplicate card(s): ${skippedNames.slice(0, 3).join(', ')}${skippedNames.length > 3 ? '…' : ''}.`);
    }
    const statusMsg = statusParts.join(' ') || 'Import finished.';
    const tone = failed.length || skippedDuplicateNames.size || skippedMaybeboardNames.size ? 'error' : 'success';

    persistDeckBuilderRecord({
      ...nextDeck,
      commander: baseCommander,
      secondCommander: baseSecondCommander,
      cards: [...existingCards, ...acceptedNewCards],
      maybeboard: mergedMaybeboard,
      tokens: mergedTokens,
    }, statusMsg, tone);
    setDeckBuilderImportStatus(statusMsg, tone);
  }

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