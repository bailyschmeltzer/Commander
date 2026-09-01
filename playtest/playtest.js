(function () {
  const DECKS_STORAGE_KEY = 'commanderDeckRecords';
  const PLAYTEST_SESSION_STORAGE_KEY = 'commanderPlaytestSessionV2';
  const DEFAULT_CANVAS_WIDTH = 1400;
  const DEFAULT_CANVAS_HEIGHT = 760;
  const HISTORY_LIMIT = 120;

  const deckSelect = document.getElementById('playtest-deck-select');
  const shuffleSeedInput = document.getElementById('playtest-shuffle-seed');
  const loadDeckButton = document.getElementById('playtest-load-deck');
  const shuffleButton = document.getElementById('playtest-shuffle');
  const drawOneButton = document.getElementById('playtest-draw-1');
  const drawSevenButton = document.getElementById('playtest-draw-7');
  const openingHandButton = document.getElementById('playtest-opening-hand');
  const mulliganButton = document.getElementById('playtest-mulligan');
  const undoButton = document.getElementById('playtest-undo');
  const untapAllButton = document.getElementById('playtest-untap-all');
  const touchMoveModeInput = document.getElementById('playtest-touch-move-mode');
  const mulliganStatusEl = document.getElementById('playtest-mulligan-status');
  const resetSessionButton = document.getElementById('playtest-reset-session');
  const statusEl = document.getElementById('playtest-status');
  const exportSessionButton = document.getElementById('playtest-export-session');
  const copySessionButton = document.getElementById('playtest-copy-session');
  const importSessionButton = document.getElementById('playtest-import-session');
  const clearSessionJsonButton = document.getElementById('playtest-clear-session-json');
  const sessionJsonInput = document.getElementById('playtest-session-json');

  const lifeInput = document.getElementById('playtest-life');
  const lifeMinusButton = document.getElementById('playtest-life-minus');
  const lifePlusButton = document.getElementById('playtest-life-plus');

  const tokenNameInput = document.getElementById('playtest-token-name');
  const tokenCountInput = document.getElementById('playtest-token-count');
  const tokenImageInput = document.getElementById('playtest-token-image');
  const createTokenButton = document.getElementById('playtest-create-token');

  const selectedNameEl = document.getElementById('playtest-selected-name');
  const toggleTapButton = document.getElementById('playtest-toggle-tap');
  const toggleFaceButton = document.getElementById('playtest-toggle-face');
  const createTokenCopyButton = document.getElementById('playtest-create-token-copy');
  const counterTypeInput = document.getElementById('playtest-counter-type');
  const counterCustomInput = document.getElementById('playtest-counter-custom');
  const counterAddButton = document.getElementById('playtest-counter-add');
  const counterRemoveButton = document.getElementById('playtest-counter-remove');
  const commanderCastButton = document.getElementById('playtest-commander-cast');
  const commanderResetButton = document.getElementById('playtest-commander-reset');
  const commanderTaxEl = document.getElementById('playtest-commander-tax');
  const moveZoneButtons = Array.from(document.querySelectorAll('[data-move-zone]'));

  const zoneButtons = Array.from(document.querySelectorAll('.playtest-zone[data-zone]'));
  const zoneCountLibrary = document.getElementById('zone-count-library');
  const zoneCountGraveyard = document.getElementById('zone-count-graveyard');
  const zoneCountExile = document.getElementById('zone-count-exile');
  const zoneCountCommand = document.getElementById('zone-count-command');
  const zonePileLibrary = document.getElementById('zone-pile-library');
  const zonePileGraveyard = document.getElementById('zone-pile-graveyard');
  const zonePileExile = document.getElementById('zone-pile-exile');
  const zonePileCommand = document.getElementById('zone-pile-command');
  const zoneToTopButton = document.getElementById('playtest-zone-to-top');
  const zoneToBottomButton = document.getElementById('playtest-zone-to-bottom');

  const battlefieldDrop = document.getElementById('playtest-battlefield-drop');
  const battlefieldCanvas = document.getElementById('playtest-battlefield-canvas');
  const battlefieldCardsEl = document.getElementById('playtest-battlefield-cards');
  const handEl = document.getElementById('playtest-hand');
  const zoneCardsEl = document.getElementById('playtest-zone-cards');
  const zoneSearchInput = document.getElementById('playtest-zone-search');
  const libraryLookCountInput = document.getElementById('playtest-library-look-count');
  const libraryLookButton = document.getElementById('playtest-library-look');
  const libraryScryButton = document.getElementById('playtest-library-scry');
  const librarySurveilButton = document.getElementById('playtest-library-surveil');
  const inspectedZoneLabelEl = document.getElementById('playtest-inspected-zone-label');
  const debugOutputEl = document.getElementById('playtest-debug-output');
  const actionLogEl = document.getElementById('playtest-action-log');

  const imageModal = document.getElementById('playtest-image-modal');
  const imageModalClose = document.getElementById('playtest-image-close');
  const imageModalImage = document.getElementById('playtest-image');

  const zoneNames = ['library', 'hand', 'battlefield', 'graveyard', 'exile', 'command'];

  const state = {
    deckCatalog: [],
    deckId: '',
    deckName: '',
    life: 40,
    selectedCardId: '',
    inspectedZone: 'graveyard',
    inspectedZoneSearch: '',
    inspectedZoneOpen: false,
    libraryPreviewCount: 0,
    libraryPreviewMode: '',
    libraryPreviewIds: [],
    mulliganCount: 0,
    openingHandSize: 7,
    shuffleSeed: '',
    touchMoveMode: false,
    commanderTaxByName: {},
    actionLog: [],
    zones: {
      library: [],
      hand: [],
      battlefield: [],
      graveyard: [],
      exile: [],
      command: [],
    },
  };

  const history = {
    undo: [],
    redo: [],
  };

  const ACTION_LOG_LIMIT = 100;

  function makeId() {
    return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 9)}`;
  }

  function setStatus(message) {
    statusEl.textContent = String(message || '');
  }

  function parseJsonSafe(raw, fallback) {
    try {
      return JSON.parse(raw);
    } catch {
      return fallback;
    }
  }

  function snapshotsEqual(first, second) {
    return JSON.stringify(first) === JSON.stringify(second);
  }

  function appendActionLog(action, details) {
    const entry = {
      id: makeId(),
      at: new Date().toISOString(),
      action: String(action || 'action'),
      details: String(details || '').trim(),
    };
    state.actionLog.push(entry);
    if (state.actionLog.length > ACTION_LOG_LIMIT) {
      state.actionLog = state.actionLog.slice(-ACTION_LOG_LIMIT);
    }
    renderActionLog();
  }

  function cloneRuntimeState(value) {
    return parseJsonSafe(JSON.stringify(value), null);
  }

  function getRuntimeSnapshot() {
    return {
      deckId: state.deckId,
      deckName: state.deckName,
      life: state.life,
      selectedCardId: state.selectedCardId,
      inspectedZone: state.inspectedZone,
      inspectedZoneSearch: state.inspectedZoneSearch,
      inspectedZoneOpen: state.inspectedZoneOpen,
      libraryPreviewCount: state.libraryPreviewCount,
      libraryPreviewMode: state.libraryPreviewMode,
      libraryPreviewIds: [...state.libraryPreviewIds],
      mulliganCount: state.mulliganCount,
      openingHandSize: state.openingHandSize,
      shuffleSeed: state.shuffleSeed,
      touchMoveMode: state.touchMoveMode,
      commanderTaxByName: { ...state.commanderTaxByName },
      actionLog: cloneRuntimeState(state.actionLog),
      zones: cloneRuntimeState(state.zones),
    };
  }

  function applyRuntimeSnapshot(snapshot) {
    if (!snapshot || typeof snapshot !== 'object') {
      return;
    }

    state.deckId = String(snapshot.deckId || '');
    state.deckName = String(snapshot.deckName || '');
    state.life = Number.isFinite(Number(snapshot.life)) ? Number(snapshot.life) : 40;
    state.selectedCardId = String(snapshot.selectedCardId || '');
    state.inspectedZone = zoneNames.includes(snapshot.inspectedZone) ? snapshot.inspectedZone : 'graveyard';
    state.inspectedZoneSearch = String(snapshot.inspectedZoneSearch || '');
    state.inspectedZoneOpen = Boolean(snapshot.inspectedZoneOpen);
    state.libraryPreviewCount = Number.isFinite(Number(snapshot.libraryPreviewCount)) ? Math.max(0, Number(snapshot.libraryPreviewCount)) : 0;
    state.libraryPreviewMode = String(snapshot.libraryPreviewMode || '');
    state.libraryPreviewIds = Array.isArray(snapshot.libraryPreviewIds)
      ? snapshot.libraryPreviewIds.map((instanceId) => String(instanceId || '')).filter(Boolean)
      : [];
    state.mulliganCount = Number.isFinite(Number(snapshot.mulliganCount)) ? Math.max(0, Number(snapshot.mulliganCount)) : 0;
    state.openingHandSize = Number.isFinite(Number(snapshot.openingHandSize)) ? Math.max(0, Number(snapshot.openingHandSize)) : 7;
    state.shuffleSeed = String(snapshot.shuffleSeed || '');
    state.touchMoveMode = Boolean(snapshot.touchMoveMode);
    state.commanderTaxByName = snapshot.commanderTaxByName && typeof snapshot.commanderTaxByName === 'object'
      ? { ...snapshot.commanderTaxByName }
      : {};
    state.actionLog = Array.isArray(snapshot.actionLog) ? snapshot.actionLog.slice(-ACTION_LOG_LIMIT) : [];

    zoneNames.forEach((zone) => {
      const cards = Array.isArray(snapshot?.zones?.[zone]) ? snapshot.zones[zone] : [];
      state.zones[zone] = cards.map(normalizeCardInstance).filter(Boolean);
    });
  }

  function saveSession() {
    localStorage.setItem(PLAYTEST_SESSION_STORAGE_KEY, JSON.stringify(getRuntimeSnapshot()));
  }

  function clearSession() {
    localStorage.removeItem(PLAYTEST_SESSION_STORAGE_KEY);
  }

  function loadSession() {
    const raw = localStorage.getItem(PLAYTEST_SESSION_STORAGE_KEY);
    if (!raw) {
      return false;
    }

    const data = parseJsonSafe(raw, null);
    if (!data || typeof data !== 'object' || !data.zones) {
      return false;
    }

    applyRuntimeSnapshot(data);
    return true;
  }

  function pushUndoSnapshot(snapshot) {
    history.undo.push(snapshot || getRuntimeSnapshot());
    if (history.undo.length > HISTORY_LIMIT) {
      history.undo.shift();
    }
    history.redo = [];
  }

  function commitMutation(callback, successMessage, actionLabel, options) {
    const beforeSnapshot = getRuntimeSnapshot();
    pushUndoSnapshot(beforeSnapshot);
    const mutationResult = callback();
    const afterSnapshot = getRuntimeSnapshot();
    if (snapshotsEqual(beforeSnapshot, afterSnapshot)) {
      history.undo.pop();
      if (options?.noopMessage) {
        setStatus(options.noopMessage);
      }
      return false;
    }

    const resolvedMessage = typeof successMessage === 'function'
      ? successMessage(mutationResult)
      : successMessage;
    if (actionLabel || resolvedMessage) {
      appendActionLog(actionLabel || 'Mutation', resolvedMessage || '');
    }
    if (resolvedMessage) {
      setStatus(resolvedMessage);
    }
    renderAll();
    saveSession();
    return true;
  }

  function undoMutation() {
    if (!history.undo.length) {
      setStatus('Nothing to undo.');
      return;
    }

    history.redo.push(getRuntimeSnapshot());
    const snapshot = history.undo.pop();
    applyRuntimeSnapshot(snapshot);
    setStatus('Undid last action.');
    appendActionLog('Undo', 'Undid last action.');
    renderAll();
    saveSession();
  }

  function redoMutation() {
    if (!history.redo.length) {
      setStatus('Nothing to redo.');
      return;
    }

    history.undo.push(getRuntimeSnapshot());
    const snapshot = history.redo.pop();
    applyRuntimeSnapshot(snapshot);
    setStatus('Redid action.');
    appendActionLog('Redo', 'Redid action.');
    renderAll();
    saveSession();
  }

  function normalizeDeckCardEntry(card) {
    if (!card || typeof card !== 'object') {
      return null;
    }

    const name = String(card.name || '').trim();
    if (!name) {
      return null;
    }

    return {
      name,
      imageUri: String(card.imageUri || '').trim(),
      imageLargeUri: String(card.imageLargeUri || '').trim(),
      typeLine: String(card.typeLine || '').trim(),
      scryfallUri: String(card.scryfallUri || '').trim(),
      count: Number.isFinite(Number(card.count)) ? Math.max(1, Number(card.count)) : 1,
      isToken: Boolean(card.isToken),
    };
  }

  function normalizeDeckRecord(deck) {
    if (!deck || typeof deck !== 'object') {
      return null;
    }

    const name = String(deck.name || '').trim() || 'Untitled Deck';
    const id = String(deck.id || makeId()).trim();
    const commander = normalizeDeckCardEntry(deck.commander);
    const secondCommander = normalizeDeckCardEntry(deck.secondCommander);
    const cards = Array.isArray(deck.cards) ? deck.cards.map(normalizeDeckCardEntry).filter(Boolean) : [];

    if (!cards.length && !commander && !secondCommander) {
      return null;
    }

    return {
      id,
      name,
      commander,
      secondCommander,
      cards,
    };
  }

  function loadDeckCatalog() {
    const mergedDecks = [];

    const addDecks = (candidate) => {
      if (Array.isArray(candidate)) {
        mergedDecks.push(...candidate);
        return;
      }
      if (candidate && typeof candidate === 'object') {
        if (Array.isArray(candidate.decks)) {
          mergedDecks.push(...candidate.decks);
        }
        if (candidate.state && typeof candidate.state === 'object' && Array.isArray(candidate.state.decks)) {
          mergedDecks.push(...candidate.state.decks);
        }
      }
    };

    addDecks(parseJsonSafe(localStorage.getItem(DECKS_STORAGE_KEY) || '[]', []));
    addDecks(parseJsonSafe(localStorage.getItem('commanderTrackerGames') || '{}', {}));

    // Cloud/bootstrap payloads may be persisted under a different app key.
    for (let index = 0; index < localStorage.length; index += 1) {
      const key = localStorage.key(index) || '';
      if (!/deck|state|cloud|sync/i.test(key)) {
        continue;
      }
      addDecks(parseJsonSafe(localStorage.getItem(key) || '', null));
    }

    const decksById = new Map();
    updateDeckCatalog(mergedDecks);
  }

  function updateDeckCatalog(candidates) {
    const decksById = new Map();

    candidates.forEach((deck) => {
      const normalized = normalizeDeckRecord(deck);
      if (normalized) {
        decksById.set(normalized.id, normalized);
      }
    });

    state.deckCatalog = [...decksById.values()];
  }

  async function refreshDeckCatalogFromCloud() {
    const user = String(localStorage.getItem('commanderTrackerSyncUser') || '').trim();
    const token = String(localStorage.getItem('commanderTrackerSyncToken') || '').trim();
    if (!user || !token) {
      return;
    }

    try {
      const response = await fetch('/api/state', {
        method: 'GET',
        headers: {
          'X-User-Name': user,
          'X-Pod-Token': token,
        },
        cache: 'no-store',
        credentials: 'same-origin',
      });
      if (!response.ok) {
        return;
      }

      const payload = await response.json();
      const cloudState = payload?.state && typeof payload.state === 'object' ? payload.state : payload;
      if (!Array.isArray(cloudState?.decks)) {
        return;
      }

      updateDeckCatalog([...state.deckCatalog, ...cloudState.decks]);
      renderDeckOptions();
    } catch {
      // Local decks remain usable when cloud refresh is unavailable.
    }
  }

  function renderDeckOptions() {
    deckSelect.innerHTML = '';
    if (!state.deckCatalog.length) {
      const emptyOption = document.createElement('option');
      emptyOption.value = '';
      emptyOption.textContent = 'No saved decks found on this device';
      deckSelect.appendChild(emptyOption);
      return;
    }

    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Choose a saved deck';
    deckSelect.appendChild(defaultOption);

    state.deckCatalog.forEach((deck) => {
      const option = document.createElement('option');
      option.value = deck.id;
      option.textContent = deck.name;
      deckSelect.appendChild(option);
    });

    if (state.deckId) {
      deckSelect.value = state.deckId;
    }
  }

  function expandDeckCards(deck) {
    const cards = [];

    deck.cards.forEach((card) => {
      for (let i = 0; i < Math.max(1, card.count); i += 1) {
        cards.push(makeCardInstance(card, 'library', { isCommanderCard: false }));
      }
    });

    if (deck.commander) {
      cards.push(makeCardInstance(deck.commander, 'command', { isCommanderCard: true }));
    }

    if (deck.secondCommander) {
      cards.push(makeCardInstance(deck.secondCommander, 'command', { isCommanderCard: true }));
    }

    return cards;
  }

  function makeCardInstance(base, zone, options) {
    const spawn = getSpawnCoordinates();
    return {
      instanceId: makeId(),
      name: String(base.name || '').trim() || 'Card',
      imageUri: String(base.imageLargeUri || base.imageUri || '').trim(),
      imageSmallUri: String(base.imageUri || '').trim(),
      typeLine: String(base.typeLine || '').trim(),
      scryfallUri: String(base.scryfallUri || '').trim(),
      isToken: Boolean(base.isToken),
      isCommanderCard: Boolean(options?.isCommanderCard),
      tapped: false,
      faceDown: false,
      counters: {},
      zone,
      x: spawn.x,
      y: spawn.y,
    };
  }

  function normalizeCardInstance(card) {
    if (!card || typeof card !== 'object') {
      return null;
    }

    const name = String(card.name || '').trim();
    if (!name) {
      return null;
    }

    const zone = zoneNames.includes(card.zone) ? card.zone : 'library';

    return {
      instanceId: String(card.instanceId || makeId()),
      name,
      imageUri: String(card.imageUri || '').trim(),
      imageSmallUri: String(card.imageSmallUri || '').trim(),
      typeLine: String(card.typeLine || '').trim(),
      scryfallUri: String(card.scryfallUri || '').trim(),
      isToken: Boolean(card.isToken),
      isCommanderCard: Boolean(card.isCommanderCard),
      tapped: Boolean(card.tapped),
      faceDown: Boolean(card.faceDown),
      counters: card.counters && typeof card.counters === 'object' ? { ...card.counters } : {},
      zone,
      x: Number.isFinite(Number(card.x)) ? Number(card.x) : getSpawnCoordinates().x,
      y: Number.isFinite(Number(card.y)) ? Number(card.y) : getSpawnCoordinates().y,
    };
  }

  function createSeededRandom(seedText) {
    let hash = 2166136261;
    const seed = String(seedText || '');
    for (let i = 0; i < seed.length; i += 1) {
      hash ^= seed.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }

    return function random() {
      hash += 0x6D2B79F5;
      let t = hash;
      t = Math.imul(t ^ (t >>> 15), 1 | t);
      t ^= t + Math.imul(t ^ (t >>> 7), 61 | t);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  function shuffleArray(values, seedText) {
    const random = String(seedText || '').trim() ? createSeededRandom(seedText) : Math.random;
    const arr = [...values];
    for (let i = arr.length - 1; i > 0; i -= 1) {
      const j = Math.floor(random() * (i + 1));
      const temp = arr[i];
      arr[i] = arr[j];
      arr[j] = temp;
    }
    return arr;
  }

  function getSpawnCoordinates() {
    const x = 60 + Math.random() * (DEFAULT_CANVAS_WIDTH - 210);
    const y = 40 + Math.random() * (DEFAULT_CANVAS_HEIGHT - 230);
    return { x, y };
  }

  function getBattlefieldClientMetrics() {
    const rect = battlefieldDrop.getBoundingClientRect();
    return {
      width: rect.width,
      height: rect.height,
      left: rect.left,
      top: rect.top,
    };
  }

  function findCardLocation(instanceId) {
    for (let i = 0; i < zoneNames.length; i += 1) {
      const zone = zoneNames[i];
      const index = state.zones[zone].findIndex((card) => card.instanceId === instanceId);
      if (index >= 0) {
        return { zone, index, card: state.zones[zone][index] };
      }
    }
    return null;
  }

  function removeCardFromZone(instanceId) {
    const located = findCardLocation(instanceId);
    if (!located) {
      return null;
    }

    const [card] = state.zones[located.zone].splice(located.index, 1);
    return card || null;
  }

  function moveCardToZone(instanceId, nextZone, options) {
    if (!zoneNames.includes(nextZone)) {
      return false;
    }

    const located = findCardLocation(instanceId);
    if (!located?.card) {
      return false;
    }

    if (located.zone === nextZone) {
      if (nextZone !== 'battlefield') {
        return false;
      }

      const nextX = Number.isFinite(Number(options?.x)) ? Number(options.x) : located.card.x;
      const nextY = Number.isFinite(Number(options?.y)) ? Number(options.y) : located.card.y;
      if (Math.abs(Number(located.card.x || 0) - nextX) < 1 && Math.abs(Number(located.card.y || 0) - nextY) < 1) {
        return false;
      }
      located.card.x = nextX;
      located.card.y = nextY;
      return true;
    }

    const card = removeCardFromZone(instanceId);
    if (!card) {
      return false;
    }

    state.libraryPreviewIds = state.libraryPreviewIds.filter((previewId) => previewId !== instanceId);
    card.zone = nextZone;

    if (nextZone === 'battlefield') {
      card.x = Number.isFinite(Number(options?.x)) ? Number(options.x) : card.x;
      card.y = Number.isFinite(Number(options?.y)) ? Number(options.y) : card.y;
    }

    state.zones[nextZone].push(card);
    return true;
  }

  function drawCards(count) {
    const amount = Math.max(1, Number(count) || 1);
    commitMutation(() => {
      let drew = 0;
      for (let i = 0; i < amount; i += 1) {
        const card = state.zones.library.pop();
        if (!card) {
          break;
        }
        card.zone = 'hand';
        card.faceDown = false;
        state.zones.hand.push(card);
        drew += 1;
      }
      return drew;
    }, (drew) => drew ? `Drew ${drew} card${drew === 1 ? '' : 's'}.` : 'Library is empty.', amount === 7 ? 'Draw Seven' : 'Draw Card', {
      noopMessage: 'Library is empty.',
    });
  }

  function previewLibraryTop(count, mode) {
    const amount = Math.max(1, Math.min(state.zones.library.length, Number(count) || 1));
    if (!state.zones.library.length) {
      setStatus('Library is empty.');
      return;
    }

    state.inspectedZone = 'library';
    state.inspectedZoneOpen = true;
    state.inspectedZoneSearch = '';
    state.libraryPreviewCount = amount;
    state.libraryPreviewMode = mode;
    state.libraryPreviewIds = state.zones.library.slice(-amount).map((card) => card.instanceId);
    renderAll();
    saveSession();
    const placementHint = mode === 'Scrying'
      ? ' Select a card, then choose Top or Bottom. Place cards in the order you want them to resolve, with the last Top choice becoming the top card.'
      : mode === 'Surveilling'
        ? ' Select a card, then choose Top or Graveyard. Place cards in the order you want them to resolve, with the last Top choice becoming the top card.'
        : '';
    setStatus(`${mode} the top ${amount} card${amount === 1 ? '' : 's'} of your library.${placementHint}`);
  }

  function initializeOpeningHand(size) {
    if (!state.deckId) {
      setStatus('Load a deck first.');
      return;
    }

    const nextSize = Math.max(0, Number(size) || 0);
    commitMutation(() => {
      let drew = 0;
      const cardsToReturn = [
        ...state.zones.hand,
        ...state.zones.battlefield,
        ...state.zones.graveyard,
        ...state.zones.exile,
      ].map((card) => ({ ...card, tapped: false, faceDown: false, counters: {}, zone: 'library' }));

      state.zones.hand = [];
      state.zones.battlefield = [];
      state.zones.graveyard = [];
      state.zones.exile = [];
      state.zones.library = shuffleArray([...state.zones.library, ...cardsToReturn], state.shuffleSeed);

      for (let i = 0; i < nextSize; i += 1) {
        const top = state.zones.library.pop();
        if (!top) {
          break;
        }
        top.zone = 'hand';
        top.faceDown = false;
        state.zones.hand.push(top);
        drew += 1;
      }

      state.openingHandSize = nextSize;
      state.selectedCardId = '';
      return drew;
    }, (drew) => `Prepared opening hand of ${nextSize}. Drew ${drew} card${drew === 1 ? '' : 's'}.`, 'Opening Hand');
  }

  function runMulligan() {
    const nextMulligan = state.mulliganCount + 1;
    const nextHandSize = nextMulligan <= 1 ? 7 : Math.max(0, 8 - nextMulligan);
    commitMutation(() => {
      state.mulliganCount = nextMulligan;
    }, `Mulligan ${nextMulligan}. ${nextMulligan === 1 ? 'First mulligan is free.' : `Next opening hand size: ${nextHandSize}.`}`, 'Mulligan');
    initializeOpeningHand(nextHandSize);
  }

  function loadDeckIntoSession(deckId) {
    const deck = state.deckCatalog.find((entry) => entry.id === deckId);
    if (!deck) {
      setStatus('Choose a valid saved deck first.');
      return;
    }

    const expandedCards = expandDeckCards(deck);
    const libraryCards = expandedCards.filter((card) => card.zone === 'library');
    const commandCards = expandedCards.filter((card) => card.zone === 'command');

    commitMutation(() => {
      state.deckId = deck.id;
      state.deckName = deck.name;
      state.life = 40;
      state.selectedCardId = '';
      state.inspectedZone = 'graveyard';
      state.mulliganCount = 0;
      state.openingHandSize = 7;
      state.shuffleSeed = String(shuffleSeedInput?.value || '').trim();
      state.commanderTaxByName = {};
      state.zones.library = shuffleArray(libraryCards, state.shuffleSeed);
      state.zones.hand = [];
      state.zones.battlefield = [];
      state.zones.graveyard = [];
      state.zones.exile = [];
      state.zones.command = commandCards;
    }, `Loaded ${deck.name}. Library ready with ${libraryCards.length} cards.${state.shuffleSeed ? ` Seed: ${state.shuffleSeed}.` : ''}`, 'Load Deck');
  }

  function adjustLife(delta) {
    commitMutation(() => {
      const parsed = Number(state.life);
      const nextLife = Number.isFinite(parsed) ? parsed + delta : 40 + delta;
      state.life = nextLife;
      return nextLife;
    }, (nextLife) => `Life set to ${nextLife}.`, 'Adjust Life');
  }

  function setLifeFromInput() {
    const parsed = Number(lifeInput.value);
    if (!Number.isFinite(parsed)) {
      lifeInput.value = String(state.life);
      return;
    }

    commitMutation(() => {
      state.life = parsed;
    }, `Life set to ${parsed}.`, 'Set Life');
  }

  function getSelectedCard() {
    if (!state.selectedCardId) {
      return null;
    }
    const located = findCardLocation(state.selectedCardId);
    return located ? located.card : null;
  }

  function isCommanderLikeCard(card) {
    if (!card) {
      return false;
    }
    return Boolean(card.isCommanderCard);
  }

  function getCommanderTaxKey(card) {
    return String(card?.instanceId || '').trim();
  }

  function updateSelectedControls() {
    const card = getSelectedCard();
    const hasSelection = Boolean(card);

    selectedNameEl.textContent = hasSelection
      ? `${card.name}${card.isToken ? ' (Token)' : ''}`
      : 'No card selected.';

    toggleTapButton.disabled = !hasSelection;
    toggleFaceButton.disabled = !hasSelection;
    createTokenCopyButton.disabled = !hasSelection;
    counterAddButton.disabled = !hasSelection;
    counterRemoveButton.disabled = !hasSelection;
    moveZoneButtons.forEach((button) => {
      button.disabled = !hasSelection;
    });

    const selectedInInspectedZone = hasSelection
      && findCardLocation(card.instanceId)?.zone === state.inspectedZone
      && state.inspectedZone !== 'battlefield';
    zoneToTopButton.disabled = !selectedInInspectedZone;
    zoneToBottomButton.disabled = !selectedInInspectedZone;

    const previewMode = state.inspectedZone === 'library' ? state.libraryPreviewMode : '';
    const isScrying = previewMode === 'Scrying';
    const isSurveilling = previewMode === 'Surveilling';
    zoneToTopButton.textContent = isScrying || isSurveilling ? 'Put on Top' : 'Top';
    zoneToBottomButton.textContent = isScrying ? 'Put on Bottom' : isSurveilling ? 'Put in Graveyard' : 'Bottom';
    zoneToTopButton.title = isScrying || isSurveilling
      ? 'Place the selected card on top of your library'
      : 'Move the selected card to the top of this zone';
    zoneToBottomButton.title = isScrying
      ? 'Place the selected scryed card on the bottom of your library'
      : isSurveilling
        ? 'Put the selected surveilled card into your graveyard'
      : 'Move the selected card to the bottom of this zone';

    const commanderCard = hasSelection && isCommanderLikeCard(card);
    commanderCastButton.disabled = !commanderCard;
    commanderResetButton.disabled = !commanderCard;

    if (commanderCard) {
      const key = getCommanderTaxKey(card);
      const castCount = Number(state.commanderTaxByName[key] || 0);
      const tax = castCount * 2;
      commanderTaxEl.textContent = `${card.name}: cast ${castCount} time${castCount === 1 ? '' : 's'}, tax +${tax}.`;
    } else {
      commanderTaxEl.textContent = 'Commander Tax: n/a';
    }
  }

  function updateMulliganStatus() {
    mulliganStatusEl.textContent = `Mulligans: ${state.mulliganCount} · Opening: ${state.openingHandSize}`;
  }

  function updateHistoryButtons() {
    undoButton.disabled = !history.undo.length;
  }

  function untapAllBattlefieldCards() {
    const tappedCards = state.zones.battlefield.filter((card) => card.tapped);
    if (!tappedCards.length) {
      setStatus('All battlefield cards are already untapped.');
      return;
    }

    commitMutation(() => {
      tappedCards.forEach((card) => {
        card.tapped = false;
      });
    }, `Untapped ${tappedCards.length} battlefield card${tappedCards.length === 1 ? '' : 's'}.`, 'Untap All');
  }

  function toggleSelectedTapState(setUntappedOnly) {
    const card = getSelectedCard();
    if (!card) {
      return;
    }

    commitMutation(() => {
      card.tapped = setUntappedOnly ? false : !card.tapped;
      return card.tapped;
    }, (isTapped) => isTapped ? `${card.name} tapped.` : `${card.name} untapped.`, 'Tap State');
  }

  function toggleSelectedFaceState() {
    const card = getSelectedCard();
    if (!card) {
      return;
    }

    commitMutation(() => {
      card.faceDown = !card.faceDown;
      return card.faceDown;
    }, (faceDown) => faceDown ? `${card.name} turned face down.` : `${card.name} turned face up.`, 'Face State');
  }

  function adjustSelectedCounter(delta) {
    const card = getSelectedCard();
    if (!card) {
      return;
    }

    const type = getSelectedCounterType();
    if (!type) {
      setStatus('Enter a counter type first.');
      return;
    }

    commitMutation(() => {
      const currentValue = Number(card.counters[type] || 0);
      const nextValue = currentValue + delta;

      if (nextValue <= 0) {
        delete card.counters[type];
      } else {
        card.counters[type] = nextValue;
      }
      return Number(card.counters[type] || 0);
    }, (nextValue) => nextValue > 0 ? `${card.name} ${type} counter now ${nextValue}.` : `${card.name} ${type} counter removed.`, 'Counters');
  }

  function getSelectedCounterType() {
    if (counterTypeInput?.value === 'custom') {
      return String(counterCustomInput?.value || '').trim();
    }
    return String(counterTypeInput?.value || '').trim();
  }

  function moveSelectedToZone(nextZone) {
    const selectedId = state.selectedCardId;
    if (!selectedId) {
      return;
    }

    commitMutation(() => {
      if (nextZone === 'battlefield') {
        const spawn = getSpawnCoordinates();
        return moveCardToZone(selectedId, nextZone, spawn);
      }
      return moveCardToZone(selectedId, nextZone);
    }, `Moved selected card to ${nextZone}.`, 'Move Card', {
      noopMessage: 'Card is already in that zone.',
    });
  }

  function reorderSelectedInInspectedZone(toTop) {
    const card = getSelectedCard();
    if (!card) {
      return;
    }

    const located = findCardLocation(card.instanceId);
    if (!located || located.zone !== state.inspectedZone || located.zone === 'battlefield') {
      return;
    }

    if (!toTop && state.inspectedZone === 'library' && state.libraryPreviewMode === 'Surveilling') {
      commitMutation(() => moveCardToZone(card.instanceId, 'graveyard'), `Put ${card.name} into your graveyard.`, 'Surveil');
      return;
    }

    commitMutation(() => {
      const zoneCards = state.zones[located.zone];
      zoneCards.splice(located.index, 1);
      if (toTop) {
        zoneCards.push(card);
      } else {
        zoneCards.unshift(card);
      }
      state.libraryPreviewIds = state.libraryPreviewIds.filter((previewId) => previewId !== card.instanceId);
    }, toTop ? `Moved ${card.name} to top of ${state.inspectedZone}.` : `Moved ${card.name} to bottom of ${state.inspectedZone}.`);
  }

  function incrementCommanderCast() {
    const card = getSelectedCard();
    if (!card || !isCommanderLikeCard(card)) {
      return;
    }

    commitMutation(() => {
      const key = getCommanderTaxKey(card);
      state.commanderTaxByName[key] = Number(state.commanderTaxByName[key] || 0) + 1;
    }, `${card.name} cast count increased.`, 'Commander Cast +1');
  }

  function resetCommanderTax() {
    const card = getSelectedCard();
    if (!card || !isCommanderLikeCard(card)) {
      return;
    }

    commitMutation(() => {
      const key = getCommanderTaxKey(card);
      state.commanderTaxByName[key] = 0;
    }, `${card.name} commander tax reset.`, 'Commander Tax Reset');
  }

  function summarizeCounters(card) {
    const entries = Object.entries(card.counters || {}).filter((entry) => Number(entry[1]) > 0);
    if (!entries.length) {
      return '';
    }

    return entries.map((entry) => `${entry[0]}:${entry[1]}`).join(' ');
  }

  function createCardElement(card, contextZone) {
    const cardEl = document.createElement('article');
    cardEl.className = `playtest-card ${contextZone === 'battlefield' ? 'battlefield' : ''}${card.tapped ? ' is-tapped' : ''}${state.selectedCardId === card.instanceId ? ' is-selected' : ''}`;
    cardEl.draggable = true;
    cardEl.dataset.instanceId = card.instanceId;
    cardEl.dataset.zone = contextZone;

    if (contextZone === 'battlefield') {
      const metrics = getBattlefieldClientMetrics();
      const x = Math.max(0, Math.min(DEFAULT_CANVAS_WIDTH - 140, Number(card.x || 0)));
      const y = Math.max(0, Math.min(DEFAULT_CANVAS_HEIGHT - 180, Number(card.y || 0)));
      const left = metrics.width > 0 ? (x / DEFAULT_CANVAS_WIDTH) * metrics.width : 0;
      const top = metrics.height > 0 ? (y / DEFAULT_CANVAS_HEIGHT) * metrics.height : 0;
      cardEl.style.left = `${left}px`;
      cardEl.style.top = `${top}px`;
    }

    if (card.faceDown) {
      const back = document.createElement('div');
      back.className = 'playtest-card-back';
      back.textContent = 'Face Down Card';
      cardEl.appendChild(back);
    } else {
      const image = document.createElement('img');
      image.src = card.imageUri || card.imageSmallUri || '';
      image.alt = card.name;
      image.loading = 'lazy';
      image.decoding = 'async';
      cardEl.appendChild(image);
    }

    const counters = summarizeCounters(card);
    if (counters) {
      const counterChip = document.createElement('span');
      counterChip.className = 'playtest-counter-chip';
      Object.entries(card.counters || {})
        .filter(([, value]) => Number(value) > 0)
        .forEach(([type, value]) => {
          const counterLine = document.createElement('span');
          counterLine.textContent = `${type}: ${value}`;
          counterChip.appendChild(counterLine);
        });
      cardEl.appendChild(counterChip);
    }

    const meta = document.createElement('div');
    meta.className = 'playtest-card-meta';
    meta.innerHTML = card.faceDown
      ? `<span>Face down</span><span>${card.tapped ? 'Tapped' : 'Ready'}</span>`
      : `<span>${escapeHtml(card.name)}</span><span>${card.tapped ? 'Tapped' : 'Ready'}</span>`;
    cardEl.appendChild(meta);

    cardEl.addEventListener('click', (event) => {
      event.stopPropagation();
      state.selectedCardId = card.instanceId;
      renderAll();
      saveSession();
    });

    cardEl.addEventListener('dblclick', (event) => {
      event.stopPropagation();
      if (card.faceDown) {
        return;
      }
      const imageSrc = card.imageUri || card.imageSmallUri;
      if (imageSrc) {
        openImageModal(imageSrc, card.name);
      }
    });

    cardEl.addEventListener('dragstart', (event) => {
      event.dataTransfer.setData('text/plain', card.instanceId);
      event.dataTransfer.effectAllowed = 'move';
    });

    return cardEl;
  }

  function renderHand() {
    handEl.innerHTML = '';
    state.zones.hand.forEach((card) => {
      handEl.appendChild(createCardElement(card, 'hand'));
    });
  }

  function renderBattlefield() {
    battlefieldCardsEl.innerHTML = '';
    state.zones.battlefield.forEach((card) => {
      battlefieldCardsEl.appendChild(createCardElement(card, 'battlefield'));
    });
  }

  function renderInspectedZoneCards() {
    zoneCardsEl.innerHTML = '';
    const cards = state.zones[state.inspectedZone] || [];
    const zoneDrawer = zoneCardsEl.closest('.playtest-zone-drawer');
    zoneDrawer?.classList.toggle('is-visible', state.inspectedZoneOpen);
    if (zoneSearchInput) {
      zoneSearchInput.value = state.inspectedZoneSearch;
      zoneSearchInput.placeholder = state.inspectedZone === 'library' ? 'Search your deck' : `Search ${state.inspectedZone}`;
    }

    const normalizedSearch = state.inspectedZoneSearch.toLowerCase();
    const previewCards = state.inspectedZone === 'library' && state.libraryPreviewCount > 0
      ? state.libraryPreviewIds
        .map((instanceId) => cards.find((card) => card.instanceId === instanceId))
        .filter(Boolean)
      : cards;
    const visibleCards = normalizedSearch
      ? previewCards.filter((card) => card.name.toLowerCase().includes(normalizedSearch) || card.typeLine.toLowerCase().includes(normalizedSearch))
      : previewCards;

    visibleCards
      .slice()
      .reverse()
      .forEach((card) => {
        const zoneViewCard = { ...card };
        // Only the library browser reveals cards; face-down battlefield/exile cards stay hidden.
        if (state.inspectedZone === 'library') {
          zoneViewCard.faceDown = false;
        }
        zoneCardsEl.appendChild(createCardElement(zoneViewCard, state.inspectedZone));
      });

    if (!visibleCards.length) {
      const empty = document.createElement('p');
      empty.className = 'playtest-muted';
      empty.textContent = normalizedSearch ? 'No matching cards.' : `No cards in ${state.inspectedZone}.`;
      zoneCardsEl.appendChild(empty);
    }
  }

  function updateZoneCounts() {
    zoneCountLibrary.textContent = String(state.zones.library.length);
    zoneCountGraveyard.textContent = String(state.zones.graveyard.length);
    zoneCountExile.textContent = String(state.zones.exile.length);
    zoneCountCommand.textContent = String(state.zones.command.length);

    zoneButtons.forEach((button) => {
      const zone = button.dataset.zone;
      button.classList.toggle('active', zone === state.inspectedZone);
    });

    renderZonePile(zonePileLibrary, state.zones.library, true);
    renderZonePile(zonePileGraveyard, state.zones.graveyard, false);
    renderZonePile(zonePileExile, state.zones.exile, false);
    renderZonePile(zonePileCommand, state.zones.command, false);
  }

  function renderZonePile(pileElement, cards, showBack) {
    if (!pileElement) {
      return;
    }

    const topCard = cards[cards.length - 1] || null;
    pileElement.classList.toggle('has-card', Boolean(topCard));
    if (!topCard) {
      pileElement.style.backgroundImage = '';
      pileElement.textContent = '';
      return;
    }

    if (showBack || topCard.faceDown) {
      pileElement.style.backgroundImage = 'linear-gradient(135deg, #303b5f, #647395)';
      pileElement.textContent = '';
      return;
    }

    pileElement.style.backgroundImage = topCard.imageUri || topCard.imageSmallUri
      ? `url("${String(topCard.imageUri || topCard.imageSmallUri).replace(/"/g, '\\"')}")`
      : 'linear-gradient(135deg, #dce4f0, #b7c5d9)';
    pileElement.textContent = '';
  }

  function renderLife() {
    lifeInput.value = String(state.life);
  }

  function renderDebugSnapshot() {
    if (!debugOutputEl) {
      return;
    }

    const selected = getSelectedCard();
    const payload = {
      deckId: state.deckId,
      deckName: state.deckName,
      selectedCardId: state.selectedCardId,
      selectedCardName: selected?.name || '',
      life: state.life,
      mulliganCount: state.mulliganCount,
      openingHandSize: state.openingHandSize,
      shuffleSeed: state.shuffleSeed,
      touchMoveMode: state.touchMoveMode,
      history: {
        undoDepth: history.undo.length,
        redoDepth: history.redo.length,
      },
      zoneCounts: {
        library: state.zones.library.length,
        hand: state.zones.hand.length,
        battlefield: state.zones.battlefield.length,
        graveyard: state.zones.graveyard.length,
        exile: state.zones.exile.length,
        command: state.zones.command.length,
      },
      commanderTaxByCard: state.commanderTaxByName,
    };

    debugOutputEl.textContent = JSON.stringify(payload, null, 2);
  }

  function renderActionLog() {
    if (!actionLogEl) {
      return;
    }

    if (!state.actionLog.length) {
      actionLogEl.innerHTML = '<span class="playtest-action-item">No actions recorded yet.</span>';
      return;
    }

    actionLogEl.innerHTML = state.actionLog
      .slice()
      .reverse()
      .map((entry) => {
        const when = new Date(entry.at).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit', second: '2-digit' });
        return `<span class="playtest-action-item">[${escapeHtml(when)}] <strong>${escapeHtml(entry.action)}</strong>${entry.details ? `: ${escapeHtml(entry.details)}` : ''}</span>`;
      })
      .join('');
  }

  function isSessionDirty() {
    if (state.deckId || state.deckName || state.mulliganCount || state.openingHandSize !== 7) {
      return true;
    }

    const hasCards = zoneNames.some((zone) => (state.zones[zone] || []).length > 0);
    if (hasCards) {
      return true;
    }

    return state.life !== 40 || Object.values(state.commanderTaxByName).some((value) => Number(value) > 0);
  }

  function renderAll() {
    updateZoneCounts();
    updateMulliganStatus();
    updateHistoryButtons();
    if (shuffleSeedInput) {
      shuffleSeedInput.value = state.shuffleSeed;
    }
    if (touchMoveModeInput) {
      touchMoveModeInput.checked = state.touchMoveMode;
    }
    renderLife();
    renderHand();
    renderBattlefield();
    renderInspectedZoneCards();
    if (inspectedZoneLabelEl) {
      inspectedZoneLabelEl.textContent = state.inspectedZone.charAt(0).toUpperCase() + state.inspectedZone.slice(1);
    }
    updateSelectedControls();
    renderDebugSnapshot();
    renderActionLog();
  }

  function toCanvasCoordinates(clientX, clientY) {
    const rect = battlefieldDrop.getBoundingClientRect();
    const renderedCard = document.querySelector('.playtest-card');
    const renderedCardRect = renderedCard?.getBoundingClientRect();
    const cardWidth = renderedCardRect?.width || 135;
    const cardHeight = renderedCardRect?.height || 190;
    const xInClient = Math.max(0, Math.min(rect.width, clientX - rect.left - cardWidth / 2));
    const yInClient = Math.max(0, Math.min(rect.height, clientY - rect.top - cardHeight / 2));

    return {
      x: Math.max(0, Math.min(DEFAULT_CANVAS_WIDTH - cardWidth, (xInClient / Math.max(1, rect.width)) * DEFAULT_CANVAS_WIDTH)),
      y: Math.max(0, Math.min(DEFAULT_CANVAS_HEIGHT - cardHeight, (yInClient / Math.max(1, rect.height)) * DEFAULT_CANVAS_HEIGHT)),
    };
  }

  function attachDropTarget(element, resolver) {
    element.addEventListener('dragover', (event) => {
      event.preventDefault();
      event.dataTransfer.dropEffect = 'move';
    });

    element.addEventListener('drop', (event) => {
      event.preventDefault();
      const instanceId = event.dataTransfer.getData('text/plain');
      if (!instanceId) {
        return;
      }

      resolver(instanceId, event);
    });
  }

  function drawBattlefieldCanvas() {
    const ctx = battlefieldCanvas.getContext('2d');
    if (!ctx) {
      return;
    }

    const width = battlefieldCanvas.width;
    const height = battlefieldCanvas.height;

    const gradient = ctx.createLinearGradient(0, 0, width, height);
    gradient.addColorStop(0, '#d9e2ef');
    gradient.addColorStop(0.55, '#c7d3e3');
    gradient.addColorStop(1, '#b9c7da');

    ctx.fillStyle = gradient;
    ctx.fillRect(0, 0, width, height);

    ctx.strokeStyle = 'rgba(44, 58, 114, 0.10)';
    ctx.lineWidth = 1;

    for (let x = 0; x <= width; x += 70) {
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, height);
      ctx.stroke();
    }

    for (let y = 0; y <= height; y += 70) {
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(width, y);
      ctx.stroke();
    }

    ctx.strokeStyle = 'rgba(44, 58, 114, 0.22)';
    ctx.lineWidth = 2;
    ctx.strokeRect(8, 8, width - 16, height - 16);
  }

  function openImageModal(src, alt) {
    imageModalImage.src = src;
    imageModalImage.alt = alt || 'Card preview';
    imageModal.hidden = false;
  }

  function closeImageModal() {
    imageModal.hidden = true;
    imageModalImage.src = '';
    imageModalImage.alt = 'Card preview';
  }

  function escapeHtml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function exportSessionJsonString() {
    return JSON.stringify(getRuntimeSnapshot(), null, 2);
  }

  function exportSessionToTextarea() {
    if (!sessionJsonInput) {
      return;
    }
    const payload = exportSessionJsonString();
    sessionJsonInput.value = payload;

    const blob = new Blob([payload], { type: 'application/json' });
    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = `playtest-session-${new Date().toISOString().replace(/[:.]/g, '-')}.json`;
    link.click();
    URL.revokeObjectURL(downloadUrl);

    appendActionLog('Session Export', 'Exported session JSON to file and textbox.');
    setStatus('Session JSON exported to file and textbox.');
    saveSession();
  }

  async function copySessionJson() {
    const payload = exportSessionJsonString();
    if (!sessionJsonInput) {
      return;
    }
    sessionJsonInput.value = payload;

    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(payload);
        appendActionLog('Session Export', 'Copied session JSON to clipboard.');
        setStatus('Session JSON copied to clipboard.');
        return;
      }
    } catch {
      // Fall back to manual copy path.
    }

    sessionJsonInput.focus();
    sessionJsonInput.select();
    appendActionLog('Session Export', 'Clipboard blocked. Selected JSON for manual copy.');
    setStatus('Clipboard unavailable. Session JSON selected for manual copy.');
  }

  function importSessionFromTextarea() {
    const raw = String(sessionJsonInput?.value || '').trim();
    if (!raw) {
      setStatus('Paste session JSON before importing.');
      return;
    }

    const parsed = parseJsonSafe(raw, null);
    if (!parsed || typeof parsed !== 'object' || !parsed.zones) {
      setStatus('Session JSON is invalid.');
      return;
    }

    commitMutation(() => {
      applyRuntimeSnapshot(parsed);
      history.undo = [];
      history.redo = [];
    }, 'Imported session JSON.', 'Session Import');
  }

  function createTokensOnBattlefield() {
    const name = String(tokenNameInput.value || '').trim();
    if (!name) {
      setStatus('Enter a token name first.');
      return;
    }

    const count = Number(tokenCountInput.value);
    const amount = Number.isFinite(count) ? Math.max(1, Math.floor(count)) : 1;
    const imageUri = String(tokenImageInput.value || '').trim();

    commitMutation(() => {
      for (let i = 0; i < amount; i += 1) {
        const spawn = getSpawnCoordinates();
        state.zones.battlefield.push({
          instanceId: makeId(),
          name,
          imageUri,
          imageSmallUri: imageUri,
          typeLine: 'Token',
          scryfallUri: '',
          isToken: true,
          isCommanderCard: false,
          tapped: false,
          faceDown: false,
          counters: {},
          zone: 'battlefield',
          x: spawn.x,
          y: spawn.y,
        });
      }
    }, `Created ${amount} ${name} token${amount === 1 ? '' : 's'} on battlefield.`);

    tokenNameInput.value = '';
    tokenCountInput.value = '1';
    tokenImageInput.value = '';
  }

  function createTokenCopyOfSelectedCard() {
    const card = getSelectedCard();
    if (!card) {
      return;
    }

    commitMutation(() => {
      const spawn = getSpawnCoordinates();
      state.zones.battlefield.push({
        instanceId: makeId(),
        name: card.name,
        imageUri: card.imageUri,
        imageSmallUri: card.imageSmallUri,
        typeLine: card.typeLine,
        scryfallUri: card.scryfallUri,
        isToken: true,
        isCommanderCard: false,
        tapped: false,
        faceDown: card.faceDown,
        counters: {},
        zone: 'battlefield',
        x: spawn.x,
        y: spawn.y,
      });
    }, `Created a token copy of ${card.name}.`, 'Copy Token');
  }

  function shouldIgnoreShortcutTarget(target) {
    if (!(target instanceof HTMLElement)) {
      return false;
    }

    return Boolean(target.closest('input, textarea, select, button'));
  }

  function handleKeyboardShortcut(event) {
    if (shouldIgnoreShortcutTarget(event.target)) {
      return;
    }

    const key = String(event.key || '').toLowerCase();

    if (event.ctrlKey || event.metaKey) {
      if (key === 'z' && !event.shiftKey) {
        event.preventDefault();
        undoMutation();
        return;
      }
      if (key === 'y' || (key === 'z' && event.shiftKey)) {
        event.preventDefault();
        redoMutation();
      }
      return;
    }

    if (key === 'escape' && !imageModal.hidden) {
      closeImageModal();
      return;
    }

    if (key === 'd') {
      event.preventDefault();
      if (event.shiftKey) {
        drawCards(7);
      } else {
        drawCards(1);
      }
      return;
    }

    if (key === 't') {
      event.preventDefault();
      toggleSelectedTapState(false);
      return;
    }

    if (key === 'u') {
      event.preventDefault();
      toggleSelectedTapState(true);
      return;
    }

    const zoneShortcuts = {
      h: 'hand',
      b: 'battlefield',
      g: 'graveyard',
      e: 'exile',
      c: 'command',
      l: 'library',
    };

    if (zoneShortcuts[key]) {
      event.preventDefault();
      moveSelectedToZone(zoneShortcuts[key]);
    }
  }

  function bindEvents() {
    loadDeckButton.addEventListener('click', () => {
      loadDeckIntoSession(deckSelect.value);
    });

    shuffleButton.addEventListener('click', () => {
      commitMutation(() => {
        state.shuffleSeed = String(shuffleSeedInput?.value || '').trim();
        state.zones.library = shuffleArray(state.zones.library, state.shuffleSeed);
      }, () => `Library shuffled.${state.shuffleSeed ? ` Seed: ${state.shuffleSeed}.` : ''}`, 'Shuffle');
    });

    drawOneButton.addEventListener('click', () => drawCards(1));
    drawSevenButton.addEventListener('click', () => drawCards(7));

    openingHandButton.addEventListener('click', () => {
      commitMutation(() => {
        state.mulliganCount = 0;
        state.openingHandSize = 7;
      }, 'Opening hand reset.', 'Opening Hand Reset');
      initializeOpeningHand(7);
    });

    mulliganButton.addEventListener('click', () => {
      runMulligan();
    });

    undoButton.addEventListener('click', undoMutation);
    untapAllButton?.addEventListener('click', untapAllBattlefieldCards);

    resetSessionButton.addEventListener('click', () => {
      if (!isSessionDirty()) {
        setStatus('Session is already empty.');
        return;
      }

      if (!window.confirm('Reset and clear this playtest session?')) {
        return;
      }

      clearSession();
      history.undo = [];
      history.redo = [];
      applyRuntimeSnapshot({
        deckId: '',
        deckName: '',
        life: 40,
        selectedCardId: '',
        inspectedZone: 'graveyard',
        mulliganCount: 0,
        openingHandSize: 7,
        shuffleSeed: '',
        touchMoveMode: state.touchMoveMode,
        commanderTaxByName: {},
        actionLog: [],
        zones: {
          library: [],
          hand: [],
          battlefield: [],
          graveyard: [],
          exile: [],
          command: [],
        },
      });
      deckSelect.value = '';
      if (sessionJsonInput) {
        sessionJsonInput.value = '';
      }
      appendActionLog('Reset Session', 'Cleared playtest session state.');
      setStatus('Session reset.');
      renderAll();
      saveSession();
    });

    shuffleSeedInput?.addEventListener('change', () => {
      const value = String(shuffleSeedInput.value || '').trim();
      commitMutation(() => {
        state.shuffleSeed = value;
      }, value ? `Shuffle seed set to ${value}.` : 'Shuffle seed cleared.', 'Shuffle Seed');
    });

    touchMoveModeInput?.addEventListener('change', () => {
      commitMutation(() => {
        state.touchMoveMode = Boolean(touchMoveModeInput.checked);
      }, state.touchMoveMode ? 'Tap-to-move mode enabled.' : 'Tap-to-move mode disabled.', 'Touch Move Mode');
    });

    exportSessionButton?.addEventListener('click', exportSessionToTextarea);
    copySessionButton?.addEventListener('click', () => {
      void copySessionJson();
    });
    importSessionButton?.addEventListener('click', importSessionFromTextarea);
    clearSessionJsonButton?.addEventListener('click', () => {
      if (sessionJsonInput) {
        sessionJsonInput.value = '';
      }
      setStatus('Import/export textbox cleared.');
    });

    lifeMinusButton.addEventListener('click', () => adjustLife(-1));
    lifePlusButton.addEventListener('click', () => adjustLife(1));
    lifeInput.addEventListener('change', setLifeFromInput);

    counterTypeInput?.addEventListener('change', () => {
      const isCustom = counterTypeInput.value === 'custom';
      if (counterCustomInput) {
        counterCustomInput.hidden = !isCustom;
        if (isCustom) {
          counterCustomInput.focus();
        }
      }
      setStatus(isCustom ? 'Enter a custom counter type.' : `Counter type selected: ${counterTypeInput.options[counterTypeInput.selectedIndex]?.text || counterTypeInput.value}.`);
    });

    createTokenButton.addEventListener('click', createTokensOnBattlefield);

    toggleTapButton.addEventListener('click', () => toggleSelectedTapState(false));
    toggleFaceButton.addEventListener('click', toggleSelectedFaceState);
    createTokenCopyButton.addEventListener('click', createTokenCopyOfSelectedCard);
    counterAddButton.addEventListener('click', () => adjustSelectedCounter(1));
    counterRemoveButton.addEventListener('click', () => adjustSelectedCounter(-1));
    commanderCastButton.addEventListener('click', incrementCommanderCast);
    commanderResetButton.addEventListener('click', resetCommanderTax);

    moveZoneButtons.forEach((button) => {
      button.addEventListener('click', () => moveSelectedToZone(button.dataset.moveZone));
    });

    zoneToTopButton.addEventListener('click', () => reorderSelectedInInspectedZone(true));
    zoneToBottomButton.addEventListener('click', () => reorderSelectedInInspectedZone(false));

    zoneButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const zone = button.dataset.zone;
        if (state.inspectedZone === zone && state.inspectedZoneOpen) {
          state.inspectedZoneOpen = false;
          renderAll();
          saveSession();
          return;
        }

        if (state.touchMoveMode && state.selectedCardId) {
          const selectedLocation = findCardLocation(state.selectedCardId);
          if (selectedLocation && selectedLocation.zone !== zone) {
            moveSelectedToZone(zone);
            return;
          }
        }

        state.inspectedZone = zone;
        state.inspectedZoneSearch = '';
        state.inspectedZoneOpen = true;
        state.libraryPreviewCount = 0;
        state.libraryPreviewMode = '';
        state.libraryPreviewIds = [];
        renderAll();
        saveSession();
      });
    });

    zoneSearchInput?.addEventListener('input', () => {
      state.inspectedZoneSearch = zoneSearchInput.value.trim();
      renderInspectedZoneCards();
      updateSelectedControls();
    });

    const previewLibrary = (mode) => {
      previewLibraryTop(libraryLookCountInput?.value, mode);
    };
    libraryLookButton?.addEventListener('click', () => previewLibrary('Looking at'));
    libraryScryButton?.addEventListener('click', () => previewLibrary('Scrying'));
    librarySurveilButton?.addEventListener('click', () => previewLibrary('Surveilling'));

    document.body.addEventListener('click', (event) => {
      if (event.target.closest('.playtest-card, .playtest-tools-panel, .playtest-zone-stack, .playtest-zone-drawer')) {
        return;
      }
      state.selectedCardId = '';
      renderAll();
      saveSession();
    });

    attachDropTarget(handEl, (instanceId) => {
      commitMutation(() => {
        return moveCardToZone(instanceId, 'hand');
      }, 'Moved card to hand.', 'Move Card');
    });

    attachDropTarget(battlefieldDrop, (instanceId, event) => {
      const coords = toCanvasCoordinates(event.clientX, event.clientY);
      commitMutation(() => {
        return moveCardToZone(instanceId, 'battlefield', coords);
      }, 'Moved card to battlefield.', 'Move Card');
    });

    zoneButtons.forEach((button) => {
      attachDropTarget(button, (instanceId) => {
        const zone = button.dataset.zone;
        commitMutation(() => {
          return moveCardToZone(instanceId, zone);
        }, `Moved card to ${zone}.`, 'Move Card');
      });
    });

    imageModalClose.addEventListener('click', closeImageModal);
    imageModal.addEventListener('click', (event) => {
      if (event.target.dataset.closeModal === 'true') {
        closeImageModal();
      }
    });

    document.addEventListener('keydown', handleKeyboardShortcut);

    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape' && state.inspectedZoneOpen && imageModal.hidden) {
        state.inspectedZoneOpen = false;
        renderAll();
        saveSession();
      }
    });

    window.addEventListener('resize', () => {
      drawBattlefieldCanvas();
      renderBattlefield();
    });
  }

  function initialize() {
    loadDeckCatalog();
    renderDeckOptions();
    void refreshDeckCatalogFromCloud();

    const restored = loadSession();
    if (!restored) {
      state.touchMoveMode = Boolean(window.matchMedia && window.matchMedia('(pointer: coarse)').matches);
    }

    if (restored) {
      if (state.deckId) {
        deckSelect.value = state.deckId;
      }
      setStatus(`Restored playtest session${state.deckName ? ` for ${state.deckName}` : ''}.`);
    }

    bindEvents();
    drawBattlefieldCanvas();
    renderAll();

    if (!restored) {
      setStatus(state.deckCatalog.length
        ? 'Choose a deck and click Load Deck.'
        : 'No saved decks found yet. Build and save a deck first.');
    }
  }

  initialize();
})();
