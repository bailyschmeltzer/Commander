// Live game page controller â€” owns all live-tracker state, setup, actions,
// rendering and wiring. Loaded only on livegame.html, after nothing else that
// shares its scope is required; app.js still provides shared services (storage,
// sync, identity helpers, the shared confirm modal) via the global scope.
(function () {
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
const livePlayerGrid = document.getElementById('live-player-grid');
const liveEventLog = document.getElementById('live-event-log');
const liveActiveActions = document.getElementById('live-active-actions');
const liveActionsToggleButton = document.getElementById('live-actions-toggle');
const liveBackButton = document.getElementById('live-back');
const liveUndoButton = document.getElementById('live-undo');
const liveSwapLifeButton = document.getElementById('live-swap-life');
const liveFinishGameButton = document.getElementById('live-finish-game');
const liveAbandonGameButton = document.getElementById('live-abandon-game');
const liveTurnTracker = document.getElementById('live-turn-tracker');
const liveTurnButton = document.getElementById('live-turn-button');
const liveTurnNumberEl = document.getElementById('live-turn-number');
const liveSourcePrompt = document.getElementById('live-source-prompt');
const liveSourceTitle = document.getElementById('live-source-title');
const liveSourceCopy = document.getElementById('live-source-copy');
const liveSourceOptions = document.getElementById('live-source-options');
const liveSourceCancelButton = document.getElementById('live-source-cancel');

let liveSetupFirstPlayerId = null;
let liveGameCompletionInFlight = false;
let liveHoldTimerId = null;
let liveHoldIntervalId = null;
let liveHoldRepeated = false;
let liveSourcePromptResolver = null;
const LIVE_HOLD_REPEAT_START_DELAY_MS = 1000;
const LIVE_HOLD_REPEAT_INTERVAL_MS = 90;
const LIVE_HOLD_ADJUSTMENT_STEP = 5;
const LIVE_HOLD_ADJUSTMENT_INTERVAL_MS = 170;

if (liveGameStatus) {
  liveGameStatus.setAttribute('role', 'status');
  liveGameStatus.setAttribute('aria-live', 'polite');
}

if (liveEventLog) {
  liveEventLog.setAttribute('role', 'log');
  liveEventLog.setAttribute('aria-live', 'polite');
  liveEventLog.setAttribute('aria-relevant', 'additions text');
}


let liveMeasurementTimerId = null;

function getLiveMobileTablePlayerCount(activeGame = activeGameState) {
  if (!window.matchMedia('(max-width: 900px)').matches || !activeGame) {
    return 0;
  }

  const playerCount = (activeGame.players || []).length;
  return playerCount >= 2 && playerCount <= 5 ? playerCount : 0;
}

function isLiveMobileTableMode() {
  return getLiveMobileTablePlayerCount() > 0;
}

function updateLiveTableModeClass() {
  const mobileTablePlayerCount = getLiveMobileTablePlayerCount();
  document.body.classList.toggle('live-table-mode', mobileTablePlayerCount > 0);
  document.body.classList.toggle('live-table-mode-2', mobileTablePlayerCount === 2);
  document.body.classList.toggle('live-table-mode-3', mobileTablePlayerCount === 3);
  document.body.classList.toggle('live-table-mode-4', mobileTablePlayerCount === 4);
  document.body.classList.toggle('live-table-mode-5', mobileTablePlayerCount === 5);
}

function getLiveMobileSeatLayout(player, playerCount = getLiveMobileTablePlayerCount()) {
  const seat = Number.parseInt(`${player?.seat || 0}`, 10);
  if (playerCount === 2) {
    return { seatClass: `live-seat-${seat}`, orientationClass: 'live-orientation-up' };
  }

  if (playerCount === 3) {
    const orientationBySeat = { 1: 'up', 2: 'right', 3: 'left' };
    return {
      seatClass: `live-seat-${seat}`,
      orientationClass: `live-orientation-${orientationBySeat[seat] || 'up'}`,
    };
  }

  if (playerCount === 5) {
    const orientationBySeat = { 1: 'right', 2: 'left', 3: 'up', 4: 'left', 5: 'right' };
    return {
      seatClass: `live-seat-${seat}`,
      orientationClass: `live-orientation-${orientationBySeat[seat] || 'up'}`,
    };
  }

  return {
    seatClass: `live-seat-${seat}`,
    orientationClass: `live-orientation-${seat}`,
  };
}

function generateLiveEventDescription(event, activeGame = activeGameState) {
  const actor = event.actorPlayerId ? getPlayerNameById(event.actorPlayerId, activeGame) : 'Environment';
  const target = event.targetPlayerId ? getPlayerNameById(event.targetPlayerId, activeGame) : 'Unknown player';

  if (event.type === 'life-loss') return `${target} lost ${event.amount} life${event.actorPlayerId ? ` from ${actor}` : ''}.`;
  if (event.type === 'life-gain') return `${target} gained ${event.amount} life.`;
  if (event.type === 'commander-damage') return `${actor} dealt ${event.amount} commander damage to ${target}.`;
  if (event.type === 'elimination') return `${actor} eliminated ${target}.`;
  if (event.type === 'automatic-win') return `${target} was marked as the winner.`;
  if (event.type === 'life-swap') return `${actor} and ${target} swapped life totals.`;
  if (event.type === 'next-turn') return `${target} started turn ${event.turnNumber}.`;
  return 'Event recorded.';
}

function renderLivePlayerGrid() {
  if (!livePlayerGrid) return;
  if (!activeGameState) {
    livePlayerGrid.innerHTML = '<p>No game in progress.</p>';
    return;
  }

  const mobileTablePlayerCount = getLiveMobileTablePlayerCount();
  const isMobileTable = mobileTablePlayerCount > 0;
  livePlayerGrid.innerHTML = activeGameState.players
    .slice()
    .sort((a, b) => a.seat - b.seat)
    .map((player) => {
      const playerName = escapeHtml(player.name);
      const playerCommander = escapeHtml(player.commander || 'No commander');
      const playerCommanderMarkup = player.commander ? buildCommanderTextHtml(player.commander) : playerCommander;
      const damageEntries = Object.entries(player.commanderDamageTaken || {}).filter(([, amount]) => amount > 0);
      const damageListMarkup = `<ul class="live-card-damage-list">${damageEntries.map(([sourceId, amount]) => `<li>${escapeHtml(getPlayerNameById(sourceId, activeGameState))}: <strong>${amount}</strong></li>`).join('')}</ul>`;
      const damageMarkup = damageEntries.length
        ? `<div class="live-player-damage-panel has-damage"><div class="live-player-damage-title">Cmdr Dmg</div>${damageListMarkup}</div>`
        : '';
      const desktopDamageMarkup = damageEntries.length
        ? `<div class="live-player-damage-inline"><div>Commander damage received:</div>${damageListMarkup}</div>`
        : '';
      const firstPlayerMarkup = player.id === activeGameState.startingPlayerId
        ? '<span class="live-first-player-indicator">First</span>'
        : '';
      const mobileSeatLayout = getLiveMobileSeatLayout(player, mobileTablePlayerCount);

      if (!isMobileTable) {
        return `
          <article class="live-player-card${player.eliminatedAt ? ' is-eliminated' : ''}" aria-label="${playerName}, ${playerCommander}">
            <div class="live-player-card-header">
              <div><h3>${playerName}</h3><p>${playerCommanderMarkup}</p></div>
              <div class="live-player-header-badges">${firstPlayerMarkup}</div>
            </div>
            <div class="live-player-life">
              <button type="button" class="live-player-life-button live-player-life-button-standard" data-action="manual-life-entry" data-player-id="${escapeHtml(player.id)}" aria-label="Set life total for ${playerName}. Current life ${player.life}.">${player.life}</button>
            </div>
            <div class="live-quick-actions live-quick-actions-standard">
              <button type="button" class="live-quick-action is-negative" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="-1" aria-label="Subtract 1 life from ${playerName}">-1</button>
              <button type="button" class="live-quick-action is-positive" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="1" aria-label="Add 1 life to ${playerName}">+1</button>
              <button type="button" class="live-quick-action" data-action="manual-commander-damage" data-player-id="${escapeHtml(player.id)}" aria-label="Add commander damage to ${playerName}">Cmdr</button>
              <button type="button" class="live-quick-action" data-action="manual-eliminate" data-player-id="${escapeHtml(player.id)}" aria-label="Mark ${playerName} out of the game">Out</button>
              <button type="button" class="live-quick-action" data-action="auto-win" data-player-id="${escapeHtml(player.id)}" aria-label="Mark ${playerName} as the winner">Win</button>
            </div>
            <div class="live-player-meta live-player-meta-standard">
              <label class="live-player-toggle"><input type="checkbox" data-action="toggle-cannot-lose" data-player-id="${escapeHtml(player.id)}" aria-label="${playerName} cannot lose the game"${player.cannotLoseTheGame ? ' checked' : ''} /><span>Cannot lose the game</span></label>
              ${desktopDamageMarkup}
            </div>
          </article>`;
      }

      return `
        <article class="live-player-card ${mobileSeatLayout.seatClass}${player.eliminatedAt ? ' is-eliminated' : ''}" aria-label="${playerName}, ${playerCommander}">
          <div class="live-player-card-body ${mobileSeatLayout.orientationClass}">
            <div class="live-player-card-topline"><div class="live-player-title-block"><h3>${playerName}</h3><p>${playerCommanderMarkup}</p></div>${firstPlayerMarkup}</div>
            <div class="live-player-main-layout">
              <button type="button" class="live-tap-half live-tap-loss" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="-1" aria-label="Subtract 1 life from ${playerName}"><span class="live-tap-hint" aria-hidden="true">−</span></button>
              <button type="button" class="live-tap-half live-tap-gain" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="1" aria-label="Add 1 life to ${playerName}"><span class="live-tap-hint" aria-hidden="true">+</span></button>
              <div class="live-player-content-overlay">
                <div class="live-player-action-column">
                  <button type="button" class="live-quick-action" data-action="manual-commander-damage" data-player-id="${escapeHtml(player.id)}" aria-label="Add commander damage to ${playerName}">Cmdr</button>
                  <button type="button" class="live-quick-action" data-action="manual-eliminate" data-player-id="${escapeHtml(player.id)}" aria-label="Mark ${playerName} out of the game">Out</button>
                  <button type="button" class="live-quick-action" data-action="auto-win" data-player-id="${escapeHtml(player.id)}" aria-label="Mark ${playerName} as the winner">Win</button>
                  <label class="live-player-toggle live-player-toggle-compact"><input type="checkbox" data-action="toggle-cannot-lose" data-player-id="${escapeHtml(player.id)}" aria-label="${playerName} cannot lose the game"${player.cannotLoseTheGame ? ' checked' : ''} /><span>Can't<br />lose</span></label>
                </div>
                <div class="live-player-counter-column"><div class="live-player-life-split"><button type="button" class="live-life-split-half live-life-split-loss" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="-1" aria-label="Subtract 1 life from ${playerName}"></button><span class="live-player-life live-life-split-display" aria-hidden="true">${player.life}</span><button type="button" class="live-life-split-half live-life-split-gain" data-action="adjust-life" data-player-id="${escapeHtml(player.id)}" data-delta="1" aria-label="Add 1 life to ${playerName}"></button></div></div>
                <div class="live-player-damage-column">${damageMarkup}</div>
              </div>
            </div>
          </div>
        </article>`;
    }).join('');

  updateLivePlayerCardMeasurements();
}

function updateLivePlayerCardMeasurements() {
  if (!livePlayerGrid) return;
  const cards = livePlayerGrid.querySelectorAll('.live-player-card');
  if (!cards.length) return;

  const measureCards = () => {
    cards.forEach((card) => {
      const rect = card.getBoundingClientRect();
      card.style.setProperty('--live-card-width', `${rect.width}px`);
      card.style.setProperty('--live-card-height', `${rect.height}px`);
    });
  };

  measureCards();
  requestAnimationFrame(measureCards);
  if (liveMeasurementTimerId) window.clearTimeout(liveMeasurementTimerId);
  liveMeasurementTimerId = window.setTimeout(measureCards, 120);
}

function updateLiveViewportHeightVariable() {
  if (!document.body.classList.contains('page-live-game')) return;
  const viewportHeight = Math.round(window.visualViewport?.height || window.innerHeight || 0);
  if (viewportHeight > 0) document.documentElement.style.setProperty('--live-vh', `${viewportHeight}px`);
}

function stabilizeLiveTableModeLayout() {
  if (!isLiveMobileTableMode()) return;
  const syncLayout = () => {
    updateLiveViewportHeightVariable();
    document.documentElement.scrollTop = 0;
    document.body.scrollTop = 0;
    window.scrollTo({ top: 0, left: 0, behavior: 'auto' });
    updateLiveTableModeClass();
    updateLivePlayerCardMeasurements();
  };
  syncLayout();
  requestAnimationFrame(syncLayout);
  window.setTimeout(syncLayout, 120);
  window.setTimeout(syncLayout, 320);
}

function renderLiveEventLog() {
  if (!liveEventLog) return;
  if (!activeGameState?.events?.length) {
    liveEventLog.innerHTML = '<p>No events yet.</p>';
    return;
  }

  liveEventLog.innerHTML = activeGameState.events
    .slice()
    .reverse()
    .slice(0, 10)
    .map((event) => `<div class="live-event-item"><p>${escapeHtml(generateLiveEventDescription(event, activeGameState))}</p><small>Turn ${event.turnNumber || activeGameState.turnNumber} · ${new Date(event.timestamp).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}</small></div>`)
    .join('');
}

function renderLiveGameStatus() {
  if (!liveActiveCard) return;
  updateLiveTableModeClass();

  if (!activeGameState) {
    if (liveGameStatus) liveGameStatus.textContent = 'No game in progress.';
    if (liveUndoButton) liveUndoButton.disabled = true;
    renderLivePlayerGrid();
    renderLiveEventLog();
    return;
  }

  const elapsedMinutes = Math.max(1, Math.round((Date.now() - new Date(activeGameState.startedAt).getTime()) / 60000));
  if (liveGameStatus) {
    liveGameStatus.textContent = `${getActiveAlivePlayers(activeGameState).length} players alive · ${elapsedMinutes} minute${elapsedMinutes === 1 ? '' : 's'} elapsed.`;
  }
  if (liveUndoButton) liveUndoButton.disabled = !activeGameUndoState.length;
  if (liveTurnTracker) liveTurnTracker.hidden = !activeGameState;
  if (liveTurnNumberEl && activeGameState) liveTurnNumberEl.textContent = String(getLiveTrackedTurnNumber());
  renderLivePlayerGrid();
  renderLiveEventLog();
}

function refreshLiveTrackerUi() {
  if (!liveSourcePromptResolver) hideLiveSourcePrompt();
  updateLiveViewportHeightVariable();
  updateLiveTableModeClass();
  renderLiveOrderPreview();
  renderLiveGameStatus();
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
    .filter((player) => player.id !== targetPlayer.id && !player.eliminatedAt)
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

function promptForLivePlayerChoice({ title, description, excludePlayerIds = [] }) {
  if (!liveSourcePrompt || !liveSourceOptions || !activeGameState) {
    return Promise.resolve(null);
  }

  const excludedIds = new Set(excludePlayerIds.filter(Boolean));
  const options = activeGameState.players
    .filter((player) => !player.eliminatedAt && !excludedIds.has(player.id))
    .map((player) => ({ id: player.id, label: player.name }));

  if (!options.length) {
    return Promise.resolve(null);
  }

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

function shouldPromptForSource(targetPlayer, projectedLife, eventType) {
  if (!activeGameState || !targetPlayer) {
    return false;
  }

  if (eventType === 'elimination' || eventType === 'commander-damage') {
    return true;
  }

  if (eventType === 'life-loss') {
    // Prompt until a real opponent is selected as first blood source.
    if (!activeGameState.firstBlood) {
      return true;
    }
    // After first blood, only prompt again if this hit would kill the player.
    return projectedLife <= 0 && !targetPlayer.cannotLoseTheGame;
  }

  return false;
}

function maybeRecordFirstBlood(sourcePlayerId, targetPlayerId, turnNumber) {
  if (!activeGameState || activeGameState.firstBlood || !sourcePlayerId || sourcePlayerId === targetPlayerId) {
    return;
  }

  const sourcePlayer = activeGameState.players.find((player) => player.id === sourcePlayerId && !player.eliminatedAt);
  if (!sourcePlayer) {
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

function getDuplicatePlayerName(rows) {
  const seenPlayers = new Map();

  for (const row of rows) {
    const playerName = String(row?.player || '').trim();
    if (!playerName) {
      continue;
    }

    const normalizedPlayerName = playerName.toLocaleLowerCase();
    if (seenPlayers.has(normalizedPlayerName)) {
      return playerName;
    }

    seenPlayers.set(normalizedPlayerName, playerName);
  }

  return '';
}

function normalizeEliminationSourcePlayerId(targetPlayerId, sourcePlayerId) {
  if (!activeGameState || !sourcePlayerId || sourcePlayerId === targetPlayerId) {
    return '';
  }

  const sourcePlayer = activeGameState.players.find((player) => player.id === sourcePlayerId && !player.eliminatedAt);
  return sourcePlayer ? sourcePlayer.id : '';
}

function eliminateLivePlayer(targetPlayer, sourcePlayerId, turnNumber, reason) {
  if (!targetPlayer || targetPlayer.eliminatedAt || targetPlayer.cannotLoseTheGame) {
    return false;
  }

  const aliveCount = getActiveAlivePlayers(activeGameState).length;
  const creditedSourcePlayerId = normalizeEliminationSourcePlayerId(targetPlayer.id, sourcePlayerId);
  targetPlayer.eliminatedAt = new Date().toISOString();
  targetPlayer.eliminatedTurnNumber = turnNumber;
  targetPlayer.eliminatedByPlayerId = creditedSourcePlayerId;
  targetPlayer.eliminationReason = reason;
  targetPlayer.place = aliveCount;

  if (creditedSourcePlayerId) {
    const sourcePlayer = activeGameState.players.find((player) => player.id === creditedSourcePlayerId);
    if (sourcePlayer) {
      sourcePlayer.kills += 1;
      sourcePlayer.killedPlayers = [...new Set([...(sourcePlayer.killedPlayers || []), targetPlayer.name])];
    }
  }

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
      ? `Who eliminated ${targetPlayer.name}? Choose self for no kill credit.`
      : `Who caused ${targetPlayer.name} to lose ${amount} life?`;
  const requireOpponent = eventType === 'commander-damage';
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

  return selectedSourceId || '';
}

async function swapLivePlayerLifeTotals() {
  if (!activeGameState) {
    return;
  }

  const eligiblePlayers = activeGameState.players.filter((player) => !player.eliminatedAt);
  if (eligiblePlayers.length < 2) {
    await promptLiveAlert('At least two active players are required to swap life totals.', 'Unable to swap life');
    return;
  }

  const firstPlayerId = await promptForLivePlayerChoice({
    title: 'Select first player',
    description: 'Choose the first player for the life-total swap.',
  });
  if (!firstPlayerId) {
    return;
  }

  const firstPlayer = activeGameState.players.find((player) => player.id === firstPlayerId && !player.eliminatedAt);
  if (!firstPlayer) {
    return;
  }

  const secondPlayerId = await promptForLivePlayerChoice({
    title: 'Select second player',
    description: `Choose who should swap life totals with ${firstPlayer.name}.`,
    excludePlayerIds: [firstPlayerId],
  });
  if (!secondPlayerId) {
    return;
  }

  const secondPlayer = activeGameState.players.find((player) => player.id === secondPlayerId && !player.eliminatedAt);
  if (!secondPlayer) {
    return;
  }

  const firstLifeBeforeSwap = firstPlayer.life;
  const secondLifeBeforeSwap = secondPlayer.life;

  saveUndoSnapshot();
  firstPlayer.life = secondLifeBeforeSwap;
  secondPlayer.life = firstLifeBeforeSwap;

  recordLiveEvent({
    type: 'life-swap',
    actorPlayerId: firstPlayer.id,
    targetPlayerId: secondPlayer.id,
    amount: 0,
    turnNumber: getLiveTrackedTurnNumber(),
    notes: `${firstPlayer.name} (${firstLifeBeforeSwap}) swapped life totals with ${secondPlayer.name} (${secondLifeBeforeSwap}).`,
  });

  persistActiveGameState(activeGameState);
  refreshLiveTrackerUi();
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

  flushQueuedActiveGamePersist();
}

function startLiveHoldRepeat(button, { applyInitialChange = false } = {}) {
  stopLiveHoldRepeat();

  const playerId = button.dataset.playerId || '';
  const delta = parseInt(button.dataset.delta || '0', 10);
  const player = activeGameState?.players?.find((entry) => entry.id === playerId);
  if (!player || Number.isNaN(delta) || delta === 0) {
    return;
  }

  if (delta < 0 && shouldPromptForSource(player, player.life + delta, 'life-loss')) {
    if (applyInitialChange) {
      // Allow one immediate life-loss tap (which may prompt for source),
      // but do not start hold-repeat while a source prompt is needed.
      applyQuickLifeChange(playerId, delta);
    }
    return;
  }

  liveHoldRepeated = false;
  if (applyInitialChange) {
    applyQuickLifeChange(playerId, delta);
  }

  const holdDelta = delta * LIVE_HOLD_ADJUSTMENT_STEP;

  liveHoldTimerId = setTimeout(() => {
    liveHoldRepeated = true;
    applyQuickLifeChange(playerId, holdDelta);
    liveHoldIntervalId = setInterval(() => {
      applyQuickLifeChange(playerId, holdDelta);
    }, LIVE_HOLD_ADJUSTMENT_INTERVAL_MS);
  }, LIVE_HOLD_REPEAT_START_DELAY_MS);
}

function getLiveTrackedTurnNumber() {
  const rawValue = parseInt(`${activeGameState?.turnNumber || 1}`, 10);
  return Math.max(1, Number.isNaN(rawValue) ? 1 : rawValue);
}

function syncActiveGameTurnFromInput() {
  const turnNumber = getLiveTrackedTurnNumber();
  if (!activeGameState) {
    return turnNumber;
  }

  if (activeGameState.turnNumber !== turnNumber) {
    saveUndoSnapshot();
  }
  activeGameState.turnNumber = turnNumber;
  persistActiveGameState(activeGameState);
  return turnNumber;
}

async function promptForTurnNumber(message, defaultValue = getLiveTrackedTurnNumber()) {
  const parsed = await promptLiveText(message, {
    title: 'Turn number',
    confirmLabel: 'Save',
    defaultValue: String(defaultValue),
    inputType: 'number',
    inputMode: 'numeric',
    validate: (rawValue) => {
      const result = parseLiveModalInteger(rawValue, 'Enter a whole number greater than 0.');
      if (result.error || result.value < 1) {
        return { error: 'Enter a whole number greater than 0.' };
      }
      return { value: result.value };
    },
  });
  if (parsed === null) {
    return null;
  }

  if (activeGameState && activeGameState.turnNumber !== parsed) {
    saveUndoSnapshot();
  }

  if (activeGameState) {
    activeGameState.turnNumber = parsed;
  }
  return parsed;
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

  const row = createLiveSetupRow(data);
  liveGamePlayerBody.appendChild(row);
  populateRowSelectors(row);
  updateLiveSetupSeatLabels();
}

function prepareLiveRematchSetup(previousPlayers = []) {
  if (!liveGamePlayerBody) {
    return;
  }

  const rematchPlayers = previousPlayers
    .slice()
    .sort((firstPlayer, secondPlayer) => {
      const firstSeat = Number.isFinite(firstPlayer?.seat) ? firstPlayer.seat : Number.MAX_SAFE_INTEGER;
      const secondSeat = Number.isFinite(secondPlayer?.seat) ? secondPlayer.seat : Number.MAX_SAFE_INTEGER;
      return firstSeat - secondSeat;
    })
    .map((player) => ({
      player: player.name || '',
      commander: '',
    }))
    .filter((player) => player.player);

  if (!rematchPlayers.length) {
    return;
  }

  liveGamePlayerBody.innerHTML = '';
  rematchPlayers.forEach((player) => {
    addLiveSetupRow(player);
  });

  liveSetupFirstPlayerId = null;
  if (liveGameDateInput) {
    liveGameDateInput.value = new Date().toISOString().slice(0, 10);
  }
  renderLiveOrderPreview();
  liveGameForm?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  liveGamePlayerBody.querySelector('input[name="commander"]')?.focus();
}

function removeLiveSetupRow() {
  if (!liveGamePlayerBody) {
    return;
  }

  const rows = liveGamePlayerBody.querySelectorAll('tr');
  if (rows.length > 2) {
    rows[rows.length - 1].remove();
    liveGamePlayerBody.querySelectorAll('tr').forEach((r) => populateRowSelectors(r));
  }
  updateLiveSetupSeatLabels();
}

function getLiveSetupRows() {
  if (!liveGamePlayerBody) {
    return [];
  }

  const playerMap = buildCanonicalIdentityMapFromValues(knownPlayers);
  const commanderMap = buildCanonicalIdentityMapFromValues(knownCommanders);

  return Array.from(liveGamePlayerBody.querySelectorAll('tr'))
    .map((row, index) => ({
      id: row.dataset.playerId || generateId(),
      player: canonicalizeIdentityValue(row.querySelector('[name="player"]')?.value || '', playerMap),
      commanderRaw: row.querySelector('[name="commander"]')?.value || '',
      commander: canonicalizeIdentityValue(row.querySelector('[name="commander"]')?.value || '', commanderMap),
      seat: index + 1,
    }))
    .filter((row) => row.player);
}

function renderLiveOrderPreview() {
  if (!liveOrderPreview) {
    return;
  }

  const rows = getLiveSetupRows();
  if (liveStartingLifeInput) {
    liveStartingLifeInput.value = rows.length === 2 ? '30' : '40';
  }
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
  const alternateWinSummary = gameState.alternateWinCondition
    ? `Alternate win: ${gameState.alternateWinCondition}`
    : '';
  const alternateLoseSummary = gameState.players
    .filter((player) => player.eliminationDetails)
    .map((player) => `${player.name} lost via ${player.eliminationDetails}`)
    .join(' | ');

  return [
    `First player: ${getPlayerNameById(gameState.startingPlayerId, gameState)}`,
    firstBloodSummary,
    alternateWinSummary,
    alternateLoseSummary,
    `Kill leader: ${killLeader ? `${killLeader.name} (${killLeader.kills})` : 'None'}`,
    `Finish order: ${finishOrderSummary}`,
  ].filter(Boolean).join(' · ');
}

function buildLiveGameRecordStats(gameState, finalPlayers, durationMinutes) {
  const playersById = new Map((gameState.players || []).map((player) => [player.id, player]));
  const runningLifeTotals = new Map((gameState.players || []).map((player) => [player.id, player.startingLife]));
  const commanderDamageByActor = new Map();
  let highestSingleHit = null;
  let highestLifeLossHit = null;
  let highestLifeTotal = null;
  let fastestElimination = null;

  const createLiveCandidate = (value, holderId, notes = '') => createRecordCandidate({
    value,
    holder: holderId ? getPlayerNameById(holderId, gameState) : '',
    commander: holderId ? playersById.get(holderId)?.commander || '' : '',
    date: gameState.date,
    notes,
  });

  (gameState.players || []).forEach((player) => {
    const candidate = createLiveCandidate(player.startingLife, player.id, 'Starting life total.');
    highestLifeTotal = chooseBetterRecordCandidate(highestLifeTotal, candidate, 'highest-life-total');
  });

  (gameState.events || []).forEach((event) => {
    const targetId = event.targetPlayerId || '';
    const actorId = event.actorPlayerId || '';
    const amount = typeof event.amount === 'number' ? event.amount : Number.parseFloat(event.amount || '0');
    if (!Number.isFinite(amount) || amount <= 0) {
      return;
    }

    if ((event.type === 'life-loss' || event.type === 'commander-damage') && actorId) {
      const targetName = targetId ? getPlayerNameById(targetId, gameState) : 'Unknown player';
      const notes = `Hit ${targetName}${event.type === 'commander-damage' ? ' with commander damage' : ''}.`;
      highestSingleHit = chooseBetterRecordCandidate(highestSingleHit, createLiveCandidate(amount, actorId, notes), 'highest-damage-dealt');
    }

    if (event.type === 'life-loss' && actorId) {
      const targetName = targetId ? getPlayerNameById(targetId, gameState) : 'Unknown player';
      highestLifeLossHit = chooseBetterRecordCandidate(highestLifeLossHit, createLiveCandidate(amount, actorId, `Hit ${targetName} to life.`), 'highest-hit-to-life');
    }

    if (event.type === 'commander-damage' && actorId) {
      commanderDamageByActor.set(actorId, (commanderDamageByActor.get(actorId) || 0) + amount);
    }

    if (targetId && runningLifeTotals.has(targetId)) {
      const currentLife = runningLifeTotals.get(targetId) || 0;
      const updatedLife = event.type === 'life-gain'
        ? currentLife + amount
        : (event.type === 'life-loss' || event.type === 'commander-damage')
          ? currentLife - amount
          : currentLife;
      runningLifeTotals.set(targetId, updatedLife);

      const lifeCandidate = createLiveCandidate(updatedLife, targetId, 'Highest life reached in a live game.');
      highestLifeTotal = chooseBetterRecordCandidate(highestLifeTotal, lifeCandidate, 'highest-life-total');
    }
  });

  let mostCommanderDamage = null;
  commanderDamageByActor.forEach((totalDamage, actorId) => {
    mostCommanderDamage = chooseBetterRecordCandidate(
      mostCommanderDamage,
      createLiveCandidate(totalDamage, actorId, 'Total commander damage dealt in one live game.'),
      'most-commander-damage',
    );
  });

  finalPlayers.forEach((player) => {
    if (typeof player.eliminatedTurnNumber === 'number' && player.eliminatedTurnNumber > 0) {
      fastestElimination = chooseBetterRecordCandidate(
        fastestElimination,
        createRecordCandidate({
          value: player.eliminatedTurnNumber,
          holder: player.eliminatedByPlayerId ? getPlayerNameById(player.eliminatedByPlayerId, gameState) : '',
          commander: player.eliminatedByPlayerId ? playersById.get(player.eliminatedByPlayerId)?.commander || '' : '',
          date: gameState.date,
          notes: `Eliminated ${player.name}.`,
        }),
        'fastest-elimination',
      );
    }
  });

  return {
    durationMinutes,
    highestSingleHit,
    highestLifeLossHit,
    highestLifeTotal,
    mostCommanderDamage,
    fastestElimination,
  };
}

function closeLiveActionsMenu() {
  if (!liveActiveActions) {
    return;
  }

  liveActiveActions.classList.remove('is-open');
  if (liveActionsToggleButton) {
    liveActionsToggleButton.setAttribute('aria-expanded', 'false');
  }
}

function toggleLiveActionsMenu() {
  if (!liveActiveActions) {
    return;
  }

  const nextOpen = !liveActiveActions.classList.contains('is-open');
  liveActiveActions.classList.toggle('is-open', nextOpen);
  if (liveActionsToggleButton) {
    liveActionsToggleButton.setAttribute('aria-expanded', String(nextOpen));
  }
}

function initializeLiveTrackerTouchGuards() {
  if (!document.body.classList.contains('page-live-game')) {
    return;
  }

  if (document.body.dataset.liveTouchGuardsInitialized === 'true') {
    return;
  }

  document.body.dataset.liveTouchGuardsInitialized = 'true';
  let lastTouchEndAt = 0;

  ['gesturestart', 'gesturechange', 'gestureend'].forEach((eventName) => {
    document.addEventListener(eventName, (event) => {
      event.preventDefault();
    }, { passive: false });
  });

  document.addEventListener('touchend', (event) => {
    const touchTarget = event.target.closest?.('.live-player-grid, .live-active-actions, .live-source-dialog, .live-player-life-button, .live-quick-action, .live-tap-half, .live-life-split-half');
    if (!touchTarget) {
      return;
    }

    const now = Date.now();
    if (now - lastTouchEndAt < 350) {
      event.preventDefault();
    }
    lastTouchEndAt = now;
  }, { passive: false });
}

async function startLiveGame() {
  const players = getLiveSetupRows();
  if (players.length < 2) {
    await promptLiveAlert('Add at least two players to start a live game.', 'Unable to start game');
    return;
  }

  const commanderValidationError = await validateCommanderEntries(players, { allowExactCardLookup: true });
  if (commanderValidationError) {
    await promptLiveAlert(commanderValidationError, 'Unable to start game');
    return;
  }

  const duplicatePlayerName = getDuplicatePlayerName(players);
  if (duplicatePlayerName) {
    await promptLiveAlert(`Each player can only appear once per match. Duplicate player: ${duplicatePlayerName}.`, 'Unable to start game');
    return;
  }

  const requestedStartingLife = parseInt(liveStartingLifeInput?.value || '40', 10);
  const startingLife = players.length === 2 ? 30 : requestedStartingLife;
  if (!startingLife || startingLife < 1) {
    await promptLiveAlert('Enter a valid starting life total.', 'Unable to start game');
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
    alternateWinCondition: '',
    startingPlayerId: liveSetupFirstPlayerId,
    turnOrder,
    firstBlood: null,
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
      eliminationDetails: '',
      place: null,
    })),
  });
  closeLiveActionsMenu();
  refreshLiveTrackerUi();
  stabilizeLiveTableModeLayout();
  acquireWakeLock();
}

async function promptForPositiveNumber(message, defaultValue = 1) {
  return promptLiveText(message, {
    title: 'Enter amount',
    confirmLabel: 'Apply',
    defaultValue: String(defaultValue),
    inputType: 'number',
    inputMode: 'numeric',
    validate: (rawValue) => {
      const result = parseLiveModalInteger(rawValue, 'Enter a whole number greater than 0.');
      if (result.error || result.value < 1) {
        return { error: 'Enter a whole number greater than 0.' };
      }
      return { value: result.value };
    },
  });
}

async function promptForSignedNumber(message, defaultValue = -1) {
  return promptLiveText(message, {
    title: 'Enter life change',
    confirmLabel: 'Apply',
    defaultValue: String(defaultValue),
    inputType: 'number',
    inputMode: 'numeric',
    validate: (rawValue) => {
      const result = parseLiveModalInteger(rawValue, 'Enter a whole number greater than 0 or less than 0.');
      if (result.error || result.value === 0) {
        return { error: 'Enter a whole number greater than 0 or less than 0.' };
      }
      return { value: result.value };
    },
  });
}

async function applyCommanderDamageToPlayer(targetPlayerId) {
  if (!activeGameState) {
    return;
  }

  const targetPlayer = activeGameState.players.find((player) => player.id === targetPlayerId);
  if (!targetPlayer || targetPlayer.eliminatedAt) {
    return;
  }

  const amount = await promptForPositiveNumber(`How much commander damage did ${targetPlayer.name} take?`, 1);
  if (amount === null) {
    return;
  }

  const sourcePlayerId = await resolveLiveSourceSelection({
    targetPlayerId,
    eventType: 'commander-damage',
    amount,
    projectedLife: targetPlayer.life - amount,
  });
  if (sourcePlayerId === null) {
    return;
  }

  const projectedLife = targetPlayer.life - amount;
  const projectedCommanderDamage = (targetPlayer.commanderDamageTaken[sourcePlayerId] || 0) + amount;
  const eventTurnNumber = getLiveTrackedTurnNumber();

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

  queueActiveGameStatePersist(activeGameState);
  refreshLiveTrackerUi();

  if (alivePlayers.length === 1 && await promptLiveConfirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`, {
    title: 'Finish game?',
    confirmLabel: 'Finish and save',
  })) {
    await completeActiveGame();
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

  const sourcePlayerId = await resolveLiveSourceSelection({
    targetPlayerId,
    eventType: 'elimination',
    amount: 0,
    projectedLife: targetPlayer.life,
  });
  if (sourcePlayerId === null) {
    return;
  }

  const eventTurnNumber = getLiveTrackedTurnNumber();

  const eliminationDetails = await promptLiveText(`Enter ${targetPlayer.name}'s alternate lose condition for the notes. Leave blank for a normal elimination.`, {
    title: 'Alternate lose condition',
    confirmLabel: 'Save',
    defaultValue: targetPlayer.eliminationDetails || '',
    placeholder: 'Optional details',
  });
  if (eliminationDetails === null) {
    return;
  }

  saveUndoSnapshot();
  targetPlayer.eliminationDetails = eliminationDetails.trim();
  eliminateLivePlayer(targetPlayer, sourcePlayerId, eventTurnNumber, 'manual');
  recordLiveEvent({
    type: 'elimination',
    actorPlayerId: sourcePlayerId,
    targetPlayerId,
    amount: 0,
    turnNumber: eventTurnNumber,
    notes: targetPlayer.eliminationDetails ? `Alternate lose: ${targetPlayer.eliminationDetails}` : '',
  });

  const alivePlayers = getActiveAlivePlayers(activeGameState);
  if (alivePlayers.length === 1) {
    alivePlayers[0].place = 1;
  }

  persistActiveGameState(activeGameState);
  refreshLiveTrackerUi();

  if (alivePlayers.length === 1 && await promptLiveConfirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`, {
    title: 'Finish game?',
    confirmLabel: 'Finish and save',
  })) {
    await completeActiveGame();
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

  const projectedLife = player.life + delta;

  if (delta < 0 && !shouldPromptForSource(player, projectedLife, 'life-loss')) {
    // No prompt needed — apply silently
    const turnNumber = getLiveTrackedTurnNumber();
    saveUndoSnapshot();
    player.life += delta;
    recordLiveEvent({
      type: 'life-loss',
      actorPlayerId: '',
      targetPlayerId: playerId,
      amount: Math.abs(delta),
      turnNumber,
    });
    const alivePlayers = getActiveAlivePlayers(activeGameState);
    if (alivePlayers.length === 1) {
      alivePlayers[0].place = 1;
    }
    queueActiveGameStatePersist(activeGameState);
    refreshLiveTrackerUi();
    if (alivePlayers.length === 1 && await promptLiveConfirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`, {
      title: 'Finish game?',
      confirmLabel: 'Finish and save',
    })) {
      await completeActiveGame();
    }
    return;
  }

  // Save undo snapshot BEFORE async prompts so the undo state always reflects
  // the true pre-action game state (shouldPromptForSource and firstBlood intact).
  saveUndoSnapshot();

  const sourcePlayerId = delta < 0
    ? await resolveLiveSourceSelection({
      targetPlayerId: playerId,
      eventType: 'life-loss',
      amount: Math.abs(delta),
      projectedLife,
    })
    : '';
  if (sourcePlayerId === null) {
    // User cancelled — pop the undo snapshot we eagerly saved
    persistActiveGameUndoState(activeGameUndoState.slice(0, -1));
    return;
  }

  const turnNumber = getLiveTrackedTurnNumber();
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

  queueActiveGameStatePersist(activeGameState);
  refreshLiveTrackerUi();

  if (alivePlayers.length === 1 && await promptLiveConfirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`, {
    title: 'Finish game?',
    confirmLabel: 'Finish and save',
  })) {
    await completeActiveGame();
  }
}

async function applyManualLifeEntry(playerId) {
  if (!activeGameState) {
    return;
  }

  const player = activeGameState.players.find((entry) => entry.id === playerId);
  if (!player || player.eliminatedAt) {
    return;
  }

  const delta = await promptForSignedNumber(`Enter life change for ${player.name}. Use negative for damage and positive for life gain.`, -1);
  if (delta === null) {
    return;
  }

  await applyQuickLifeChange(playerId, delta);
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

  const turnNumber = getLiveTrackedTurnNumber();

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

  if (alivePlayers.length === 1 && await promptLiveConfirm(`${alivePlayers[0].name} is the last player alive. Finish and save this game now?`, {
    title: 'Finish game?',
    confirmLabel: 'Finish and save',
  })) {
    await completeActiveGame();
  }
}

async function markPlayerAutomaticWinner(playerId) {
  if (!activeGameState) {
    return;
  }

  const winner = activeGameState.players.find((entry) => entry.id === playerId);
  if (!winner || winner.eliminatedAt) {
    return;
  }

  if (!await promptLiveConfirm(`Mark ${winner.name} as the winner and finish the game?`, {
    title: 'Confirm winner',
    confirmLabel: 'Mark winner',
  })) {
    return;
  }

  const turnNumber = getLiveTrackedTurnNumber();

  const alternateWinCondition = await promptLiveText(`Enter ${winner.name}'s alternate win condition for the notes. Leave blank for a normal win.`, {
    title: 'Alternate win condition',
    confirmLabel: 'Save',
    defaultValue: activeGameState.alternateWinCondition || '',
    placeholder: 'Optional details',
  });
  if (alternateWinCondition === null) {
    return;
  }

  saveUndoSnapshot();
  activeGameState.turnNumber = turnNumber;
  activeGameState.alternateWinCondition = alternateWinCondition.trim();
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
    turnNumber,
    notes: activeGameState.alternateWinCondition ? `Alternate win: ${activeGameState.alternateWinCondition}` : '',
  });
  await completeActiveGame();
}

function undoLastLiveAction() {
  if (!activeGameUndoState.length) {
    return;
  }

  const previousState = activeGameUndoState[activeGameUndoState.length - 1] || null;
  persistActiveGameState(cloneActiveGameState(previousState));
  persistActiveGameUndoState(activeGameUndoState.slice(0, -1));
  refresh();
}

function buildCompletedLiveGameNotes(autoNotes = '', finalComments = '') {
  const trimmedAutoNotes = String(autoNotes || '').trim();
  const trimmedFinalComments = String(finalComments || '').trim();

  if (!trimmedAutoNotes) {
    return trimmedFinalComments;
  }

  if (!trimmedFinalComments) {
    return trimmedAutoNotes;
  }

  return `${trimmedAutoNotes} · ${trimmedFinalComments}`;
}

async function completeActiveGame() {
  if (!activeGameState || liveGameCompletionInFlight) {
    return;
  }

  flushQueuedActiveGamePersist();
  const gameToSave = cloneActiveGameState(activeGameState);
  if (!gameToSave) {
    return;
  }

  liveGameCompletionInFlight = true;

  try {
    const rematchPlayers = gameToSave.players.map((player) => ({
      name: player.name,
      seat: player.seat,
    }));

    if (!gameToSave.events.length && !await promptLiveConfirm('This live game has no recorded events yet. Save it anyway?', {
      title: 'Save empty live game?',
      confirmLabel: 'Save anyway',
    })) {
      return;
    }

    const alivePlayers = getActiveAlivePlayers(gameToSave);
    if (alivePlayers.length > 1 && !await promptLiveConfirm('More than one player is still alive. Finish and score remaining players by life total?', {
      title: 'Finish active game?',
      confirmLabel: 'Finish game',
    })) {
      return;
    }

    const completedPlayers = gameToSave.players
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
    const durationMinutes = Math.max(1, Math.round((Date.now() - new Date(gameToSave.startedAt).getTime()) / 60000));
    const recordStats = buildLiveGameRecordStats(gameToSave, finalPlayers, durationMinutes);
    const playerRows = finalPlayers.map((player) => ({
      player: player.name,
      commander: player.commander,
      place: player.place,
      kills: player.kills,
      turnKilled: player.eliminatedTurnNumber || null,
      killed: player.killedPlayers || [],
    })).map((row) => enrichSavedGamePlayerRow(row));
    const finishOrder = finalPlayers.map((player) => player.name);

    const autoNotes = buildActiveGameSummary({ ...gameToSave, players: finalPlayers });

    const games = loadGames();
    games.push(buildCompactSavedGameEntry({
      id: gameToSave.id,
      date: gameToSave.date,
      playerRows,
      finishOrder,
      notes: autoNotes,
      liveSummary: {
        startingPlayer: getPlayerNameById(gameToSave.startingPlayerId, gameToSave),
        alternateWinCondition: gameToSave.alternateWinCondition || '',
        alternateLoseConditions: finalPlayers
          .filter((player) => player.eliminationDetails)
          .map((player) => ({ player: player.name, details: player.eliminationDetails })),
        durationMinutes,
        firstBlood: gameToSave.firstBlood
          ? {
            actorPlayer: getPlayerNameById(gameToSave.firstBlood.actorPlayerId, gameToSave),
            actorCommander: finalPlayers.find((player) => player.id === gameToSave.firstBlood.actorPlayerId)?.commander || '',
            targetPlayer: getPlayerNameById(gameToSave.firstBlood.targetPlayerId, gameToSave),
            turnNumber: gameToSave.firstBlood.turnNumber,
          }
          : null,
        recordStats,
        turnNumber: gameToSave.turnNumber,
        eventCount: gameToSave.events.length,
      },
    }));
    saveGames(games, { queueSync: false });
    appState = normalizeAppStateData({
      ...appState,
      activeGame: null,
      activeGameUndo: [],
    });
    persistLocalState(appState);
    activeGameState = null;
    activeGameUndoState = [];
    if (activeGamePersistTimer) {
      clearTimeout(activeGamePersistTimer);
      activeGamePersistTimer = null;
    }
    queueCloudSync(0);
    releaseWakeLock();
    refresh();
    refreshLiveTrackerUi();

    liveGameCompletionInFlight = false;

    const finalComments = await promptLiveText('Any final comments for this game? These will be saved with the game notes.', {
      title: 'Final comments',
      confirmLabel: 'Save game',
      defaultValue: '',
      placeholder: 'Optional notes',
      multiline: true,
    });
    if (finalComments !== null) {
      const trimmedComments = String(finalComments || '').trim();
      const updatedGames = loadGames();
      const savedGameIndex = updatedGames.findIndex((savedGame) => savedGame.id === gameToSave.id);
      if (savedGameIndex >= 0) {
        updatedGames[savedGameIndex] = buildCompactSavedGameEntry({
          ...updatedGames[savedGameIndex],
          notes: buildCompletedLiveGameNotes(autoNotes, trimmedComments),
        });
        saveGames(updatedGames, { queueSync: false });
        refresh();
        queueCloudSync(0);
      }
    }

    await promptLiveAlert('Live game saved to history.', 'Game saved');

    if (await promptLiveConfirm('Start a rematch with the same players in the same seats? You can pick new commanders before starting the next game.', {
      title: 'Start rematch?',
      confirmLabel: 'Set up rematch',
      cancelLabel: 'Not now',
    })) {
      prepareLiveRematchSetup(rematchPlayers);
    }
  } finally {
    if (liveGameCompletionInFlight) {
      liveGameCompletionInFlight = false;
    }
  }
}

async function abandonActiveGame() {
  if (!activeGameState) {
    return;
  }

  const eliminatedCount = activeGameState.players.filter((player) => player.eliminatedAt).length;
  const aliveCount = getActiveAlivePlayers(activeGameState).length;
  const summary = [`${activeGameState.events.length} logged event${activeGameState.events.length === 1 ? '' : 's'}`, `${eliminatedCount} eliminated`, `${aliveCount} still alive`].join(', ');

  if (!await promptLiveConfirm(`Abandon the current live game? This removes the in-progress tracker without saving to history. Current state: ${summary}.`, {
    title: 'Abandon live game?',
    confirmLabel: 'Abandon game',
  })) {
    return;
  }

  persistActiveGameState(null, { syncDelay: 0 });
  persistActiveGameUndoState(null, { syncDelay: 0 });
  releaseWakeLock();
  refreshLiveTrackerUi();
}


  window.CommanderLiveGame = {
    hasView: () => Boolean(liveGameForm || livePlayerGrid || liveGameStatus || document.body.classList.contains('page-live-game')),
    refresh: () => refreshLiveTrackerUi(),
    refreshViewport: () => {
      updateLiveViewportHeightVariable();
      refreshLiveTrackerUi();
    },
    handleViewportChange: () => {
      closeLiveActionsMenu();
      updateLiveViewportHeightVariable();
      refreshLiveTrackerUi();
    },
    init: () => {
      hideLiveSourcePrompt();
      initializeLiveTrackerTouchGuards();
      if (liveGameForm) {
        liveGameDateInput.value = new Date().toISOString().slice(0, 10);
        if (!liveGamePlayerBody.children.length) {
          for (let index = 0; index < 4; index += 1) {
            addLiveSetupRow();
          }
        }
        renderLiveOrderPreview();
      }
    },
  };


  liveSourceOptions?.addEventListener('click', (event) => {
    const button = event.target.closest('[data-source-id]');
    if (button) hideLiveSourcePrompt(button.dataset.sourceId || '');
  });

  liveSourceCancelButton?.addEventListener('click', () => hideLiveSourcePrompt(null));
  liveGamePlayerBody?.addEventListener('input', () => {
    updateLiveSetupSeatLabels();
    renderLiveOrderPreview();
  });

  liveAddPlayerButton?.addEventListener('click', () => {
    addLiveSetupRow();
    renderLiveOrderPreview();
  });

  liveRemovePlayerButton?.addEventListener('click', () => {
    removeLiveSetupRow();
    renderLiveOrderPreview();
  });

  liveRandomizeFirstButton?.addEventListener('click', () => randomizeLiveFirstPlayer());

  liveGameForm?.addEventListener('submit', (event) => {
    event.preventDefault();
    startLiveGame();
  });

  if (livePlayerGrid) {
    const beginLiveLifeAdjustment = (event, adjustButton) => {
      event.preventDefault();
      if (
        event.type === 'pointerdown'
        && typeof adjustButton.setPointerCapture === 'function'
        && Number.isInteger(event.pointerId)
      ) {
        adjustButton.setPointerCapture(event.pointerId);
      }
      startLiveHoldRepeat(adjustButton, { applyInitialChange: true });
    };

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
      if (!button || event.detail !== 0) return;
      const playerId = button.dataset.playerId || '';
      const delta = parseInt(button.dataset.delta || '0', 10);
      if (!playerId || Number.isNaN(delta) || delta === 0) return;
      applyQuickLifeChange(playerId, delta);
    });

    if (typeof window !== 'undefined' && 'PointerEvent' in window) {
      livePlayerGrid.addEventListener('pointerdown', (event) => {
        const adjustButton = event.target.closest('[data-action="adjust-life"]');
        if (adjustButton) beginLiveLifeAdjustment(event, adjustButton);
      });

      ['pointerup', 'pointercancel'].forEach((eventName) => {
        livePlayerGrid.addEventListener(eventName, () => stopLiveHoldRepeat());
        document.addEventListener(eventName, () => stopLiveHoldRepeat());
      });
    } else {
      livePlayerGrid.addEventListener('touchstart', (event) => {
        const adjustButton = event.target.closest('[data-action="adjust-life"]');
        if (adjustButton) beginLiveLifeAdjustment(event, adjustButton);
      }, { passive: false });

      ['touchend', 'touchcancel'].forEach((eventName) => {
        livePlayerGrid.addEventListener(eventName, () => stopLiveHoldRepeat(), { passive: true });
        document.addEventListener(eventName, () => stopLiveHoldRepeat(), { passive: true });
      });
    }

    livePlayerGrid.addEventListener('contextmenu', (event) => {
      if (event.target.closest('[data-action="adjust-life"]')) event.preventDefault();
    });
  }

  liveActionsToggleButton?.addEventListener('click', () => toggleLiveActionsMenu());

  liveBackButton?.addEventListener('click', () => {
    closeLiveActionsMenu();
    if (window.history.length > 1) {
      window.history.back();
      return;
    }
    window.location.href = 'index.html';
  });

  liveUndoButton?.addEventListener('click', () => {
    closeLiveActionsMenu();
    undoLastLiveAction();
  });

  liveSwapLifeButton?.addEventListener('click', () => {
    closeLiveActionsMenu();
    swapLivePlayerLifeTotals();
  });

  liveFinishGameButton?.addEventListener('click', () => {
    closeLiveActionsMenu();
    completeActiveGame();
  });

  liveAbandonGameButton?.addEventListener('click', () => {
    closeLiveActionsMenu();
    abandonActiveGame();
  });

  liveTurnButton?.addEventListener('click', () => {
    if (!activeGameState) return;
    saveUndoSnapshot();
    activeGameState.turnNumber = getLiveTrackedTurnNumber() + 1;
    persistActiveGameState(activeGameState);
    if (liveTurnNumberEl) liveTurnNumberEl.textContent = String(activeGameState.turnNumber);
  });

  document.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape') return;
    if (liveSourcePrompt && !liveSourcePrompt.hidden) hideLiveSourcePrompt(null);
  });

  document.addEventListener('click', (event) => {
    if (!liveActiveActions || !liveActiveActions.classList.contains('is-open')) return;
    if (!event.target.closest('#live-active-actions')) closeLiveActionsMenu();
  });
})();