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

(function () {
  const setupBody = document.getElementById('live-game-player-body');
  const addPlayerButton = document.getElementById('live-add-player');
  const removePlayerButton = document.getElementById('live-remove-player');
  const randomizeFirstButton = document.getElementById('live-randomize-first');
  const gameForm = document.getElementById('live-game-form');
  const playerGrid = document.getElementById('live-player-grid');
  const actionsToggleButton = document.getElementById('live-actions-toggle');
  const backButton = document.getElementById('live-back');
  const undoButton = document.getElementById('live-undo');
  const swapLifeButton = document.getElementById('live-swap-life');
  const finishGameButton = document.getElementById('live-finish-game');
  const abandonGameButton = document.getElementById('live-abandon-game');
  const turnButton = document.getElementById('live-turn-button');
  const turnNumberEl = document.getElementById('live-turn-number');
  const sourceOptions = document.getElementById('live-source-options');
  const sourceCancelButton = document.getElementById('live-source-cancel');
  const modalCancelButton = document.getElementById('live-modal-cancel');
  const modalConfirmButton = document.getElementById('live-modal-confirm');
  const modalInput = document.getElementById('live-modal-input');

  sourceOptions?.addEventListener('click', (event) => {
    const button = event.target.closest('[data-source-id]');
    if (button) hideLiveSourcePrompt(button.dataset.sourceId || '');
  });

  sourceCancelButton?.addEventListener('click', () => hideLiveSourcePrompt(null));
  modalCancelButton?.addEventListener('click', () => hideLiveModal(null));

  modalConfirmButton?.addEventListener('click', () => {
    if (!liveModalConfig) {
      hideLiveModal(true);
      return;
    }

    const activeControl = getActiveLiveModalControl();
    if (!activeControl) {
      hideLiveModal(true);
      return;
    }

    const rawValue = activeControl.value;
    const validationResult = typeof liveModalConfig.validate === 'function'
      ? liveModalConfig.validate(rawValue)
      : { value: rawValue };

    if (validationResult && validationResult.error) {
      setLiveModalError(validationResult.error);
      activeControl.focus();
      return;
    }

    hideLiveModal(validationResult && Object.prototype.hasOwnProperty.call(validationResult, 'value') ? validationResult.value : rawValue);
  });

  modalInput?.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      modalConfirmButton?.click();
    }
  });

  setupBody?.addEventListener('input', () => {
    updateLiveSetupSeatLabels();
    renderLiveOrderPreview();
  });

  addPlayerButton?.addEventListener('click', () => {
    addLiveSetupRow();
    renderLiveOrderPreview();
  });

  removePlayerButton?.addEventListener('click', () => {
    removeLiveSetupRow();
    renderLiveOrderPreview();
  });

  randomizeFirstButton?.addEventListener('click', () => randomizeLiveFirstPlayer());

  gameForm?.addEventListener('submit', (event) => {
    event.preventDefault();
    startLiveGame();
  });

  if (playerGrid) {
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

    playerGrid.addEventListener('click', (event) => {
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
      playerGrid.addEventListener('pointerdown', (event) => {
        const adjustButton = event.target.closest('[data-action="adjust-life"]');
        if (adjustButton) beginLiveLifeAdjustment(event, adjustButton);
      });

      ['pointerup', 'pointercancel'].forEach((eventName) => {
        playerGrid.addEventListener(eventName, () => stopLiveHoldRepeat());
        document.addEventListener(eventName, () => stopLiveHoldRepeat());
      });
    } else {
      playerGrid.addEventListener('touchstart', (event) => {
        const adjustButton = event.target.closest('[data-action="adjust-life"]');
        if (adjustButton) beginLiveLifeAdjustment(event, adjustButton);
      }, { passive: false });

      ['touchend', 'touchcancel'].forEach((eventName) => {
        playerGrid.addEventListener(eventName, () => stopLiveHoldRepeat(), { passive: true });
        document.addEventListener(eventName, () => stopLiveHoldRepeat(), { passive: true });
      });
    }

    playerGrid.addEventListener('contextmenu', (event) => {
      if (event.target.closest('[data-action="adjust-life"]')) event.preventDefault();
    });
  }

  actionsToggleButton?.addEventListener('click', () => toggleLiveActionsMenu());

  backButton?.addEventListener('click', () => {
    closeLiveActionsMenu();
    if (window.history.length > 1) {
      window.history.back();
      return;
    }
    window.location.href = 'index.html';
  });

  undoButton?.addEventListener('click', () => {
    closeLiveActionsMenu();
    undoLastLiveAction();
  });

  swapLifeButton?.addEventListener('click', () => {
    closeLiveActionsMenu();
    swapLivePlayerLifeTotals();
  });

  finishGameButton?.addEventListener('click', () => {
    closeLiveActionsMenu();
    completeActiveGame();
  });

  abandonGameButton?.addEventListener('click', () => {
    closeLiveActionsMenu();
    abandonActiveGame();
  });

  turnButton?.addEventListener('click', () => {
    if (!activeGameState) return;
    saveUndoSnapshot();
    activeGameState.turnNumber = getLiveTrackedTurnNumber() + 1;
    persistActiveGameState(activeGameState);
    if (turnNumberEl) turnNumberEl.textContent = String(activeGameState.turnNumber);
  });
})();