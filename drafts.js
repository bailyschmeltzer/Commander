(function () {
  const draftEntryForm = document.getElementById('draft-entry-form');
  const draftPlayerInput = document.getElementById('draft-player');
  const draftSetInput = document.getElementById('draft-set');
  const draftDateInput = document.getElementById('draft-date');
  const draftPlacementInput = document.getElementById('draft-placement');
  const draftWinsInput = document.getElementById('draft-wins');
  const draftLossesInput = document.getElementById('draft-losses');
  const draftTiesInput = document.getElementById('draft-ties');
  const draftEntryStatus = document.getElementById('draft-entry-status');
  const draftSummaryBody = document.getElementById('draft-summary-body');
  const draftSummaryEmpty = document.getElementById('draft-summary-empty');
  const draftResultsBody = document.getElementById('draft-results-body');
  const draftResultsEmpty = document.getElementById('draft-results-empty');
  const colorOptions = [
    { code: 'W', name: 'White' },
    { code: 'U', name: 'Blue' },
    { code: 'B', name: 'Black' },
    { code: 'R', name: 'Red' },
    { code: 'G', name: 'Green' },
  ];

  function getLocalDateValue() {
    const now = new Date();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${now.getFullYear()}-${month}-${day}`;
  }

  function getDraftResults() {
    return window.CommanderDraftData?.load?.() || [];
  }

  function getEventPoints(result) {
    return (result.wins * 3) + result.ties;
  }

  function formatRecord(wins, losses, ties) {
    return `${wins}-${losses}-${ties}`;
  }

  function formatAverage(value) {
    return Number(value).toFixed(1);
  }

  function renderColorChips(colors) {
    const selectedColors = colorOptions.filter((option) => colors.includes(option.code));
    if (!selectedColors.length) {
      return '<span aria-label="No colors recorded">—</span>';
    }

    return `<span class="draft-color-list">${selectedColors.map((option) => `<span class="draft-color-chip draft-color-chip-${option.code}" title="${option.name}" aria-label="${option.name}">${option.code}</span>`).join('')}</span>`;
  }

  function groupResultsByPlayer(results) {
    const groups = new Map();
    results.forEach((result) => {
      const player = String(result.player || '').trim();
      if (!player) {
        return;
      }

      const playerKey = typeof getIdentityKey === 'function'
        ? getIdentityKey(player)
        : player.toLocaleLowerCase();
      if (!groups.has(playerKey)) {
        groups.set(playerKey, {
          player,
          events: 0,
          wins: 0,
          losses: 0,
          ties: 0,
          totalPlacement: 0,
          bestPlacement: Infinity,
          totalPoints: 0,
          bestRecord: null,
          colorCounts: new Map(),
        });
      }

      const group = groups.get(playerKey);
      const eventPoints = getEventPoints(result);
      group.events += 1;
      group.wins += result.wins;
      group.losses += result.losses;
      group.ties += result.ties;
      group.totalPlacement += result.placement;
      group.bestPlacement = Math.min(group.bestPlacement, result.placement);
      group.totalPoints += eventPoints;
      result.colors.forEach((color) => group.colorCounts.set(color, (group.colorCounts.get(color) || 0) + 1));

      const currentBest = group.bestRecord;
      if (
        !currentBest
        || eventPoints > currentBest.points
        || (eventPoints === currentBest.points && result.wins > currentBest.wins)
        || (eventPoints === currentBest.points && result.wins === currentBest.wins && result.losses < currentBest.losses)
      ) {
        group.bestRecord = {
          wins: result.wins,
          losses: result.losses,
          ties: result.ties,
          points: eventPoints,
        };
      }
    });

    return [...groups.values()];
  }

  function getFavoriteColors(group) {
    const highestFrequency = Math.max(0, ...group.colorCounts.values());
    return colorOptions
      .filter((option) => highestFrequency > 0 && group.colorCounts.get(option.code) === highestFrequency)
      .map((option) => option.code);
  }

  function renderSummary(results) {
    if (!draftSummaryBody) {
      return;
    }

    const groups = groupResultsByPlayer(results)
      .sort((first, second) => {
        const firstAveragePoints = first.totalPoints / first.events;
        const secondAveragePoints = second.totalPoints / second.events;
        return secondAveragePoints - firstAveragePoints
          || (first.totalPlacement / first.events) - (second.totalPlacement / second.events)
          || first.player.localeCompare(second.player);
      });

    draftSummaryBody.innerHTML = groups.map((group) => {
      const averageRecord = [group.wins, group.losses, group.ties]
        .map((value) => formatAverage(value / group.events))
        .join('-');
      const bestRecord = group.bestRecord
        ? `${formatRecord(group.bestRecord.wins, group.bestRecord.losses, group.bestRecord.ties)} (${group.bestRecord.points} pts)`
        : '—';

      return `<tr>
        <td>${escapeHtml(group.player)}</td>
        <td>${group.events}</td>
        <td>${averageRecord}</td>
        <td>${formatAverage(group.totalPlacement / group.events)}</td>
        <td>${group.bestPlacement}</td>
        <td class="draft-best-record">${bestRecord}</td>
        <td>${renderColorChips(getFavoriteColors(group))}</td>
        <td>${formatAverage(group.totalPoints / group.events)}</td>
      </tr>`;
    }).join('');

    draftSummaryEmpty.hidden = groups.length > 0;
  }

  function renderResults(results) {
    if (!draftResultsBody) {
      return;
    }

    const sortedResults = results.slice().sort((first, second) =>
      String(second.date || '').localeCompare(String(first.date || ''))
      || first.setName.localeCompare(second.setName)
      || first.player.localeCompare(second.player));

    draftResultsBody.innerHTML = sortedResults.map((result) => `
      <tr>
        <td>${escapeHtml(result.date || '—')}</td>
        <td>${escapeHtml(result.player)}</td>
        <td>${escapeHtml(result.setName)}</td>
        <td>${formatRecord(result.wins, result.losses, result.ties)}</td>
        <td>${result.placement}</td>
        <td>${renderColorChips(result.colors)}</td>
        <td>${getEventPoints(result)}</td>
        <td><button type="button" class="draft-delete-button" data-draft-delete="${escapeHtml(result.id)}" aria-label="Delete ${escapeHtml(result.player)}'s ${escapeHtml(result.setName)} result">Delete</button></td>
      </tr>`).join('');

    draftResultsEmpty.hidden = sortedResults.length > 0;
  }

  function render() {
    const results = getDraftResults();
    renderSummary(results);
    renderResults(results);
  }

  function setStatus(message, tone = 'muted') {
    if (!draftEntryStatus) {
      return;
    }

    draftEntryStatus.textContent = message;
    draftEntryStatus.className = `draft-entry-status status-${tone}`;
  }

  async function handleEntrySubmit(event) {
    event.preventDefault();
    const wins = Number(draftWinsInput?.value);
    const losses = Number(draftLossesInput?.value);
    const ties = Number(draftTiesInput?.value);
    const placement = Number(draftPlacementInput?.value);
    const player = String(draftPlayerInput?.value || '').trim();
    const setName = String(draftSetInput?.value || '').trim();

    if (
      !player
      || !setName
      || !Number.isSafeInteger(wins) || wins < 0
      || !Number.isSafeInteger(losses) || losses < 0
      || !Number.isSafeInteger(ties) || ties < 0
      || !Number.isSafeInteger(placement) || placement < 1
    ) {
      setStatus('Check the player, set, record, and placement values.', 'error');
      return;
    }

    const colors = Array.from(draftEntryForm.querySelectorAll('input[name="draft-color"]:checked'))
      .map((input) => input.value);
    const results = getDraftResults();
    results.push({
      id: generateId(),
      date: draftDateInput?.value || getLocalDateValue(),
      player,
      setName,
      wins,
      losses,
      ties,
      placement,
      colors,
    });

    window.CommanderDraftData.save(results);
    draftEntryForm.reset();
    if (draftDateInput) {
      draftDateInput.value = getLocalDateValue();
    }
    setStatus(`Saved ${formatRecord(wins, losses, ties)} for ${player}.`, 'success');
  }

  async function handleResultsClick(event) {
    const deleteButton = event.target.closest('[data-draft-delete]');
    if (!deleteButton) {
      return;
    }

    const resultId = deleteButton.dataset.draftDelete || '';
    const results = getDraftResults();
    const result = results.find((entry) => entry.id === resultId);
    if (!result) {
      return;
    }

    const shouldDelete = typeof promptLiveConfirm === 'function'
      ? await promptLiveConfirm(`Delete ${result.player}'s ${result.setName} result?`, {
        title: 'Delete draft result?',
        confirmLabel: 'Delete result',
      })
      : window.confirm(`Delete ${result.player}'s ${result.setName} result?`);
    if (!shouldDelete) {
      return;
    }

    window.CommanderDraftData.save(results.filter((entry) => entry.id !== resultId));
    setStatus(`Deleted ${result.player}'s ${result.setName} result.`, 'success');
  }

  window.CommanderDraftTracker = {
    hasView: () => Boolean(draftEntryForm || draftSummaryBody || draftResultsBody),
    init: () => {
      if (!draftEntryForm) {
        return;
      }
      if (draftDateInput && !draftDateInput.value) {
        draftDateInput.value = getLocalDateValue();
      }
      draftEntryForm.addEventListener('submit', handleEntrySubmit);
      draftResultsBody?.addEventListener('click', handleResultsClick);
    },
    render,
  };
})();