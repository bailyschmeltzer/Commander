(function () {
  const draftEntryForm = document.getElementById('draft-entry-form');
  const draftPlayerInput = document.getElementById('draft-player');
  const draftSetInput = document.getElementById('draft-set');
  const draftSetOptions = document.getElementById('draft-set-options');
  const draftSetStatus = document.getElementById('draft-set-status');
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
  const draftFilterPlayer = document.getElementById('draft-filter-player');
  const draftFilterSet = document.getElementById('draft-filter-set');
  const draftFilterSetOptions = document.getElementById('draft-filter-set-options');
  const draftClearFiltersButton = document.getElementById('draft-clear-filters');
  const draftPodEvents = document.getElementById('draft-pod-events');
  const draftPodPlayers = document.getElementById('draft-pod-players');
  const draftPodRecord = document.getElementById('draft-pod-record');
  const draftPodPoints = document.getElementById('draft-pod-points');
  let officialSets = [];
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

  function getPlayerKey(player) {
    return typeof getIdentityKey === 'function'
      ? getIdentityKey(player)
      : String(player || '').trim().toLocaleLowerCase();
  }

  function updateFilterOptions(results) {
    if (draftFilterPlayer) {
      const selectedPlayer = draftFilterPlayer.value;
      const players = new Map();
      results.forEach((result) => {
        const key = getPlayerKey(result.player);
        if (key && !players.has(key)) {
          players.set(key, result.player);
        }
      });
      const playerOptions = [...players.entries()].sort((first, second) => first[1].localeCompare(second[1]));
      draftFilterPlayer.innerHTML = [
        '<option value="">All players</option>',
        ...playerOptions.map(([key, name]) => `<option value="${escapeHtml(key)}">${escapeHtml(name)}</option>`),
      ].join('');
      if (players.has(selectedPlayer)) {
        draftFilterPlayer.value = selectedPlayer;
      }
    }

    if (draftFilterSet) {
      const selectedSet = draftFilterSet.value;
      const setsByName = new Map(officialSets.map((set) => [set.name.toLocaleLowerCase(), set.name]));
      results.forEach((result) => {
        const name = String(result.setName || '').trim();
        const normalizedName = name.toLocaleLowerCase();
        if (name && !setsByName.has(normalizedName)) {
          setsByName.set(normalizedName, name);
        }
      });
      if (draftFilterSetOptions) {
        draftFilterSetOptions.innerHTML = [...setsByName.values()]
          .map((name) => `<option value="${escapeHtml(name)}"></option>`)
          .join('');
      }
      draftFilterSet.value = selectedSet;
    }
  }

  function getFilteredResults(results) {
    const selectedPlayer = draftFilterPlayer?.value || '';
    const selectedSet = String(draftFilterSet?.value || '').trim().toLocaleLowerCase();
    return results.filter((result) => {
      if (selectedPlayer && getPlayerKey(result.player) !== selectedPlayer) {
        return false;
      }
      return !selectedSet || String(result.setName || '').toLocaleLowerCase().includes(selectedSet);
    });
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

      const playerKey = getPlayerKey(player);
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

    draftResultsEmpty.textContent = results.length ? 'No events match these filters.' : 'No events recorded yet.';
    draftResultsEmpty.hidden = sortedResults.length > 0;
  }

  function renderPodOverview(results) {
    if (!draftPodEvents) {
      return;
    }

    const players = new Set(results.map((result) => getPlayerKey(result.player)).filter(Boolean));
    const wins = results.reduce((total, result) => total + result.wins, 0);
    const losses = results.reduce((total, result) => total + result.losses, 0);
    const ties = results.reduce((total, result) => total + result.ties, 0);
    const points = results.reduce((total, result) => total + getEventPoints(result), 0);
    draftPodEvents.textContent = String(results.length);
    draftPodPlayers.textContent = String(players.size);
    draftPodRecord.textContent = formatRecord(wins, losses, ties);
    draftPodPoints.textContent = String(points);
  }

  async function loadOfficialSets() {
    if (!draftSetInput) {
      return;
    }

    draftSetInput.disabled = true;
    try {
      const response = await fetch('/api/sets');
      if (!response.ok) {
        throw new Error(`Set request failed (${response.status}).`);
      }

      const payload = await response.json();
      officialSets = (Array.isArray(payload?.sets) ? payload.sets : [])
        .filter((set) => set?.code && set?.name)
        .map((set) => ({
          code: String(set.code).toLowerCase(),
          name: String(set.name),
          releasedAt: String(set.releasedAt || ''),
        }));
      if (!officialSets.length) {
        throw new Error('The official set list was empty.');
      }

      if (draftSetOptions) {
        draftSetOptions.innerHTML = officialSets
          .map((set) => `<option value="${escapeHtml(set.name)}"></option>`)
          .join('');
      }
      if (draftSetStatus) {
        draftSetStatus.textContent = `${officialSets.length} draftable sets`;
        draftSetStatus.className = 'status-muted';
      }
      render();
    } catch (error) {
      if (draftSetStatus) {
        draftSetStatus.textContent = 'Could not load official sets. Check your connection and reload this page.';
        draftSetStatus.className = 'status-error';
      }
    } finally {
      draftSetInput.disabled = false;
    }
  }

  function render() {
    const allResults = getDraftResults();
    updateFilterOptions(allResults);
    const filteredResults = getFilteredResults(allResults);
    renderPodOverview(filteredResults);
    renderSummary(filteredResults);
    renderResults(filteredResults);
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
    const requestedSetName = String(draftSetInput?.value || '').trim();
    const selectedSet = officialSets.find((set) => set.name.toLocaleLowerCase() === requestedSetName.toLocaleLowerCase());
    const setCode = selectedSet?.code || '';
    const setName = selectedSet?.name || '';

    if (
      !player
      || !setName
      || !setCode
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
      setCode,
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
      draftFilterPlayer?.addEventListener('change', render);
      draftFilterSet?.addEventListener('input', render);
      draftClearFiltersButton?.addEventListener('click', () => {
        if (draftFilterPlayer) draftFilterPlayer.value = '';
        if (draftFilterSet) draftFilterSet.value = '';
        render();
      });
      void loadOfficialSets();
    },
    render,
  };
})();