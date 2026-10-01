(function () {
  const searchInput = document.getElementById('commander-search');
  const tableBody = document.getElementById('commander-stats-body');

  function getCommanderStatsData(games) {
    const cacheBucket = getDerivedCacheBucket(games);
    if (cacheBucket.commanderStatsData) {
      return cacheBucket.commanderStatsData;
    }

    const rawCommanders = [];
    games.forEach((game) => {
      getGameRows(game).forEach((row) => {
        const commander = (row.commander || '').trim();
        if (commander) rawCommanders.push(commander);
      });
    });

    const commanderMap = buildCanonicalIdentityMapFromValues([
      ...rawCommanders,
      ...getKnownCommanderOptions(),
    ]);
    const stats = {};

    games.forEach((game) => {
      const firstBlood = getGameFirstBloodInfo(game);
      getGameRows(game).forEach((row) => {
        const commander = canonicalizeIdentityValue((row.commander || '').trim(), commanderMap);
        if (!commander) return;

        if (!stats[commander]) {
          stats[commander] = {
            games: 0,
            wins: 0,
            points: 0,
            kills: 0,
            firstBloods: 0,
            placementTotal: 0,
            placementScoreTotal: 0,
            placementGames: 0,
            winTurnTotal: 0,
            winTurnCount: 0,
            winTurns: [],
          };
        }

        stats[commander].games += 1;
        if (Array.isArray(game.finishOrder) && game.finishOrder[0] === row.player) {
          stats[commander].wins += 1;
          const winningTurn = parseOptionalPositiveInteger(game?.liveSummary?.turnNumber);
          if (winningTurn) {
            stats[commander].winTurnTotal += winningTurn;
            stats[commander].winTurnCount += 1;
            stats[commander].winTurns.push(winningTurn);
          }
        }

        if (Array.isArray(game.finishOrder) && game.finishOrder.length) {
          const place = game.finishOrder.indexOf(row.player) + 1;
          if (place > 0) {
            stats[commander].points += getPlacementPoints(place);
            stats[commander].placementTotal += place;
            stats[commander].placementGames += 1;

            const playerCount = game.finishOrder.length;
            const placementScore = playerCount > 1 ? 1 - ((place - 1) / (playerCount - 1)) : 1;
            stats[commander].placementScoreTotal += placementScore;
          }
        }

        const kills = typeof row.kills === 'number' && !Number.isNaN(row.kills) ? row.kills : 0;
        stats[commander].kills += kills;
        if (firstBlood?.actorPlayer === row.player) {
          stats[commander].firstBloods += 1;
        }
      });
    });

    cacheBucket.commanderStatsData = stats;
    return stats;
  }

  function getCommanderActualPower(commanderStats) {
    const entries = Object.entries(commanderStats);
    if (!entries.length) return {};

    const maxKillsPerGame = Math.max(...entries.map(([, stat]) => (stat.games ? stat.kills / stat.games : 0)));
    const actual = {};
    entries.forEach(([commander, stat]) => {
      const winRate = stat.games ? stat.wins / stat.games : 0;
      const killsPerGame = stat.games ? stat.kills / stat.games : 0;
      const killScore = maxKillsPerGame ? killsPerGame / maxKillsPerGame : 0;
      const placementScore = stat.placementGames ? stat.placementScoreTotal / stat.placementGames : 0;
      const avgWinTurn = stat.winTurnCount ? (stat.winTurnTotal / stat.winTurnCount) : null;
      const singleGameTurnScore = Array.isArray(stat.winTurns) && stat.winTurns.length
        ? getMean(stat.winTurns.map((turn) => getNormalizedTurnWinScore(turn)))
        : 0;
      const averageTurnScore = avgWinTurn ? getNormalizedTurnWinScore(avgWinTurn) : 0;
      const turnSpeedScore = stat.winTurnCount
        ? ((singleGameTurnScore * 0.65) + (averageTurnScore * 0.35))
        : (stat.wins ? getNeutralTurnWinScore() : 0);
      const rawScore = 0.4 * winRate + 0.2 * killScore + 0.2 * placementScore + 0.2 * turnSpeedScore;
      actual[commander] = Math.round(rawScore * 100) / 10;
    });

    return actual;
  }

  function getCalibratedCommanderActualPower(relativeActualPowers) {
    const commanders = Object.keys(relativeActualPowers);
    const expectedByCommander = {};
    commanders.forEach((commander) => {
      expectedByCommander[commander] = getCommanderExpectedPower(commander);
    });

    const overlap = commanders.filter((commander) => typeof expectedByCommander[commander] === 'number');
    const calibrated = {};
    if (overlap.length >= 2) {
      const relativeValues = overlap.map((commander) => relativeActualPowers[commander]);
      const expectedValues = overlap.map((commander) => expectedByCommander[commander]);
      const relativeMean = getMean(relativeValues);
      const expectedMean = getMean(expectedValues);
      const relativeStd = getStandardDeviation(relativeValues, relativeMean);
      const expectedStd = getStandardDeviation(expectedValues, expectedMean);

      commanders.forEach((commander) => {
        const relative = relativeActualPowers[commander];
        const normalized = relativeStd > 0 ? (relative - relativeMean) / relativeStd : 0;
        const targetStd = expectedStd > 0 ? expectedStd : 1;
        calibrated[commander] = Math.round(clamp(expectedMean + (normalized * targetStd), 0, 10) * 10) / 10;
      });
      return calibrated;
    }

    if (overlap.length === 1) {
      const singleExpected = expectedByCommander[overlap[0]];
      const singleRelative = relativeActualPowers[overlap[0]];
      const shift = singleExpected - singleRelative;
      commanders.forEach((commander) => {
        calibrated[commander] = Math.round(clamp(relativeActualPowers[commander] + shift, 0, 10) * 10) / 10;
      });
      return calibrated;
    }

    commanders.forEach((commander) => {
      calibrated[commander] = relativeActualPowers[commander];
    });
    return calibrated;
  }

  function render(games) {
    if (!tableBody) return;

    const sortState = getTableSort('commanderStats', 'games', true);
    const searchTerm = searchInput?.value.trim().toLowerCase() || '';
    const commanderStats = getCommanderStatsData(games);
    const actualPowers = getCommanderActualPower(commanderStats);
    const calibratedPowers = getCalibratedCommanderActualPower(actualPowers);

    const entries = Object.entries(commanderStats)
      .filter(([commander]) => commander.toLowerCase().includes(searchTerm))
      .map(([commander, stat]) => {
        const winRate = stat.games ? (stat.wins / stat.games) * 100 : 0;
        const pointsPerGame = stat.games ? stat.points / stat.games : 0;
        const killsPerGame = stat.games ? stat.kills / stat.games : 0;
        const averagePlacement = stat.placementGames ? stat.placementTotal / stat.placementGames : 0;
        const avgWinTurn = stat.winTurnCount ? (stat.winTurnTotal / stat.winTurnCount) : 0;
        const expected = getCommanderExpectedPower(commander);
        const actual = typeof actualPowers[commander] === 'number' ? actualPowers[commander] : 0;
        const actualCal = typeof calibratedPowers[commander] === 'number' ? calibratedPowers[commander] : 0;
        return {
          commander,
          games: stat.games,
          wins: stat.wins,
          avgWinTurn,
          pointsPerGame,
          winRate,
          kills: stat.kills,
          firstBloods: stat.firstBloods,
          kd: killsPerGame,
          averagePlacement,
          expected,
          actual,
          actualCal,
          delta: typeof expected === 'number' ? actualCal - expected : 0,
          confidence: getCommanderConfidence(stat.games),
          stat,
        };
      });

    entries.sort((first, second) => {
      let firstValue;
      let secondValue;
      switch (sortState.column) {
        case 'commander': {
          firstValue = first.commander.toLowerCase();
          secondValue = second.commander.toLowerCase();
          return sortState.descending ? secondValue.localeCompare(firstValue) : firstValue.localeCompare(secondValue);
        }
        case 'games': firstValue = first.games; secondValue = second.games; break;
        case 'wins': firstValue = first.wins; secondValue = second.wins; break;
        case 'avgWinTurn': firstValue = first.avgWinTurn; secondValue = second.avgWinTurn; break;
        case 'pointsPerGame': firstValue = first.pointsPerGame; secondValue = second.pointsPerGame; break;
        case 'winRate': firstValue = first.winRate; secondValue = second.winRate; break;
        case 'kills': firstValue = first.kills; secondValue = second.kills; break;
        case 'firstBloods': firstValue = first.firstBloods; secondValue = second.firstBloods; break;
        case 'kd': firstValue = first.kd; secondValue = second.kd; break;
        case 'avgPlace': firstValue = first.averagePlacement; secondValue = second.averagePlacement; break;
        case 'expected': firstValue = first.expected || 0; secondValue = second.expected || 0; break;
        case 'actual': firstValue = first.actual; secondValue = second.actual; break;
        case 'actualCal': firstValue = first.actualCal; secondValue = second.actualCal; break;
        case 'delta': firstValue = first.delta; secondValue = second.delta; break;
        case 'confidence': firstValue = first.confidence.score; secondValue = second.confidence.score; break;
        default: return 0;
      }
      const difference = firstValue - secondValue;
      return sortState.descending ? -difference : difference;
    });

    const rows = entries.map(({ commander, stat, avgWinTurn, pointsPerGame, actual, actualCal, expected, delta, confidence }) => {
      const winRate = stat.games ? (stat.wins / stat.games) * 100 : 0;
      const killsPerGame = stat.games ? stat.kills / stat.games : 0;
      const averagePlacement = stat.placementGames ? stat.placementTotal / stat.placementGames : 0;
      const roundedExpected = typeof expected === 'number' ? expected.toFixed(1) : '';
      const actualPower = typeof actual === 'number' ? actual.toFixed(1) : '0.0';
      const actualCalibrated = typeof actualCal === 'number' ? actualCal.toFixed(1) : '0.0';
      const deltaClass = delta > 0.15 ? 'delta-positive' : delta < -0.15 ? 'delta-negative' : 'delta-neutral';
      const deltaText = typeof expected === 'number' ? `${delta > 0 ? '+' : ''}${delta.toFixed(1)}` : '—';

      return `
        <tr>
          <td>${buildDeckListLinkOrText(commander)}</td>
          <td>${stat.games}</td>
          <td>${stat.wins}</td>
          <td>${avgWinTurn ? avgWinTurn.toFixed(2) : '—'}</td>
          <td>${formatPointsPerGame(pointsPerGame)}</td>
          <td>${formatPercent(winRate)}</td>
          <td>${stat.kills}</td>
          <td>${stat.firstBloods || 0}</td>
          <td>${killsPerGame.toFixed(1)}</td>
          <td>${averagePlacement ? averagePlacement.toFixed(2) : '—'}</td>
          <td>
            <input
              type="number"
              class="commander-expected-input"
              data-commander="${escapeHtml(commander)}"
              min="0"
              max="10"
              step="0.1"
              value="${roundedExpected}"
              placeholder="0.0"
            />
          </td>
          <td>${actualPower}</td>
          <td>${actualCalibrated}</td>
          <td><span class="delta-pill ${deltaClass}">${deltaText}</span></td>
          <td>
            <div class="confidence-cell">
              <span class="confidence-label confidence-${confidence.levelClass}">${confidence.label}</span>
              <div class="confidence-bar"><span style="width: ${confidence.percent}%;"></span></div>
            </div>
          </td>
        </tr>`;
    }).join('');

    tableBody.innerHTML = rows || '<tr><td colspan="15">No commanders match your search.</td></tr>';
    updateSortableTableIndicators('commanderStats');
  }

  function handleExpectedInput(event) {
    const input = event.target.closest('.commander-expected-input');
    if (!input) return;

    const commander = input.dataset.commander;
    const value = parseFloat(input.value);
    if (Number.isNaN(value)) {
      setCommanderExpectedPower(commander, null);
    } else {
      setCommanderExpectedPower(commander, value);
      input.value = Math.round(value * 10) / 10;
    }
  }

  searchInput?.addEventListener('input', () => {
    render(loadGames());
  });
  tableBody?.addEventListener('change', handleExpectedInput);

  window.CommanderStats = {
    hasView: () => Boolean(tableBody),
    render,
  };
})();