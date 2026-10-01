(function () {
  const tableBody = document.getElementById('player-stats-body');

  function getPlayerStatsData(games) {
    const cacheBucket = getDerivedCacheBucket(games);
    if (cacheBucket.playerStatsData) {
      return cacheBucket.playerStatsData;
    }

    const rawCommanders = [];
    games.forEach((game) => {
      getGameRows(game).forEach((row) => {
        const commander = (row.commander || '').trim();
        if (commander) {
          rawCommanders.push(commander);
        }
      });
    });

    const commanderMap = buildCanonicalIdentityMapFromValues([
      ...rawCommanders,
      ...getKnownCommanderOptions(),
    ]);
    const stats = {};

    games.forEach((game) => {
      const rows = getGameRows(game);
      const winner = Array.isArray(game.finishOrder) && game.finishOrder.length ? game.finishOrder[0] : null;
      const winningTurn = parseOptionalPositiveInteger(game?.liveSummary?.turnNumber);
      const firstBlood = getGameFirstBloodInfo(game);

      if (firstBlood?.actorPlayer) {
        const firstBloodStat = ensurePlayerStats(stats, firstBlood.actorPlayer);
        firstBloodStat.firstBloods = (firstBloodStat.firstBloods || 0) + 1;
      }

      rows.forEach((row) => {
        const player = (row.player || '').trim();
        if (!player) {
          return;
        }

        const playerStat = ensurePlayerStats(stats, player);
        playerStat.games += 1;
        if (winner === player) {
          playerStat.wins += 1;
          if (winningTurn) {
            playerStat.winTurnTotal += winningTurn;
            playerStat.winTurnCount += 1;
            playerStat.winTurns.push(winningTurn);
          }
        }

        const commander = canonicalizeIdentityValue((row.commander || '').trim(), commanderMap);
        if (commander) {
          playerStat.commanders[commander] = (playerStat.commanders[commander] || 0) + 1;
          if (!playerStat.commanderStats[commander]) {
            playerStat.commanderStats[commander] = { played: 0, wins: 0 };
          }
          playerStat.commanderStats[commander].played += 1;
          if (winner === player) {
            playerStat.commanderStats[commander].wins += 1;
          }
        }

        const killedList = getCleanKilledList(row.killed);
        const killsCount = typeof row.kills === 'number' && !Number.isNaN(row.kills)
          ? row.kills
          : killedList.length;
        playerStat.kills += killsCount;

        killedList.forEach((target) => {
          if (!target) {
            return;
          }

          playerStat.victimCounts[target] = (playerStat.victimCounts[target] || 0) + 1;
          const targetStat = ensurePlayerStats(stats, target);
          targetStat.killerCounts[player] = (targetStat.killerCounts[player] || 0) + 1;
        });
      });
    });

    cacheBucket.playerStatsData = stats;
    return stats;
  }

  function render(games) {
    if (!tableBody) {
      return;
    }

    const sortState = getTableSort('playerStats', 'games', true);
    const stats = getPlayerStatsData(games);
    const players = Object.keys(stats);

    if (!players.length) {
      tableBody.innerHTML = '<tr><td colspan="12">No player stats available.</td></tr>';
      return;
    }

    const rows = players.map((player) => {
      const stat = stats[player];
      const winRateValue = stat.games ? (stat.wins / stat.games) * 100 : 0;
      const favoriteCommander = getMaxCountKey(stat.commanders);
      const nemesis = getMaxCountKey(stat.killerCounts);
      const victim = getMaxCountKey(stat.victimCounts);
      const killAverage = stat.games ? (stat.kills / stat.games).toFixed(1) : '0.0';
      const avgWinTurn = stat.winTurnCount ? (stat.winTurnTotal / stat.winTurnCount) : 0;

      let bestDeck = '—';
      let bestDeckCommander = '';
      let bestDeckWinRate = 0;
      let bestDeckGames = 0;
      Object.entries(stat.commanderStats).forEach(([commander, data]) => {
        if (!data.played) {
          return;
        }
        const successRate = data.wins / data.played;
        const currentBest = bestDeck === '—' ? null : bestDeck;
        if (currentBest === null) {
          bestDeck = `${commander} (${formatPercent(successRate * 100)} from ${data.played})`;
          bestDeckCommander = commander;
          bestDeckWinRate = successRate * 100;
          bestDeckGames = data.played;
          return;
        }

        const [, bestRatePart] = currentBest.match(/\(([-\d.]+)%/) || [null, '0'];
        const bestRate = Number(bestRatePart);
        if (successRate * 100 > bestRate || (successRate * 100 === bestRate && data.played > Number(currentBest.match(/from (\d+)/)?.[1] || 0))) {
          bestDeck = `${commander} (${formatPercent(successRate * 100)} from ${data.played})`;
          bestDeckCommander = commander;
          bestDeckWinRate = successRate * 100;
          bestDeckGames = data.played;
        }
      });

      return {
        player,
        games: stat.games,
        wins: stat.wins,
        winRate: winRateValue,
        firstBloods: stat.firstBloods || 0,
        favoriteCommander: favoriteCommander || '',
        nemesis: nemesis || '',
        victim: victim || '',
        kills: stat.kills,
        kd: Number.parseFloat(killAverage),
        avgWinTurn,
        bestDeck,
        bestDeckCommander,
        bestDeckWinRate,
        bestDeckGames,
      };
    });

    rows.sort((a, b) => {
      let result = 0;
      switch (sortState.column) {
        case 'player': result = compareTextValues(a.player, b.player); break;
        case 'games': result = compareNumberValues(a.games, b.games); break;
        case 'wins': result = compareNumberValues(a.wins, b.wins); break;
        case 'winRate': result = compareNumberValues(a.winRate, b.winRate); break;
        case 'firstBlood': result = compareNumberValues(a.firstBloods, b.firstBloods); break;
        case 'favoriteCommander': result = compareTextValues(a.favoriteCommander, b.favoriteCommander); break;
        case 'nemesis': result = compareTextValues(a.nemesis, b.nemesis); break;
        case 'victim': result = compareTextValues(a.victim, b.victim); break;
        case 'kills': result = compareNumberValues(a.kills, b.kills); break;
        case 'kd': result = compareNumberValues(a.kd, b.kd); break;
        case 'avgWinTurn': result = compareNumberValues(a.avgWinTurn, b.avgWinTurn); break;
        case 'bestDeck': result = compareTextValues(a.bestDeck, b.bestDeck); break;
        default: result = compareNumberValues(a.games, b.games); break;
      }

      if (result === 0) {
        result = compareTextValues(a.player, b.player);
      }

      return finalizeSortResult(result, sortState.descending);
    });

    tableBody.innerHTML = rows.map((row) => `
        <tr>
          <td>${buildHistoryFilterLink(row.player, { player: row.player })}</td>
          <td>${row.games}</td>
          <td>${row.wins}</td>
          <td>${formatPercent(row.winRate)}</td>
          <td>${row.firstBloods}</td>
          <td>${buildDeckListLinkOrText(row.favoriteCommander)}</td>
          <td>${row.nemesis ? escapeHtml(row.nemesis) : '—'}</td>
          <td>${row.victim ? escapeHtml(row.victim) : '—'}</td>
          <td>${row.kills}</td>
          <td>${row.kd.toFixed(1)}</td>
          <td>${row.avgWinTurn ? row.avgWinTurn.toFixed(2) : '—'}</td>
          <td>${row.bestDeckCommander ? `${buildDeckListLinkOrText(row.bestDeckCommander)} (${formatPercent(row.bestDeckWinRate)} from ${row.bestDeckGames})` : row.bestDeck}</td>
        </tr>`).join('');
    updateSortableTableIndicators('playerStats');
  }

  window.CommanderPlayerStats = {
    hasView: () => Boolean(tableBody),
    render,
  };
})();