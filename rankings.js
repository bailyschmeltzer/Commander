(function () {
  const summary = document.getElementById('rankings-summary');
  const tableBody = document.getElementById('rankings-table-body');
  const trendsSummary = document.getElementById('recent-trends-summary');
  const recentPlayerBody = document.getElementById('recent-player-trends-body');
  const recentCommanderBody = document.getElementById('recent-commander-trends-body');
  const streaksSummary = document.getElementById('streaks-summary');
  const playerStreaksBody = document.getElementById('player-streaks-body');
  const commanderStreaksBody = document.getElementById('commander-streaks-body');

  function renderLoading() {
    if (summary) {
      summary.innerHTML = Array.from({ length: 6 }, () => `
      <article class="stats-card stats-card--skeleton" aria-hidden="true">
        <span class="stats-skeleton stats-skeleton--title"></span>
        <span class="stats-skeleton stats-skeleton--body"></span>
      </article>`).join('');
    }

    if (trendsSummary) {
      trendsSummary.innerHTML = Array.from({ length: 4 }, () => `
      <article class="stats-card stats-card--skeleton" aria-hidden="true">
        <span class="stats-skeleton stats-skeleton--title"></span>
        <span class="stats-skeleton stats-skeleton--body"></span>
      </article>`).join('');
    }

    if (streaksSummary) {
      streaksSummary.innerHTML = Array.from({ length: 4 }, () => `
      <article class="stats-card stats-card--skeleton" aria-hidden="true">
        <span class="stats-skeleton stats-skeleton--title"></span>
        <span class="stats-skeleton stats-skeleton--body"></span>
      </article>`).join('');
    }

    if (tableBody) tableBody.innerHTML = buildSkeletonTableRows(13, 5);
    if (recentPlayerBody) recentPlayerBody.innerHTML = buildSkeletonTableRows(7, 4);
    if (recentCommanderBody) recentCommanderBody.innerHTML = buildSkeletonTableRows(7, 4);
    if (playerStreaksBody) playerStreaksBody.innerHTML = buildSkeletonTableRows(6, 4);
    if (commanderStreaksBody) commanderStreaksBody.innerHTML = buildSkeletonTableRows(6, 4);
  }

  function renderMain(games) {
    if (!summary || !tableBody) {
      return;
    }

    const sortState = getTableSort('rankingsMain', 'rank', false);
    const entries = buildPlayerRankingEntries(games);
    const recentEntries = buildPlayerRankingEntries(getGamesSortedByDateAscending(games).slice(-10));
    const commanderEntries = buildCommanderRankingEntries(games);
    const playerStreaks = buildStreakEntries(games, (row) => row.player, 'playerStreakEntries');
    const playerStreaksByName = new Map(playerStreaks.map((entry) => [entry.name, entry]));

    if (!entries.length) {
      renderStatCardGroup(summary, [
        { title: 'No rankings yet', body: 'Save a few games to generate pod standings.' },
      ]);
      tableBody.innerHTML = '<tr><td colspan="13">No rankings available yet.</td></tr>';
      return;
    }

    const topPlayer = entries[0];
    const bestEfficiency = entries
      .filter((entry) => entry.games >= 3)
      .sort((a, b) => b.pointsPerGame - a.pointsPerGame)[0] || topPlayer;
    const hottestRecent = recentEntries[0] || topPlayer;
    const activeStreakLeader = playerStreaks.find((entry) => entry.currentWinStreak > 0) || null;
    const topCommander = commanderEntries
      .filter((entry) => entry.games >= 3)
      .sort((a, b) => b.pointsPerGame - a.pointsPerGame || compareAggregateEntries(a, b))[0]
      || commanderEntries[0]
      || null;

    renderStatCardGroup(summary, [
      {
        title: '#1 Player',
        body: `${topPlayer.name} leads at ${formatRating(topPlayer.rating)} ELO across ${topPlayer.games} games.`,
      },
      {
        title: 'Best Efficiency',
        body: `${bestEfficiency.name} averages ${formatPointsPerGame(bestEfficiency.pointsPerGame)} placement points per game.`,
      },
      {
        title: 'Hottest Last 10',
        body: `${hottestRecent.name} tops the last 10 at ${formatRating(hottestRecent.rating)} ELO.`,
      },
      {
        title: 'Best Commander Form',
        body: topCommander ? `${topCommander.name} averages ${formatPointsPerGame(topCommander.pointsPerGame)} placement points per game.` : 'No commander data yet.',
      },
      {
        title: 'Active Win Streak',
        body: activeStreakLeader ? `${activeStreakLeader.name} is on ${activeStreakLeader.currentWinStreak} straight wins.` : 'No active player win streak right now.',
      },
      {
        title: 'Read With Caution',
        body: `${getSampleSizeLabel(topPlayer.games)}. Pod standings still react quickly when only a few dated games are on file.`,
      },
    ]);

    const rankingRows = entries.map((entry, index) => {
      const streakEntry = playerStreaksByName.get(entry.name) || { currentWinStreak: 0, bestWinStreak: 0 };
      return {
        rank: index + 1,
        player: entry.name,
        rating: entry.rating,
        perfGame: entry.pointsPerGame,
        games: entry.games,
        wins: entry.wins,
        currentStreak: streakEntry.currentWinStreak,
        longestStreak: streakEntry.bestWinStreak,
        kills: entry.kills,
        firstBlood: entry.firstBloods,
        avgPlace: entry.avgPlace,
        favoriteCommander: entry.favoriteCommander || '',
        entry,
        streakEntry,
      };
    });

    const playerRecentForm = new Map();
    getGamesSortedByDateAscending(games).forEach((game) => {
      getGameRows(game).forEach((row) => {
        const player = String(row.player || '').trim();
        if (!player) return;
        const place = getGameRowPlace(row, game);
        if (!playerRecentForm.has(player)) playerRecentForm.set(player, []);
        playerRecentForm.get(player).push(place === 1 ? 'W' : 'L');
      });
    });
    playerRecentForm.forEach((results, player) => {
      playerRecentForm.set(player, results.slice(-5));
    });

    rankingRows.sort((a, b) => {
      let result = 0;
      switch (sortState.column) {
        case 'rank': result = compareNumberValues(a.rank, b.rank); break;
        case 'player': result = compareTextValues(a.player, b.player); break;
        case 'rating': result = compareNumberValues(a.rating, b.rating); break;
        case 'perfGame': result = compareNumberValues(a.perfGame, b.perfGame); break;
        case 'games': result = compareNumberValues(a.games, b.games); break;
        case 'wins': result = compareNumberValues(a.wins, b.wins); break;
        case 'currentStreak': result = compareNumberValues(a.currentStreak, b.currentStreak); break;
        case 'longestStreak': result = compareNumberValues(a.longestStreak, b.longestStreak); break;
        case 'kills': result = compareNumberValues(a.kills, b.kills); break;
        case 'firstBlood': result = compareNumberValues(a.firstBlood, b.firstBlood); break;
        case 'avgPlace': result = compareNumberValues(a.avgPlace, b.avgPlace); break;
        case 'favoriteCommander': result = compareTextValues(a.favoriteCommander, b.favoriteCommander); break;
        default: result = compareNumberValues(a.rank, b.rank); break;
      }

      if (result === 0) {
        result = compareNumberValues(a.rank, b.rank);
      }

      return finalizeSortResult(result, sortState.descending);
    });

    tableBody.innerHTML = rankingRows.map((row) => {
      const { entry, streakEntry } = row;
      const form = playerRecentForm.get(entry.name) || [];
      const formDots = form.map((result) => `<span class="form-dot form-dot--${result === 'W' ? 'win' : 'loss'}" title="${result === 'W' ? 'Win' : 'Loss'}">${result === 'W' ? 'W' : 'L'}</span>`).join('');
      return `
      <tr>
        <td>${row.rank}</td>
        <td>${buildHistoryFilterLink(entry.name, { player: entry.name })}</td>
        <td>${formatRating(entry.rating)}</td>
        <td>${formatPointsPerGame(entry.pointsPerGame)}</td>
        <td class="rankings-games-cell">${entry.games}</td>
        <td>${entry.wins}</td>
        <td>${streakEntry.currentWinStreak}</td>
        <td>${streakEntry.bestWinStreak}</td>
        <td>${entry.kills}</td>
        <td>${entry.firstBloods}</td>
        <td>${formatAveragePlace(entry.avgPlace)}</td>
        <td>${entry.favoriteCommander ? buildCommanderDisplayHtml(entry.favoriteCommander, buildHistoryFilterLink(entry.favoriteCommander, { commander: entry.favoriteCommander })) : '—'}</td>
        <td><span class="form-dots">${formDots}</span></td>
      </tr>`;
    }).join('');

    updateSortableTableIndicators('rankingsMain');
  }

  function renderRecentTrends(games) {
    if (!trendsSummary || !recentPlayerBody || !recentCommanderBody) {
      return;
    }

    const playerSortState = getTableSort('recentPlayerTrends', 'points', true);
    const commanderSortState = getTableSort('recentCommanderTrends', 'points', true);
    const recentGames = getGamesSortedByDateAscending(games).slice(-10);
    const playerEntries = buildPlayerRankingEntries(recentGames);
    const commanderEntries = buildCommanderRankingEntries(recentGames);

    if (!recentGames.length) {
      renderStatCardGroup(trendsSummary, [
        { title: 'No recent trend data', body: 'Play a few games and the last-10-game view will appear here.' },
      ]);
      recentPlayerBody.innerHTML = '<tr><td colspan="7">No recent player trend data yet.</td></tr>';
      recentCommanderBody.innerHTML = '<tr><td colspan="7">No recent commander trend data yet.</td></tr>';
      return;
    }

    const hottestPlayer = playerEntries
      .filter((entry) => entry.games >= 2)
      .sort((a, b) => b.winRate - a.winRate || compareAggregateEntries(a, b))[0] || playerEntries[0] || null;
    const hottestCommander = commanderEntries
      .filter((entry) => entry.games >= 2)
      .sort((a, b) => b.winRate - a.winRate || compareAggregateEntries(a, b))[0] || commanderEntries[0] || null;
    const killLeader = playerEntries.slice().sort((a, b) => b.kills - a.kills || compareAggregateEntries(a, b))[0] || null;
    const firstBloodLeader = playerEntries.slice().sort((a, b) => b.firstBloods - a.firstBloods || compareAggregateEntries(a, b))[0] || null;

    renderStatCardGroup(trendsSummary, [
      {
        title: 'Hottest Player',
        body: hottestPlayer ? `${hottestPlayer.name} is at ${formatPercent(hottestPlayer.winRate)} over the last 10 games.` : 'No player trend data yet.',
      },
      {
        title: 'Hottest Commander',
        body: hottestCommander ? `${hottestCommander.name} is at ${formatPercent(hottestCommander.winRate)} lately.` : 'No commander trend data yet.',
      },
      {
        title: 'Recent Kill Leader',
        body: killLeader ? `${killLeader.name} logged ${killLeader.kills} kills in the last 10 games.` : 'No recent kills logged yet.',
      },
      {
        title: 'First Blood Leader',
        body: firstBloodLeader ? `${firstBloodLeader.name} opened ${firstBloodLeader.firstBloods} games with first blood.` : 'No recent first blood leader yet.',
      },
      {
        title: 'Trend Window',
        body: getRecentWindowSummary(games, 10),
      },
    ]);

    const sortedPlayerEntries = playerEntries.slice().sort((a, b) => {
      let result = 0;
      switch (playerSortState.column) {
        case 'player': result = compareTextValues(a.name, b.name); break;
        case 'games': result = compareNumberValues(a.games, b.games); break;
        case 'points': result = compareNumberValues(a.points, b.points); break;
        case 'winRate': result = compareNumberValues(a.winRate, b.winRate); break;
        case 'kills': result = compareNumberValues(a.kills, b.kills); break;
        case 'firstBlood': result = compareNumberValues(a.firstBloods, b.firstBloods); break;
        case 'avgPlace': result = compareNumberValues(a.avgPlace, b.avgPlace); break;
        default:
          result = compareAggregateEntries(a, b);
          return playerSortState.descending ? result : -result;
      }
      if (result === 0) {
        result = compareAggregateEntries(a, b);
        return playerSortState.descending ? result : -result;
      }
      return finalizeSortResult(result, playerSortState.descending);
    });

    recentPlayerBody.innerHTML = sortedPlayerEntries.map((entry) => `
      <tr>
        <td>${buildHistoryFilterLink(entry.name, { player: entry.name })}</td>
        <td>${entry.games}</td>
        <td>${entry.points}</td>
        <td>${formatPercent(entry.winRate)}</td>
        <td>${entry.kills}</td>
        <td>${entry.firstBloods}</td>
        <td>${formatAveragePlace(entry.avgPlace)}</td>
      </tr>`).join('') || '<tr><td colspan="7">No recent player trend data yet.</td></tr>';

    const sortedCommanderEntries = commanderEntries.slice().sort((a, b) => {
      let result = 0;
      switch (commanderSortState.column) {
        case 'commander': result = compareTextValues(a.name, b.name); break;
        case 'games': result = compareNumberValues(a.games, b.games); break;
        case 'points': result = compareNumberValues(a.points, b.points); break;
        case 'winRate': result = compareNumberValues(a.winRate, b.winRate); break;
        case 'kills': result = compareNumberValues(a.kills, b.kills); break;
        case 'firstBlood': result = compareNumberValues(a.firstBloods, b.firstBloods); break;
        case 'avgPlace': result = compareNumberValues(a.avgPlace, b.avgPlace); break;
        default:
          result = compareAggregateEntries(a, b);
          return commanderSortState.descending ? result : -result;
      }
      if (result === 0) {
        result = compareAggregateEntries(a, b);
        return commanderSortState.descending ? result : -result;
      }
      return finalizeSortResult(result, commanderSortState.descending);
    });

    recentCommanderBody.innerHTML = sortedCommanderEntries.map((entry) => `
      <tr>
        <td>${buildDeckListLinkOrText(entry.name)}</td>
        <td>${entry.games}</td>
        <td>${entry.points}</td>
        <td>${formatPercent(entry.winRate)}</td>
        <td>${entry.kills}</td>
        <td>${entry.firstBloods}</td>
        <td>${formatAveragePlace(entry.avgPlace)}</td>
      </tr>`).join('') || '<tr><td colspan="7">No recent commander trend data yet.</td></tr>';

    updateSortableTableIndicators('recentPlayerTrends');
    updateSortableTableIndicators('recentCommanderTrends');
  }

  function renderStreaks(games) {
    if (!streaksSummary || !playerStreaksBody || !commanderStreaksBody) {
      return;
    }

    const playerSortState = getTableSort('playerStreaks', 'currentWins', true);
    const commanderSortState = getTableSort('commanderStreaks', 'currentWins', true);
    const playerEntries = buildStreakEntries(games, (row) => row.player, 'playerStreakEntries');
    const rawStreakCommanders = games.flatMap((game) => getGameRows(game).map((row) => String(row.commander || '').trim()).filter(Boolean));
    const streakCommanderMap = buildCanonicalIdentityMapFromValues([...getKnownCommanderOptions(), ...rawStreakCommanders]);
    const commanderEntries = buildStreakEntries(games, (row) => canonicalizeIdentityValue(String(row.commander || '').trim(), streakCommanderMap), 'commanderStreakEntriesCanonical');

    if (!playerEntries.length && !commanderEntries.length) {
      renderStatCardGroup(streaksSummary, [
        { title: 'No streak data yet', body: 'Streaks appear once you have saved a few games.' },
      ]);
      playerStreaksBody.innerHTML = '<tr><td colspan="6">No player streaks available yet.</td></tr>';
      commanderStreaksBody.innerHTML = '<tr><td colspan="6">No commander streaks available yet.</td></tr>';
      return;
    }

    const activePlayerStreak = playerEntries.find((entry) => entry.currentWinStreak > 0) || null;
    const bestPlayerStreak = playerEntries.slice().sort((a, b) => b.bestWinStreak - a.bestWinStreak || a.name.localeCompare(b.name))[0] || null;
    const activeCommanderStreak = commanderEntries.find((entry) => entry.currentWinStreak > 0) || null;
    const longestDrought = playerEntries.slice().sort((a, b) => b.drought - a.drought || a.name.localeCompare(b.name))[0] || null;

    renderStatCardGroup(streaksSummary, [
      {
        title: 'Active Player Streak',
        body: activePlayerStreak ? `${activePlayerStreak.name} has ${activePlayerStreak.currentWinStreak} consecutive wins.` : 'No active player win streak right now.',
      },
      {
        title: 'Best Player Run',
        body: bestPlayerStreak ? `${bestPlayerStreak.name} peaked at ${bestPlayerStreak.bestWinStreak} straight wins.` : 'No player win streaks yet.',
      },
      {
        title: 'Active Commander Streak',
        body: activeCommanderStreak ? `${activeCommanderStreak.name} has ${activeCommanderStreak.currentWinStreak} consecutive wins.` : 'No active commander streak right now.',
      },
      {
        title: 'Longest Drought',
        body: longestDrought ? `${longestDrought.name} has gone ${longestDrought.drought} appearances without a win.` : 'No drought data yet.',
      },
      {
        title: 'Date Ordering',
        body: getRecentWindowSummary(games, 10),
      },
    ]);

    const sortedPlayerEntries = playerEntries.slice().sort((a, b) => {
      let result = 0;
      switch (playerSortState.column) {
        case 'player': result = compareTextValues(a.name, b.name); break;
        case 'games': result = compareNumberValues(a.games, b.games); break;
        case 'currentWins': result = compareNumberValues(a.currentWinStreak, b.currentWinStreak); break;
        case 'bestWinStreak': result = compareNumberValues(a.bestWinStreak, b.bestWinStreak); break;
        case 'drought': result = compareNumberValues(a.drought, b.drought); break;
        case 'lastWin': result = compareDateValues(a.lastWinDate, b.lastWinDate); break;
        default: result = compareNumberValues(a.currentWinStreak, b.currentWinStreak); break;
      }
      if (result === 0) {
        result = compareTextValues(a.name, b.name);
      }
      return finalizeSortResult(result, playerSortState.descending);
    });

    playerStreaksBody.innerHTML = sortedPlayerEntries.length
      ? sortedPlayerEntries.map((entry) => `
        <tr>
          <td>${buildHistoryFilterLink(entry.name, { player: entry.name })}</td>
          <td>${entry.games}</td>
          <td>${entry.currentWinStreak}</td>
          <td>${entry.bestWinStreak}</td>
          <td>${entry.drought}</td>
          <td>${escapeHtml(entry.lastWinDate || '—')}</td>
        </tr>`).join('')
      : '<tr><td colspan="6">No player streaks available yet.</td></tr>';

    const sortedCommanderEntries = commanderEntries.slice().sort((a, b) => {
      let result = 0;
      switch (commanderSortState.column) {
        case 'commander': result = compareTextValues(a.name, b.name); break;
        case 'games': result = compareNumberValues(a.games, b.games); break;
        case 'currentWins': result = compareNumberValues(a.currentWinStreak, b.currentWinStreak); break;
        case 'bestWinStreak': result = compareNumberValues(a.bestWinStreak, b.bestWinStreak); break;
        case 'drought': result = compareNumberValues(a.drought, b.drought); break;
        case 'lastWin': result = compareDateValues(a.lastWinDate, b.lastWinDate); break;
        default: result = compareNumberValues(a.currentWinStreak, b.currentWinStreak); break;
      }
      if (result === 0) {
        result = compareTextValues(a.name, b.name);
      }
      return finalizeSortResult(result, commanderSortState.descending);
    });

    commanderStreaksBody.innerHTML = sortedCommanderEntries.length
      ? sortedCommanderEntries.map((entry) => `
        <tr>
          <td>${buildDeckListLinkOrText(entry.name)}</td>
          <td>${entry.games}</td>
          <td>${entry.currentWinStreak}</td>
          <td>${entry.bestWinStreak}</td>
          <td>${entry.drought}</td>
          <td>${escapeHtml(entry.lastWinDate || '—')}</td>
        </tr>`).join('')
      : '<tr><td colspan="6">No commander streaks available yet.</td></tr>';

    updateSortableTableIndicators('playerStreaks');
    updateSortableTableIndicators('commanderStreaks');
  }

  window.CommanderRankings = {
    hasView: () => Boolean(summary || tableBody || recentPlayerBody || recentCommanderBody || playerStreaksBody || commanderStreaksBody),
    renderLoading,
    renderMain,
    renderAll(games) {
      renderMain(games);
      renderRecentTrends(games);
      renderStreaks(games);
    },
  };
})();