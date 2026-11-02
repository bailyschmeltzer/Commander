(function () {
  const authAuditSection = document.getElementById('auth-audit-section');
  const authAuditStatus = document.getElementById('auth-audit-status');
  const authAuditList = document.getElementById('auth-audit-list');
  const authAuditRefreshButton = document.getElementById('auth-audit-refresh');
  const authAuditExportJsonButton = document.getElementById('auth-audit-export-json');
  const authAuditExportCsvButton = document.getElementById('auth-audit-export-csv');
  const authAuditFilterResultSelect = document.getElementById('auth-audit-filter-result');
  const authAuditFilterActionSelect = document.getElementById('auth-audit-filter-action');
  const authAuditFilterUserInput = document.getElementById('auth-audit-filter-user');
  const authAuditFilterFromInput = document.getElementById('auth-audit-filter-from');
  const authAuditFilterToInput = document.getElementById('auth-audit-filter-to');
  const authAuditFilterClearButton = document.getElementById('auth-audit-filter-clear');
  const authAuditPageSizeSelect = document.getElementById('auth-audit-page-size');
  const authAuditPagePrevButton = document.getElementById('auth-audit-page-prev');
  const authAuditPageNextButton = document.getElementById('auth-audit-page-next');
  const authAuditPageSummary = document.getElementById('auth-audit-page-summary');
  const syncDebugSection = document.getElementById('sync-debug-section');
  const syncDebugStatus = document.getElementById('sync-debug-status');
  const syncDebugList = document.getElementById('sync-debug-list');
  const syncDebugClearButton = document.getElementById('sync-debug-clear');
  const registeredAccountsSection = document.getElementById('registered-accounts-section');
  const registeredAccountsStatus = document.getElementById('registered-accounts-status');
  const registeredAccountsBody = document.getElementById('registered-accounts-body');

  let authAuditLogs = [];
  let authAuditLoading = false;
  let authAuditLastLoadedAt = 0;
  let authAuditLoadedForUserId = '';
  let authAuditFilterState = { result: 'all', action: 'all', user: '', from: '', to: '' };
  let authAuditPageSize = 10;
  let authAuditCurrentPage = 1;
  let registeredAccounts = [];
  let registeredAccountsLoading = false;
  let registeredAccountsLastLoadedAt = 0;
  let registeredAccountsLoadedForUserId = '';

  function setAuthAuditStatus(message, tone = 'muted') {
    if (!authAuditStatus) return;
    authAuditStatus.textContent = message;
    authAuditStatus.classList.remove('status-neutral', 'status-success', 'status-error', 'status-muted');
    authAuditStatus.classList.add(`status-${tone}`);
  }

  function setRegisteredAccountsStatus(message, tone = 'muted') {
    if (!registeredAccountsStatus) return;
    registeredAccountsStatus.textContent = message;
    registeredAccountsStatus.classList.remove('status-neutral', 'status-success', 'status-error', 'status-muted');
    registeredAccountsStatus.classList.add(`status-${tone}`);
  }

  function formatAuthAuditTimestamp(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return 'Unknown time';
    return date.toLocaleString([], {
      month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit', second: '2-digit',
    });
  }

  function normalizeAuditDateInputValue(value) {
    const normalized = String(value || '').trim();
    return /^\d{4}-\d{2}-\d{2}$/.test(normalized) ? normalized : '';
  }

  function updateAuthAuditFilterStateFromInputs() {
    authAuditFilterState = {
      result: authAuditFilterResultSelect?.value || 'all',
      action: authAuditFilterActionSelect?.value || 'all',
      user: String(authAuditFilterUserInput?.value || '').trim().toLowerCase(),
      from: normalizeAuditDateInputValue(authAuditFilterFromInput?.value || ''),
      to: normalizeAuditDateInputValue(authAuditFilterToInput?.value || ''),
    };
    authAuditCurrentPage = 1;
  }

  function resetAuthAuditFilters() {
    if (authAuditFilterResultSelect) authAuditFilterResultSelect.value = 'all';
    if (authAuditFilterActionSelect) authAuditFilterActionSelect.value = 'all';
    if (authAuditFilterUserInput) authAuditFilterUserInput.value = '';
    if (authAuditFilterFromInput) authAuditFilterFromInput.value = '';
    if (authAuditFilterToInput) authAuditFilterToInput.value = '';
    updateAuthAuditFilterStateFromInputs();
  }

  function getAuthAuditAction(entry) {
    const explicitAction = String(entry?.action || '').trim().toLowerCase();
    if (explicitAction) return explicitAction;
    const path = String(entry?.path || '').trim().toLowerCase();
    const reason = String(entry?.reason || '').trim().toLowerCase();
    const status = Number(entry?.status || 0);
    if (!entry?.success || status >= 400 || /error|failed|invalid|unable/.test(reason)) return 'error';
    if (reason.startsWith('saved game:')) return 'game-save';
    if (reason.startsWith('updated game:')) return 'game-update';
    if (reason.startsWith('deleted game:')) return 'game-delete';
    if (reason.startsWith('saved deck:')) return 'deck-save';
    if (reason.startsWith('updated deck:')) return 'deck-update';
    if (reason.startsWith('deleted deck:')) return 'deck-delete';
    if (path.startsWith('/api/session') || /authenticated|authentication|incorrect pod access code|display name is required|pod access code is required/.test(reason)) return 'authentication';
    return 'authentication';
  }

  function updateAuthAuditPaginationState() {
    const parsedPageSize = Number.parseInt(String(authAuditPageSizeSelect?.value || authAuditPageSize || '10'), 10);
    authAuditPageSize = Number.isFinite(parsedPageSize) && parsedPageSize > 0 ? parsedPageSize : 10;
    authAuditCurrentPage = Math.max(1, authAuditCurrentPage);
  }

  function getFilteredAuthAuditLogs(logs) {
    if (!Array.isArray(logs) || !logs.length) return [];
    const { result, action, user, from, to } = authAuditFilterState;
    return logs.filter((entry) => {
      const isSuccess = Boolean(entry?.success);
      if (result === 'failed' && isSuccess) return false;
      if (result === 'success' && !isSuccess) return false;
      if (action !== 'all' && getAuthAuditAction(entry) !== action) return false;
      if (user) {
        const haystack = [entry?.auth?.displayName, entry?.auth?.userId, entry?.user, entry?.normalizedUser].join(' ').toLowerCase();
        if (!haystack.includes(user)) return false;
      }
      const isoDate = String(entry?.timestamp || '').slice(0, 10);
      if ((from && (!isoDate || isoDate < from)) || (to && (!isoDate || isoDate > to))) return false;
      return true;
    });
  }

  function getPagedAuthAuditLogs(logs) {
    const visibleLogs = getFilteredAuthAuditLogs(logs);
    const totalPages = Math.max(1, Math.ceil(visibleLogs.length / authAuditPageSize));
    authAuditCurrentPage = Math.min(Math.max(1, authAuditCurrentPage), totalPages);
    const startIndex = (authAuditCurrentPage - 1) * authAuditPageSize;
    const endIndex = startIndex + authAuditPageSize;
    return { visibleLogs, pagedLogs: visibleLogs.slice(startIndex, endIndex), totalPages, startIndex, endIndex: Math.min(endIndex, visibleLogs.length) };
  }

  function updateAuthAuditPaginationControls(totalVisible = 0, totalPages = 1, startIndex = 0, endIndex = 0) {
    if (authAuditPageSizeSelect) authAuditPageSizeSelect.value = String(authAuditPageSize);
    if (authAuditPagePrevButton) authAuditPagePrevButton.disabled = authAuditCurrentPage <= 1 || totalVisible === 0;
    if (authAuditPageNextButton) authAuditPageNextButton.disabled = authAuditCurrentPage >= totalPages || totalVisible === 0;
    if (authAuditPageSummary) {
      authAuditPageSummary.textContent = totalVisible ? `Page ${authAuditCurrentPage} of ${totalPages} · ${startIndex + 1}-${endIndex} of ${totalVisible}` : 'Page 1 of 1';
    }
  }

  function updateAuthAuditStatusSummary(tone = 'neutral') {
    if (!authAuditSection || authAuditSection.hidden) return;
    const visibleCount = getFilteredAuthAuditLogs(authAuditLogs).length;
    setAuthAuditStatus(`Showing ${visibleCount} of ${authAuditLogs.length} stored authentication events.`, tone);
  }

  function renderRegisteredAccounts() {
    if (!registeredAccountsSection || !registeredAccountsBody) return;
    if (!hasSyncCredentials() || !isCurrentSyncUserAdmin()) {
      registeredAccountsSection.hidden = true;
      registeredAccountsBody.innerHTML = '';
      return;
    }
    registeredAccountsSection.hidden = false;
    if (!registeredAccounts.length) {
      registeredAccountsBody.innerHTML = '<tr><td colspan="4">No registered accounts found yet.</td></tr>';
      applyResponsiveTableLabels();
      return;
    }
    registeredAccountsBody.innerHTML = registeredAccounts.map((entry) => `<tr><td>${escapeHtml(entry?.displayName || 'Unknown')}</td><td>${escapeHtml(entry?.userId || '')}</td><td>${escapeHtml(entry?.role || 'member')}</td><td>${escapeHtml(entry?.source || 'unknown')}</td></tr>`).join('');
    applyResponsiveTableLabels();
  }

  async function refreshRegisteredAccounts({ force = false } = {}) {
    if (!registeredAccountsSection || !registeredAccountsBody || !registeredAccountsStatus) return;
    if (!hasSyncCredentials() || !isCurrentSyncUserAdmin()) {
      registeredAccountsSection.hidden = true;
      registeredAccounts = [];
      registeredAccountsLoadedForUserId = '';
      registeredAccountsLastLoadedAt = 0;
      setRegisteredAccountsStatus('Registered accounts are only visible to admins.', 'muted');
      renderRegisteredAccounts();
      return;
    }
    const currentUserId = getCurrentSyncUserId();
    const shouldRefresh = force || registeredAccountsLoadedForUserId !== currentUserId
      || !registeredAccounts.length || Date.now() - registeredAccountsLastLoadedAt >= AUTH_AUDIT_REFRESH_INTERVAL_MS;
    registeredAccountsSection.hidden = false;
    renderRegisteredAccounts();
    if (!shouldRefresh || registeredAccountsLoading) return;
    registeredAccountsLoading = true;
    setRegisteredAccountsStatus('Loading registered accounts...', 'neutral');
    try {
      const payload = await cloudRequest(REGISTERED_ACCOUNTS_ENDPOINT, { method: 'GET' });
      registeredAccounts = Array.isArray(payload?.accounts) ? payload.accounts : [];
      registeredAccountsLoadedForUserId = currentUserId;
      registeredAccountsLastLoadedAt = Date.now();
      setRegisteredAccountsStatus(`Showing ${registeredAccounts.length} registered account${registeredAccounts.length === 1 ? '' : 's'}.`, 'success');
      renderRegisteredAccounts();
    } catch (error) {
      if (error?.status === 403) {
        registeredAccounts = [];
        renderRegisteredAccounts();
        setRegisteredAccountsStatus('Connected as non-admin on the server. Sign in as Baily to view registered accounts.', 'error');
        return;
      }
      setRegisteredAccountsStatus(`Unable to load registered accounts: ${error.message}`, 'error');
    } finally {
      registeredAccountsLoading = false;
    }
  }

  function renderSyncDebugLog(entries = []) {
    if (!syncDebugList || !syncDebugStatus) return;
    if (!entries.length) {
      syncDebugStatus.textContent = 'No sync events yet.';
      syncDebugStatus.className = 'sync-status status-muted';
      syncDebugList.innerHTML = '<tr><td colspan="5" class="history-empty-state">No sync events yet.</td></tr>';
      return;
    }
    syncDebugStatus.textContent = `Last sync: ${entries[0]?.timestamp || 'unknown'}`;
    syncDebugStatus.className = 'sync-status status-success';
    syncDebugList.innerHTML = entries.map((entry) => `<tr><td>${escapeHtml(entry?.timestamp ? new Date(entry.timestamp).toLocaleString() : 'Unknown')}</td><td>${escapeHtml(entry?.event || 'Unknown event')}</td><td>${escapeHtml(String(entry?.decks ?? '-'))}</td><td>${escapeHtml(String(entry?.lists ?? '-'))}</td><td>${escapeHtml(entry?.user || 'Unknown')}</td></tr>`).join('');
    applyResponsiveTableLabels();
  }

  function renderAuthAuditLogs() {
    if (!authAuditSection || !authAuditList) return;
    if (!hasSyncCredentials() || !isCurrentSyncUserAdmin()) {
      authAuditSection.hidden = true;
      authAuditList.innerHTML = '';
      return;
    }
    authAuditSection.hidden = false;
    updateAuthAuditPaginationState();
    const { visibleLogs, pagedLogs, totalPages, startIndex, endIndex } = getPagedAuthAuditLogs(authAuditLogs);
    updateAuthAuditPaginationControls(visibleLogs.length, totalPages, startIndex, endIndex);
    if (!visibleLogs.length) {
      authAuditList.innerHTML = '<p class="history-empty-state">No login attempts logged yet.</p>';
      return;
    }
    authAuditList.innerHTML = pagedLogs.map((entry) => {
      const status = entry?.success ? 'success' : 'failed';
      const reason = entry?.reason || (entry?.success ? 'Authenticated.' : 'Authentication failed.');
      const user = entry?.auth?.displayName || entry?.user || 'Unknown user';
      const userId = entry?.auth?.userId || entry?.normalizedUser || '';
      const statusCode = Number.isFinite(Number(entry?.status)) ? Number(entry.status) : 0;
      return `<article class="auth-audit-item auth-audit-item--${status}"><div class="auth-audit-item-topline"><strong>${escapeHtml(status.toUpperCase())}</strong><span>${escapeHtml(formatAuthAuditTimestamp(entry?.timestamp))}</span></div><p class="auth-audit-item-reason">${escapeHtml(reason)}</p><div class="auth-audit-item-meta"><span class="history-item-fact">Action: ${escapeHtml(getAuthAuditAction(entry))}</span><span class="history-item-fact">User: ${escapeHtml(user)}</span>${userId ? `<span class="history-item-fact">ID: ${escapeHtml(userId)}</span>` : ''}<span class="history-item-fact">HTTP: ${escapeHtml(entry?.method || 'UNKNOWN')} ${escapeHtml(entry?.path || '/api/state')}</span>${statusCode ? `<span class="history-item-fact">Status: ${statusCode}</span>` : ''}<span class="history-item-fact">IP: ${escapeHtml(entry?.ip || 'Unknown IP')}</span>${entry?.tokenHint ? `<span class="history-item-fact">Token: ${escapeHtml(entry.tokenHint)}</span>` : ''}</div></article>`;
    }).join('');
  }

  async function refreshAuthAuditLogs({ force = false } = {}) {
    if (!authAuditSection || !authAuditList || !authAuditStatus) return;
    if (!hasSyncCredentials() || !isCurrentSyncUserAdmin()) {
      authAuditSection.hidden = true;
      authAuditLogs = [];
      authAuditLoadedForUserId = '';
      authAuditLastLoadedAt = 0;
      setAuthAuditStatus('Authentication audit logs are only visible to admins.', 'muted');
      renderAuthAuditLogs();
      return;
    }
    const currentUserId = getCurrentSyncUserId();
    const shouldRefresh = force || authAuditLoadedForUserId !== currentUserId
      || !authAuditLogs.length || Date.now() - authAuditLastLoadedAt >= AUTH_AUDIT_REFRESH_INTERVAL_MS;
    authAuditSection.hidden = false;
    renderAuthAuditLogs();
    if (!shouldRefresh || authAuditLoading) return;
    authAuditLoading = true;
    setAuthAuditStatus('Loading authentication logs...', 'neutral');
    try {
      const payload = await cloudRequest(`${AUTH_AUDIT_ENDPOINT}?limit=1000`, { method: 'GET' });
      authAuditLogs = Array.isArray(payload?.logs) ? payload.logs : [];
      authAuditCurrentPage = 1;
      authAuditLoadedForUserId = currentUserId;
      authAuditLastLoadedAt = Date.now();
      updateAuthAuditStatusSummary('success');
      renderAuthAuditLogs();
    } catch (error) {
      if (error?.status === 403) {
        authAuditSection.hidden = false;
        authAuditLogs = [];
        renderAuthAuditLogs();
        setAuthAuditStatus('Connected as non-admin on the server. Sign in as Baily with the matching pod credentials, then refresh logs.', 'error');
        return;
      }
      setAuthAuditStatus(`Unable to load authentication logs: ${error.message}`, 'error');
    } finally {
      authAuditLoading = false;
    }
  }

  function buildExportBaseName() {
    const now = new Date();
    const part = (value) => String(value).padStart(2, '0');
    return `auth-audit-${now.getFullYear()}${part(now.getMonth() + 1)}${part(now.getDate())}-${part(now.getHours())}${part(now.getMinutes())}${part(now.getSeconds())}`;
  }

  function downloadFile(filename, mimeType, content) {
    const blobUrl = URL.createObjectURL(new Blob([content], { type: mimeType }));
    const link = document.createElement('a');
    link.href = blobUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(blobUrl);
  }

  function csvCell(value) {
    return `"${String(value ?? '').replace(/"/g, '""')}"`;
  }

  function exportFilteredAuthAuditAsJson() {
    const logs = getFilteredAuthAuditLogs(authAuditLogs);
    if (!logs.length) return updateAuthAuditStatusSummary('muted');
    downloadFile(`${buildExportBaseName()}.json`, 'application/json;charset=utf-8', `${JSON.stringify({
      exportedAt: new Date().toISOString(), filters: { ...authAuditFilterState }, count: logs.length, totalAvailable: authAuditLogs.length, logs,
    }, null, 2)}\n`);
    setAuthAuditStatus(`Exported ${logs.length} authentication events to JSON.`, 'success');
  }

  function exportFilteredAuthAuditAsCsv() {
    const logs = getFilteredAuthAuditLogs(authAuditLogs);
    if (!logs.length) return updateAuthAuditStatusSummary('muted');
    const header = ['timestamp', 'action', 'result', 'reason', 'statusCode', 'method', 'path', 'displayName', 'userId', 'rawUser', 'normalizedUser', 'ip', 'tokenHint', 'authMode', 'userAgent'];
    const lines = [header.map(csvCell).join(',')];
    logs.forEach((entry) => lines.push([
      entry?.timestamp || '', getAuthAuditAction(entry), entry?.success ? 'success' : 'failed', entry?.reason || '',
      Number.isFinite(Number(entry?.status)) ? Number(entry.status) : '', entry?.method || '', entry?.path || '',
      entry?.auth?.displayName || '', entry?.auth?.userId || '', entry?.user || '', entry?.normalizedUser || '',
      entry?.ip || '', entry?.tokenHint || '', entry?.auth?.mode || '', entry?.userAgent || '',
    ].map(csvCell).join(',')));
    downloadFile(`${buildExportBaseName()}.csv`, 'text/csv;charset=utf-8', `${lines.join('\n')}\n`);
    setAuthAuditStatus(`Exported ${logs.length} authentication events to CSV.`, 'success');
  }

  function resetForUserChange() {
    authAuditLogs = [];
    authAuditLoading = false;
    authAuditLastLoadedAt = 0;
    authAuditLoadedForUserId = '';
    registeredAccounts = [];
    registeredAccountsLoading = false;
    registeredAccountsLastLoadedAt = 0;
    registeredAccountsLoadedForUserId = '';
    renderAuthAuditLogs();
    renderRegisteredAccounts();
  }

  function clearSyncDebugLogs() {
    window.clearCommanderSyncDebugLogs?.();
  }

  function refreshView() {
    updateAuthAuditFilterStateFromInputs();
    updateAuthAuditPaginationState();
    renderAuthAuditLogs();
    renderRegisteredAccounts();
    renderSyncDebugLog(window.getCommanderSyncDebugLogs?.() || []);
    updateAuthAuditStatusSummary('neutral');
  }

  authAuditRefreshButton?.addEventListener('click', () => {
    void refreshAuthAuditLogs({ force: true });
    void refreshRegisteredAccounts({ force: true });
  });
  authAuditExportJsonButton?.addEventListener('click', exportFilteredAuthAuditAsJson);
  authAuditExportCsvButton?.addEventListener('click', exportFilteredAuthAuditAsCsv);
  [authAuditFilterResultSelect, authAuditFilterActionSelect, authAuditFilterUserInput, authAuditFilterFromInput, authAuditFilterToInput]
    .forEach((input) => {
      input?.addEventListener('input', refreshView);
      input?.addEventListener('change', refreshView);
    });
  authAuditFilterClearButton?.addEventListener('click', () => {
    resetAuthAuditFilters();
    refreshView();
  });
  authAuditPageSizeSelect?.addEventListener('change', () => {
    authAuditCurrentPage = 1;
    refreshView();
  });
  authAuditPagePrevButton?.addEventListener('click', () => {
    authAuditCurrentPage -= 1;
    renderAuthAuditLogs();
  });
  authAuditPageNextButton?.addEventListener('click', () => {
    authAuditCurrentPage += 1;
    renderAuthAuditLogs();
  });
  syncDebugClearButton?.addEventListener('click', () => {
    clearSyncDebugLogs();
    renderSyncDebugLog(window.getCommanderSyncDebugLogs?.() || []);
  });
  document.addEventListener('commander:auth-changed', () => {
    resetForUserChange();
    void refreshAuthAuditLogs({ force: true });
    void refreshRegisteredAccounts({ force: true });
  });

  window.CommanderAdminLogs = {
    hasView: () => Boolean(authAuditSection || syncDebugSection || registeredAccountsSection),
    refreshAuthAuditLogs,
    refreshRegisteredAccounts,
    renderAuthAuditLogs,
    renderRegisteredAccounts,
    renderSyncDebugLog,
    updateAuthAuditFilterStateFromInputs,
    updateAuthAuditPaginationState,
    updateAuthAuditStatusSummary,
    resetAuthAuditFilters,
    resetForUserChange,
    exportFilteredAuthAuditAsJson,
    exportFilteredAuthAuditAsCsv,
  };

  void refreshAuthAuditLogs({ force: true });
  void refreshRegisteredAccounts({ force: true });
  renderSyncDebugLog(window.getCommanderSyncDebugLogs?.() || []);
})();
