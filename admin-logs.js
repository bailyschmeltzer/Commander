(function () {
  const api = window.CommanderAdminLogs;
  if (!api) {
    return;
  }

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
  const syncDebugClearButton = document.getElementById('sync-debug-clear');

  function refreshView() {
    api.updateAuthAuditFilterStateFromInputs();
    api.updateAuthAuditPaginationState();
    api.renderAuthAuditLogs();
    api.renderRegisteredAccounts();
    api.renderSyncDebugLog();
    api.updateAuthAuditStatusSummary('neutral');
  }

  authAuditRefreshButton?.addEventListener('click', () => {
    void api.refreshAuthAuditLogs({ force: true });
    void api.refreshRegisteredAccounts({ force: true });
  });

  authAuditExportJsonButton?.addEventListener('click', () => {
    api.exportFilteredAuthAuditAsJson();
  });

  authAuditExportCsvButton?.addEventListener('click', () => {
    api.exportFilteredAuthAuditAsCsv();
  });

  [
    authAuditFilterResultSelect,
    authAuditFilterActionSelect,
    authAuditFilterUserInput,
    authAuditFilterFromInput,
    authAuditFilterToInput,
  ].forEach((input) => {
    input?.addEventListener('input', refreshView);
    input?.addEventListener('change', refreshView);
  });

  authAuditFilterClearButton?.addEventListener('click', () => {
    api.resetAuthAuditFilters();
    refreshView();
  });

  authAuditPageSizeSelect?.addEventListener('change', () => {
    const state = api.getAuthAuditState();
    state.setCurrentPage(1);
    refreshView();
  });

  authAuditPagePrevButton?.addEventListener('click', () => {
    const state = api.getAuthAuditState();
    state.setCurrentPage(state.currentPage - 1);
    api.renderAuthAuditLogs();
  });

  authAuditPageNextButton?.addEventListener('click', () => {
    const state = api.getAuthAuditState();
    state.setCurrentPage(state.currentPage + 1);
    api.renderAuthAuditLogs();
  });

  syncDebugClearButton?.addEventListener('click', () => {
    api.getAuthAuditState().clearSyncDebugLogs();
    api.renderSyncDebugLog();
  });

  document.addEventListener('commander:auth-changed', () => {
    void api.refreshAuthAuditLogs({ force: true });
    void api.refreshRegisteredAccounts({ force: true });
  });

  // app.js performs shared bootstrap first; this controller owns the page load.
  void api.refreshAuthAuditLogs({ force: true });
  void api.refreshRegisteredAccounts({ force: true });
  api.renderSyncDebugLog();
})();
