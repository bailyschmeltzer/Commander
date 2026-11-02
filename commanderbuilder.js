(function () {
  const api = window.CommanderBuilderPage;
  const form = document.getElementById('commander-builder-form');
  const keywordSearch = document.getElementById('commander-builder-keyword-search');
  const keywordSelection = document.getElementById('commander-builder-keyword-selection');
  const keywordStatus = document.getElementById('commander-builder-keyword-status');
  const keywordGrid = document.getElementById('commander-builder-keyword-grid');
  const rerollButton = document.getElementById('commander-builder-reroll');
  const result = document.getElementById('commander-builder-result');
  const status = document.getElementById('commander-builder-status');
  const count = document.getElementById('commander-builder-count');
  const modeInputs = Array.from(document.querySelectorAll('input[name="commander-builder-mode"]'));
  const identityPanel = document.getElementById('commander-builder-identity-panel');
  const identityCopy = document.getElementById('commander-builder-identity-copy');
  const keywordPanel = document.getElementById('commander-builder-keyword-panel');

  if (!api || !form) {
    return;
  }

  if (status) {
    status.setAttribute('role', 'status');
    status.setAttribute('aria-live', 'polite');
  }

  let keywordsLoading = false;
  let keywordsLoaded = false;
  let keywordErrorMessage = '';
  let keywordCatalog = [];
  let keywordSearchTerm = '';
  let selectedKeywords = [];
  let poolQueryKey = '';
  let poolLoading = false;
  let poolRequestId = 0;
  let lastCardName = '';

  function getMode() {
    const selectedInput = modeInputs.find((input) => input.checked);
    const value = String(selectedInput?.value || '').trim().toLowerCase();
    return value === 'keywords' || value === 'both' ? value : 'identity';
  }

  function getSelectedIdentity() {
    const selectedValues = Array.from(form.querySelectorAll('input[name="commander-color"]'))
      .filter((input) => input.checked)
      .map((input) => String(input.value || '').trim().toUpperCase())
      .filter(Boolean);
    if (!selectedValues.length) {
      return '';
    }
    if (selectedValues.includes('C')) {
      return selectedValues.length === 1 ? 'c' : '';
    }
    return ['W', 'U', 'B', 'R', 'G'].filter((code) => selectedValues.includes(code)).join('').toLowerCase();
  }

  function getQuery() {
    const mode = getMode();
    const keywords = selectedKeywords.slice();
    if (mode === 'keywords') {
      return keywords.length ? { mode, keywords } : null;
    }
    if (mode === 'both') {
      const identity = getSelectedIdentity();
      return keywords.length && identity ? { mode, keywords, identity } : null;
    }
    const identity = getSelectedIdentity();
    return identity ? { mode, identity } : null;
  }

  function getQueryKey(query) {
    if (!query || typeof query !== 'object') {
      return '';
    }
    if (query.mode === 'keywords' || query.mode === 'both') {
      const keywordKey = (query.keywords || []).map((keyword) => normalizeKeyword(keyword).toLowerCase()).join('|');
      const identityKey = String(query.identity || '').trim().toLowerCase();
      return `${query.mode}:${keywordKey}:identity:${identityKey || 'any'}`;
    }
    return `identity:${String(query.identity || '').trim().toLowerCase()}`;
  }

  function getCriteriaLabel(query) {
    if (!query || typeof query !== 'object') {
      return '';
    }
    if (query.mode === 'keywords' || query.mode === 'both') {
      const keywordLabel = (query.keywords || []).join(', ');
      const identityLabel = query.identity ? api.getCommanderIdentityLabel(query.identity) : '';
      return identityLabel ? `${keywordLabel} within ${identityLabel}` : keywordLabel;
    }
    return api.getCommanderIdentityLabel(query.identity);
  }

  function getTagLabel(query) {
    if (!query || typeof query !== 'object') {
      return '';
    }
    const criteriaLabel = getCriteriaLabel(query);
    return query.mode === 'keywords' || query.mode === 'both' ? `Keywords: ${criteriaLabel}` : criteriaLabel;
  }

  function getIdlePlaceholder(mode = getMode()) {
    if (mode === 'keywords') {
      return 'Select at least one keyword ability to get started.';
    }
    if (mode === 'both') {
      return 'Choose a color identity and at least one keyword ability to get started.';
    }
    return 'Choose a color identity to get started.';
  }

  function getIdleCount(mode = getMode()) {
    if (mode === 'keywords') {
      return 'Choose keyword abilities to load the pool.';
    }
    if (mode === 'both') {
      return 'Choose colors and keyword abilities to load the pool.';
    }
    return 'Choose colors to load the pool.';
  }

  function getIdleStatus(mode = getMode()) {
    if (mode === 'keywords') {
      return 'Select at least one keyword ability, then roll for a commander.';
    }
    if (mode === 'both') {
      return 'Choose an exact color identity and at least one keyword ability, then roll for a commander.';
    }
    return 'Choose at least one color, or select Colorless, then roll for a commander.';
  }

  function getReadyPlaceholder(query) {
    if (!query || typeof query !== 'object') {
      return getIdlePlaceholder();
    }
    return query.mode === 'keywords' || query.mode === 'both'
      ? `Roll to load a commander that matches ${getCriteriaLabel(query)}.`
      : `Roll to load a random ${getCriteriaLabel(query)} commander.`;
  }

  function getReadyCount(query) {
    if (!query || typeof query !== 'object') {
      return getIdleCount();
    }
    return query.mode === 'keywords' || query.mode === 'both'
      ? `Ready to search ${query.keywords.length} keyword abil${query.keywords.length === 1 ? 'ity' : 'ities'}.`
      : `Ready to load ${getCriteriaLabel(query)} commanders.`;
  }

  function getReadyStatus(query) {
    if (!query || typeof query !== 'object') {
      return getIdleStatus();
    }
    return query.mode === 'keywords' || query.mode === 'both'
      ? `Roll to load commanders matching ${getCriteriaLabel(query)}.`
      : `Roll to load the ${getCriteriaLabel(query)} pool.`;
  }

  function updateModeUi() {
    const mode = getMode();
    const showIdentityPanel = mode === 'identity' || mode === 'both';
    const showKeywordPanel = mode === 'keywords' || mode === 'both';
    if (identityPanel) {
      identityPanel.hidden = !showIdentityPanel;
    }
    if (identityCopy) {
      identityCopy.textContent = mode === 'both'
        ? 'This stays exact. Choosing blue and black only returns Dimir commanders that also match your selected keyword abilities.'
        : 'The selection is exact. Choosing blue and black only returns Dimir commanders, not mono-blue or Grixis options.';
    }
    if (keywordPanel) {
      keywordPanel.hidden = !showKeywordPanel;
    }
    if (showKeywordPanel) {
      void loadKeywordCatalog();
    }
  }

  function syncPreviewState() {
    const mode = getMode();
    const query = getQuery();
    const queryKey = getQueryKey(query);
    if (!query) {
      renderPlaceholder(getIdlePlaceholder(mode));
      setCount(getIdleCount(mode));
      setStatus(getIdleStatus(mode), 'muted');
      updateControls();
      return;
    }
    if (queryKey !== poolQueryKey) {
      renderPlaceholder(getReadyPlaceholder(query));
      setCount(getReadyCount(query));
      setStatus(getReadyStatus(query), 'neutral');
    }
    updateControls();
  }

  function loadCache() {
    const raw = api.readLocalStorageValue(api.cacheStorageKey);
    const cache = api.parseJsonSafe(raw || '{}', {});
    return cache && typeof cache === 'object' ? cache : {};
  }

  function saveCache(key, cards) {
    if (!key) {
      return;
    }
    const cache = loadCache();
    cache[key] = {
      fetchedAt: Date.now(),
      totalCards: Number.isFinite(Number(cards?.totalCards)) ? Number(cards.totalCards) : 0,
    };
    api.writeLocalStorageValue(api.cacheStorageKey, JSON.stringify(cache));
  }

  function getCachedCards(key, { allowExpired = false } = {}) {
    if (!key) {
      return null;
    }
    const entry = loadCache()[key];
    if (!entry || !Number.isFinite(Number(entry.totalCards))) {
      return null;
    }
    const fetchedAt = Number(entry.fetchedAt || 0);
    const isFresh = Number.isFinite(fetchedAt) && (Date.now() - fetchedAt) < api.cacheTtlMs;
    if (!allowExpired && !isFresh) {
      return null;
    }
    return { totalCards: Number(entry.totalCards) };
  }

  async function fetchSelection(query) {
    const queryKey = getQueryKey(query);
    const cachedSummary = getCachedCards(queryKey);
    const staleSummary = getCachedCards(queryKey, { allowExpired: true });
    let response;
    try {
      const requestUrl = new URL(api.commanderBuilderEndpoint, window.location.origin);
      requestUrl.searchParams.set('mode', query?.mode === 'keywords' || query?.mode === 'both' ? 'keywords' : 'identity');
      if (query?.mode === 'keywords' || query?.mode === 'both') {
        (query.keywords || []).forEach((keyword) => requestUrl.searchParams.append('keyword', keyword));
        if (query.identity) {
          requestUrl.searchParams.set('identity', query.identity);
        }
      } else {
        requestUrl.searchParams.set('identity', query?.identity || '');
      }
      requestUrl.searchParams.set('_', String(Date.now()));
      response = await fetch(requestUrl.toString(), { cache: 'no-store' });
    } catch (error) {
      if (staleSummary) {
        throw new Error('Unable to refresh the commander pick right now. Try again in a moment.');
      }
      throw error;
    }
    if (!response.ok) {
      let message = `Request failed (${response.status})`;
      try {
        const errorBody = await response.json();
        if (errorBody?.detail) {
          message = `${errorBody.error || 'Request failed'} ${errorBody.detail}`.trim();
        } else if (errorBody?.error) {
          message = errorBody.error;
        }
      } catch {}
      throw new Error(message);
    }
    const payload = await response.json();
    const selectedCard = payload?.card && payload.card.name && payload.card.scryfallUri ? payload.card : null;
    const totalCards = Number.isFinite(Number(payload?.totalCards))
      ? Number(payload.totalCards)
      : Number(cachedSummary?.totalCards || staleSummary?.totalCards || 0);
    saveCache(queryKey, { totalCards });
    return {
      card: selectedCard,
      totalCards,
      source: 'network',
      identity: String(payload?.identity || query?.identity || '').trim().toLowerCase(),
      keywords: Array.isArray(payload?.keywords) ? payload.keywords : (query?.keywords || []),
      mode: query?.mode || (payload?.mode === 'keywords' ? 'keywords' : 'identity'),
    };
  }

  function normalizeKeyword(value) {
    return String(value || '').replace(/\s+/g, ' ').trim();
  }

  function buildKeywordCatalog(payload) {
    const seen = new Set();
    return (Array.isArray(payload?.keywordAbilities) ? payload.keywordAbilities : [])
      .map((name) => ({ name: normalizeKeyword(name), category: 'ability' }))
      .filter(({ name }) => name)
      .filter(({ name }) => {
        const key = name.toLowerCase();
        if (seen.has(key)) {
          return false;
        }
        seen.add(key);
        return true;
      })
      .sort((left, right) => api.compareTextValues(left.name, right.name));
  }

  function renderKeywordSelection() {
    if (!keywordSelection) {
      return;
    }
    if (!selectedKeywords.length) {
      keywordSelection.innerHTML = '<p class="commander-builder-keyword-empty status-muted">No keyword abilities selected yet.</p>';
      return;
    }
    keywordSelection.innerHTML = selectedKeywords.map((keyword) => `
      <button type="button" class="commander-builder-selected-keyword" data-remove-commander-keyword="${api.escapeHtml(keyword)}" aria-label="Remove ${api.escapeHtml(keyword)}">
        <span>${api.escapeHtml(keyword)}</span>
        <span aria-hidden="true">×</span>
      </button>`).join('');
  }

  function renderKeywordCatalog() {
    if (!keywordGrid || !keywordStatus) {
      return;
    }
    const filteredKeywords = keywordCatalog.filter(({ name }) => name.toLowerCase().includes(keywordSearchTerm.toLowerCase()));
    if (keywordsLoading && !keywordCatalog.length) {
      keywordStatus.textContent = 'Loading keyword ability catalog...';
      keywordStatus.className = 'commander-builder-keyword-status status-muted';
      keywordGrid.innerHTML = '<p class="status-muted">Loading keyword ability catalog...</p>';
      renderKeywordSelection();
      return;
    }
    if (keywordErrorMessage && !keywordCatalog.length) {
      keywordStatus.textContent = keywordErrorMessage;
      keywordStatus.className = 'commander-builder-keyword-status status-error';
      keywordGrid.innerHTML = '<p class="status-error">Keyword ability search is unavailable right now.</p>';
      renderKeywordSelection();
      return;
    }
    if (!filteredKeywords.length) {
      keywordStatus.textContent = keywordCatalog.length
        ? 'No keyword abilities match your search.'
        : 'No keyword abilities are available right now.';
      keywordStatus.className = `commander-builder-keyword-status ${keywordCatalog.length ? 'status-muted' : 'status-error'}`;
      keywordGrid.innerHTML = '<p class="status-muted">Try a different keyword ability search.</p>';
      renderKeywordSelection();
      return;
    }
    keywordStatus.textContent = `Showing ${filteredKeywords.length} of ${keywordCatalog.length} keyword abilities.`;
    keywordStatus.className = 'commander-builder-keyword-status status-muted';
    keywordGrid.innerHTML = filteredKeywords.map(({ name, category }) => {
      const isSelected = selectedKeywords.includes(name);
      return `
        <label class="commander-builder-keyword-option${isSelected ? ' is-selected' : ''}">
          <input type="checkbox" name="commander-builder-keyword" value="${api.escapeHtml(name)}"${isSelected ? ' checked' : ''} />
          <span class="commander-builder-keyword-name">${api.escapeHtml(name)}</span>
          <span class="commander-builder-keyword-badge commander-builder-keyword-badge-${api.escapeHtml(category)}">Ability</span>
        </label>`;
    }).join('');
    renderKeywordSelection();
  }

  async function loadKeywordCatalog() {
    if (!keywordGrid || keywordsLoading || keywordsLoaded) {
      return;
    }
    keywordsLoading = true;
    keywordErrorMessage = '';
    renderKeywordCatalog();
    try {
      const requestUrl = new URL(api.keywordsEndpoint, window.location.origin);
      requestUrl.searchParams.set('_', String(Date.now()));
      const response = await fetch(requestUrl.toString(), { cache: 'no-store' });
      if (!response.ok) {
        throw new Error(`Unable to load keywords (${response.status}).`);
      }
      keywordCatalog = buildKeywordCatalog(await response.json());
      keywordsLoaded = true;
    } catch (error) {
      keywordErrorMessage = error instanceof Error ? error.message : 'Unable to load keyword catalog right now.';
    } finally {
      keywordsLoading = false;
      renderKeywordCatalog();
      updateControls();
    }
  }

  function setStatus(message, tone = 'muted') {
    if (!status) {
      return;
    }
    status.textContent = message;
    status.classList.remove('status-neutral', 'status-success', 'status-error', 'status-muted');
    status.classList.add(`status-${tone}`);
  }

  function setCount(message) {
    if (count) {
      count.textContent = message;
    }
  }

  function updateControls() {
    const query = getQuery();
    const queryKey = getQueryKey(query);
    const submitButton = form.querySelector('button[type="submit"]');
    if (submitButton) {
      submitButton.disabled = poolLoading || !query;
    }
    if (rerollButton) {
      rerollButton.disabled = poolLoading || !query || poolQueryKey !== queryKey;
    }
  }

  function renderPlaceholder(message) {
    if (result) {
      result.innerHTML = `<p>${api.escapeHtml(message)}</p>`;
    }
  }

  function formatRichText(text) {
    return api.escapeHtml(String(text || '').trim()).replace(/\n/g, '<br />');
  }

  function getStatsLine(card) {
    if (!card || typeof card !== 'object') {
      return '';
    }
    const parts = [];
    const power = String(card.power || '').trim();
    const toughness = String(card.toughness || '').trim();
    const loyalty = String(card.loyalty || '').trim();
    const defense = String(card.defense || '').trim();
    if (power || toughness) {
      parts.push(`${api.escapeHtml(power || '?')}/${api.escapeHtml(toughness || '?')}`);
    }
    if (loyalty) {
      parts.push(`Loyalty ${api.escapeHtml(loyalty)}`);
    }
    if (defense) {
      parts.push(`Defense ${api.escapeHtml(defense)}`);
    }
    return parts.join(' · ');
  }

  function getImageSources(card) {
    if (!card || typeof card !== 'object') {
      return [];
    }
    const faceCards = Array.isArray(card.cardFaces) ? card.cardFaces : [];
    const sources = [card.imageUri, card.imageLargeUri, card.imagePngUri,
      ...faceCards.flatMap((face) => [face.imageUri, face.imageLargeUri, face.imagePngUri])];
    return [...new Set(sources.map((value) => String(value || '').trim()).filter(Boolean))];
  }

  function attachImageFallback(card) {
    const image = result?.querySelector('.commander-builder-image');
    if (!image || !card) {
      return;
    }
    const sources = getImageSources(card);
    if (!sources.length) {
      image.closest('.commander-builder-media')?.remove();
      return;
    }
    let sourceIndex = sources.findIndex((source) => source === image.src);
    if (sourceIndex < 0) {
      sourceIndex = 0;
      image.src = sources[0];
    }
    image.addEventListener('error', () => {
      sourceIndex += 1;
      if (sourceIndex < sources.length) {
        image.src = sources[sourceIndex];
      } else {
        image.closest('.commander-builder-media')?.remove();
      }
    });
  }

  function toggleKeyword(keyword, isSelected) {
    const normalizedKeyword = normalizeKeyword(keyword);
    if (!normalizedKeyword) {
      return;
    }
    if (isSelected) {
      if (!selectedKeywords.includes(normalizedKeyword)) {
        selectedKeywords = [...selectedKeywords, normalizedKeyword];
      }
    } else {
      selectedKeywords = selectedKeywords.filter((value) => value !== normalizedKeyword);
    }
    renderKeywordCatalog();
  }

  function renderResultCard(card, criteriaLabel, totalCards, source) {
    if (!result || !card) {
      return;
    }
    const safeName = api.escapeHtml(card.name);
    const safeTypeLine = api.escapeHtml(card.typeLine || 'Commander-eligible card');
    const safeManaCost = api.escapeHtml(card.manaCost || 'No mana cost');
    const safeCriteriaLabel = api.escapeHtml(criteriaLabel || 'Commander');
    const safePoolSize = api.escapeHtml(String(totalCards));
    const imageSources = getImageSources(card);
    const safeImageUri = api.escapeHtml(imageSources[0] || '');
    const faceCards = Array.isArray(card.cardFaces) ? card.cardFaces : [];
    const fallbackFace = faceCards[0] || null;
    const rulesText = String(card.oracleText || fallbackFace?.oracleText || '').trim();
    const statsLine = getStatsLine(card) || getStatsLine(fallbackFace);
    const sourceLabel = source === 'network'
      ? 'Live Scryfall pool'
      : source === 'cache'
        ? 'Cached pool from this device'
        : 'Cached pool from this device due to a fetch issue';
    const imageMarkup = safeImageUri
      ? `<div class="commander-builder-media"><img class="commander-builder-image" src="${safeImageUri}" alt="${safeName}" loading="lazy" /></div>`
      : '';
    const rulesMarkup = faceCards.length > 1
      ? faceCards.map((face) => {
        const faceName = api.escapeHtml(face.name || card.name);
        const faceTypeLine = api.escapeHtml(face.typeLine || '');
        const faceManaCost = api.escapeHtml(face.manaCost || 'No mana cost');
        const faceStats = getStatsLine(face);
        const faceRulesText = String(face.oracleText || '').trim();
        return `
          <section class="commander-builder-rules-block commander-builder-face-block">
            <div class="commander-builder-rules-header">
              <h4>${faceName}</h4>
              <p class="commander-builder-meta">${faceTypeLine || 'Commander face'}</p>
            </div>
            <p class="commander-builder-meta">Mana cost: ${faceManaCost}</p>
            ${faceStats ? `<p class="commander-builder-stats">Stats: ${faceStats}</p>` : ''}
            <div class="commander-builder-rules-text">${formatRichText(faceRulesText || 'No rules text available.')}</div>
          </section>`;
      }).join('')
      : `
        <section class="commander-builder-rules-block">
          <div class="commander-builder-rules-header"><h4>Card text</h4></div>
          ${statsLine ? `<p class="commander-builder-stats">Stats: ${statsLine}</p>` : ''}
          <div class="commander-builder-rules-text">${formatRichText(rulesText || 'No rules text available.')}</div>
        </section>`;
    result.innerHTML = `
      <article class="commander-builder-result-card">
        ${imageMarkup}
        <div class="commander-builder-details">
          <p class="commander-builder-tag">${safeCriteriaLabel}</p>
          <h3>${api.buildCommanderDisplayHtml(card.name, safeName, card)}</h3>
          <p class="commander-builder-meta">${safeTypeLine}</p>
          <p class="commander-builder-meta">Mana cost: ${safeManaCost}</p>
          <p class="commander-builder-pool">Pulled from ${safePoolSize} eligible commanders.</p>
          ${rulesMarkup}
          <p class="commander-builder-source">${api.escapeHtml(sourceLabel)}</p>
          <div class="actions">
            <a href="${api.escapeHtml(card.scryfallUri)}" target="_blank" rel="noopener noreferrer">View on Scryfall</a>
            <a href="${api.escapeHtml(`deckbuilder.html?new=1&commander=${encodeURIComponent(card.name)}`)}">Build This Deck</a>
          </div>
        </div>
      </article>`;
    attachImageFallback(card);
  }

  function syncExclusiveSelection(changedInput) {
    const colorInputs = Array.from(form.querySelectorAll('input[name="commander-color"]'));
    if (!changedInput || !colorInputs.length) {
      return;
    }
    if (changedInput.value === 'C' && changedInput.checked) {
      colorInputs.forEach((input) => {
        if (input !== changedInput) {
          input.checked = false;
        }
      });
      return;
    }
    if (changedInput.checked) {
      const colorlessInput = colorInputs.find((input) => input.value === 'C');
      if (colorlessInput) {
        colorlessInput.checked = false;
      }
    }
  }

  async function runRoll() {
    const mode = getMode();
    const query = getQuery();
    if (!query) {
      renderPlaceholder(getIdlePlaceholder(mode));
      setCount(getIdleCount(mode));
      setStatus(getIdleStatus(mode), 'error');
      updateControls();
      return;
    }
    const queryKey = getQueryKey(query);
    const criteriaLabel = getCriteriaLabel(query);
    const requestId = poolRequestId + 1;
    poolRequestId = requestId;
    poolLoading = true;
    updateControls();
    renderPlaceholder('Loading commanders...');
    setStatus(
      query.mode === 'keywords' || query.mode === 'both'
        ? `Loading commanders matching ${criteriaLabel} from Scryfall...`
        : `Loading ${criteriaLabel} commanders from Scryfall...`,
      'neutral',
    );
    try {
      const payload = await fetchSelection(query);
      if (requestId !== poolRequestId) {
        return;
      }
      poolQueryKey = queryKey;
      setCount(`${payload.totalCards} eligible commanders`);
      if (!payload.card || !payload.totalCards) {
        lastCardName = '';
        renderPlaceholder(query.mode === 'keywords' || query.mode === 'both'
          ? `No commander-eligible cards were found matching ${criteriaLabel}.`
          : `No commander-eligible cards were found for ${criteriaLabel}.`);
        setStatus(query.mode === 'keywords' || query.mode === 'both'
          ? 'No commanders found for that keyword combination.'
          : 'No commanders found for that exact color identity.', 'error');
        return;
      }
      lastCardName = payload.card?.name || '';
      renderResultCard(payload.card, getTagLabel(query), payload.totalCards, payload.source);
      setStatus(`${lastCardName} selected for ${criteriaLabel}.`, 'success');
    } catch (error) {
      if (requestId !== poolRequestId) {
        return;
      }
      poolQueryKey = '';
      lastCardName = '';
      renderPlaceholder('Unable to load commanders right now. Try again in a moment.');
      setCount('Commander pool unavailable right now.');
      setStatus(error instanceof Error ? error.message : 'Unable to load commanders right now.', 'error');
    } finally {
      if (requestId === poolRequestId) {
        poolLoading = false;
        updateControls();
      }
    }
  }

  function rerollCard() {
    const query = getQuery();
    if (!query || poolQueryKey !== getQueryKey(query)) {
      setStatus('Load a commander pool for the current filter before rerolling.', 'error');
      return;
    }
    void runRoll();
  }

  function render() {
    if (!result) {
      return;
    }
    const mode = getMode();
    updateModeUi();
    renderKeywordCatalog();
    if (!result.dataset.initialized) {
      renderPlaceholder(getIdlePlaceholder());
      result.dataset.initialized = 'true';
    }
    if (status && !status.dataset.initialized) {
      setStatus(getIdleStatus(), 'muted');
      status.dataset.initialized = 'true';
    }
    if (count && !count.dataset.initialized) {
      setCount(getIdleCount());
      count.dataset.initialized = 'true';
    }
    if (mode === 'keywords' || mode === 'both') {
      void loadKeywordCatalog();
    }
    syncPreviewState();
  }

  window.CommanderBuilderController = {
    getCriteriaLabel,
    getIdleCount,
    getIdlePlaceholder,
    getIdleStatus,
    getMode,
    getQuery,
    getQueryKey,
    getReadyCount,
    getReadyPlaceholder,
    getReadyStatus,
    getSelectedIdentity,
    getSelectedKeywords: () => selectedKeywords.slice(),
    fetchSelection,
    getTagLabel,
    loadKeywordCatalog,
    normalizeKeyword,
    render,
    renderKeywordCatalog,
    renderPlaceholder,
    renderResultCard,
    rerollCard,
    runRoll,
    syncExclusiveSelection,
    syncPreviewState,
    updateControls,
    updateModeUi,
  };

  form.addEventListener('change', (event) => {
    const modeInput = event.target.closest('input[name="commander-builder-mode"]');
    if (modeInput) {
      updateModeUi();
      syncPreviewState();
      return;
    }

    const colorInput = event.target.closest('input[name="commander-color"]');
    if (colorInput) {
      syncExclusiveSelection(colorInput);
      syncPreviewState();
      return;
    }

    const keywordInput = event.target.closest('input[name="commander-builder-keyword"]');
    if (!keywordInput) {
      return;
    }

    toggleKeyword(keywordInput.value, keywordInput.checked);
    syncPreviewState();
  });

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    await runRoll();
  });

  keywordSearch?.addEventListener('input', (event) => {
    keywordSearchTerm = String(event.target.value || '').trim();
    renderKeywordCatalog();
  });

  keywordSelection?.addEventListener('click', (event) => {
    const button = event.target.closest('[data-remove-commander-keyword]');
    if (!button) {
      return;
    }

    toggleKeyword(button.dataset.removeCommanderKeyword, false);
    syncPreviewState();
  });

  rerollButton?.addEventListener('click', () => {
    rerollCard();
  });

  render();
})();
