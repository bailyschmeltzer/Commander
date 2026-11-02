(function () {
  const CATEGORY_LABELS = {
    ability: 'Keyword Ability',
    action: 'Keyword Action',
    word: 'Ability Word',
    jargon: 'Common Term',
  };

  let allKeywords = [];
  let activeCategory = 'all';
  let searchQuery = '';
  let scryfallObserver = null;
  let scryfallRequestQueue = Promise.resolve();
  let lastScryfallRequestAt = 0;
  const scryfallDetailsCache = new Map();
  const SCRYFALL_REQUEST_INTERVAL_MS = 110;
  const grid = document.getElementById('keyword-grid');
  const status = document.getElementById('keyword-status');
  const search = document.getElementById('keyword-search');

  function escapeHtml(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function getDescriptionFor(name) {
    return KEYWORD_DESCRIPTIONS[name.toLowerCase()] || null;
  }

  function getScryfallCardDetails(name) {
    const cacheKey = name.toLowerCase();
    if (scryfallDetailsCache.has(cacheKey)) {
      return scryfallDetailsCache.get(cacheKey);
    }

    const request = scryfallRequestQueue.then(async () => {
      const waitMs = Math.max(0, SCRYFALL_REQUEST_INTERVAL_MS - (Date.now() - lastScryfallRequestAt));
      if (waitMs) {
        await new Promise((resolve) => window.setTimeout(resolve, waitMs));
      }
      lastScryfallRequestAt = Date.now();
      const escapedName = name.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
      const requestUrl = new URL('https://api.scryfall.com/cards/search');
      requestUrl.searchParams.set('q', `oracle:"${escapedName}"`);
      requestUrl.searchParams.set('order', 'edhrec');
      requestUrl.searchParams.set('unique', 'cards');
      const response = await fetch(requestUrl.toString(), {
        cache: 'force-cache',
        headers: { Accept: 'application/json' },
      });
      if (!response.ok) {
        throw new Error(`Scryfall request failed (${response.status}).`);
      }
      const payload = await response.json();
      const card = Array.isArray(payload.data) ? payload.data[0] : null;
      if (!card) {
        return null;
      }
      const firstFace = Array.isArray(card.card_faces) ? card.card_faces[0] : null;
      const oracleText = String(card.oracle_text || (Array.isArray(card.card_faces)
        ? card.card_faces.map((face) => face.oracle_text || '').filter(Boolean).join('\n\n')
        : '')).trim();
      return {
        name: String(card.name || ''),
        typeLine: String(card.type_line || firstFace?.type_line || ''),
        oracleText,
        imageUrl: String(card.image_uris?.small || firstFace?.image_uris?.small || ''),
        scryfallUri: String(card.scryfall_uri || ''),
        setName: String(card.set_name || ''),
        rarity: String(card.rarity || ''),
      };
    });

    scryfallRequestQueue = request.then(() => undefined, () => undefined);
    scryfallDetailsCache.set(cacheKey, request);
    return request;
  }

  function renderScryfallDetails(container, card) {
    if (!card) {
      container.innerHTML = '<p class="keyword-scryfall-status">No matching card details were returned.</p>';
      return;
    }
    const imageMarkup = card.imageUrl
      ? `<img class="keyword-scryfall-image" src="${escapeHtml(card.imageUrl)}" alt="${escapeHtml(card.name)}" loading="lazy" />`
      : '';
    const cardLink = card.scryfallUri
      ? `<a class="keyword-scryfall-card-name" href="${escapeHtml(card.scryfallUri)}" target="_blank" rel="noopener noreferrer">${escapeHtml(card.name)}</a>`
      : `<strong class="keyword-scryfall-card-name">${escapeHtml(card.name)}</strong>`;
    const setDetails = [card.setName, card.rarity].filter(Boolean).join(' · ');
    const rulesText = card.oracleText
      ? escapeHtml(card.oracleText).replace(/\n/g, '<br />')
      : 'No rules text available.';
    container.innerHTML = `
      <div class="keyword-scryfall-example">
        ${imageMarkup}
        <div class="keyword-scryfall-copy">
          ${cardLink}
          ${card.typeLine ? `<p class="keyword-scryfall-type">${escapeHtml(card.typeLine)}</p>` : ''}
          ${setDetails ? `<p class="keyword-scryfall-set">${escapeHtml(setDetails)}</p>` : ''}
          <p class="keyword-scryfall-oracle">${rulesText}</p>
        </div>
      </div>`;
  }

  async function loadScryfallDetails(container) {
    const name = container.dataset.scryfallKeyword || '';
    if (!name) {
      return;
    }
    container.innerHTML = '<p class="keyword-scryfall-status">Loading Scryfall card information…</p>';
    try {
      const card = await getScryfallCardDetails(name);
      if (container.isConnected) {
        renderScryfallDetails(container, card);
      }
    } catch (error) {
      scryfallDetailsCache.delete(name.toLowerCase());
      if (container.isConnected) {
        container.innerHTML = '<p class="keyword-scryfall-status">Scryfall card information is unavailable right now.</p>';
      }
    }
  }

  function observeScryfallDetails() {
    scryfallObserver?.disconnect();
    const containers = grid.querySelectorAll('[data-scryfall-keyword]');
    if (!containers.length) {
      return;
    }
    if (!('IntersectionObserver' in window)) {
      containers.forEach((container) => void loadScryfallDetails(container));
      return;
    }
    if (!scryfallObserver) {
      scryfallObserver = new IntersectionObserver((entries, observer) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            observer.unobserve(entry.target);
            void loadScryfallDetails(entry.target);
          }
        });
      }, { rootMargin: '240px' });
    }
    containers.forEach((container) => scryfallObserver.observe(container));
  }

  function renderGrid() {
    scryfallObserver?.disconnect();
    const filtered = allKeywords.filter((keyword) => {
      if (activeCategory !== 'all' && keyword.category !== activeCategory) return false;
      if (!searchQuery) return true;
      const query = searchQuery.toLowerCase();
      return keyword.name.toLowerCase().includes(query) || (keyword.description || '').toLowerCase().includes(query);
    });

    if (!filtered.length) {
      grid.innerHTML = '';
      status.textContent = allKeywords.length ? 'No keywords match your search.' : '';
      status.className = 'status-muted';
      return;
    }

    status.textContent = `Showing ${filtered.length} keyword${filtered.length === 1 ? '' : 's'}`;
    status.className = 'status-muted';
    grid.innerHTML = filtered.map((keyword) => `
      <article class="keyword-card">
        <div class="keyword-card-header">
          <strong class="keyword-card-name">${escapeHtml(keyword.name)}</strong>
          <span class="keyword-card-badge keyword-card-badge--${escapeHtml(keyword.category)}">${escapeHtml(CATEGORY_LABELS[keyword.category] || keyword.category)}</span>
        </div>
        <div class="keyword-card-description">${keyword.description
          ? escapeHtml(keyword.description)
          : `<div class="keyword-scryfall-details" data-scryfall-keyword="${escapeHtml(keyword.name)}" aria-live="polite"><p class="keyword-scryfall-status">Scryfall card information will load when visible.</p></div>`
        }</div>
      </article>`).join('');
    observeScryfallDetails();
  }

  function loadKeywordsFromLocal() {
    const actionNames = new Set(['adapt', 'amass', 'bolster', 'clash', 'discover', 'explore', 'exert', 'fight', 'investigate', 'learn', 'manifest', 'meld', 'monstrous', 'populate', 'proliferate', 'regenerate', 'scry', 'support', 'surveil', 'transform', 'venture into the dungeon', 'vote']);
    const wordNames = new Set(['addendum', 'alliance', 'battalion', 'channel', 'constellation', 'converge', 'coven', 'delirium', 'domain', 'eminence', 'enrage', 'fateful hour', 'ferocious', 'formidable', 'grandeur', 'hellbent', 'heroic', 'imprint', 'kinship', 'landfall', 'lieutenant', 'magecraft', 'metalcraft', 'morbid', 'pack tactics', 'parley', 'radiance', 'raid', 'rally', 'revolt', 'spell mastery', 'sweep', 'tempting offer', 'threshold', 'undergrowth', 'will of the council']);

    return Object.entries(KEYWORD_DESCRIPTIONS).map(([key, description]) => ({
      name: key.charAt(0).toUpperCase() + key.slice(1),
      category: actionNames.has(key) ? 'action' : wordNames.has(key) ? 'word' : 'ability',
      description,
    })).sort((a, b) => a.name.localeCompare(b.name));
  }

  async function loadKeywords() {
    try {
      const response = await fetch('/api/keywords', { cache: 'no-store' });
      if (!response.ok) throw new Error(`Request failed (${response.status})`);
      const data = await response.json();
      if (data.error) throw new Error(data.error);

      const abilities = (data.keywordAbilities || []).map((name) => ({ name, category: 'ability', description: getDescriptionFor(name) }));
      const actions = (data.keywordActions || []).map((name) => ({ name, category: 'action', description: getDescriptionFor(name) }));
      const words = (data.abilityWords || []).map((name) => ({ name, category: 'word', description: getDescriptionFor(name) }));
      const jargon = Object.entries(COMMON_TERMS).map(([key, description]) => ({ name: key.charAt(0).toUpperCase() + key.slice(1), category: 'jargon', description }));
      allKeywords = [...abilities, ...actions, ...words, ...jargon].sort((a, b) => a.name.localeCompare(b.name));
      status.textContent = '';
    } catch (_) {
      const jargon = Object.entries(COMMON_TERMS).map(([key, description]) => ({ name: key.charAt(0).toUpperCase() + key.slice(1), category: 'jargon', description }));
      allKeywords = [...loadKeywordsFromLocal(), ...jargon].sort((a, b) => a.name.localeCompare(b.name));
      status.textContent = 'Showing cached keyword data (live catalog unavailable).';
    }
    renderGrid();
  }

  search.addEventListener('input', (event) => {
    searchQuery = event.target.value.trim();
    renderGrid();
  });

  document.querySelectorAll('.keyword-tab').forEach((tab) => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.keyword-tab').forEach((item) => {
        item.classList.remove('is-active');
        item.setAttribute('aria-selected', 'false');
      });
      tab.classList.add('is-active');
      tab.setAttribute('aria-selected', 'true');
      activeCategory = tab.dataset.category;
      renderGrid();
    });
  });

  void loadKeywords();

  const pageSwitch = document.querySelector('.page-switch');
  const toggleButton = document.querySelector('.page-switch-toggle');
  const panel = document.querySelector('.page-switch-panel');
  if (pageSwitch && toggleButton && panel) {
    const currentPage = location.pathname.split('/').pop() || 'index.html';
    panel.querySelectorAll('.page-link').forEach((link) => {
      const isCurrent = (link.getAttribute('href') || '').trim().toLowerCase() === currentPage.toLowerCase();
      link.classList.toggle('is-current', isCurrent);
      if (isCurrent) link.setAttribute('aria-current', 'page');
    });
    toggleButton.addEventListener('click', () => {
      const next = !pageSwitch.classList.contains('is-open');
      pageSwitch.classList.toggle('is-open', next);
      toggleButton.setAttribute('aria-expanded', String(next));
    });
    panel.addEventListener('click', (event) => {
      if (event.target.closest('.page-link')) {
        pageSwitch.classList.remove('is-open');
        toggleButton.setAttribute('aria-expanded', 'false');
      }
    });
    document.addEventListener('click', (event) => {
      if (!pageSwitch.contains(event.target)) {
        pageSwitch.classList.remove('is-open');
        toggleButton.setAttribute('aria-expanded', 'false');
      }
    });
  }
})();
