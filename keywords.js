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

  function renderGrid() {
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
        <p class="keyword-card-description">${keyword.description
          ? escapeHtml(keyword.description)
          : `<a class="keyword-rules-link" href="https://scryfall.com/search?q=oracle%3A%22${encodeURIComponent(keyword.name)}%22&order=edhrec" target="_blank" rel="noopener noreferrer">Search on Scryfall</a>`
        }</p>
      </article>`).join('');
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
