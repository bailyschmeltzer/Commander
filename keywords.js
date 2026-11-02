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
  let reminderObserver = null;
  let scryfallRequestQueue = Promise.resolve();
  let lastScryfallRequestAt = 0;
  const scryfallReminderCache = new Map();
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

  function extractKeywordReminder(oracleText, keywordName) {
    const text = String(oracleText || '');
    const normalizedText = text.toLowerCase();
    const normalizedName = String(keywordName || '').replace(/\s+/g, ' ').trim().toLowerCase();
    const reminders = [];
    let searchFrom = 0;

    while (normalizedName) {
      const matchIndex = normalizedText.indexOf(normalizedName, searchFrom);
      if (matchIndex < 0) {
        break;
      }
      searchFrom = matchIndex + normalizedName.length;
      const before = normalizedText[matchIndex - 1] || ' ';
      const after = normalizedText[matchIndex + normalizedName.length] || ' ';
      if (/[a-z0-9]/i.test(before) || /[a-z0-9]/i.test(after)) {
        continue;
      }

      let cursor = matchIndex + normalizedName.length;
      while (true) {
        while (/\s/.test(text[cursor] || '')) {
          cursor += 1;
        }
        if (text[cursor] !== '{') {
          break;
        }
        const symbolEnd = text.indexOf('}', cursor + 1);
        if (symbolEnd < 0) {
          break;
        }
        cursor = symbolEnd + 1;
      }

      const numberMatch = text.slice(cursor).match(/^\s+\d+\b/);
      if (numberMatch) {
        cursor += numberMatch[0].length;
      }
      while (/\s/.test(text[cursor] || '')) {
        cursor += 1;
      }
      if (text[cursor] !== '(') {
        continue;
      }

      const reminderStart = cursor + 1;
      let depth = 1;
      cursor = reminderStart;
      while (cursor < text.length && depth > 0) {
        if (text[cursor] === '(') depth += 1;
        if (text[cursor] === ')') depth -= 1;
        cursor += 1;
      }
      if (depth === 0) {
        const reminder = text.slice(reminderStart, cursor - 1).replace(/\s+/g, ' ').trim();
        if (reminder) {
          reminders.push(reminder);
        }
      }
    }
    return reminders;
  }

  function getScryfallReminderText(name) {
    const cacheKey = name.toLowerCase();
    if (scryfallReminderCache.has(cacheKey)) {
      return scryfallReminderCache.get(cacheKey);
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
      const reminders = (Array.isArray(payload.data) ? payload.data : []).flatMap((card) => {
        const texts = [card.oracle_text, ...(Array.isArray(card.card_faces) ? card.card_faces.map((face) => face.oracle_text) : [])];
        return texts.flatMap((text) => extractKeywordReminder(text, name));
      });
      const reminderCounts = new Map();
      reminders.forEach((reminder) => {
        reminderCounts.set(reminder, (reminderCounts.get(reminder) || 0) + 1);
      });
      return [...reminderCounts.entries()]
        .sort((first, second) => second[1] - first[1] || second[0].length - first[0].length)[0]?.[0] || '';
    });

    scryfallRequestQueue = request.then(() => undefined, () => undefined);
    scryfallReminderCache.set(cacheKey, request);
    return request;
  }

  function renderKeywordReminder(container, reminder) {
    container.innerHTML = reminder
      ? `<p class="keyword-reminder-text">${escapeHtml(reminder)}</p>`
      : '<p class="keyword-reminder-status">No keyword reminder text found on Scryfall.</p>';
  }

  async function loadKeywordReminder(container) {
    const name = container.dataset.scryfallKeyword || '';
    if (!name) {
      return;
    }
    container.innerHTML = '<p class="keyword-reminder-status">Loading keyword reminder…</p>';
    try {
      const reminder = await getScryfallReminderText(name);
      if (container.isConnected) {
        renderKeywordReminder(container, reminder);
      }
    } catch (error) {
      scryfallReminderCache.delete(name.toLowerCase());
      if (container.isConnected) {
        container.innerHTML = '<p class="keyword-reminder-status">Keyword reminder text is unavailable right now.</p>';
      }
    }
  }

  function observeKeywordReminders() {
    reminderObserver?.disconnect();
    const containers = grid.querySelectorAll('[data-scryfall-keyword]');
    if (!containers.length) {
      return;
    }
    if (!('IntersectionObserver' in window)) {
      containers.forEach((container) => void loadKeywordReminder(container));
      return;
    }
    if (!reminderObserver) {
      reminderObserver = new IntersectionObserver((entries, observer) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            observer.unobserve(entry.target);
            void loadKeywordReminder(entry.target);
          }
        });
      }, { rootMargin: '240px' });
    }
    containers.forEach((container) => reminderObserver.observe(container));
  }

  function renderGrid() {
    reminderObserver?.disconnect();
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
          : `<div class="keyword-reminder" data-scryfall-keyword="${escapeHtml(keyword.name)}" aria-live="polite"><p class="keyword-reminder-status">Keyword reminder will load when visible.</p></div>`
        }</div>
      </article>`).join('');
    observeKeywordReminders();
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
