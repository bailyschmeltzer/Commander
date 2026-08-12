(function () {
  function escapeHtml(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function formatDate(iso) {
    if (!iso) return '';
    try {
      return new Date(iso + 'T00:00:00').toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric',
      });
    } catch (_) {
      return iso;
    }
  }

  function renderManaSymbols(text) {
    if (!text) return '';
    return escapeHtml(text).replace(/\{([^}]+)\}/g, (_, sym) => {
      const key = sym.toUpperCase();
      return `<span class="mana-symbol mana-${key.replace(/\//g, '')}" title="{${sym}}">{${escapeHtml(sym)}}</span>`;
    });
  }

  function renderOracleText(text) {
    if (!text) return '';
    return escapeHtml(text)
      .replace(/\{([^}]+)\}/g, (_, sym) => `<span class="mana-sym">{${escapeHtml(sym)}}</span>`)
      .replace(/\n/g, '<br>');
  }

  const searchInput = document.getElementById('rulings-search');
  const autocompleteList = document.getElementById('rulings-autocomplete');
  const form = document.getElementById('rulings-form');
  const statusEl = document.getElementById('rulings-status');
  const resultEl = document.getElementById('rulings-result');
  const imageCol = document.getElementById('rulings-image-col');
  const cardInfoEl = document.getElementById('rulings-card-info');
  const rulingsListEl = document.getElementById('rulings-list');
  let autocompleteDebounce = null;
  let selectedIndex = -1;
  let autocompleteItems = [];

  function closeAutocomplete() {
    autocompleteList.hidden = true;
    autocompleteList.innerHTML = '';
    searchInput.setAttribute('aria-expanded', 'false');
    selectedIndex = -1;
    autocompleteItems = [];
  }

  function renderAutocomplete(names) {
    autocompleteItems = names;
    selectedIndex = -1;
    if (!names.length) {
      closeAutocomplete();
      return;
    }
    autocompleteList.innerHTML = names.map((name, index) =>
      `<li role="option" class="rulings-autocomplete-item" data-index="${index}" aria-selected="false">${escapeHtml(name)}</li>`
    ).join('');
    autocompleteList.hidden = false;
    searchInput.setAttribute('aria-expanded', 'true');
  }

  function highlightAutocompleteItem(index) {
    autocompleteList.querySelectorAll('.rulings-autocomplete-item').forEach((element, itemIndex) => {
      element.classList.toggle('is-highlighted', itemIndex === index);
      element.setAttribute('aria-selected', String(itemIndex === index));
    });
  }

  async function loadRulings(name) {
    statusEl.textContent = 'Looking up card...';
    resultEl.hidden = true;

    try {
      const response = await fetch(`/api/card-rulings?name=${encodeURIComponent(name)}`);
      const data = await response.json();
      if (!response.ok || data.error) {
        statusEl.textContent = data.error || 'Card not found. Try a different spelling.';
        return;
      }

      renderResult(data.card, data.rulings);
      statusEl.textContent = '';
      resultEl.hidden = false;
      history.replaceState(null, '', `?${new URLSearchParams({ card: name })}`);
    } catch (_) {
      statusEl.textContent = 'Unable to load rulings right now. Please try again.';
    }
  }

  function renderResult(card, rulings) {
    const hasFaces = Array.isArray(card.cardFaces) && card.cardFaces.length >= 2 && card.cardFaces[0].imageUri;
    imageCol.innerHTML = '';

    if (hasFaces) {
      const wrap = document.createElement('div');
      wrap.className = 'rulings-dfc-images';
      card.cardFaces.forEach((face) => {
        if (face.imageUri) {
          const image = document.createElement('img');
          image.src = face.imageLargeUri || face.imageUri;
          image.alt = face.name || card.name;
          image.className = 'rulings-card-image';
          image.loading = 'lazy';
          wrap.appendChild(image);
        }
      });
      imageCol.appendChild(wrap);
    } else if (card.imageUri) {
      const image = document.createElement('img');
      image.src = card.imageLargeUri || card.imageUri;
      image.alt = card.name;
      image.className = 'rulings-card-image';
      image.loading = 'lazy';
      imageCol.appendChild(image);
    }

    const stats = [
      card.power && card.toughness ? `${card.power}/${card.toughness}` : '',
      card.loyalty ? `Loyalty: ${card.loyalty}` : '',
      card.defense ? `Defense: ${card.defense}` : '',
    ].filter(Boolean).join(' · ');
    const oracleHtml = card.oracleText
      ? `<p class="rulings-oracle-text">${renderOracleText(card.oracleText)}</p>`
      : hasFaces
        ? card.cardFaces.map((face) => face.oracleText ? `<div class="rulings-face-oracle"><p class="rulings-face-name">${escapeHtml(face.name)}</p><p class="rulings-oracle-text">${renderOracleText(face.oracleText)}</p>${face.power && face.toughness ? `<p class="rulings-stat">${escapeHtml(face.power)}/${escapeHtml(face.toughness)}</p>` : ''}</div>` : '').join('')
        : '';

    cardInfoEl.innerHTML = `
      <div class="rulings-card-header"><h2 class="rulings-card-name">${escapeHtml(card.name)}</h2>${card.manaCost ? `<span class="rulings-mana-cost">${renderManaSymbols(card.manaCost)}</span>` : ''}</div>
      <p class="rulings-type-line">${escapeHtml(card.typeLine)}</p>
      ${stats ? `<p class="rulings-stat">${escapeHtml(stats)}</p>` : ''}
      ${oracleHtml}
      <p class="rulings-set-line">${escapeHtml(card.setName)}${card.releasedAt ? ` · ${escapeHtml(card.releasedAt)}` : ''}${card.scryfallUri ? ` · <a href="${escapeHtml(card.scryfallUri)}" target="_blank" rel="noopener noreferrer" class="rulings-scryfall-link">View on Scryfall</a>` : ''}</p>`;

    rulingsListEl.innerHTML = rulings.length
      ? rulings.map((ruling) => `<div class="ruling-item"><p class="ruling-date">${escapeHtml(formatDate(ruling.published_at))}</p><p class="ruling-text">${escapeHtml(ruling.comment)}</p></div>`).join('')
      : '<p class="status-muted">No official rulings for this card.</p>';
  }

  searchInput.addEventListener('input', () => {
    clearTimeout(autocompleteDebounce);
    const query = searchInput.value.trim();
    if (query.length < 2) {
      closeAutocomplete();
      return;
    }
    autocompleteDebounce = setTimeout(async () => {
      try {
        const response = await fetch(`/api/deck-search?q=${encodeURIComponent(query)}`);
        if (!response.ok) return;
        const data = await response.json();
        renderAutocomplete(Array.isArray(data.results) ? data.results.map((entry) => entry.name || entry).filter(Boolean) : []);
      } catch (_) {
        closeAutocomplete();
      }
    }, 200);
  });

  searchInput.addEventListener('keydown', (event) => {
    if (autocompleteList.hidden) return;
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      selectedIndex = Math.min(selectedIndex + 1, autocompleteItems.length - 1);
      highlightAutocompleteItem(selectedIndex);
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      selectedIndex = Math.max(selectedIndex - 1, -1);
      highlightAutocompleteItem(selectedIndex);
    } else if (event.key === 'Enter' && selectedIndex >= 0) {
      event.preventDefault();
      searchInput.value = autocompleteItems[selectedIndex];
      closeAutocomplete();
      void loadRulings(searchInput.value);
    } else if (event.key === 'Escape') {
      closeAutocomplete();
    }
  });

  autocompleteList.addEventListener('mousedown', (event) => {
    const item = event.target.closest('.rulings-autocomplete-item');
    if (!item) return;
    event.preventDefault();
    searchInput.value = autocompleteItems[Number(item.dataset.index)];
    closeAutocomplete();
    void loadRulings(searchInput.value);
  });

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    closeAutocomplete();
    if (searchInput.value.trim()) void loadRulings(searchInput.value.trim());
  });

  document.addEventListener('click', (event) => {
    if (!event.target.closest('.rulings-search-wrap')) closeAutocomplete();
  });

  const cardParam = new URLSearchParams(location.search).get('card');
  if (cardParam) {
    searchInput.value = cardParam;
    void loadRulings(cardParam);
  }

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
