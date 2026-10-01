(function () {
  const form = document.getElementById('deck-selector-form');
  const ownerList = document.getElementById('deck-selector-owner-list');
  const wheel = document.getElementById('deck-selector-wheel');
  const wheelDisc = document.getElementById('deck-selector-wheel-disc');
  const wheelStatus = document.getElementById('deck-selector-wheel-status');
  const results = document.getElementById('deck-selector-results');
  const submitButton = document.querySelector('#deck-selector-form button[type="submit"]');
  let spinTimer = null;
  let rotation = 0;

  function getSelectedOwners() {
    if (!ownerList) {
      return [];
    }

    return Array.from(ownerList.querySelectorAll('input[name="deck-selector-owner"]:checked'))
      .map((input) => input.value)
      .filter(Boolean);
  }

  function shuffleList(items) {
    const shuffled = [...items];
    for (let index = shuffled.length - 1; index > 0; index -= 1) {
      const swapIndex = Math.floor(Math.random() * (index + 1));
      [shuffled[index], shuffled[swapIndex]] = [shuffled[swapIndex], shuffled[index]];
    }
    return shuffled;
  }

  function getDeckSelectorPool(selectedOwners) {
    const ownerGroups = getDeckOwnerGroups();
    return shuffleList(selectedOwners.flatMap((owner) => ownerGroups[owner] || []));
  }

  function getDeckWheelPalette(count) {
    const palette = [
      '#6aa9ff',
      '#7fd4b8',
      '#ffd36d',
      '#ff9f7a',
      '#c8a8ff',
      '#8fd3ff',
      '#ffb6c7',
      '#b7df76',
      '#ffc280',
      '#8ec5a4',
    ];

    return Array.from({ length: count }, (_, index) => palette[index % palette.length]);
  }

  function polarToCartesian(centerX, centerY, radius, angleInDegrees) {
    const angleInRadians = ((angleInDegrees - 90) * Math.PI) / 180;
    return {
      x: centerX + (radius * Math.cos(angleInRadians)),
      y: centerY + (radius * Math.sin(angleInRadians)),
    };
  }

  function describeWheelSegment(startAngle, endAngle) {
    const start = polarToCartesian(50, 50, 48, endAngle);
    const end = polarToCartesian(50, 50, 48, startAngle);
    const largeArcFlag = endAngle - startAngle <= 180 ? '0' : '1';
    return [
      'M 50 50',
      `L ${start.x} ${start.y}`,
      `A 48 48 0 ${largeArcFlag} 0 ${end.x} ${end.y}`,
      'Z',
    ].join(' ');
  }

  function getDeckWheelLabelFontSize(labelLength, segmentSize, radialPathLength = 26) {
    const minFontSize = 1.55;
    const maxFontSize = 2;
    const averageRadius = 31;
    const segmentRadians = (segmentSize * Math.PI) / 180;
    const wedgeWidth = averageRadius * segmentRadians;
    const estimatedLengthUnits = (labelLength * 0.62) + 1.2;
    const lengthConstrainedSize = radialPathLength / Math.max(estimatedLengthUnits, 1);
    const wedgeConstrainedSize = wedgeWidth * 0.72;
    return Math.max(minFontSize, Math.min(maxFontSize, lengthConstrainedSize, wedgeConstrainedSize));
  }

  function truncateDeckWheelLabel(labelText, maxCharacters) {
    if (labelText.length <= maxCharacters) {
      return labelText;
    }

    if (maxCharacters <= 4) {
      return `${labelText.slice(0, Math.max(1, maxCharacters - 1))}…`;
    }

    const hardTrimmed = labelText.slice(0, maxCharacters - 1).trimEnd();
    const lastSpaceIndex = hardTrimmed.lastIndexOf(' ');
    const charactersDroppedForWordBoundary = lastSpaceIndex >= 0
      ? hardTrimmed.length - lastSpaceIndex
      : Number.POSITIVE_INFINITY;
    const trimmed = charactersDroppedForWordBoundary <= 2
      ? hardTrimmed.slice(0, lastSpaceIndex)
      : hardTrimmed;

    return `${trimmed.trim()}…`;
  }

  function fitDeckWheelLabelText(labelText, segmentSize, radialPathLength = 26) {
    return {
      fontSize: getDeckWheelLabelFontSize(labelText.length, segmentSize, radialPathLength),
      text: labelText,
    };
  }

  function getDeckWheelLabelPath(midAngle, innerRadius, outerRadius) {
    const innerPoint = polarToCartesian(50, 50, innerRadius, midAngle);
    const outerPoint = polarToCartesian(50, 50, outerRadius, midAngle);
    const isLeftSide = midAngle > 90 && midAngle < 270;
    return {
      start: isLeftSide ? outerPoint : innerPoint,
      end: isLeftSide ? innerPoint : outerPoint,
    };
  }

  function getDeckWheelSvgMarkup(deckPool) {
    const count = deckPool.length;
    const colors = getDeckWheelPalette(count);
    const segmentSize = 360 / count;

    const segments = deckPool.map((deck, index) => {
      const startAngle = index * segmentSize;
      const endAngle = startAngle + segmentSize;
      const midAngle = startAngle + (segmentSize / 2);
      const pathId = `deck-wheel-label-path-${index}`;
      const path = getDeckWheelLabelPath(midAngle, 18, 44);
      const fittedLabel = fitDeckWheelLabelText(deck.commander, segmentSize);

      return `
      <path d="${describeWheelSegment(startAngle, endAngle)}" fill="${colors[index]}" class="deck-wheel-segment" />
      <path id="${pathId}" d="M ${path.start.x} ${path.start.y} L ${path.end.x} ${path.end.y}" class="deck-wheel-label-guide" />
      <text class="deck-wheel-segment-text" data-full-label="${escapeHtml(deck.commander)}" data-segment-size="${segmentSize}" style="font-size: ${fittedLabel.fontSize.toFixed(2)}px;">
        <textPath href="#${pathId}" startOffset="8%">${escapeHtml(fittedLabel.text)}</textPath>
      </text>`;
    }).join('');

    return `<svg viewBox="0 0 100 100" class="deck-wheel-svg" aria-hidden="true">${segments}</svg>`;
  }

  function fitDeckWheelSvgLabels(rootElement) {
    const minFontSize = 1.55;
    const guaranteedLabelLength = 31;
    const guaranteedMinFontSize = 0.95;
    const labelElements = rootElement?.querySelectorAll?.('.deck-wheel-segment-text') || [];

    labelElements.forEach((labelElement) => {
      const textPath = labelElement.querySelector('textPath');
      const fullLabel = labelElement.dataset.fullLabel || textPath?.textContent || '';
      const segmentSize = Number(labelElement.dataset.segmentSize || '0');
      const pathReference = textPath?.getAttribute('href');

      if (!textPath || !fullLabel || !pathReference) {
        return;
      }

      const guidePath = rootElement.querySelector(pathReference);

      if (!guidePath || typeof guidePath.getTotalLength !== 'function' || typeof labelElement.getComputedTextLength !== 'function') {
        textPath.textContent = fullLabel;
        return;
      }

      const availableLength = guidePath.getTotalLength() * 0.9;
      const preferredFontSize = getDeckWheelLabelFontSize(fullLabel.length, segmentSize);
      textPath.textContent = fullLabel;
      labelElement.style.fontSize = `${preferredFontSize.toFixed(2)}px`;

      let renderedLength = labelElement.getComputedTextLength();
      if (renderedLength <= availableLength) {
        return;
      }

      if (fullLabel.length <= guaranteedLabelLength) {
        const guaranteedSize = Math.max(guaranteedMinFontSize, preferredFontSize * (availableLength / renderedLength));
        labelElement.style.fontSize = `${guaranteedSize.toFixed(2)}px`;
        return;
      }

      const scaledFontSize = Math.max(minFontSize, preferredFontSize * (availableLength / renderedLength));
      labelElement.style.fontSize = `${scaledFontSize.toFixed(2)}px`;
      renderedLength = labelElement.getComputedTextLength();

      if (renderedLength <= availableLength || scaledFontSize > minFontSize) {
        return;
      }

      labelElement.style.fontSize = `${minFontSize.toFixed(2)}px`;

      let bestFit = truncateDeckWheelLabel(fullLabel, 3);
      let low = 3;
      let high = fullLabel.length;

      while (low <= high) {
        const middle = Math.floor((low + high) / 2);
        const candidate = truncateDeckWheelLabel(fullLabel, middle);
        textPath.textContent = candidate;

        if (labelElement.getComputedTextLength() <= availableLength) {
          bestFit = candidate;
          low = middle + 1;
        } else {
          high = middle - 1;
        }
      }

      textPath.textContent = bestFit;
    });
  }

  function renderWheel(deckPool, centerLabel = 'Ready to Spin') {
    if (!wheel || !wheelDisc) {
      return;
    }

    if (!deckPool.length) {
      wheel.classList.add('is-empty');
      wheelDisc.innerHTML = `<div class="deck-wheel-center-label">${escapeHtml(centerLabel)}</div>`;
      wheelDisc.style.transform = 'rotate(0deg)';
      return;
    }

    wheel.classList.remove('is-empty');
    wheelDisc.innerHTML = `
    ${getDeckWheelSvgMarkup(deckPool)}
    <div class="deck-wheel-center-label">${escapeHtml(centerLabel)}</div>
  `;
    fitDeckWheelSvgLabels(wheelDisc);
    wheelDisc.style.transform = `rotate(${rotation}deg)`;
  }

  function renderResult(selectedOwners, deck) {
    if (!results) {
      return;
    }

    const safeUrl = escapeHtml(deck.url);
    const safeOwner = escapeHtml(deck.owner || 'Unassigned');
    const safePool = escapeHtml(selectedOwners.join(', '));
    const buildHref = deck.deckId
      ? `deckbuilder.html?deckId=${encodeURIComponent(deck.deckId)}`
      : `deckbuilder.html?der=${encodeURIComponent(deck.commander)}`;
    const deckUrlMarkup = deck.url
      ? `<a href="${safeUrl}" target="_blank" rel="noopener noreferrer">Open deck list</a>`
      : '<p class="status-muted">No deck URL saved for this deck yet.</p>';

    results.innerHTML = `
    <article class="deck-selector-card">
      <p class="deck-selector-owner">From pool: ${safePool}</p>
      <h3>${buildCommanderTextHtml(deck.commander)}</h3>
      <p>Owned by ${safeOwner}</p>
      ${deckUrlMarkup}
      <a href="${escapeHtml(buildHref)}" class="secondary-button">Build This Deck</a>
    </article>`;
  }

  function render() {
    if (!ownerList || !results || !wheelStatus) {
      return;
    }

    const ownerGroups = getDeckOwnerGroups();
    const owners = Object.keys(ownerGroups).sort(compareTextValues);
    const previouslySelected = new Set(getSelectedOwners());
    const shouldDefaultSelection = !previouslySelected.size;
    const currentDisplayName = normalizeIdentityLabel(getCurrentSyncDisplayName());
    const defaultOwner = owners.includes(currentDisplayName) ? currentDisplayName : '';

    if (!owners.length) {
      ownerList.innerHTML = '<p class="status-muted">Add in-rotation decks with owners to use the selector.</p>';
      results.innerHTML = '<p>Add in-rotation decks first.</p>';
      wheelStatus.textContent = 'No in-rotation decks available yet.';
      renderWheel([], 'No Decks');
      return;
    }

    ownerList.innerHTML = owners.map((owner) => {
      const checked = previouslySelected.has(owner) || (shouldDefaultSelection && defaultOwner && owner === defaultOwner) ? ' checked' : '';
      const count = ownerGroups[owner]?.length || 0;
      return `
      <label class="deck-selector-option">
        <input type="checkbox" name="deck-selector-owner" value="${escapeHtml(owner)}"${checked} />
        <span>${escapeHtml(owner)} (${count})</span>
      </label>`;
    }).join('');

    const selectedOwners = getSelectedOwners();
    if (!selectedOwners.length) {
      results.innerHTML = '<p>Select at least one player to randomize decks.</p>';
      wheelStatus.textContent = 'Select players and click Spin the wheel.';
      renderWheel([], 'Select Players');
      return;
    }

    const pooledDecks = getDeckSelectorPool(selectedOwners);
    if (!pooledDecks.length) {
      results.innerHTML = '<p>No in-rotation decks were found for the selected players.</p>';
      wheelStatus.textContent = 'No eligible decks found in the selected pool.';
      renderWheel([], 'No Decks');
      return;
    }

    results.innerHTML = '<p>Click Spin the wheel to randomize a deck from the selected players.</p>';
    wheelStatus.textContent = `${pooledDecks.length} eligible decks ready.`;
    renderWheel(pooledDecks, 'Ready to Spin');
  }

  function spin(selectedOwners) {
    if (!results || !wheelDisc) {
      return;
    }

    if (spinTimer) {
      clearTimeout(spinTimer);
      spinTimer = null;
    }

    if (!selectedOwners.length) {
      results.innerHTML = '<p>Select at least one player to randomize decks.</p>';
      if (wheelStatus) {
        wheelStatus.textContent = 'Select players and click Randomize decks.';
      }
      renderWheel([], 'Select Players');
      return;
    }

    const pooledDecks = getDeckSelectorPool(selectedOwners);
    if (!pooledDecks.length) {
      results.innerHTML = '<p>No in-rotation decks were found for the selected players.</p>';
      if (wheelStatus) {
        wheelStatus.textContent = 'No eligible decks found in the selected pool.';
      }
      renderWheel([], 'No Decks');
      return;
    }

    const winningIndex = Math.floor(Math.random() * pooledDecks.length);
    const deck = pooledDecks[winningIndex];
    if (!deck) {
      results.innerHTML = '<p>No in-rotation decks were found for the selected players.</p>';
      if (wheelStatus) {
        wheelStatus.textContent = 'No eligible decks found in the selected pool.';
      }
      return;
    }

    const segmentSize = 360 / pooledDecks.length;
    const winningCenterAngle = (winningIndex * segmentSize) + (segmentSize / 2);
    const extraTurns = 5 + Math.floor(Math.random() * 2);
    const normalizedRotation = ((rotation % 360) + 360) % 360;
    const neededOffset = ((360 - winningCenterAngle - normalizedRotation) + 360) % 360;
    const targetRotation = rotation + (extraTurns * 360) + neededOffset;
    const spinDuration = 5200;

    renderWheel(pooledDecks, 'Spinning');
    results.innerHTML = '<p>The wheel is spinning...</p>';
    if (wheelStatus) {
      wheelStatus.textContent = `Spinning through ${pooledDecks.length} eligible decks...`;
    }
    if (submitButton) {
      submitButton.disabled = true;
    }

    wheelDisc.style.transition = 'none';
    wheelDisc.style.transform = `rotate(${rotation}deg)`;
    wheelDisc.getBoundingClientRect();
    wheelDisc.style.transition = `transform ${spinDuration}ms cubic-bezier(0.16, 1, 0.3, 1)`;
    wheelDisc.style.transform = `rotate(${targetRotation}deg)`;
    rotation = targetRotation;

    spinTimer = window.setTimeout(() => {
      spinTimer = null;
      renderWheel(pooledDecks, 'Winner');
      wheelDisc.style.transition = 'none';
      wheelDisc.style.transform = `rotate(${rotation}deg)`;
      renderResult(selectedOwners, deck);
      if (wheelStatus) {
        wheelStatus.textContent = `${deck.commander} selected.`;
      }
      if (submitButton) {
        submitButton.disabled = false;
      }
    }, spinDuration + 80);
  }

  if (form) {
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      spin(getSelectedOwners());
    });
  }

  window.CommanderDeckSelector = {
    hasView: () => Boolean(results || ownerList),
    render,
  };
})();