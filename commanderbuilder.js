(function () {
  const api = window.CommanderBuilderPage;
  const form = document.getElementById('commander-builder-form');
  const keywordSearch = document.getElementById('commander-builder-keyword-search');
  const keywordSelection = document.getElementById('commander-builder-keyword-selection');
  const rerollButton = document.getElementById('commander-builder-reroll');

  if (!api || !form) {
    return;
  }

  form.addEventListener('change', (event) => {
    const modeInput = event.target.closest('input[name="commander-builder-mode"]');
    if (modeInput) {
      api.updateCommanderBuilderModeUi();
      api.syncCommanderBuilderPreviewState();
      return;
    }

    const colorInput = event.target.closest('input[name="commander-color"]');
    if (colorInput) {
      api.syncCommanderBuilderExclusiveSelection(colorInput);
      api.syncCommanderBuilderPreviewState();
      return;
    }

    const keywordInput = event.target.closest('input[name="commander-builder-keyword"]');
    if (!keywordInput) {
      return;
    }

    api.toggleCommanderBuilderKeyword(keywordInput.value, keywordInput.checked);
    api.syncCommanderBuilderPreviewState();
  });

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    await api.runCommanderBuilderRoll();
  });

  keywordSearch?.addEventListener('input', (event) => {
    api.setCommanderBuilderKeywordSearchTerm?.(String(event.target.value || '').trim());
    api.renderCommanderBuilderKeywordCatalog();
  });

  keywordSelection?.addEventListener('click', (event) => {
    const button = event.target.closest('[data-remove-commander-keyword]');
    if (!button) {
      return;
    }

    api.toggleCommanderBuilderKeyword(button.dataset.removeCommanderKeyword, false);
    api.syncCommanderBuilderPreviewState();
  });

  rerollButton?.addEventListener('click', () => {
    api.rerollCommanderBuilderCard();
  });

  api.renderCommanderBuilder();
})();
