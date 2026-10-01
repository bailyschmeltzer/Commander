(function () {
  const recordsForm = document.getElementById('records-form');
  const recordsTableBody = document.getElementById('records-table-body');
  const customRecordForm = document.getElementById('custom-record-form');
  const customRecordTitleInput = document.getElementById('custom-record-title');
  const customRecordUnitInput = document.getElementById('custom-record-unit');
  const customRecordValueInput = document.getElementById('custom-record-value');
  const customRecordHolderInput = document.getElementById('custom-record-holder');
  const customRecordCommanderInput = document.getElementById('custom-record-commander');
  const customRecordDateInput = document.getElementById('custom-record-date');
  const customRecordNotesInput = document.getElementById('custom-record-notes');
  const customRecordHolderMenu = document.getElementById('custom-record-holder-menu');
  const customRecordCommanderMenu = document.getElementById('custom-record-commander-menu');

  function populateLookupMenus() {
    if (recordsTableBody) {
      recordsTableBody.querySelectorAll('.record-holder-menu').forEach((menu) => {
        buildDropdownMenu(menu, knownPlayers);
      });

      recordsTableBody.querySelectorAll('.record-commander-menu').forEach((menu) => {
        buildDropdownMenu(menu, knownCommanders);
      });
    }

    if (customRecordHolderMenu) {
      buildDropdownMenu(customRecordHolderMenu, knownPlayers);
    }

    if (customRecordCommanderMenu) {
      buildDropdownMenu(customRecordCommanderMenu, knownCommanders);
    }

    attachLookupWrapperHandlers(recordsForm || document);
    attachLookupWrapperHandlers(customRecordForm || document);
  }

  function collectRecordsFromTable() {
    if (!recordsTableBody) {
      return loadRecords();
    }

    return Array.from(recordsTableBody.querySelectorAll('tr[data-record-id]'))
      .map((row, index) => {
        const isCustom = row.dataset.custom === 'true';
        const valueInput = row.querySelector('[name="value"]');
        const holderInput = row.querySelector('[name="holder"]');
        const commanderInput = row.querySelector('[name="commander"]');
        const dateField = row.querySelector('[name="date"]');
        const notesInput = row.querySelector('[name="notes"]');

        return normalizeRecordEntry({
          id: row.dataset.recordId,
          key: row.dataset.key || '',
          title: row.dataset.title || '',
          unit: row.dataset.unit || '',
          value: valueInput?.value || '',
          holder: holderInput?.value || '',
          commander: commanderInput?.value || '',
          date: dateField?.value || '',
          notes: notesInput?.value || '',
          isCustom,
        }, index);
      })
      .filter(Boolean);
  }

  function render() {
    if (!recordsTableBody) {
      return;
    }

    const records = loadRecords();
    recordsTableBody.innerHTML = records
      .map((record) => `
          <tr
            data-record-id="${escapeHtml(record.id)}"
            data-key="${escapeHtml(record.key || '')}"
            data-title="${escapeHtml(record.title)}"
            data-unit="${escapeHtml(record.unit || '')}"
            data-custom="${record.isCustom ? 'true' : 'false'}"
          >
            <td class="record-title-cell">
              <strong>${escapeHtml(record.title)}</strong>
            </td>
            <td class="record-value-cell"><input type="text" name="value" value="${escapeHtml(record.value)}" placeholder="Record" /></td>
            <td class="record-unit-cell">
              <span class="record-unit-badge">${escapeHtml(record.unit || 'open')}</span>
            </td>
            <td>
              <div class="combined-input-wrapper record-lookup-wrapper">
                <input class="lookup-input" type="text" name="holder" list="player-list" value="${escapeHtml(record.holder)}" placeholder="Player" autocomplete="off" autocapitalize="none" autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" />
                <button type="button" class="dropdown-button" title="Show players">▼</button>
                <div class="dropdown-menu record-holder-menu"></div>
              </div>
            </td>
            <td class="record-commander-cell">
              <div class="combined-input-wrapper record-lookup-wrapper">
                <input class="lookup-input" type="text" name="commander" list="commander-list" value="${escapeHtml(record.commander)}" placeholder="Commander or deck" autocomplete="off" autocapitalize="none" autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" />
                <button type="button" class="dropdown-button" title="Show commanders">▼</button>
                <div class="dropdown-menu record-commander-menu"></div>
              </div>
            </td>
            <td><input type="date" name="date" value="${escapeHtml(record.date)}" /></td>
            <td><textarea name="notes" rows="2" placeholder="How it happened, matchup, table notes...">${escapeHtml(record.notes)}</textarea></td>
          </tr>`)
      .join('');

    populateLookupMenus();
  }

  if (recordsForm) {
    recordsForm.addEventListener('submit', (event) => {
      event.preventDefault();

      const records = collectRecordsFromTable();
      saveRecords(records);
      render();
    });
  }

  if (customRecordForm) {
    customRecordForm.addEventListener('submit', async (event) => {
      event.preventDefault();

      const title = customRecordTitleInput?.value.trim() || '';
      if (!title) {
        await promptLiveAlert('Please enter a title for the custom record.', 'Unable to add custom record');
        return;
      }

      const records = collectRecordsFromTable();
      records.push({
        id: generateId(),
        key: '',
        title,
        unit: customRecordUnitInput?.value.trim() || '',
        value: customRecordValueInput?.value.trim() || '',
        holder: customRecordHolderInput?.value.trim() || '',
        commander: customRecordCommanderInput?.value.trim() || '',
        date: customRecordDateInput?.value || '',
        notes: customRecordNotesInput?.value.trim() || '',
        isCustom: true,
      });

      saveRecords(records);
      customRecordForm.reset();
      render();
    });
  }

  window.CommanderRecords = {
    hasView: () => Boolean(recordsTableBody || recordsForm),
    populateLookupMenus,
    render,
  };
})();