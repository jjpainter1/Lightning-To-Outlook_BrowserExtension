// Popup script for user interface interactions

document.addEventListener('DOMContentLoaded', async () => {
  const loginBtn = document.getElementById('login-btn');
  const syncBtn = document.getElementById('sync-btn');
  const getRowsBtn = document.getElementById('get-rows-btn');
  const calendarSelect = document.getElementById('calendar-select');
  const refreshCalendarsBtn = document.getElementById('refresh-calendars-btn');
  const statusDiv = document.getElementById('status');
  const rowsPreview = document.getElementById('rows-preview');
  const diffPreview = document.getElementById('diff-preview');
  const optionsLink = document.getElementById('options-link');
  const darkModeToggle = document.getElementById('dark-mode-toggle');
  const syncSpinner = document.getElementById('sync-spinner');

  let selectedRows = [];
  let accessToken = null;
  let userCalendars = [];
  let lastDiffResults = null;

  // Load theme + auth state on load
  await loadThemePreference();
  // Ensure spinner is hidden and not animating on initial load
  if (syncSpinner) {
    syncSpinner.classList.add('hidden');
    syncSpinner.classList.remove('active');
  }
  await checkAuthStatus();
  await loadCalendars();

  // Event listeners
  loginBtn.addEventListener('click', handleLogin);
  syncBtn.addEventListener('click', handleSync);
  getRowsBtn.addEventListener('click', handleGetRows);
  refreshCalendarsBtn.addEventListener('click', loadCalendars);
  optionsLink.addEventListener('click', (e) => {
    e.preventDefault();
    chrome.runtime.openOptionsPage();
  });
  darkModeToggle.addEventListener('change', handleThemeToggle);

  /**
   * Check if user is authenticated
   */
  async function checkAuthStatus() {
    const result = await chrome.storage.local.get(['accessToken', 'userInfo']);
    
    if (result.accessToken) {
      accessToken = result.accessToken;
      updateAuthUI(true, result.userInfo);
    } else {
      updateAuthUI(false);
    }
  }

  async function loadThemePreference() {
    const stored = await chrome.storage.local.get(['popupDarkMode']);
    const isDark = !!stored.popupDarkMode;
    document.body.classList.toggle('dark', isDark);
    if (darkModeToggle) {
      darkModeToggle.checked = isDark;
    }
  }

  async function handleThemeToggle() {
    const isDark = !!darkModeToggle.checked;
    document.body.classList.toggle('dark', isDark);
    await chrome.storage.local.set({ popupDarkMode: isDark });
  }

  /**
   * Update authentication UI
   */
  function updateAuthUI(isAuthenticated, userInfo = null) {
    const authStatus = document.getElementById('auth-status');
    
    if (isAuthenticated) {
      loginBtn.textContent = 'Sign Out';
      loginBtn.onclick = handleLogout;
      authStatus.textContent = userInfo ? `Signed in as ${userInfo.displayName || userInfo.userPrincipalName}` : 'Signed in';
      authStatus.classList.remove('hidden');
      calendarSelect.disabled = false;
      refreshCalendarsBtn.disabled = false;
    } else {
      loginBtn.textContent = 'Sign in with Microsoft';
      loginBtn.onclick = handleLogin;
      authStatus.classList.add('hidden');
      calendarSelect.disabled = true;
      refreshCalendarsBtn.disabled = true;
    }
  }

  /**
   * Handle login
   */
  async function handleLogin() {
    try {
      showStatus('Initiating Microsoft sign-in...', 'info');
      
      // Send message to background script to start auth flow
      const response = await chrome.runtime.sendMessage({ action: 'authenticate' });
      
      if (response.success) {
        accessToken = response.accessToken;
        await checkAuthStatus();
        await loadCalendars();
        showStatus('Successfully signed in!', 'success');
      } else {
        showStatus(`Authentication failed: ${response.error}`, 'error');
      }
    } catch (error) {
      console.error('Login error:', error);
      showStatus(`Login error: ${error.message}`, 'error');
    }
  }

  /**
   * Handle logout
   */
  async function handleLogout() {
    await chrome.storage.local.remove(['accessToken', 'refreshToken', 'userInfo']);
    accessToken = null;
    updateAuthUI(false);
    calendarSelect.innerHTML = '<option value="">Select a calendar</option>';
    showStatus('Signed out successfully', 'info');
  }

  /**
   * Load user calendars
   */
  async function loadCalendars() {
    if (!accessToken) {
      calendarSelect.innerHTML = '<option value="">Please sign in first</option>';
      return;
    }

    try {
      calendarSelect.innerHTML = '<option value="">Loading calendars...</option>';
      calendarSelect.disabled = true;

      const response = await chrome.runtime.sendMessage({ 
        action: 'getCalendars',
        accessToken: accessToken 
      });

      if (response.success && response.calendars) {
        userCalendars = response.calendars;
        
        // Get saved calendar preference
        const saved = await chrome.storage.local.get(['selectedCalendarId']);
        
        calendarSelect.innerHTML = '<option value="">Select a calendar</option>';
        userCalendars.forEach(cal => {
          const option = document.createElement('option');
          option.value = cal.id;
          option.textContent = cal.name;
          if (saved.selectedCalendarId === cal.id) {
            option.selected = true;
          }
          calendarSelect.appendChild(option);
        });

        // If no saved preference, select default calendar
        if (!saved.selectedCalendarId) {
          const defaultCal = userCalendars.find(cal => cal.isDefault);
          if (defaultCal) {
            calendarSelect.value = defaultCal.id;
          }
        }

        calendarSelect.disabled = false;
      } else {
        calendarSelect.innerHTML = '<option value="">Error loading calendars</option>';
        showStatus(`Error: ${response.error || 'Failed to load calendars'}`, 'error');
      }
    } catch (error) {
      console.error('Error loading calendars:', error);
      calendarSelect.innerHTML = '<option value="">Error loading calendars</option>';
      showStatus(`Error loading calendars: ${error.message}`, 'error');
    }
  }

  /**
   * Get selected rows from schedule page
   */
  async function handleGetRows() {
    try {
      // Get active tab
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      
      const url = tab.url || '';
      const onOldSchedule = url.includes('prestigeav.ielightning.net/reports/general/mySchedule');
      const onNewSchedule = url.includes('prestigeav.ielightning.net/laborSchedule');

      if (!onOldSchedule && !onNewSchedule) {
        showStatus('Please navigate to the schedule page first', 'error');
        return;
      }

      // Send message to content script
      const response = await chrome.tabs.sendMessage(tab.id, { action: 'getSelectedRows' });
      
      if (response && response.success) {
        selectedRows = response.data || [];
        
        if (selectedRows.length === 0) {
          showStatus('No rows selected. Please check some rows on the schedule page.', 'info');
          rowsPreview.classList.add('hidden');
          diffPreview.classList.add('hidden');
        } else {
          displayRowsPreview(selectedRows);
          showStatus(`Found ${selectedRows.length} row(s). Comparing with Outlook...`, 'info');
          await previewDifferences();
          updateSyncButton();
        }
      } else {
        showStatus('Could not read schedule data. Make sure you are on the schedule page.', 'error');
      }
    } catch (error) {
      console.error('Error getting rows:', error);
      showStatus(`Error: ${error.message}`, 'error');
    }
  }

  /**
   * Display preview of selected rows
   */
  function displayRowsPreview(rows) {
    rowsPreview.innerHTML = '';
    rowsPreview.classList.remove('hidden');

    rows.forEach((row, index) => {
      const displayName = row.jobName || row.name || 'Unnamed';
      const refLabel = row.refNumber || row.jobNumber || 'N/A';
      const descParts = [];
      if (row.talent) descParts.push(row.talent);
      if (row.task) descParts.push(row.task);
      const descriptionText = descParts.join(' | ') || row.description || 'No description';

      const rowDiv = document.createElement('div');
      rowDiv.className = 'row-item';
      rowDiv.innerHTML = `
        <strong>${displayName}</strong><br>
        <small>Ref #: ${refLabel} | ${descriptionText}</small><br>
        <small>${formatDate(row.startDate)} - ${formatDate(row.endDate)}</small>
      `;
      rowsPreview.appendChild(rowDiv);
    });
  }

  /**
   * Ask background script to compare selected rows with Outlook events
   * and render a human-readable diff summary.
   */
  async function previewDifferences() {
    diffPreview.innerHTML = '';
    diffPreview.classList.add('hidden');
    lastDiffResults = null;

    if (!accessToken || !calendarSelect.value || selectedRows.length === 0) {
      return;
    }

    try {
      const response = await chrome.runtime.sendMessage({
        action: 'previewDifferences',
        rows: selectedRows,
        calendarId: calendarSelect.value,
        accessToken: accessToken
      });

      if (!response || !response.success) {
        console.warn('previewDifferences failed', response?.error);
        return;
      }

      lastDiffResults = response.results || [];

      if (!lastDiffResults.length) {
        return;
      }

      const groups = document.createElement('div');

      lastDiffResults.forEach((res) => {
        const groupDiv = document.createElement('div');
        groupDiv.className = 'diff-group';

        const title = document.createElement('div');
        title.className = 'diff-group-title';
        const statusLabel =
          res.status === 'identical'
            ? '✓ In sync (skipped)'
            : res.status === 'missing'
            ? '＋ Will be created'
            : res.status === 'different'
            ? '⟳ Will be updated'
            : '⚠ Error';

        title.textContent = `${res.name || 'Unnamed'} (Ref # ${res.refNumber || 'N/A'}) — ${statusLabel}`;
        groupDiv.appendChild(title);

        if (res.status === 'error' && res.error) {
          const errDiv = document.createElement('div');
          errDiv.className = 'diff-item';
          errDiv.textContent = `Error: ${res.error}`;
          groupDiv.appendChild(errDiv);
        } else if (res.differences && res.differences.length > 0) {
          res.differences.forEach((d) => {
            const item = document.createElement('div');
            item.className = 'diff-item';
            item.innerHTML = `<span class=\"label\">${d.field}:</span> schedule → ${escapeHtml(
              d.schedule ?? ''
            )}, outlook → ${escapeHtml(d.outlook ?? '')}`;
            groupDiv.appendChild(item);
          });
        }

        groups.appendChild(groupDiv);
      });

      diffPreview.innerHTML = '';
      diffPreview.appendChild(groups);
      diffPreview.classList.remove('hidden');

      const identicalCount = lastDiffResults.filter((r) => r.status === 'identical').length;
      if (identicalCount > 0) {
        showStatus(
          `Comparison complete. ${identicalCount} event(s) already match Outlook and will be skipped.`,
          'info'
        );
      } else {
        showStatus('Comparison complete. All selected rows will create or update events.', 'info');
      }
    } catch (error) {
      console.error('previewDifferences error:', error);
    }
  }

  /**
   * Format date for display
   */
  function formatDate(date) {
    if (!date) return 'N/A';
    try {
      const d = new Date(date);
      return d.toLocaleString();
    } catch {
      return date.toString();
    }
  }

  /**
   * Handle sync to Outlook
   */
  async function handleSync() {
    if (selectedRows.length === 0) {
      showStatus('No rows selected. Click "Get Selected Rows" first.', 'error');
      return;
    }

    if (!accessToken) {
      showStatus('Please sign in first', 'error');
      return;
    }

    const calendarId = calendarSelect.value;
    if (!calendarId) {
      showStatus('Please select a calendar', 'error');
      return;
    }

    const updateExisting = document.getElementById('update-existing').checked;

    try {
      // If we have a diff preview, only send rows that are not identical.
      // previewDifferences returns results in the same order as rows were sent,
      // so we can safely use the index to align rows with diff results.
      let rowsToSync = selectedRows;
      let skippedCountFromDiff = 0;
      if (lastDiffResults && lastDiffResults.length === selectedRows.length) {
        rowsToSync = selectedRows.filter((row, idx) => {
          const res = lastDiffResults[idx];
          if (res && res.status === 'identical') {
            skippedCountFromDiff += 1;
            return false;
          }
          return true;
        });
      }

      // If everything is already in sync, don't call the backend at all; just
      // show a summary that reflects the skipped identical events.
      if (rowsToSync.length === 0) {
        const total = selectedRows.length;
        const message = `Sync complete! Created: 0, Updated: 0, Skipped: ${skippedCountFromDiff || total}`;
        showStatus(message, 'success');
        return;
      }

      syncBtn.disabled = true;
      document.getElementById('sync-btn-text').textContent = 'Syncing...';
      if (syncSpinner) {
        syncSpinner.classList.remove('hidden');
        syncSpinner.classList.add('active');
      }
      showStatus('Syncing to Outlook...', 'info');

      const response = await chrome.runtime.sendMessage({
        action: 'syncToOutlook',
        rows: rowsToSync,
        calendarId: calendarId,
        accessToken: accessToken,
        updateExisting: updateExisting
      });

      if (response.success) {
        const summary = response.summary || {};
        const message = `Sync complete! Created: ${summary.created || 0}, Updated: ${summary.updated || 0}, Skipped: ${summary.skipped || 0}`;
        showStatus(message, 'success');
      } else {
        showStatus(`Sync failed: ${response.error}`, 'error');
      }
    } catch (error) {
      console.error('Sync error:', error);
      showStatus(`Sync error: ${error.message}`, 'error');
    } finally {
      syncBtn.disabled = false;
      document.getElementById('sync-btn-text').textContent = 'Sync to Outlook';
      if (syncSpinner) {
        syncSpinner.classList.add('hidden');
        syncSpinner.classList.remove('active');
      }
    }
  }

  /**
   * Update sync button state
   */
  function updateSyncButton() {
    const canSync = selectedRows.length > 0 && accessToken && calendarSelect.value;
    syncBtn.disabled = !canSync;
  }

  // Update sync button when calendar selection changes
  calendarSelect.addEventListener('change', () => {
    updateSyncButton();
    // Save calendar preference
    chrome.storage.local.set({ selectedCalendarId: calendarSelect.value });
  });

  /**
   * Show status message
   */
  function showStatus(message, type = 'info') {
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    statusDiv.classList.remove('hidden');
    
    // Auto-hide after 5 seconds for success/info
    if (type !== 'error') {
      setTimeout(() => {
        statusDiv.classList.add('hidden');
      }, 5000);
    }
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
});

