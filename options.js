// Options page script

document.addEventListener('DOMContentLoaded', async () => {
  const clientIdInput = document.getElementById('client-id');
  const redirectUriInput = document.getElementById('redirect-uri');
  const saveBtn = document.getElementById('save-btn');
  const clearMappingsBtn = document.getElementById('clear-mappings-btn');
  const saveStatus = document.getElementById('save-status');
  const mappingsStatus = document.getElementById('mappings-status');
  const initialsInput = document.getElementById('initials');
  const defaultReminderInput = document.getElementById('default-reminder');

  // Get redirect URI
  const redirectUri = chrome.identity.getRedirectURL();
  redirectUriInput.value = redirectUri;

  // Load saved settings
  const stored = await chrome.storage.local.get(['clientId', 'initials', 'defaultReminderMinutes']);
  if (stored.clientId) {
    clientIdInput.value = stored.clientId;
  }
  if (stored.initials) {
    initialsInput.value = stored.initials;
  }
  if (typeof stored.defaultReminderMinutes === 'number') {
    defaultReminderInput.value = stored.defaultReminderMinutes;
  }

  // Save configuration & preferences
  saveBtn.addEventListener('click', async () => {
    const clientId = clientIdInput.value.trim();
    
    if (!clientId) {
      showStatus(saveStatus, 'Please enter a Client ID', 'error');
      return;
    }

    // Validate UUID format (Azure AD Client IDs are UUIDs)
    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!uuidRegex.test(clientId)) {
      showStatus(saveStatus, 'Client ID should be a valid UUID format', 'error');
      return;
    }

    // Read preferences
    const initials = (initialsInput.value || '').trim();
    const reminderRaw = defaultReminderInput.value.trim();
    let defaultReminderMinutes = null;
    if (reminderRaw !== '') {
      const parsed = parseInt(reminderRaw, 10);
      if (Number.isNaN(parsed) || parsed < 0) {
        showStatus(saveStatus, 'Default reminder must be a non-negative number of minutes, or left blank.', 'error');
        return;
      }
      defaultReminderMinutes = parsed;
    }

    await chrome.storage.local.set({
      clientId: clientId,
      initials: initials,
      defaultReminderMinutes: defaultReminderMinutes
    });
    
    // Update background script
    chrome.runtime.sendMessage({ action: 'updateClientId', clientId: clientId });
    
    showStatus(saveStatus, 'Configuration saved successfully!', 'success');
  });

  // Clear event mappings
  clearMappingsBtn.addEventListener('click', async () => {
    if (confirm('Are you sure you want to clear all event mappings? This will prevent the extension from updating existing events.')) {
      await chrome.storage.local.remove(['eventMappings']);
      showStatus(mappingsStatus, 'Event mappings cleared', 'success');
    }
  });
});

function showStatus(element, message, type) {
  element.textContent = message;
  element.className = `status ${type}`;
  element.classList.remove('hidden');
  
  setTimeout(() => {
    element.classList.add('hidden');
  }, 5000);
}

