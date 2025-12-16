// Background service worker for handling authentication and API calls

// Microsoft Graph API configuration
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/Calendars.ReadWrite https://graph.microsoft.com/User.Read';
const MS_GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0';
const MS_AUTHORITY = 'https://login.microsoftonline.com/common';

// You'll need to register an Azure AD app and get these values.
// Prefer hard-coding the Client ID here for your organization, and optionally
// allow overrides via the Options page for development/testing.
const HARDCODED_CLIENT_ID = 'e60acb5e-d7be-4543-9864-7bc0199b5e80'; // e.g. '11111111-2222-3333-4444-555555555555'
let CLIENT_ID = HARDCODED_CLIENT_ID || null;

// Load client ID from storage (useful for development or overrides).
// If HARDCODED_CLIENT_ID is set, it will take precedence unless you change this.
chrome.storage.local.get(['clientId'], (result) => {
  if (result.clientId && !HARDCODED_CLIENT_ID) {
    CLIENT_ID = result.clientId;
  }
});

// Listen for messages from popup and content scripts
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'getClientId') {
    sendResponse({ 
      clientId: CLIENT_ID,
      isHardcoded: !!HARDCODED_CLIENT_ID,
      hardcodedClientId: HARDCODED_CLIENT_ID || null
    });
    return false;
  }

  if (request.action === 'updateClientId') {
    // Only allow override if no hardcoded client ID exists
    if (HARDCODED_CLIENT_ID) {
      sendResponse({ success: false, error: 'Client ID is hardcoded and cannot be overridden' });
      return false;
    }
    CLIENT_ID = request.clientId;
    sendResponse({ success: true });
    return false;
  }

  if (request.action === 'authenticate') {
    handleAuthentication().then(sendResponse).catch(error => {
      sendResponse({ success: false, error: error.message });
    });
    return true; // Keep channel open for async
  }

  if (request.action === 'getCalendars') {
    getCalendars(request.accessToken).then(sendResponse).catch(error => {
      sendResponse({ success: false, error: error.message });
    });
    return true;
  }

  if (request.action === 'previewDifferences') {
    previewDifferences(request.rows, request.calendarId, request.accessToken)
      .then(sendResponse)
      .catch(error => {
        sendResponse({ success: false, error: error.message });
      });
    return true;
  }

  if (request.action === 'syncToOutlook') {
    syncToOutlook(request.rows, request.calendarId, request.accessToken, request.updateExisting)
      .then(sendResponse)
      .catch(error => {
        sendResponse({ success: false, error: error.message });
      });
    return true;
  }
});

/**
 * Handle Microsoft authentication using OAuth2 implicit flow
 * Note: This is a simplified version. For production, you'd use MSAL.js library
 */
async function handleAuthentication() {
  if (!CLIENT_ID) {
    throw new Error('Client ID not configured. Please set it in the options page.');
  }

  // Generate state for security
  const state = generateRandomString();
  const redirectUri = chrome.identity.getRedirectURL();
  
  // Build authorization URL for implicit flow (access token directly in fragment)
  const authUrl = new URL('https://login.microsoftonline.com/common/oauth2/v2.0/authorize');
  authUrl.searchParams.set('client_id', CLIENT_ID);
  // Request access token directly (implicit flow)
  authUrl.searchParams.set('response_type', 'token');
  authUrl.searchParams.set('redirect_uri', redirectUri);
  authUrl.searchParams.set('scope', MS_GRAPH_SCOPE);
  authUrl.searchParams.set('state', state);
  // Tokens are returned in the fragment part for implicit flow
  authUrl.searchParams.set('response_mode', 'fragment');

  try {
    // Launch auth flow
    const responseUrl = await chrome.identity.launchWebAuthFlow({
      url: authUrl.toString(),
      interactive: true
    });

    // responseUrl will contain the access token in the URL fragment for implicit flow
    // Example: https://<redirectUri>#access_token=...&token_type=Bearer&expires_in=...
    const fragment = responseUrl.split('#')[1] || '';
    const params = new URLSearchParams(fragment);
    const accessToken = params.get('access_token');
    const returnedState = params.get('state');
    const expiresIn = parseInt(params.get('expires_in') || '0', 10);

    if (!accessToken) {
      throw new Error('No access token received');
    }

    if (returnedState !== state) {
      throw new Error('State mismatch - possible CSRF attack');
    }

    // Store token (no refresh token in implicit flow)
    await chrome.storage.local.set({
      accessToken: accessToken,
      tokenExpiry: expiresIn ? Date.now() + (expiresIn * 1000) : null
    });

    // Get user info
    const userInfo = await getUserInfo(accessToken);
    await chrome.storage.local.set({ userInfo: userInfo });

    return {
      success: true,
      accessToken: accessToken,
      userInfo: userInfo
    };
  } catch (error) {
    console.error('Authentication error:', error);
    throw error;
  }
}

// Note: With implicit flow we don't get refresh tokens, so if we wanted
// to auto-refresh we'd need to re-run the auth flow. For now we rely on
// the user re-authenticating if needed.

/**
 * Get user information
 */
async function getUserInfo(accessToken) {
  const response = await fetch(`${MS_GRAPH_ENDPOINT}/me`, {
    headers: {
      'Authorization': `Bearer ${accessToken}`
    }
  });

  if (!response.ok) {
    throw new Error('Failed to get user info');
  }

  return await response.json();
}

/**
 * Get user calendars
 */
async function getCalendars(accessToken) {
  const response = await fetch(`${MS_GRAPH_ENDPOINT}/me/calendars`, {
    headers: {
      'Authorization': `Bearer ${accessToken}`
    }
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to get calendars: ${error}`);
  }

  const data = await response.json();
  return {
    success: true,
    calendars: data.value || []
  };
}

/**
 * Sync schedule rows to Outlook calendar
 */
async function syncToOutlook(rows, calendarId, accessToken, updateExisting) {
  const summary = {
    created: 0,
    updated: 0,
    skipped: 0,
    errors: []
  };

  // Load existing event mappings and user preferences
  const stored = await chrome.storage.local.get(['eventMappings', 'initials', 'defaultReminderMinutes']);
  // Mapping key format (new): `${refNumber}|${startDateIso}`
  // This ensures each schedule row (even with same Ref # on different days)
  // is treated as a separate event.
  const eventMappings = stored.eventMappings || {};

  const prefs = {
    initials: (stored.initials || '').trim(),
    defaultReminderMinutes: typeof stored.defaultReminderMinutes === 'number'
      ? stored.defaultReminderMinutes
      : null
  };

  for (const row of rows) {
    try {
      const mappingKey = buildMappingKey(row);

      // Determine logical identifier for this row (Ref #, Job #, etc.)
      const rowId = getRowIdentifier(row);

      // Check if event already exists (from prior sync)
      let eventId = eventMappings[mappingKey];

      // If we have no mapping yet but are allowed to update, try to locate
      // an existing event in Outlook by logical ID and approximate start time.
      if (!eventId && updateExisting && rowId) {
        const found = await findEventByRef(calendarId, rowId, row.startDate, accessToken);
        if (found && found.id) {
          eventId = found.id;
          eventMappings[mappingKey] = eventId;
        }
      }

      if (eventId && updateExisting) {
        // Update existing event
        const updated = await updateCalendarEvent(
          calendarId,
          eventId,
          row,
          accessToken,
          prefs
        );

        if (updated) {
          summary.updated++;
        } else {
          // Event mapping exists but event is gone (e.g., user deleted it) – create a new one
          const newEventId = await createCalendarEvent(calendarId, row, accessToken, prefs);
          if (newEventId) {
            eventMappings[mappingKey] = newEventId;
            summary.created++;
          } else {
            summary.skipped++;
          }
        }
      } else if (eventId && !updateExisting) {
        // Skip if exists and update not requested
        summary.skipped++;
      } else {
        // No matching event found in Outlook; create a brand new one.
        const newEventId = await createCalendarEvent(calendarId, row, accessToken, prefs);

        if (newEventId) {
          eventMappings[mappingKey] = newEventId;
          summary.created++;
        } else {
          summary.skipped++;
        }
      }
    } catch (error) {
      console.error(`Error syncing row ${row.refNumber}:`, error);
      summary.errors.push({
        refNumber: row.refNumber,
        error: error.message
      });
      summary.skipped++;
    }
  }

  // Save updated mappings
  await chrome.storage.local.set({ eventMappings: eventMappings });

  return {
    success: summary.errors.length === 0,
    summary: summary
  };
}

/**
 * Preview differences between schedule rows and existing Outlook events.
 * Returns, for each row, whether an event exists and what fields differ.
 */
async function previewDifferences(rows, calendarId, accessToken) {
  const stored = await chrome.storage.local.get(['eventMappings', 'initials', 'defaultReminderMinutes']);
  const eventMappings = stored.eventMappings || {};

  const prefs = {
    initials: (stored.initials || '').trim(),
    defaultReminderMinutes: typeof stored.defaultReminderMinutes === 'number'
      ? stored.defaultReminderMinutes
      : null
  };

  const results = [];

  for (const row of rows) {
    const mappingKey = buildMappingKey(row);
    let event = null;
    let status = 'missing';
    let source = 'none';

    try {
      const existingEventId = eventMappings[mappingKey];

      if (existingEventId) {
        // Fetch the mapped event
        const resp = await fetch(
          `${MS_GRAPH_ENDPOINT}/me/calendars/${calendarId}/events/${existingEventId}`,
          {
            headers: { 'Authorization': `Bearer ${accessToken}` }
          }
        );
        if (resp.ok) {
          const candidate = await resp.json();
          // Validate that this mapping still corresponds to the same schedule row
          if (sameCalendarDay(row.startDate, candidate.start)) {
            event = candidate;
            status = 'found';
            source = 'mapping';
          } else {
            // Stale/incorrect mapping; remove it and fall back to lookup by Ref #
            delete eventMappings[mappingKey];
          }
        } else if (resp.status === 404) {
          // Event no longer exists; clear mapping
          delete eventMappings[mappingKey];
        }
      }

      const rowId = getRowIdentifier(row);

      // If mapping lookup failed, try to locate an event by logical ID and start time
      if (!event && rowId) {
        event = await findEventByRef(calendarId, rowId, row.startDate, accessToken);
        if (event && event.id) {
          status = 'found';
          source = 'search';
          // Store mapping so subsequent runs are faster and more precise
          eventMappings[mappingKey] = event.id;
        }
      }

      if (!event) {
        results.push({
          mappingKey,
          refNumber: row.refNumber || row.jobNumber || null,
          name: row.jobName || row.name || null,
          status: 'missing',
          differences: [{ field: 'event', schedule: 'exists', outlook: 'none' }]
        });
        continue;
      }

      const diffs = computeDifferences(row, event, prefs);

      results.push({
        mappingKey,
        refNumber: row.refNumber || row.jobNumber || null,
        name: row.jobName || row.name || null,
        status: diffs.length === 0 ? 'identical' : 'different',
        source,
        differences: diffs
      });
    } catch (error) {
      results.push({
        mappingKey,
        refNumber: row.refNumber || row.jobNumber || null,
        name: row.jobName || row.name || null,
        status: 'error',
        error: error.message,
        differences: []
      });
    }
  }

  // Persist any cleaned-up or newly added mappings
  await chrome.storage.local.set({ eventMappings });

  return { success: true, results };
}

/**
 * Create a calendar event from schedule row
 */
async function createCalendarEvent(calendarId, row, accessToken, prefs) {
  const event = buildEventFromRow(row, prefs);

  const response = await fetch(
    `${MS_GRAPH_ENDPOINT}/me/calendars/${calendarId}/events`,
    {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(event)
    }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to create event: ${error}`);
  }

  const createdEvent = await response.json();
  return createdEvent.id;
}

/**
 * Update an existing calendar event
 */
async function updateCalendarEvent(calendarId, eventId, row, accessToken, prefs) {
  const event = buildEventFromRow(row, prefs);

  const response = await fetch(
    `${MS_GRAPH_ENDPOINT}/me/calendars/${calendarId}/events/${eventId}`,
    {
      method: 'PATCH',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(event)
    }
  );

  if (!response.ok) {
    // If event not found, return false to indicate it should be created
    if (response.status === 404) {
      return false;
    }
    const error = await response.text();
    throw new Error(`Failed to update event: ${error}`);
  }

  return true;
}

/**
 * Build Microsoft Graph event object from schedule row
 */
function buildEventFromRow(row, prefs = {}) {
  const startDate = new Date(row.startDate);
  const endDate = new Date(row.endDate);

  // The labor schedule times are in Eastern Time (UTC-5).
  // We send the local wall-clock time and specify the time zone explicitly
  // so Outlook shows exactly the same times as the schedule.
  const timeZoneId = 'Eastern Standard Time';

  // Build base subject:
  // - Prefer Job Name when available (token laborSchedule page)
  // - Fallback to original Name/Description combo
  const baseJobName = row.jobName || row.name || 'Schedule Item';
  let subject = baseJobName;
  if (!row.jobName && row.description) {
    subject += ` - ${row.description}`;
  }

  // Prefix subject with user initials if configured
  const initials = (prefs.initials || '').trim();
  if (initials) {
    subject = `${initials} - ${subject}`;
  }

  // Build body with all details
  const bodyParts = [];
  if (row.refNumber) {
    bodyParts.push(`Ref #: ${row.refNumber}`);
  }
  if (row.jobNumber) {
    bodyParts.push(`Job #: ${row.jobNumber}`);
  }
  if (row.jobName) {
    bodyParts.push(`Job Name: ${row.jobName}`);
  }
  if (row.projectNumber) {
    bodyParts.push(`Project #: ${row.projectNumber}`);
  }
  if (row.type) {
    bodyParts.push(`Type: ${row.type}`);
  }
  if (row.description) {
    bodyParts.push(`Description: ${row.description}`);
  }
  if (row.talent) {
    bodyParts.push(`Talent: ${row.talent}`);
  }
  if (row.task) {
    bodyParts.push(`Task: ${row.task}`);
  }
  if (row.client) {
    bodyParts.push(`Client: ${row.client}`);
  }
  if (row.venueName) {
    bodyParts.push(`Venue Name: ${row.venueName}`);
  }
  if (row.venueRoom) {
    bodyParts.push(`Venue Room: ${row.venueRoom}`);
  }
  if (row.address) {
    bodyParts.push(`Address: ${row.address}`);
  }
  if (row.office) {
    bodyParts.push(`Office: ${row.office}`);
  }
  if (row.salesperson) {
    bodyParts.push(`Salesperson: ${row.salesperson}`);
  }
  if (row.orderStatus) {
    bodyParts.push(`Order Status: ${row.orderStatus}`);
  }
  if (row.status) {
    bodyParts.push(`Status: ${row.status}`);
  }
  if (row.laborCustom) {
    bodyParts.push(`Labor Custom: ${row.laborCustom}`);
  }

  const body = bodyParts.join('\n');

  const event = {
    subject: subject,
    body: {
      contentType: 'Text',
      content: body
    },
    start: {
      dateTime: formatLocalDateTime(startDate),
      timeZone: timeZoneId
    },
    end: {
      dateTime: formatLocalDateTime(endDate),
      timeZone: timeZoneId
    },
    location: {
      // Prefer Address when available, then Venue Name, then Office
      displayName: row.address || row.venueName || row.office || ''
    },
    // Store logical identifier in extended properties for matching
    singleValueExtendedProperties: [
      {
        id: 'String {66f5a359-4659-4830-9070-00047ec6ac6e} Name LightningRefNumber',
        value: getRowIdentifier(row) || ''
      }
    ]
  };

  // Apply default reminder if configured; otherwise let Outlook use calendar default
  if (typeof prefs.defaultReminderMinutes === 'number') {
    if (prefs.defaultReminderMinutes === 0) {
      event.isReminderOn = false;
      event.reminderMinutesBeforeStart = 0;
    } else if (prefs.defaultReminderMinutes > 0) {
      event.isReminderOn = true;
      event.reminderMinutesBeforeStart = prefs.defaultReminderMinutes;
    }
  }

  return event;
}

/**
 * Parse body text into a map of field names to values.
 * Handles both plain text and HTML-wrapped content.
 */
function parseBodyFields(bodyText) {
  const fields = {};
  if (!bodyText) return fields;

  let text = String(bodyText || '');

  // If the body looks like HTML, strip tags and decode entities
  if (/<[a-z][\s\S]*>/i.test(text)) {
    // Convert <br> to newlines first
    text = text.replace(/<br\s*\/?>/gi, '\n');
    // Remove all other tags
    text = text.replace(/<[^>]+>/g, ' ');
    // Decode HTML entities
    text = text
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"')
      .replace(/&#39;/g, "'");
  }

  // Normalize line endings
  text = text.replace(/\r\n/g, '\n');
  text = text.replace(/\r/g, '\n');

  // Focus on the portion starting with "Ref #:" or "Job #:" if present
  let idx = text.indexOf('Ref #:');
  if (idx === -1) {
    idx = text.indexOf('Job #:');
  }
  if (idx >= 0) {
    text = text.slice(idx);
  }

  // Parse lines like "Field Name: value"
  const lines = text.split('\n');
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    
    const match = trimmed.match(/^([^:]+):\s*(.+)$/);
    if (match) {
      const fieldName = match[1].trim();
      const fieldValue = match[2].trim();
      if (fieldName && fieldValue) {
        fields[fieldName] = fieldValue;
      }
    }
  }

  return fields;
}

/**
 * Compute differences between a schedule row and an Outlook event,
 * using the same transformation logic that will be applied on sync.
 */
function computeDifferences(row, outlookEvent, prefs = {}) {
  const desired = buildEventFromRow(row, prefs);
  const diffs = [];

  // Subject (normalize by trimming)
  const desiredSubject = (desired.subject || '').trim();
  const actualSubject = (outlookEvent.subject || '').trim();
  if (desiredSubject !== actualSubject) {
    diffs.push({
      field: 'subject',
      schedule: desiredSubject,
      outlook: actualSubject
    });
  }

  // Start / End (compare schedule Date vs Outlook dateTime/timeZone by instant)
  const desiredStart = desired.start?.dateTime || null;
  const desiredEnd = desired.end?.dateTime || null;
  const actualStart = outlookEvent.start?.dateTime || null;
  const actualEnd = outlookEvent.end?.dateTime || null;

  if (!sameInstant(row.startDate, outlookEvent.start)) {
    diffs.push({
      field: 'start',
      schedule: `${desiredStart} (${desired.start?.timeZone || ''})`,
      outlook: `${actualStart} (${outlookEvent.start?.timeZone || ''})`
    });
  }

  if (!sameInstant(row.endDate, outlookEvent.end)) {
    diffs.push({
      field: 'end',
      schedule: `${desiredEnd} (${desired.end?.timeZone || ''})`,
      outlook: `${actualEnd} (${outlookEvent.end?.timeZone || ''})`
    });
  }

  // Location
  const desiredLoc = desired.location?.displayName || '';
  const actualLoc = outlookEvent.location?.displayName || '';
  if (desiredLoc !== actualLoc) {
    diffs.push({
      field: 'location',
      schedule: desiredLoc,
      outlook: actualLoc
    });
  }

  // Body/content - compare individual fields instead of entire body
  const desiredBodyFields = parseBodyFields(desired.body?.content || '');
  const actualBodyFields = parseBodyFields(outlookEvent.body?.content || '');

  // Compare each field individually and only report differences
  const allFieldKeys = new Set([
    ...Object.keys(desiredBodyFields),
    ...Object.keys(actualBodyFields)
  ]);

  for (const key of allFieldKeys) {
    const desiredValue = (desiredBodyFields[key] || '').trim();
    const actualValue = (actualBodyFields[key] || '').trim();
    
    if (desiredValue !== actualValue) {
      diffs.push({
        field: key.toLowerCase(),
        schedule: desiredValue || '(empty)',
        outlook: actualValue || '(empty)'
      });
    }
  }

  // Reminder settings
  const desiredReminderOn = desired.isReminderOn ?? null;
  const actualReminderOn = outlookEvent.isReminderOn ?? null;
  const desiredReminderMinutes = desired.reminderMinutesBeforeStart ?? null;
  const actualReminderMinutes = outlookEvent.reminderMinutesBeforeStart ?? null;

  if (desiredReminderOn !== actualReminderOn ||
      desiredReminderMinutes !== actualReminderMinutes) {
    diffs.push({
      field: 'reminder',
      schedule: desiredReminderOn ? `${desiredReminderMinutes} min` : 'off',
      outlook: actualReminderOn ? `${actualReminderMinutes} min` : 'off'
    });
  }

  return diffs;
}

/**
 * Compare a schedule Date (local) vs a Graph dateTime/timeZone pair
 * by instant in time (epoch ms), ignoring formatting/timezone representation.
 */
function sameInstant(scheduleDate, graphDateTimeObj) {
  if (!scheduleDate && !graphDateTimeObj) return true;
  if (!scheduleDate || !graphDateTimeObj || !graphDateTimeObj.dateTime) return false;

  const schedMs = new Date(scheduleDate).getTime();

  const dt = graphDateTimeObj.dateTime;
  const tz = (graphDateTimeObj.timeZone || 'UTC').toUpperCase();

  let outlookMs;
  try {
    if (tz === 'UTC' || tz === 'UTC+00:00') {
      // Treat as UTC
      const normalized = dt.endsWith('Z') ? dt : dt + 'Z';
      outlookMs = Date.parse(normalized);
    } else {
      // Fallback: interpret as local time
      outlookMs = new Date(dt).getTime();
    }
  } catch {
    return false;
  }

  // Consider equal if within 1 minute
  return Math.abs(schedMs - outlookMs) < 60 * 1000;
}

/**
 * Check whether a schedule Date and a Graph dateTime/timeZone pair
 * represent the same calendar day (in local time).
 */
function sameCalendarDay(scheduleDate, graphDateTimeObj) {
  if (!scheduleDate || !graphDateTimeObj || !graphDateTimeObj.dateTime) return false;

  const sched = new Date(scheduleDate);
  const gdt = new Date(graphDateTimeObj.dateTime);

  return (
    sched.getFullYear() === gdt.getFullYear() &&
    sched.getMonth() === gdt.getMonth() &&
    sched.getDate() === gdt.getDate()
  );
}

function normalizeBody(body) {
  let text = String(body || '');

  // If the body looks like HTML, strip tags and decode a few common entities
  if (/<[a-z][\s\S]*>/i.test(text)) {
    // Convert <br> to newlines first
    text = text.replace(/<br\s*\/?>/gi, '\n');
    // Remove all other tags
    text = text.replace(/<[^>]+>/g, ' ');

    // Decode a few frequent HTML entities
    text = text
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"')
      .replace(/&#39;/g, "'");
  }

  // Normalize whitespace and line endings
  text = text.replace(/\r\n/g, '\n');
  text = text.replace(/[ \t]+/g, ' ');
  text = text.replace(/\s+$/gm, '');
  text = text.trim();

  // Focus comparison on the portion we control (starting at "Ref #:" if present)
  const idx = text.indexOf('Ref #:');
  if (idx >= 0) {
    text = text.slice(idx).trim();
  }

  return text;
}

function truncateForDiff(text, maxLen = 160) {
  if (text.length <= maxLen) return text;
  return text.slice(0, maxLen) + '…';
}

/**
 * Determine the logical identifier for a schedule row for matching events.
 * Prefer Ref #, then Job #, then Job Name.
 */
function getRowIdentifier(row) {
  if (row.refNumber && row.refNumber.trim()) return row.refNumber.trim();
  if (row.jobNumber && row.jobNumber.trim()) return row.jobNumber.trim();
  if (row.jobName && row.jobName.trim()) return row.jobName.trim();
  return null;
}

/**
 * Build a stable mapping key for a schedule row so each row
 * (even with the same Ref # on different dates) is its own event.
 */
function buildMappingKey(row) {
  const logicalId = getRowIdentifier(row) || 'NOID';
  const name = (row.name || '').trim();
  let startKey = '';

  try {
    if (row.startDate) {
      const d = new Date(row.startDate);
      // Use ISO string for uniqueness; time-zone conversion here doesn't
      // matter as long as it's consistent for the same row.
      startKey = d.toISOString();
    }
  } catch (e) {
    // Fallback: leave startKey empty if parsing fails
    startKey = '';
  }

  const indexPart = typeof row.index === 'number' ? `#${row.index}` : '';

  return `${logicalId}|${startKey}|${name}${indexPart}`;
}

/**
 * Try to find an existing Outlook event in the given calendar
 * that has the same LightningRefNumber extended property.
 */
async function findEventByRef(calendarId, refNumber, scheduleStartDate, accessToken) {
  const encodedRef = encodeURIComponent(refNumber);
  const extendedPropId = 'String {66f5a359-4659-4830-9070-00047ec6ac6e} Name LightningRefNumber';

  const url =
    `${MS_GRAPH_ENDPOINT}/me/calendars/${calendarId}/events` +
    `?$top=50` +
    `&$expand=singleValueExtendedProperties($filter=id eq '${encodeURIComponent(extendedPropId)}')` +
    `&$filter=singleValueExtendedProperties/Any(ep: ep/id eq '${extendedPropId}' and ep/value eq '${encodedRef}')`;

  const response = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`
    }
  });

  if (!response.ok) {
    // If the search fails, just return null so caller can fall back to creating a new event
    return null;
  }

  const data = await response.json();
  const events = data.value || [];
  if (!events.length) {
    return null;
  }

  // If we have a schedule start, only consider events that occur on the same
  // calendar day (local), then pick the one whose start is closest in time.
  if (scheduleStartDate) {
    const target = new Date(scheduleStartDate);
    const targetY = target.getFullYear();
    const targetM = target.getMonth();
    const targetD = target.getDate();

    let best = null;
    let bestDiff = Number.POSITIVE_INFINITY;

    events.forEach((ev) => {
      const evStartStr = ev.start?.dateTime;
      if (!evStartStr) return;
      const evStart = new Date(evStartStr);

      // Require same calendar day to avoid matching a different day's shift
      if (
        evStart.getFullYear() !== targetY ||
        evStart.getMonth() !== targetM ||
        evStart.getDate() !== targetD
      ) {
        return;
      }

      const diff = Math.abs(evStart.getTime() - target.getTime());
      if (diff < bestDiff) {
        bestDiff = diff;
        best = ev;
      }
    });

    // If we didn't find any event on the same calendar day, treat as not found.
    if (best) return best;
    return null;
  }

  // No schedule start provided: fall back to the first matching event
  return events[0];
}

/**
 * Format a Date as a local datetime string without timezone (YYYY-MM-DDTHH:mm:ss)
 * suitable for Microsoft Graph when accompanied by a timeZone field.
 */
function formatLocalDateTime(date) {
  const pad = (n) => (n < 10 ? '0' + n : '' + n);

  const year = date.getFullYear();
  const month = pad(date.getMonth() + 1);
  const day = pad(date.getDate());
  const hours = pad(date.getHours());
  const minutes = pad(date.getMinutes());
  const seconds = pad(date.getSeconds());

  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}`;
}

/**
 * Generate random string for state parameter
 */
function generateRandomString() {
  const array = new Uint32Array(28);
  crypto.getRandomValues(array);
  return Array.from(array, dec => ('0' + dec.toString(16)).substr(-2)).join('');
}

/**
 * Refresh access token if expired
 */
async function refreshAccessToken() {
  const stored = await chrome.storage.local.get(['refreshToken', 'tokenExpiry']);
  
  if (!stored.refreshToken) {
    throw new Error('No refresh token available');
  }

  // Check if token is still valid (with 5 minute buffer)
  if (stored.tokenExpiry && Date.now() < (stored.tokenExpiry - 5 * 60 * 1000)) {
    const tokenData = await chrome.storage.local.get(['accessToken']);
    return tokenData.accessToken;
  }

  // Refresh token
  const tokenUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
  
  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      refresh_token: stored.refreshToken,
      grant_type: 'refresh_token',
      scope: MS_GRAPH_SCOPE
    })
  });

  if (!response.ok) {
    throw new Error('Failed to refresh token');
  }

  const tokenResponse = await response.json();
  
  await chrome.storage.local.set({
    accessToken: tokenResponse.access_token,
    refreshToken: tokenResponse.refresh_token || stored.refreshToken,
    tokenExpiry: Date.now() + (tokenResponse.expires_in * 1000)
  });

  return tokenResponse.access_token;
}

