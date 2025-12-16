# Chrome Web Store Permission Justifications

## Storage Permission

**Justification:**
The extension requires the `storage` permission to persist user data and authentication state locally on the user's device. Specifically, the extension stores:

1. **OAuth Access Tokens**: Microsoft Graph API access tokens are stored securely in `chrome.storage.local` to maintain user authentication sessions. This prevents users from having to re-authenticate every time they use the extension.

2. **User Preferences**: User-configured settings including:
   - User initials (for event title prefixes)
   - Default reminder minutes for calendar events
   - Selected calendar preference

3. **Event Mappings**: A mapping between schedule row identifiers (Ref #, Job #) and Outlook calendar event IDs. This mapping is essential to prevent duplicate events and enable updates to existing calendar events when schedule data changes.

4. **User Information**: Cached Microsoft account information (display name, email) to show authentication status in the extension UI.

All data is stored locally using `chrome.storage.local` (which is encrypted by the browser) and never transmitted to third-party servers except Microsoft's Graph API for calendar operations.

**Code References:**
- `background.js`: Lines 16, 111-114, 125, 132, 200, 280, 293, 386, 950, 958, 984
- `options.js`: Lines 39, 73, 97, 108
- `popup.js`: Lines 46, 57, 68, 121, 150, 449

---

## ActiveTab Permission

**Justification:**
The extension requires the `activeTab` permission to interact with the schedule page (`prestigeav.ielightning.net`) when the user activates the extension. Specifically:

1. **Content Script Communication**: The extension needs to query the active tab to send messages to the content script that extracts schedule data from the page.

2. **Schedule Data Extraction**: When a user clicks "Get Selected Rows" in the extension popup, the extension must communicate with the content script running on the active schedule page to:
   - Read selected checkboxes on schedule rows
   - Extract schedule data (dates, job names, descriptions, etc.) from table cells
   - Return the extracted data to the popup for syncing to Outlook

The `activeTab` permission is scoped to only work when the user explicitly activates the extension (clicks the extension icon), ensuring the extension cannot access tabs without user interaction.

**Code References:**
- `popup.js`: Lines 189, 201 - Uses `chrome.tabs.query()` and `chrome.tabs.sendMessage()` to communicate with the active tab's content script

---

## Identity Permission

**Justification:**
The extension requires the `identity` permission to authenticate users with Microsoft Azure AD using OAuth 2.0 implicit flow. This is essential for the extension's core functionality:

1. **Microsoft Authentication**: The extension uses `chrome.identity.launchWebAuthFlow()` to initiate the Microsoft OAuth authentication flow, allowing users to sign in with their Microsoft 365 account.

2. **Access Token Retrieval**: After authentication, the extension receives an OAuth access token from Microsoft's identity platform, which is required to make API calls to Microsoft Graph API.

3. **Calendar Access**: The access token enables the extension to create and update calendar events in the user's Outlook calendar via Microsoft Graph API.

The `identity` permission is necessary because Chrome extensions cannot perform OAuth flows without this permission. The extension only requests access to the user's own calendar data (Calendars.ReadWrite) and basic profile information (User.Read), and users must explicitly consent to these permissions during the OAuth flow.

**Code References:**
- `background.js`: Lines 88, 103 - Uses `chrome.identity.getRedirectURL()` and `chrome.identity.launchWebAuthFlow()`
- `options.js`: Line 15 - Uses `chrome.identity.getRedirectURL()` to display redirect URI for Azure AD configuration

---

## Host Permissions

### `https://prestigeav.ielightning.net/*`

**Justification:**
This host permission is required for the content script to access and extract schedule data from the labor schedule pages. The extension:

1. **Content Script Injection**: Injects a content script (`content.js`) into schedule pages to read table data and selected checkboxes.

2. **Schedule Data Extraction**: Extracts schedule information including:
   - Job/Ref numbers
   - Dates and times
   - Job names and descriptions
   - Location information
   - Other schedule metadata

3. **User Interaction**: Allows users to select specific schedule rows via checkboxes, which the extension then syncs to their Outlook calendar.

The extension only accesses pages that match the content script patterns:
- `https://prestigeav.ielightning.net/reports/general/mySchedule*`
- `https://prestigeav.ielightning.net/laborSchedule*`

**Code References:**
- `manifest.json`: Lines 19-27 - Content script matches for schedule pages
- `content.js`: Entire file - Extracts data from schedule page DOM

### `https://graph.microsoft.com/*`

**Justification:**
This host permission is required to make API calls to Microsoft Graph API for calendar operations. The extension:

1. **Calendar Operations**: Creates and updates calendar events in the user's Outlook calendar.

2. **Calendar Listing**: Retrieves the list of user's calendars so they can select which calendar to sync to.

3. **Event Queries**: Searches for existing events to determine if an event should be created or updated.

All API calls are authenticated using the OAuth access token obtained through the `identity` permission, ensuring only the authenticated user's own calendar data is accessed.

**Code References:**
- `background.js`: Lines 139, 156, 303, 384, 411, 843 - All Microsoft Graph API calls

### `https://login.microsoftonline.com/*`

**Justification:**
This host permission is required for the OAuth 2.0 authentication flow with Microsoft Azure AD. The extension:

1. **OAuth Flow**: Redirects users to Microsoft's login page for authentication.

2. **Token Retrieval**: Receives OAuth access tokens after successful authentication.

This is a standard requirement for any extension that integrates with Microsoft services using OAuth 2.0. The extension uses the implicit flow via `chrome.identity.launchWebAuthFlow()`, which requires access to Microsoft's authentication endpoints.

**Code References:**
- `background.js`: Line 77 - OAuth authorization URL construction
- `background.js`: Line 103 - `chrome.identity.launchWebAuthFlow()` redirects to this domain

---

## Remote Code Usage

**Answer: NO, the extension does NOT use remote code.**

The extension does not download, fetch, or execute any JavaScript code from remote servers. All code is bundled within the extension package:

1. **All JavaScript files are local**:
   - `background.js` - Service worker (bundled)
   - `content.js` - Content script (bundled)
   - `popup.js` - Popup script (bundled)
   - `options.js` - Options page script (bundled)

2. **No dynamic code execution**:
   - No use of `eval()`
   - No use of `new Function()`
   - No dynamic script injection from remote sources
   - No fetching of `.js` files from external servers

3. **Only API calls for data**:
   - The extension makes HTTP requests to Microsoft Graph API (`https://graph.microsoft.com/*`) to:
     - Retrieve calendar lists
     - Create/update calendar events
     - Query existing events
   - These are REST API calls that return JSON data, not executable code.

4. **DOM manipulation only**:
   - Uses of `innerHTML` in the code (e.g., `popup.js`) are for DOM manipulation only, inserting static HTML strings, not executing remote code.

**Verification:**
- All code files are included in the extension package
- No `fetch()` calls for JavaScript files
- No `XMLHttpRequest` calls for JavaScript files
- No dynamic script tag creation with remote sources

---

## Summary

All permissions requested by this extension are necessary for its core functionality of syncing labor schedule data from ielightning.net to Outlook calendars. The extension:

- Stores data locally (Storage)
- Accesses the active schedule page (ActiveTab)
- Authenticates with Microsoft (Identity)
- Communicates with Microsoft Graph API and the schedule website (Host Permissions)

The extension does not use remote code and all functionality is provided by code bundled within the extension package.

