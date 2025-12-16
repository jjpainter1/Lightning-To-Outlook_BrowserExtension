# Lightning Schedule to Outlook Extension for Chrome

A small helper extension that takes your labor schedule from ielightning.net and turns it into Outlook calendar events using the Microsoft Graph API.

## Features

- Select specific schedule rows to sync
- Create new Outlook calendar events
- Update existing events (never deletes)
- Choose which calendar to sync to
- One-time sync operation
- Automatic event detection and matching

## Prerequisites

1. **Azure AD application**
   - You’ll need an app registration in Azure AD.
   - From that app registration you’ll use the **Application (client) ID**.
   - Redirect URI and API permissions are covered in the Developer Setup section.

2. **Microsoft 365 account**
   - Access to at least one of these schedule pages:
     - `https://prestigeav.ielightning.net/reports/general/mySchedule`
     - `https://prestigeav.ielightning.net/laborSchedule?token=...`
   - Access to an Outlook calendar where events will be created/updated.

## User Setup Instructions

These steps assume someone has already created and configured the Azure AD app (see **Developer Setup Instructions** below). In a typical organization setup, users will only need to install the extension and, if asked, paste a Client ID that an admin provides.

### 1. Download the extension

1. Download the latest packaged build from the **Releases** section of the GitHub repository (or clone this repo and use it directly).
2. Unzip it to a local directory.

### 2. Install the extension

1. Open Chrome/Edge and go to `chrome://extensions/` (or `edge://extensions/`)
2. Enable "Developer mode" (toggle in top right)
3. Click "Load unpacked"
4. Select this project folder
5. The extension should now appear in your extensions list

### 3. Configure the extension

1. Click the extension icon in your browser toolbar.
2. Click **Settings** at the bottom.
3. If the Client ID field is disabled, it means it's hardcoded in the extension and you don't need to enter it.
4. Fill out the fields for Initials and Default Reminder values
5. Click **Save Configuration**.

## Usage

1. **Navigate to a schedule page**
   - Option A (original schedule): `https://prestigeav.ielightning.net/reports/general/mySchedule`
   - Option B (token schedule): `https://prestigeav.ielightning.net/laborSchedule?token=...`
   - Make sure you’re logged in (and, for the token URL, that the token itself is valid).

2. **Select rows to sync**
   - Check the checkboxes next to the schedule rows you want to sync.
   - If you don’t check anything, the extension can fall back to “all visible rows” (useful for testing).

3. **Open the extension popup**
   - Click the extension icon in your browser toolbar.

4. **Sign in**
   - Click **Sign in with Microsoft**.
   - Complete the authentication flow and grant the requested permissions.

5. **Pick a calendar**
   - Choose which Outlook calendar you want to sync into.
   - Your default calendar will be pre-selected if available.

6. **Sync**
   - Click **Get Selected Rows** to preview what will be synced and see which rows will create/update vs. skip.
   - Click **Sync to Outlook** to actually create/update the events.

## How It Works

- **Event matching**
  - The extension stores a mapping between a **logical row identifier** and Outlook event IDs.
  - On the original schedule page this is usually **Ref #**.
  - On the token labor schedule page this is **Ref #**, or **Job #**, or **Job Name** (whichever is available, in that order).
  - When syncing, if an event already exists with that ID on the same day, it’s updated instead of duplicated.

- **Never deletes**
  - The extension only creates or updates events; it will never delete anything from your calendar.

- **Data mapping (original `mySchedule` page)**
  - Event subject: `Initials - Name - Description`
  - Event body: `Ref #`, `Project #`, `Type`, `Description`, `Office`
  - Start/End: schedule Start Date / End Date
  - Location: `Office`

- **Data mapping (token `laborSchedule` page)**
  - Event subject: `Initials - Job Name`
  - Event location: `Address` (or `Venue Name`, or `Office` as fallback)
  - Event body (details): `Ref #` (if present), `Job #`, `Job Name`, `Project #`, `Type`, `Description`, `Talent`, `Task`, `Client`, `Venue Name`, `Venue Room`, `Address`, `Salesperson`, `Order Status`, `Status`, `Labor Custom`, `Office`

## Developer Setup Instructions

If you’re the one wiring this up for yourself or a team, these are the steps you care about.

### 1. Register Azure AD application

1. Go to [Azure Portal – App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade).
2. Click **New registration**.
3. Fill in:
   - **Name**: Lightning to Outlook Extension (or any name you like).
   - **Supported account types**: Accounts in this organizational directory only (or as needed).
4. Click **Register**.
5. Note your **Application (client) ID** – you’ll paste this into the extension later.

### 2. Configure API permissions

1. In your app registration, go to **API permissions**.
2. Click **Add a permission**.
3. Select **Microsoft Graph** → **Delegated permissions**.
4. Add:
   - `Calendars.ReadWrite`
   - `User.Read`
5. Click **Add permissions**.
6. **Important**: Click **Grant admin consent** (or ask your admin to) so users don’t see consent prompts every time.

### 3. Configure redirect URI as a Single-page application

The extension uses the Microsoft identity platform v2.0 and the **SPA/implicit flow** via `chrome.identity.launchWebAuthFlow`. You must configure it as a **Single-page application** and allow access tokens (and optionally ID tokens).

**⚠️ Important: Redirect URI Distribution Issue**

When distributing an unpacked extension, each installation gets a unique extension ID, which means each user would have a different redirect URI (e.g., `https://[unique-id].chromiumapp.org/`). This makes it impractical to add each user's redirect URI to Azure AD.

**Solutions:**

#### Option A: Publish to Chrome Web Store (Recommended)

1. **Publish your extension to the Chrome Web Store**
   - This gives you a **fixed extension ID** that never changes
   - You only need to add **one redirect URI** to Azure AD
   - Users can install from the store without developer mode

2. **Get your fixed extension ID**
   - After publishing, your extension will have a permanent ID
   - The redirect URI will be: `https://[fixed-extension-id].chromiumapp.org/`

3. **Add the redirect URI to Azure AD**
   - Go to **Authentication** in your app registration
   - Under **Platform configurations**, click **Add a platform**
   - Select **Single-page application**
   - Add: `https://[fixed-extension-id].chromiumapp.org/`
   - Under **Implicit grant and hybrid flows**:
     - Enable **Access tokens**
     - Optionally enable **ID tokens**
   - Click **Save**

#### Option B: Use a Custom Redirect Page (Advanced)

If you can't publish to Chrome Web Store, you can set up a redirect page on your own domain:

1. **Create a redirect page** on your domain (e.g., `https://yourdomain.com/oauth-redirect.html`)
2. **Modify the extension** to use your custom redirect URI instead of `chrome.identity.getRedirectURL()`
3. **The redirect page** should extract the token from the URL and communicate it back to the extension using `postMessage` or similar
4. **Add your custom redirect URI** to Azure AD as a Single-page application

**For Development/Testing:**

If you're only testing locally or with a small team:

1. Go to **Authentication** in your app registration.
2. Under **Platform configurations**, click **Add a platform**.
3. Select **Single-page application**.
4. In **Redirect URIs**, add the redirect URI shown in the extension's **Settings** page:
   - It will look like: `https://[extension-id].chromiumapp.org/`.
   - You can copy it directly from the Options page after loading the extension.
5. Under **Implicit grant and hybrid flows** (inside the SPA configuration):
   - Enable **Access tokens**.
   - Optionally enable **ID tokens**.
6. Click **Save**.


## File Structure

```
├── manifest.json          # Extension manifest
├── background.js          # Service worker (auth & API calls)
├── content.js            # Content script (extracts schedule data)
├── content.css           # Styles for content script
├── popup.html            # Extension popup UI
├── popup.js              # Popup logic
├── popup.css             # Popup styles
├── options.html          # Settings page
├── options.js            # Settings logic
├── options.css           # Settings styles
└── README.md             # This file
```

## Troubleshooting

### Authentication Issues

- **"Client ID not configured"**: Make sure you've entered the Client ID in the options page
- **"Redirect URI mismatch"**: Ensure the redirect URI in Azure AD matches exactly what's shown in the options page
- **"Insufficient permissions"**: Make sure admin consent has been granted for the API permissions

### Sync Issues

- **"No rows selected"**: Make sure you've checked some rows on the schedule page, or click "Get Selected Rows" to sync all visible rows
- **"Failed to create event"**: Check that you have write permissions to the selected calendar
- **Events not updating**: Clear event mappings in settings and re-sync (this will create new events)

### Data Extraction Issues

- If the extension can't read schedule data, the table structure might have changed
- Check the browser console for errors
- The content script looks for table rows with checkboxes - adjust selectors in `content.js` if needed

## Development Notes

- Built as a Manifest V3 extension.
- Uses OAuth 2.0 with the Microsoft identity platform (via `chrome.identity.launchWebAuthFlow`).
- Event matching is based on the logical identifier stored in a single-value extended property on each event.

## Future Enhancements

- Better error handling and retry logic
- Sync history/logs

## Security Considerations

- Access tokens are stored in `chrome.storage.local` (encrypted by browser)
- Client ID is stored locally (not sensitive)
- All API calls use HTTPS
- Tokens are refreshed automatically when expired

## License

This project is licensed under the **MIT License**. See the `LICENSE` file for details.

