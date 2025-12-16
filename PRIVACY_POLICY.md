# Privacy Policy for Lightning Schedule to Outlook Extension

**Last Updated:** January 2025

This privacy policy describes how the Lightning Schedule to Outlook browser extension ("the Extension") collects, uses, and protects your information.

## Overview

The Lightning Schedule to Outlook Extension is designed to help you sync labor schedule data from ielightning.net to your Microsoft Outlook calendar. We are committed to protecting your privacy and being transparent about our data practices.

## Data Collection

The Extension collects the following types of data:

### 1. Authentication Information
- **OAuth Access Tokens**: Microsoft OAuth access tokens required to authenticate with Microsoft Graph API
- **Microsoft Account Information**: Display name and email address obtained from Microsoft Graph API (`/me` endpoint)
- **Purpose**: To authenticate with Microsoft services and access your Outlook calendar

### 2. Personal Communication
- **Calendar Events**: Calendar events created or updated in your Outlook calendar as part of the sync functionality
- **Purpose**: To sync schedule items from ielightning.net to your Outlook calendar

### 3. User Activity Data
- **Event Mappings**: A mapping between schedule row identifiers (Ref #, Job #) and Outlook calendar event IDs
- **User Preferences**: 
  - Selected calendar preference
  - User initials (optional, user-configured)
  - Default reminder minutes (optional, user-configured)
  - Dark mode preference (optional, user-configured)
- **Purpose**: To prevent duplicate calendar events, maintain sync state, and remember your preferences

### 4. Website Context Data
- **Schedule Data**: Information extracted from schedule pages on `prestigeav.ielightning.net`, including:
  - Job/Ref numbers
  - Job names and descriptions
  - Dates and times
  - Location information (addresses, venue names, offices)
  - Other schedule metadata (talent, tasks, clients, etc.)
- **Purpose**: To extract schedule information that you select to sync to your Outlook calendar
- **When Collected**: Only when you explicitly activate the extension and click "Get Selected Rows" or "Sync to Outlook"

## How Data is Stored

All data collected by the Extension is stored **locally on your device** using Chrome's `chrome.storage.local` API, which is encrypted by the browser. No data is transmitted to or stored on any third-party servers except:

- **Microsoft Graph API**: Calendar events are created/updated in your Microsoft Outlook calendar, which is stored according to Microsoft's privacy policy
- **Microsoft Authentication**: Authentication tokens are obtained through Microsoft's OAuth 2.0 service

## Data Usage

The Extension uses collected data **solely** for the following purposes:

1. **Calendar Synchronization**: To create and update calendar events in your Outlook calendar based on schedule data you select
2. **Authentication**: To maintain your Microsoft authentication session and access your calendar
3. **Preventing Duplicates**: To identify existing calendar events and update them instead of creating duplicates
4. **User Preferences**: To remember your settings and preferences for a better user experience

## Data Sharing

**The Extension does NOT share your data with any third parties** except:

- **Microsoft Corporation**: When you use the Extension, you authenticate with Microsoft and grant the Extension permission to access your Outlook calendar. Calendar events are stored in your Microsoft Outlook account, which is subject to Microsoft's privacy policy. The Extension only accesses your calendar data that you explicitly sync.

The Extension does NOT:
- Sell your data
- Share your data with advertisers
- Use your data for analytics or tracking
- Transmit your data to any servers other than Microsoft's services

## Third-Party Services

The Extension integrates with the following third-party services:

### Microsoft Graph API
- **Purpose**: To create, read, and update calendar events in your Outlook calendar
- **Data Shared**: Calendar event data (subject, dates, locations, descriptions) that you choose to sync
- **Privacy Policy**: [Microsoft Privacy Statement](https://privacy.microsoft.com/en-us/privacystatement)

### Microsoft Azure AD
- **Purpose**: To authenticate you with your Microsoft account
- **Data Shared**: OAuth authentication tokens and basic account information (display name, email)
- **Privacy Policy**: [Microsoft Privacy Statement](https://privacy.microsoft.com/en-us/privacystatement)

## Data Retention

- **Local Storage**: Data stored locally on your device remains until you:
  - Clear event mappings via the Extension's settings
  - Uninstall the Extension
  - Clear browser data for the Extension
- **Microsoft Calendar**: Calendar events created by the Extension are stored in your Microsoft Outlook account according to Microsoft's data retention policies
- **OAuth Tokens**: Access tokens expire according to Microsoft's token expiration policies (typically 1 hour for implicit flow tokens)

## Your Rights and Choices

You have the following rights regarding your data:

1. **Access**: You can view your stored preferences and event mappings through the Extension's settings page
2. **Deletion**: You can clear all event mappings at any time via the Extension's settings ("Clear Event Mappings" button)
3. **Uninstall**: Uninstalling the Extension will remove all locally stored data (you may need to manually clear browser data)
4. **Control**: You choose which schedule rows to sync and which calendar to sync to
5. **Authentication**: You can sign out at any time, which will remove stored authentication tokens

## Security

The Extension implements the following security measures:

- All data is stored locally using Chrome's encrypted storage API
- OAuth 2.0 authentication is used for secure Microsoft account access
- All API communications use HTTPS encryption
- No sensitive data is transmitted to unauthorized servers

## Children's Privacy

The Extension is not intended for use by children under the age of 13. We do not knowingly collect personal information from children under 13.

## Changes to This Privacy Policy

We may update this privacy policy from time to time. We will notify you of any changes by:
- Updating the "Last Updated" date at the top of this policy
- Posting the updated policy on the Extension's GitHub repository
- For significant changes, we may provide additional notification through the Extension or GitHub releases

## Contact Information

If you have questions or concerns about this privacy policy or the Extension's data practices, please contact us through:

- **GitHub Repository**: [https://github.com/jjpainter1/Lightning-To-Outlook_BrowserExtension](https://github.com/jjpainter1/Lightning-To-Outlook_BrowserExtension)
- **Issues**: Open an issue on the GitHub repository

## Open Source

This Extension is open source. You can review the source code on GitHub to verify our data practices. The Extension's code does not contain any hidden data collection or transmission mechanisms.

## Consent

By using the Lightning Schedule to Outlook Extension, you consent to this privacy policy. If you do not agree with this policy, please do not use the Extension.

---

**Note**: This Extension is provided "as is" and is not affiliated with, endorsed by, or sponsored by Microsoft Corporation, ielightning.net, or PrestigeAV. The Extension is an independent tool created to help users sync their schedule data.

