# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2025-01-XX

### Fixed
- **Client ID Configuration**: Fixed issue where options page required Client ID entry even when it was hardcoded in `background.js`. The Client ID field now automatically detects and displays hardcoded values as read-only, preventing user confusion and configuration errors.

### Changed
- **Options Page UI**: Client ID input field is now automatically disabled and marked as read-only when a hardcoded Client ID is detected, with an informational message explaining the state.
- **Options Page Validation**: Client ID validation is now skipped when a hardcoded Client ID is present, allowing users to save other preferences without errors.

### Added
- **Redirect URI Distribution Documentation**: Added comprehensive documentation in README.md explaining the redirect URI distribution challenge and providing two solutions:
  - **Option A (Recommended)**: Publishing to Chrome Web Store for a fixed extension ID
  - **Option B (Advanced)**: Using a custom redirect page on your own domain
- **Background Script API**: Added `getClientId` message handler to allow options page to query whether Client ID is hardcoded
- **Options Page Info Messages**: Added informational message display when Client ID is hardcoded

### Technical Details
- Modified `background.js` to expose hardcoded Client ID status via message handler
- Updated `options.js` to check for hardcoded Client ID on page load and adjust UI accordingly
- Enhanced `options.html` with informational elements for better user guidance
- Improved `options.css` styling for disabled input fields

## [0.1.0] - Initial Release

### Added
- Initial proof of concept release
- Schedule data extraction from ielightning.net
- Microsoft Graph API integration for Outlook calendar sync
- Event creation and update functionality
- User preferences (initials, default reminders)
- Event mapping system to prevent duplicates
- Support for both original schedule page and token-based labor schedule page

