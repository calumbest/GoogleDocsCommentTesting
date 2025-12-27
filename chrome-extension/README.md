# Google Docs Comment Injector - Chrome Extension

**Test 7: Chrome Extension for Internal API Access**

This extension attempts to inject anchored comments into Google Docs by calling the internal API from an authenticated browser context.

## Installation

1. Open Chrome and go to `chrome://extensions/`
2. Enable **Developer mode** (toggle in top right)
3. Click **Load unpacked**
4. Select this `chrome-extension` folder

## Usage

1. Open a Google Doc in Chrome
2. Click the extension icon in the toolbar
3. Enter the **target text** you want to highlight
4. Enter your **comment text**
5. Click **Add Comment**

## How It Works

1. **Content script** runs in the Google Docs page context with authenticated session
2. Extracts document text to find target position (character indices)
3. Calls `/save` endpoint to create anchor at those positions
4. Calls `/docos/p/sync` endpoint to create comment linked to anchor

## Known Challenges

- **Session extraction**: Need to find where Google stores session tokens in the page
- **Revision tracking**: May need to get current document revision number
- **Text extraction**: Google Docs uses canvas rendering; text position mapping may be complex

## Debug

Open Chrome DevTools (F12) on the Google Doc page and check the Console tab for `[Comment Injector]` logs.

## Files

- `manifest.json` - Extension configuration
- `content.js` - Main logic running in Google Docs context
- `popup.html` - UI
- `popup.js` - Popup logic
