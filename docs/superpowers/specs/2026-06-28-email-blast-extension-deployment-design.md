# Email Blast — Multi-User Deployment via Browser Extension

**Date:** 2026-06-28
**Status:** Draft

## Overview

Refactors the current local-only tool into a deployed web application usable by 2-10 colleagues in the same organization. The core problem: the current Playwright-based approach launches a Chromium browser on the **server machine**, which does not work when the server is remote from users. The solution: replace Playwright entirely with a **Chrome/Edge browser extension** that runs on each user's machine, handling all Outlook Web interaction locally inside the user's own browser profile.

## Constraints

- No Python installation on user machines (company laptops, locked down)
- No Azure AD app registration (bypass IT entirely)
- No shared VM or Remote Desktop infrastructure
- No database — data flows through memory
- Users are on Windows, using Chrome or Edge
- Extension can be sideloaded if Web Store unavailable

## Architecture

### Three-Part Split

```
┌─────────────────────────────────────────────────────────┐
│  Deployed Server (e.g., company internal host)          │
│  ┌───────────┐  ┌──────────────┐                       │
│  │  Frontend │  │  Template    │  Serves UI             │
│  │  (static) │  │  Renderer    │  Renders MJML → HTML   │
│  └───────────┘  └──────┬───────┘  No Playwright         │
│                        │                                │
│  Routes:               │                                │
│  GET  /               → index.html                      │
│  GET  /api/fields     → field definitions               │
│  POST /api/preview    → render template → HTML           │
│  POST /api/render     → render final HTML per recipient  │
└────────────────────────┼────────────────────────────────┘
                         │
                         │ chrome.runtime.sendMessage
                         ▼
┌─────────────────────────────────────────────────────────┐
│  Chrome Extension (one per user, on their PC)           │
│  ┌────────────────┐  ┌──────────────────┐               │
│  │  Background    │  │  Content Script  │               │
│  │  Worker        │──│  (Outlook DOM    │               │
│  │  (orchestrates)│  │   automation)    │               │
│  └───────┬────────┘  └──────────────────┘               │
│          │                                               │
│  ┌───────┴────────┐                                     │
│  │  Content Script│  Bridge: relays messages between    │
│  │  (on web app)  │  web page ↔ background worker       │
│  └────────────────┘                                     │
└─────────────────────────────────────────────────────────┘
```

| Piece | Where it runs | What it does |
|---|---|---|
| **Web App** (FE+BE) | Deployed server | UI, template rendering, image embedding |
| **Chrome Extension** | Each user's browser | Outlook auth, compose, send, status |
| **Content Script Bridge** | Injected into web app pages | Relays messages between page ↔ extension |

## Communication Protocol

Uses a content-script bridge pattern (no hardcoded extension ID, no Chrome-specific code in the web app):

```
Web Page JS              Content Script            Background Worker
     │  window.postMessage()  │                         │
     │ ──────────────────────▶│                         │
     │                        │ chrome.runtime.sendMsg()│
     │                        │ ───────────────────────▶│
     │                        │                         │
     │                        │ chrome.runtime.onMsg()  │
     │                        │◀────────────────────────│
     │ window.postMessage()   │                         │
     │◀───────────────────────│                         │
```

**Message types:**

| Direction | Action | Payload |
|---|---|---|
| Page → Extension | `connect` | `{}` |
| Page → Extension | `send` | `{to, subject, html}` |
| Page → Extension | `disconnect` | `{}` |
| Page → Extension | `status` | `{}` (request current state) |
| Extension → Page | `status` | `{ready, connecting, user, error?}` |
| Extension → Page | `sendResult` | `{success, error?}` |
| Extension → Page | `disconnected` | `{}` |

## Extension Components

### 1. Background Worker (Service Worker, MV3)

Orchestrates all Outlook interaction. Manages state in `chrome.storage.session`:

```js
{
  status: "disconnected" | "connecting" | "connected",
  user: "user@tm.com.my" | null,
  outlookTabId: number | null,
  pendingSend: { to, subject, html } | null
}
```

Key functions:

- `connect()` — opens Outlook tab (`outlook.office.com/mail/`), injects login-detection content script, polls until logged in
- `send(to, subject, html)` — navigates to compose deep link, injects compose content script, fills and clicks Send
- `disconnect()` — closes Outlook tab, clears state

### 2. Outlook Content Script

Injected into `outlook.office.com/mail/*` pages. Mimics the current `sender.py` Playwright logic in vanilla JS running directly in the Outlook DOM:

**Login detection** (same checks as `sender.py:_is_logged_in_dom`):
- Poll for inbox indicators: `[data-test-id="user-display-name"]`, `[aria-label="New mail"]`, etc.
- Report back to background worker

**Compose automation** (same logic as `sender.py:_do_send`):
- Navigate to `outlook.office.com/mail/deeplink/compose`
- Fill `To:` input field
- Fill `Subject:` input field
- Inject HTML body into the compose iframe (Outlook Web uses a sandboxed iframe for the editor)
- Click Send button
- Report success/error back to background worker

**Email extraction:**
- Read user email from DOM after login for display in web UI

### 3. Content Script Bridge

Injected into the deployed web app's pages. Manifest declares the web app's origin in `content_scripts.matches`. Sole purpose: relay `window.postMessage` ↔ `chrome.runtime.sendMessage`.

```js
// Bridge content script
window.addEventListener("message", (e) => {
  if (e.source !== window) return;
  chrome.runtime.sendMessage(e.data, (response) => {
    window.postMessage({ ...response, _msgId: e.data._msgId }, "*");
  });
});
chrome.runtime.onMessage.addListener((msg) => {
  window.postMessage(msg, "*");
});
```

### Extension Manifest (manifest.json, MV3)

```json
{
  "manifest_version": 3,
  "name": "Email Blast — Outlook Connector",
  "version": "1.0",
  "permissions": ["tabs", "scripting", "storage"],
  "host_permissions": [
    "https://outlook.office.com/*",
    "https://outlook.live.com/*",
    "https://login.microsoftonline.com/*"
  ],
  "background": {
    "service_worker": "background.js"
  },
  "content_scripts": [
    {
      "matches": ["<DEPLOYED_WEB_APP_URL>"],
      "js": ["bridge.js"],
      "run_at": "document_idle"
    }
  ]
}
```

## Server Changes

### Routes Removed
- `POST /api/session/connect` — no longer opens browser on server
- `POST /api/session/logout` — no browser to close on server
- `GET /api/session/status` — state lives in extension, not server

### Routes Kept (unchanged)
- `GET /` — serve SPA
- `GET /api/fields` — field definitions
- `POST /api/preview` — render template with test values
- `POST /api/render` — render final HTML per recipient (renamed from `/api/send`; no longer sends, just renders)

### Dependencies Removed
- `playwright` — no longer needed on server
- `pywin32` — no longer needed on server

### New: `/api/render` (replaces `/api/send`)
```python
@app.route("/api/render", methods=["POST"])
def render():
    data = request.get_json()
    rendered = renderer.render(data["fields"])
    return {"html": rendered}
```

## Frontend Changes

| Change | Detail |
|---|---|
| Status bar | Polls extension via bridge instead of `/api/session/status` |
| "Connect" button | Sends `connect` message to extension instead of calling `/api/session/connect` |
| "Disconnect" button | Sends `disconnect` message instead of `/api/session/logout` |
| "Send" flow | `POST /api/render` → gets HTML → sends `send` message to extension with `{to, subject, html}` |
| Missing extension warning | Shows banner if extension doesn't respond within 5s of page load |
| Status messages | Updated for extension context: "Connect Outlook" → "Extension not found" → "Sign in to Outlook" → "Connected as X" |

## Session & State Management

### Why Extension State is Better

| Current (server-side) | Extension (browser-side) |
|---|---|
| State lost on server restart | Persists in browser profile |
| `--guest` mode → login every launch | User's normal browser → stays logged in across sessions |
| One shared session per server | Each user has independent session |
| Polling hits the server | Polling is local (instant) |

The user's Outlook session persists because the extension opens Outlook in the user's **normal browser profile** — not a guest Chromium instance. Microsoft session cookies are retained, so re-authentication is rare.

## Error Handling

| Scenario | Detection | Recovery |
|---|---|---|
| User closes Outlook tab manually | `chrome.tabs.onRemoved` listener | Update status to "disconnected", prompt user to reconnect |
| Session expires / logged out | Content script polls DOM login indicators every 30s | Notify web page → user clicks "Connect" again |
| Send fails (compose iframe not ready) | Content script timeout after 10s | Retry 2x with 2s delay, then report failure to web page |
| Send fails (button not found, selector broke) | Content script error | Fallback selectors, then alert user to update extension |
| Extension not installed/disabled | Web page heartbeat check — no response within 5s | Show "Extension not found" banner with install instructions |
| Outlook DOM changes (Microsoft redesign) | Send returns unexpected DOM error | Log error, prompt user to check for extension updates |
| Network failure between web page ↔ extension | Bridge script errors | web page retries message 3x before showing error |

## Extension Distribution

### Phase 1: Sideload (Immediate)
- Users enable Developer Mode in `chrome://extensions`
- Load unpacked extension folder
- Chrome shows "Disable developer mode extensions" warning on launch (annoying but functional)

### Phase 2: Chrome Web Store — Unlisted (Stable)
- One-time $5 Chrome Web Store developer fee
- Publish as "unlisted" (not searchable, accessible via direct link only)
- Auto-updates work, no launch warnings
- ~2 days for initial review

### Phase 3: Enterprise Policy Push (If IT cooperates)
- IT pushes extension via Group Policy / registry (`ExtensionInstallForcelist`)
- Users see nothing — extension appears automatically
- Requires IT to package and deploy

## Project Structure After Refactor

```
email_tm_blast/
├── server.py              # Flask app (no Playwright/no pywin32)
├── renderer.py            # MJML → HTML (unchanged)
├── requirements.txt       # flask only
├── static/
│   └── style.css
├── templates/
│   ├── index.html         # Updated UI (extension bridge, no session API calls)
│   └── fields.json
├── extension/             # NEW — Chrome extension
│   ├── manifest.json
│   ├── background.js      # Service worker — orchestrator
│   ├── bridge.js          # Content script — page ↔ extension relay
│   ├── outlook.js         # Content script — Outlook DOM automation
│   └── icons/
│       ├── icon16.png
│       ├── icon48.png
│       └── icon128.png
└── docs/
    └── superpowers/
        └── specs/
            └── 2026-06-28-email-blast-extension-deployment-design.md
```

## Non-Goals

- No token extraction or Graph API integration
- No multi-tenant support (users are in the same org, same Outlook domain)
- No email tracking or analytics
- No template editor — templates remain file-based
- No offline support — extension requires web app to be accessible
- No Firefox/Safari extension (Chrome/Edge only, matching user base)
