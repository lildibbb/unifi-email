# Email Blast — Browser Extension Refactor Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace server-side Playwright-based Outlook automation with a Chrome extension that handles Outlook Web interaction on each user's machine, enabling multi-user deployment.

**Architecture:** Three-part split — deployed server handles UI + template rendering; Chrome extension on each user's PC handles Outlook auth, compose, and send; a content script bridge relays messages between the web page and the extension via `window.postMessage` ↔ `chrome.runtime.sendMessage`.

**Tech Stack:** Flask (Python), vanilla JS (frontend), Chrome Extension MV3 (background service worker + content scripts), no new Python dependencies.

## Global Constraints

- No Python installation on user machines
- No Azure AD app registration
- No database — data flows through memory
- Flask>=3.0 only server-side dependency (playwright and pywin32 removed)
- Chrome/Edge MV3 extension only
- Extension sideloaded initially, Chrome Web Store unlisted listing later

---

## File Structure After Refactor

```
email_tm_blast/
├── server.py              # Flask app — render only, no sender/session logic
├── renderer.py            # MJML → HTML (unchanged)
├── requirements.txt       # flask>=3.0 only
├── sender.py              # KEPT but NOT imported (local-mode reference only)
├── templates/
│   ├── index.html         # Updated — extension bridge, no server polling
│   └── fields.json
├── extension/             # NEW
│   ├── manifest.json
│   ├── background.js      # Service worker — orchestrator
│   ├── bridge.js          # Content script — page ↔ extension relay
│   ├── outlook.js         # Content script — Outlook DOM automation
│   └── icons/
│       ├── icon16.png
│       ├── icon48.png
│       └── icon128.png
└── static/
    └── style.css
```

---

### Task 1: Clean Server — Remove Sender/Session Logic

**Files:**
- Modify: `server.py`
- Modify: `requirements.txt`

**Interfaces:**
- Produces: Flask app with only render routes (`/`, `/api/fields`, `/api/preview`, `/api/render`)
- Removes: `SenderDispatcher` import, session routes, Playwright initialization

**Description:** Strip `server.py` down to a pure rendering server. Remove all Playwright/session/sender code. Rename `/api/send` to `/api/render` (returns HTML only, does not send). Remove `playwright` and `pywin32` from requirements.

- [ ] **Step 1: Rewrite `server.py`**

Replace all content of `server.py` with the render-only version:

```python
from flask import Flask, request, jsonify, render_template_string
import os
from pathlib import Path

from renderer import get_fields, render

app = Flask(__name__, static_folder="static", static_url_path="/static")

_index_html = None


def _get_index_html():
    global _index_html
    if _index_html is None:
        path = Path(__file__).parent / "templates" / "index.html"
        _index_html = path.read_text(encoding="utf-8")
    return _index_html


@app.route("/")
def index():
    return render_template_string(_get_index_html())


@app.route("/api/fields")
def api_fields():
    return jsonify(get_fields())


@app.route("/api/preview", methods=["POST"])
def api_preview():
    data = request.get_json()
    if not data or "fields" not in data:
        return jsonify({"error": "fields required"}), 400

    lang = data.get("lang", "ENG").upper()
    if lang not in ("ENG", "BM"):
        return jsonify({"error": "lang must be ENG or BM"}), 400

    try:
        html = render(data["fields"], lang=lang)
        return jsonify({"html": html})
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/render", methods=["POST"])
def api_render():
    data = request.get_json()
    if not data:
        return jsonify({"error": "request body required"}), 400

    required = ["fields"]
    for key in required:
        if key not in data:
            return jsonify({"error": f"'{key}' is required"}), 400

    lang = data.get("lang", "ENG").upper()
    if lang not in ("ENG", "BM"):
        return jsonify({"error": "lang must be ENG or BM"}), 400

    try:
        html = render(data["fields"], lang=lang)
        return jsonify({"html": html})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        print(f"\n  NextGen VIP Email Blast — Render Server")
        print(f"  Open http://localhost:5000 in your browser\n")
    app.run(debug=False, port=5000, use_reloader=False)
```

- [ ] **Step 2: Update `requirements.txt`**

Replace content:
```
flask>=3.0
```

- [ ] **Step 3: Verify server starts cleanly**

Run: `python server.py`
Expected: No import errors, prints "NextGen VIP Email Blast — Render Server"

- [ ] **Step 4: Test `/api/render` endpoint**

With server running, use curl or fetch to test:
```
POST /api/render
Content-Type: application/json
{"fields": {"customer_name": "Test User", "crew_name": "Test Crew"}, "lang": "ENG"}
```
Expected: Returns `{"html": "<!DOCTYPE html>..."}` with substituted values.

---

### Task 2: Create Extension Manifest

**Files:**
- Create: `extension/manifest.json`

**Interfaces:**
- Produces: MV3 manifest declaring permissions, background worker, content scripts, host permissions

- [ ] **Step 1: Create `extension/` directory**

```powershell
New-Item -ItemType Directory -Path "extension" -Force
```

- [ ] **Step 2: Create `extension/manifest.json`**

```json
{
  "manifest_version": 3,
  "name": "Email Blast — Outlook Connector",
  "version": "1.0",
  "description": "Connects the Email Blast web app to your Outlook Web session for sending emails.",
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
      "matches": ["<all_urls>"],
      "js": ["bridge.js"],
      "run_at": "document_idle"
    }
  ],
  "icons": {
    "16": "icons/icon16.png",
    "48": "icons/icon48.png",
    "128": "icons/icon128.png"
  }
}
```

---

### Task 3: Create Background Worker

**Files:**
- Create: `extension/background.js`

**Interfaces:**
- Consumes: Messages from bridge content script via `chrome.runtime.onMessage`
- Produces: Status updates, send results, tab management
- State: `chrome.storage.session` keys: `status`, `user`, `outlookTabId`

**Description:** The service worker orchestrates all Outlook interaction. It manages an Outlook tab, injects the outlook.js content script, and processes connect/send/disconnect/status commands.

- [ ] **Step 1: Create `extension/background.js`**

```javascript
const OUTLOOK_URL = 'https://outlook.office.com/mail/';
const COMPOSE_URL = 'https://outlook.office.com/mail/deeplink/compose';
const LOGIN_TIMEOUT = 300000; // 5 min

function getState() {
  return chrome.storage.session.get({
    status: 'disconnected',
    user: null,
    outlookTabId: null
  });
}

function setState(update) {
  return chrome.storage.session.set(update);
}

function notifyPage(msg) {
  chrome.runtime.sendMessage(msg).catch(() => {});
}

// ── Bridge message handler ──────────────────────────────────────────
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || !msg.action) return;

  switch (msg.action) {
    case 'connect':
      handleConnect().then(() => sendResponse({}));
      return true; // async response

    case 'disconnect':
      handleDisconnect().then(() => sendResponse({}));
      return true;

    case 'send':
      handleSend(msg.to, msg.subject, msg.html)
        .then(result => sendResponse(result))
        .catch(err => sendResponse({ success: false, error: err.message }));
      return true;

    case 'status':
      getState().then(state => sendResponse(state));
      return true;
  }
});

// ── Connect flow ────────────────────────────────────────────────────
async function handleConnect() {
  await setState({ status: 'connecting', user: null });
  notifyPage({ action: 'status', status: 'connecting', connecting: true });

  try {
    const tabs = await chrome.tabs.query({ url: [
      'https://outlook.office.com/*',
      'https://outlook.live.com/*'
    ]});

    let tabId;
    if (tabs.length > 0) {
      tabId = tabs[0].id;
      await chrome.tabs.update(tabId, { active: true });
      await chrome.tabs.reload(tabId);
    } else {
      const tab = await chrome.tabs.create({ url: OUTLOOK_URL, active: true });
      tabId = tab.id;
    }

    await setState({ outlookTabId: tabId });

    // Poll for login
    const startTime = Date.now();
    while (Date.now() - startTime < LOGIN_TIMEOUT) {
      const state = await getState();
      if (state.status === 'disconnected') return; // cancelled

      try {
        const results = await chrome.scripting.executeScript({
          target: { tabId },
          func: checkLoggedIn
        });

        if (results[0]?.result) {
          await setState({
            status: 'connected',
            user: results[0].result
          });
          notifyPage({
            action: 'status',
            status: 'connected',
            ready: true,
            user: results[0].result
          });
          return;
        }
      } catch (e) {
        // Tab may have closed or navigated — retry
      }

      await sleep(2000);
    }

    await setState({ status: 'disconnected' });
    notifyPage({
      action: 'status',
      status: 'disconnected',
      error: 'Login timed out (5 min)'
    });

  } catch (err) {
    await setState({ status: 'disconnected' });
    notifyPage({
      action: 'status',
      status: 'disconnected',
      error: err.message
    });
  }
}

// ── Send flow ───────────────────────────────────────────────────────
async function handleSend(to, subject, html) {
  const state = await getState();

  if (state.status !== 'connected') {
    return { success: false, error: 'Not connected — click Connect Outlook first' };
  }

  let tabId = state.outlookTabId;
  if (!tabId) {
    return { success: false, error: 'No Outlook tab found — click Connect again' };
  }

  // Verify session still valid
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId },
      func: checkLoggedIn
    });
    if (!results[0]?.result) {
      await setState({ status: 'disconnected' });
      notifyPage({ action: 'status', status: 'disconnected' });
      return { success: false, error: 'Session expired — click Connect Outlook to sign in again' };
    }
  } catch (e) {
    await setState({ status: 'disconnected' });
    notifyPage({ action: 'status', status: 'disconnected' });
    return { success: false, error: 'Outlook tab lost — click Connect again' };
  }

  // Open compose
  try {
    await chrome.tabs.update(tabId, { url: COMPOSE_URL, active: true });
    await sleep(3000); // Wait for compose to load
  } catch (e) {
    return { success: false, error: 'Could not open compose window' };
  }

  // Run the compose automation
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId },
      func: doComposeAndSend,
      args: [{ to, subject, html }]
    });

    const result = results[0]?.result;
    if (result?.success) {
      await sleep(2000); // Wait for send to complete
    }
    return result || { success: false, error: 'Compose automation returned no result' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── Disconnect ──────────────────────────────────────────────────────
async function handleDisconnect() {
  const state = await getState();
  if (state.outlookTabId) {
    try {
      await chrome.tabs.remove(state.outlookTabId);
    } catch (e) {}
  }
  await setState({ status: 'disconnected', user: null, outlookTabId: null });
  notifyPage({ action: 'status', status: 'disconnected' });
}

// ── Tab close detection ─────────────────────────────────────────────
chrome.tabs.onRemoved.addListener(async (tabId) => {
  const state = await getState();
  if (state.outlookTabId === tabId && state.status !== 'disconnected') {
    await setState({ status: 'disconnected', user: null, outlookTabId: null });
    notifyPage({ action: 'status', status: 'disconnected' });
  }
});

// ── Utilities ───────────────────────────────────────────────────────
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ── DOM check functions (run in page context via executeScript) ─────

function checkLoggedIn() {
  if (window.location.href.includes('login.microsoftonline.com') ||
      window.location.href.includes('login.live.com')) {
    return null;
  }
  if (document.querySelector('[data-test-id="user-display-name"]')) {
    const el = document.querySelector('[data-test-id="user-display-name"]');
    return (el.getAttribute('title') || el.textContent || 'Connected').trim();
  }
  if (document.querySelector('[aria-label="New mail"]')) return 'Connected';
  if (document.querySelector('[data-app-section="MailModule"]')) return 'Connected';
  const inboxLabel = document.querySelector('[title="Inbox"], [aria-label="Inbox"]');
  if (inboxLabel) return 'Connected';

  return null; // still on login page or unknown state
}
```

- [ ] **Step 2 (continued): Append compose automation to `background.js`**

```javascript
function doComposeAndSend(params) {
  const { to, subject, html } = params;

  // Wait for compose form to appear
  const TO_SELECTORS = [
    'input[aria-label="To"]',
    'div[aria-label="To"] input',
    'div[aria-label="To"]',
    '[role="combobox"][aria-label*="To"]',
    'input[type="text"]:visible'
  ];

  let toEl = null;
  for (let i = 0; i < 30; i++) {
    for (const sel of TO_SELECTORS) {
      const el = document.querySelector(sel);
      if (el && el.offsetParent !== null) {
        toEl = el;
        break;
      }
    }
    if (toEl) break;
    // No sleep in page context — caller handles retries
    return { success: false, error: 'Compose window did not load — selector not found' };
  }

  if (!toEl) {
    return { success: false, error: 'Compose window did not load' };
  }

  // Fill To
  toEl.focus();
  toEl.click();
  toEl.value = to;
  toEl.dispatchEvent(new Event('input', { bubbles: true }));
  toEl.dispatchEvent(new Event('change', { bubbles: true }));

  // Simulate Enter to confirm recipient
  setTimeout(() => {
    toEl.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  }, 300);

  // Fill Subject
  const subjectEl = document.querySelector(
    'input[placeholder="Add a subject"], input[aria-label="Subject"], [aria-label="Add a subject"]'
  );
  if (subjectEl) {
    subjectEl.focus();
    subjectEl.value = subject;
    subjectEl.dispatchEvent(new Event('input', { bubbles: true }));
    subjectEl.dispatchEvent(new Event('change', { bubbles: true }));
  } else {
    return { success: false, error: 'Could not find Subject field' };
  }

  // Inject HTML body — try contenteditable editors
  const bodySelectors = [
    '[aria-label="Message body"]',
    '[aria-label*="body" i]',
    '[data-testid="compose-editor"]',
    '[contenteditable="true"][role="textbox"]',
    '#editorBody',
    '.ms-composer [contenteditable="true"]',
    '[contenteditable="true"]'
  ];

  let bodyEl = null;
  for (const sel of bodySelectors) {
    const editors = [...document.querySelectorAll(sel)];
    const el = editors[editors.length - 1]; // last match = body, not To/CC
    if (el && el.contentEditable === 'true') {
      bodyEl = el;
      break;
    }
  }

  if (!bodyEl) {
    return { success: false, error: 'Could not find message body editor' };
  }

  bodyEl.focus();
  document.execCommand('selectAll', false, null);
  const ok = document.execCommand('insertHTML', false, html);
  if (!ok || bodyEl.innerHTML.length < 10) {
    bodyEl.innerHTML = html;
    bodyEl.dispatchEvent(new Event('input', { bubbles: true }));
    bodyEl.dispatchEvent(new Event('change', { bubbles: true }));
  }

  // Click Send
  const sendBtn = document.querySelector(
    'button[aria-label="Send"], button[title="Send"]'
  );
  if (sendBtn) {
    sendBtn.click();
    return { success: true, error: null };
  }

  // Fallback: try finding any send-like button
  const buttons = [...document.querySelectorAll('button')];
  const send = buttons.find(b =>
    b.textContent.trim().toLowerCase() === 'send' ||
    (b.getAttribute('aria-label') || '').toLowerCase() === 'send'
  );
  if (send) {
    send.click();
    return { success: true, error: null };
  }

  return { success: false, error: 'Could not find Send button — Outlook UI may have changed' };
}
```

Note: The `doComposeAndSend` function runs in the page context via `chrome.scripting.executeScript`. It cannot use `await` or `setTimeout` for synchronous DOM operations — those are handled by the caller retrying.

---

### Task 4: Create Bridge Content Script

**Files:**
- Create: `extension/bridge.js`

**Interfaces:**
- Consumes: `window.postMessage` from web page
- Produces: `chrome.runtime.sendMessage` to background worker
- Relays responses back to web page via `window.postMessage`

- [ ] **Step 1: Create `extension/bridge.js`**

```javascript
(function () {
  let pendingCallbacks = {};
  let msgCounter = 0;

  window.addEventListener('message', function (event) {
    if (event.source !== window) return;
    if (!event.data || event.data.source !== 'email-blast') return;

    const msg = event.data;
    const msgId = msg._msgId;

    chrome.runtime.sendMessage(msg, function (response) {
      if (chrome.runtime.lastError) {
        window.postMessage({
          source: 'email-blast-extension',
          _msgId: msgId,
          error: chrome.runtime.lastError.message
        }, '*');
        return;
      }

      window.postMessage({
        source: 'email-blast-extension',
        _msgId: msgId,
        response: response
      }, '*');
    });
  });

  // Forward unsolicited messages from background to page
  chrome.runtime.onMessage.addListener(function (msg) {
    window.postMessage({
      source: 'email-blast-extension',
      unsolicited: true,
      data: msg
    }, '*');
  });
})();
```

---

### Task 5: Refactor Frontend — Replace Server Polling with Extension Bridge

**Files:**
- Modify: `templates/index.html` (lines 350-546, the `<script>` block)

**Interfaces:**
- Consumes: Extension bridge via `window.postMessage`
- Produces: Updated UI that communicates with extension instead of server for session/send

**Description:** Replace the entire JavaScript block. The new code:
- Sends messages to extension via `window.postMessage({ source: 'email-blast', action: '...' })`
- Listens for responses via `window.addEventListener('message', ...)`
- Routes `/api/render` for template rendering (server does this, not extension)
- Routes send commands to extension after getting rendered HTML from server
- Handles extension-not-found case with warning banner

- [ ] **Step 1: Replace the `<script>` block in `templates/index.html`**

Remove lines 350-546 and replace with the new script:

```javascript
<script>
// ── State ───────────────────────────────────────────────────────────────
let lang      = 'ENG';
let allFields = [];
let extReady  = false;
let extUser   = null;
let extStatus = 'disconnected';
let extError  = null;
let pendingIds = {};
let msgCounter = 0;

const GROUP_ORDER = {
  customer: ['customer_name','customer_phone','home_office'],
  service:  ['unifi_login','package','address','connection','router','internetusage'],
  crew:     ['crew_name','crew_phone'],
};

// ── Extension Communication ────────────────────────────────────────────
function sendToExtension(msg) {
  return new Promise((resolve, reject) => {
    const id = ++msgCounter;
    pendingIds[id] = { resolve, reject };
    window.postMessage({ source: 'email-blast', _msgId: id, ...msg }, '*');

    setTimeout(() => {
      if (pendingIds[id]) {
        delete pendingIds[id];
        reject(new Error('Extension response timed out'));
      }
    }, 30000);
  });
}

function sendToExtensionFire(msg) {
  // Fire-and-forget
  window.postMessage({ source: 'email-blast', _msgId: ++msgCounter, ...msg }, '*');
}

window.addEventListener('message', function (event) {
  if (event.source !== window) return;
  if (!event.data || event.data.source !== 'email-blast-extension') return;

  const data = event.data;

  // Unsolicited status updates from background worker
  if (data.unsolicited && data.data) {
    const update = data.data;
    if (update.action === 'status') {
      extStatus = update.status;
      extUser   = update.user || null;
      extError  = update.error || null;
      updateStatusUI();
    }
    return;
  }

  // Response to a pending request
  if (data._msgId && pendingIds[data._msgId]) {
    const { resolve, reject } = pendingIds[data._msgId];
    delete pendingIds[data._msgId];
    if (data.error) {
      reject(new Error(data.error));
    } else {
      resolve(data.response || {});
    }
  }
});

// ── Status Check (heartbeat) ──────────────────────────────────────────
async function checkExtension() {
  try {
    const state = await sendToExtension({ action: 'status' });
    extStatus = state.status || 'disconnected';
    extUser   = state.user || null;
    extError  = null;
  } catch (e) {
    extStatus = 'missing';
  }
  updateStatusUI();
}

// ── Update Status UI ──────────────────────────────────────────────────
function updateStatusUI() {
  const pill  = document.getElementById('status-pill');
  const label = document.getElementById('status-label');
  const btnC  = document.getElementById('btn-connect');
  const btnD  = document.getElementById('btn-disconnect');
  const btnS  = document.getElementById('btn-send');

  if (extStatus === 'connected') {
    pill.className    = 'status-pill ok';
    label.textContent = extUser || 'Connected';
    btnC.style.display = 'none';
    btnD.style.display = 'inline-flex';
    btnS.disabled      = false;
  } else if (extStatus === 'connecting') {
    pill.className    = 'status-pill busy';
    label.textContent = 'Signing in…';
    btnC.style.display = btnD.style.display = 'none';
    btnS.disabled      = true;
  } else if (extStatus === 'missing') {
    pill.className    = 'status-pill err';
    label.textContent = 'Extension not found';
    btnC.style.display = 'none';
    btnD.style.display = 'none';
    btnS.disabled      = true;
    notify('Browser extension not detected — install the Outlook Connector extension', 'error', 8000);
  } else {
    // disconnected or error
    pill.className    = 'status-pill';
    if (extError && extStatus === 'disconnected') {
      pill.className  = 'status-pill err';
      label.textContent = 'Auth Error';
      btnC.textContent  = 'Reconnect';
    } else {
      label.textContent = 'Disconnected';
      btnC.textContent  = 'Connect Outlook';
    }
    btnC.style.display = 'inline-flex';
    btnD.style.display = 'none';
    btnS.disabled      = true;
  }
}

// ── Notify Toast ──────────────────────────────────────────────────────
function notify(msg, type = 'info', ms = 4000) {
  const el   = document.getElementById('notify');
  const icon = document.getElementById('n-icon');
  const txt  = document.getElementById('n-msg');
  icon.textContent = type === 'success' ? '\u2713' : type === 'error' ? '\u2715' : '\u2139';
  txt.textContent  = msg;
  el.className     = 'show ' + type;
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.className = ''; }, ms);
}

// ── Language Toggle ───────────────────────────────────────────────────
function setLang(l) {
  lang = l;
  document.getElementById('btn-eng').classList.toggle('active', l === 'ENG');
  document.getElementById('btn-bm').classList.toggle('active',  l === 'BM');

  const subj = document.getElementById('f-subject');
  if (!subj.value.trim()) {
    subj.placeholder = l === 'ENG'
      ? 'Your Personalized Unifi Experience \u2014 NEXTGen Crew'
      : 'Alami Pengalaman Unifi Dengan NEXTGen Crew';
  }

  allFields.forEach(f => {
    const lbl = document.querySelector('[data-label-for="' + f.name + '"]');
    if (lbl) lbl.textContent = (l === 'BM' && f.label_bm) ? f.label_bm : f.label;
  });
}

// ── Form Building ─────────────────────────────────────────────────────
function buildField(f) {
  const labelText = (lang === 'BM' && f.label_bm) ? f.label_bm : f.label;
  const req       = f.required ? '<span class="req">*</span>' : '';
  let input;

  if (f.type === 'select') {
    const opts = f.options.map(o => '<option value="' + o.value + '">' + o.label + '</option>').join('');
    input = '<select id="f-' + f.name + '">' + opts + '</select>';
  } else {
    input = '<input type="text" id="f-' + f.name + '" placeholder="' + (f.placeholder || '') + '">';
  }

  return '' +
    '<div class="field">' +
      '<label for="f-' + f.name + '">' +
        '<span data-label-for="' + f.name + '">' + labelText + '</span>' + req +
      '</label>' +
      input +
    '</div>';
}

async function loadFields() {
  const res = await fetch('/api/fields');
  allFields  = await res.json();

  for (const [group, keys] of Object.entries(GROUP_ORDER)) {
    const el = document.getElementById('group-' + group);
    if (!el) continue;
    const ordered = keys
      .map(k => allFields.find(f => f.name === k))
      .filter(Boolean);
    el.innerHTML = ordered.map(buildField).join('');
  }
  setLang(lang);
}

function getValues() {
  const v = {};
  allFields.forEach(f => {
    const el = document.getElementById('f-' + f.name);
    v[f.name] = el ? el.value.trim() : '';
  });
  return v;
}

// ── Send ──────────────────────────────────────────────────────────────
async function doSend() {
  const to      = document.getElementById('f-to').value.trim();
  const subject = document.getElementById('f-subject').value.trim();

  if (!to || !subject) {
    notify('To address and Subject line are required', 'error');
    return;
  }

  const missing = allFields
    .filter(f => f.required)
    .filter(f => !document.getElementById('f-' + f.name)?.value.trim());
  if (missing.length) {
    notify('Required: ' + missing.map(f => f.label).join(', '), 'error');
    return;
  }

  if (!confirm('Send email to ' + to + '?')) return;

  const btn = document.getElementById('btn-send');
  btn.innerHTML = '<span class="spin"></span> Sending\u2026';
  btn.disabled  = true;

  try {
    // Step 1: Render template on server
    const renderResp = await fetch('/api/render', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields: getValues(), lang }),
    }).then(r => r.json());

    if (renderResp.error) {
      notify('Render failed: ' + renderResp.error, 'error');
      return;
    }

    // Step 2: Send via extension
    const sendResult = await sendToExtension({
      action: 'send',
      to: to,
      subject: subject,
      html: renderResp.html
    });

    if (sendResult.success) {
      notify('\u2713 Sent to ' + to, 'success', 6000);
    } else {
      notify('Failed: ' + (sendResult.error || 'Unknown error'), 'error', 8000);
    }
  } catch(e) {
    if (e.message && e.message.includes('Extension')) {
      notify('Extension communication error — is the Outlook Connector installed?', 'error', 8000);
    } else {
      notify('Send error: ' + e.message, 'error');
    }
  } finally {
    btn.textContent = 'Send Email';
    btn.disabled    = false;
    checkExtension();
  }
}

// ── Connect / Disconnect ──────────────────────────────────────────────
async function doConnect() {
  document.getElementById('btn-connect').disabled = true;
  notify('Opening Outlook — sign in on the browser window that appears.', 'info', 8000);

  try {
    sendToExtensionFire({ action: 'connect' });
    // Status updates will come via unsolicited messages
    extStatus = 'connecting';
    updateStatusUI();

    // Poll for completion
    let attempts = 0;
    while (attempts < 60) { // 2 min max
      await new Promise(r => setTimeout(r, 2000));
      attempts++;
      await checkExtension();
      if (extStatus === 'connected' || extStatus === 'disconnected') break;
    }
  } catch (e) {
    notify('Failed to connect: ' + e.message, 'error');
  }
  document.getElementById('btn-connect').disabled = false;
}

async function doDisconnect() {
  try {
    sendToExtensionFire({ action: 'disconnect' });
  } catch (e) {}
  extStatus = 'disconnected';
  extUser = null;
  updateStatusUI();
  notify('Outlook session closed.', 'info');
}

// ── Init ──────────────────────────────────────────────────────────────
loadFields();
setTimeout(checkExtension, 500);
// Heartbeat every 30s to detect extension install/removal
setInterval(checkExtension, 30000);
</script>
```

- [ ] **Step 2: Verify file is valid HTML**

Check that the closing `</script>` tag appears exactly once at the end and the file structure (head/body/scripts) is intact.

---

### Task 6: Create Placeholder Extension Icons

**Files:**
- Create: `extension/icons/icon16.png`
- Create: `extension/icons/icon48.png`
- Create: `extension/icons/icon128.png`

**Description:** Create minimal 1-pixel PNG files as placeholders. The extension will load but show blank icons until replaced with proper artwork.

- [ ] **Step 1: Create icons directory**

```powershell
New-Item -ItemType Directory -Path "extension\icons" -Force
```

- [ ] **Step 2: Generate placeholder PNGs**

Create minimal orange-colored 16x16, 48x48, and 128x128 PNG files. Use a simple script or pre-made minimal PNG bytes.

```python
# Run this once to generate placeholder icons
import struct, zlib

def create_png(width, height, r, g, b, filepath):
    def chunk(chunk_type, data):
        c = chunk_type + data
        return struct.pack('>I', len(data)) + c + struct.pack('>I', zlib.crc32(c) & 0xffffffff)

    header = b'\x89PNG\r\n\x1a\n'
    ihdr = chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))

    raw = b''
    for y in range(height):
        raw += b'\x00'  # filter byte
        for x in range(width):
            raw += struct.pack('BBB', r, g, b)

    idat = chunk(b'IDAT', zlib.compress(raw))
    iend = chunk(b'IEND', b'')

    with open(filepath, 'wb') as f:
        f.write(header + ihdr + idat + iend)

# Orange #FF5E00
for s, p in [(16,'icon16.png'),(48,'icon48.png'),(128,'icon128.png')]:
    create_png(s, s, 255, 94, 0, f'extension/icons/{p}')
```

---

### Task 7: Verify Everything Works End-to-End

**Description:** Load the extension, start the server, test the full flow.

- [ ] **Step 1: Load extension in Chrome**

1. Go to `chrome://extensions`
2. Enable "Developer mode"
3. Click "Load unpacked"
4. Select the `extension/` folder
5. Verify extension appears with no errors

- [ ] **Step 2: Start server**

Run: `python server.py`
Expected: "NextGen VIP Email Blast — Render Server"

- [ ] **Step 3: Open web app**

Navigate to `http://localhost:5000`
Expected: Page loads, status shows "Disconnected" with "Connect Outlook" button visible

- [ ] **Step 4: Test Connect flow**

Click "Connect Outlook"
Expected: A new Outlook tab opens, user signs in, status updates to green with email shown

- [ ] **Step 5: Test Send flow**

Fill form, click Send
Expected: Template renders on server, HTML sent to extension, extension automates Outlook compose+send

- [ ] **Step 6: Test Disconnect flow**

Click "Disconnect"
Expected: Outlook tab closes (if opened by extension), status returns to "Disconnected"

---

## Completion Checklist

- [ ] Extension loads without errors in chrome://extensions
- [ ] `python server.py` starts cleanly, no import errors
- [ ] `pip install -r requirements.txt` installs only flask
- [ ] Web page loads at localhost:5000
- [ ] "Connect Outlook" opens Outlook tab via extension
- [ ] Status updates correctly (connected/connecting/missing)
- [ ] Send flow: render → extension → Outlook compose → send
- [ ] Disconnect flow: closes Outlook tab
- [ ] Extension missing detection shows warning banner
