# Email Blast Tool — Implementation Plan

> **For agentic workers:** Inline execution — build tasks sequentially.

**Goal:** A Flask web tool for sending MJML-templated emails via Outlook COM or Playwright browser automation.

**Architecture:** Flask serves a single-page UI. Sender dispatcher selects COM (Windows+Outlook desktop) or Playwright (Outlook Web). MJML templates rendered via npx mjml subprocess. No database.

**Tech Stack:** Python 3, Flask, Playwright, pywin32 (Windows), MJML (Node/npx)

## Global Constraints

- No database — data flows through memory, destroyed on shutdown
- No Azure AD app registration
- No SMTP — COM and Playwright only
- Must work on Windows and Mac
- Single-page UI, vanilla JS, no framework
- Templates in project folder with {{placeholder}} variables

---

### Task 1: Project Scaffolding

**Files:**
- Create: `requirements.txt`
- Create: `templates/email.mjml`
- Create: `templates/fields.json`

**Interfaces:**
- Produces: Project skeleton ready for modules

- [ ] **Step 1: Create requirements.txt**

```
flask>=3.0
playwright>=1.40
pywin32>=306; sys_platform == "win32"
```

- [ ] **Step 2: Create sample email.mjml**

```xml
<mjml>
  <mj-head>
    <mj-title>Hello {{customer_name}}</mj-title>
  </mj-head>
  <mj-body>
    <mj-section>
      <mj-column>
        <mj-text font-size="20px">Dear {{customer_name}},</mj-text>
        <mj-text>{{message_body}}</mj-text>
        <mj-text>Best regards,</mj-text>
        <mj-text>{{sender_name}}</mj-text>
      </mj-column>
    </mj-section>
  </mj-body>
</mjml>
```

- [ ] **Step 3: Create fields.json**

```json
[
  {"name": "customer_name", "label": "Customer Name", "type": "text"},
  {"name": "message_body", "label": "Message", "type": "textarea"},
  {"name": "sender_name", "label": "Your Name", "type": "text"}
]
```

---

### Task 2: Renderer Module

**Files:**
- Create: `renderer.py`

**Interfaces:**
- Produces: `load_template(path) -> str`, `get_fields(path) -> list[dict]`, `substitute(template, values) -> str`, `render_mjml(mjml) -> str`

- [ ] **Step 1: Write renderer.py**

```python
import subprocess
import tempfile
import json
import os
from pathlib import Path

TEMPLATES_DIR = Path(__file__).parent / "templates"

def load_template(filename: str = "email.mjml") -> str:
    path = TEMPLATES_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"Template not found: {path}")
    return path.read_text(encoding="utf-8")

def get_fields(filename: str = "fields.json") -> list[dict]:
    path = TEMPLATES_DIR / filename
    if not path.exists():
        return []
    return json.loads(path.read_text(encoding="utf-8"))

def substitute(template: str, values: dict[str, str]) -> str:
    result = template
    for key, val in values.items():
        result = result.replace(f"{{{{{key}}}}}", str(val))
    return result

def render_mjml(mjml_source: str) -> str:
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".mjml", delete=False, encoding="utf-8"
    ) as f:
        f.write(mjml_source)
        mjml_path = f.name

    html_path = mjml_path.replace(".mjml", ".html")

    try:
        result = subprocess.run(
            ["npx", "mjml", mjml_path, "-o", html_path],
            capture_output=True, text=True, timeout=30,
            shell=True if os.name == "nt" else False
        )
        if result.returncode != 0:
            raise RuntimeError(f"MJML render failed: {result.stderr}")

        html = Path(html_path).read_text(encoding="utf-8")
        return html
    finally:
        for p in [mjml_path, html_path]:
            if os.path.exists(p):
                os.unlink(p)
```

---

### Task 3: Sender — Desktop (COM)

**Files:**
- Create: `sender.py` (DesktopSender class only)

**Interfaces:**
- Produces: `DesktopSender` class with `is_available() -> bool`, `send(to, subject, html) -> dict`, `shutdown()`

- [ ] **Step 1: Write sender.py with DesktopSender**

```python
import sys
import platform

class DesktopSender:
    def __init__(self):
        self.outlook = None

    @staticmethod
    def is_available() -> bool:
        if platform.system() != "Windows":
            return False
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            outlook.Quit()
            return True
        except Exception:
            return False

    def connect(self):
        import win32com.client
        self.outlook = win32com.client.Dispatch("Outlook.Application")

    def send(self, to: str, subject: str, html: str) -> dict:
        try:
            if not self.outlook:
                self.connect()
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = to
            mail.Subject = subject
            mail.HTMLBody = html
            mail.Send()
            return {"success": True, "error": None}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def shutdown(self):
        self.outlook = None

    def __str__(self):
        return "Outlook Desktop (COM)"
```

---

### Task 4: Sender — Web (Playwright)

**Files:**
- Modify: `sender.py` (add WebSender class and SenderDispatcher)

**Interfaces:**
- Produces: `WebSender` class, `SenderDispatcher` class
- Consumes: `DesktopSender` from Task 3

- [ ] **Step 1: Implement test with pytest**

```python
# tests/test_sender.py
import pytest
from sender import SenderDispatcher

def test_sender_dispatcher_exists():
    dispatcher = SenderDispatcher()
    assert dispatcher is not None
    assert dispatcher.mode in ("desktop", "web", "unavailable")
```

Run: `pytest tests/test_sender.py -v`
Expected: FAIL (WebSender not yet implemented)

- [ ] **Step 2: Add WebSender and SenderDispatcher to sender.py**

```python
import sys
import platform
import time
import threading

class DesktopSender:
    def __init__(self):
        self.outlook = None

    @staticmethod
    def is_available() -> bool:
        if platform.system() != "Windows":
            return False
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            outlook.Quit()
            return True
        except Exception:
            return False

    def connect(self):
        import win32com.client
        self.outlook = win32com.client.Dispatch("Outlook.Application")

    def send(self, to: str, subject: str, html: str) -> dict:
        try:
            if not self.outlook:
                self.connect()
            mail = self.outlook.CreateItem(0)
            mail.To = to
            mail.Subject = subject
            mail.HTMLBody = html
            mail.Send()
            return {"success": True, "error": None}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def shutdown(self):
        self.outlook = None

    def __str__(self):
        return "Outlook Desktop (COM)"


class WebSender:
    def __init__(self):
        self.browser = None
        self.page = None
        self.playwright = None
        self.logged_in = False
        self._lock = threading.Lock()

    @staticmethod
    def is_available() -> bool:
        try:
            import playwright.sync_api
            return True
        except ImportError:
            return False

    def connect(self):
        from playwright.sync_api import sync_playwright
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=False,
            args=["--no-sandbox"]
        )
        self.page = self.browser.new_page()
        self.page.goto("https://outlook.office.com/mail/", timeout=60000)

    def wait_for_login(self, timeout_seconds: int = 120) -> bool:
        if not self.page:
            return False
        start = time.time()
        while time.time() - start < timeout_seconds:
            url = self.page.url
            if "/mail/inbox" in url or "/mail/" in url and "login" not in url.lower():
                self.logged_in = True
                return True
            time.sleep(1)
        return False

    def send(self, to: str, subject: str, html: str) -> dict:
        try:
            if not self.logged_in or not self.page:
                return {"success": False, "error": "Not logged in"}

            with self._lock:
                self.page.goto(
                    "https://outlook.office.com/mail/deeplink/compose",
                    timeout=30000
                )
                self.page.wait_for_timeout(2000)

                self.page.fill('input[aria-label="To"]', to)
                self.page.fill('input[placeholder="Add a subject"]', subject)

                self.page.evaluate(f"""
                    const html = `{html}`;
                    const editor = document.querySelector('[role="textbox"]');
                    if (editor) {{
                        editor.focus();
                        document.execCommand('selectAll');
                        document.execCommand('insertHTML', false, html);
                    }}
                """)

                self.page.click('button[aria-label="Send"]')
                self.page.wait_for_timeout(3000)

            return {"success": True, "error": None}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_user_email(self) -> str:
        try:
            self.page.goto("https://outlook.office.com/mail/", timeout=15000)
            self.page.wait_for_timeout(2000)
            email = self.page.evaluate("""
                const el = document.querySelector('[data-test-id="user-display-name"]');
                return el ? el.getAttribute('title') || el.textContent : null;
            """)
            return email or "Connected"
        except Exception:
            return "Connected"

    def shutdown(self):
        try:
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass
        self.logged_in = False

    def __str__(self):
        return "Outlook Web (Playwright)"


class SenderDispatcher:
    def __init__(self):
        self.sender = None
        self.mode = None
        self._detect()

    def _detect(self):
        if DesktopSender.is_available():
            self.mode = "desktop"
            self.sender = DesktopSender()
            self.sender.connect()
        elif WebSender.is_available():
            self.mode = "web"
            self.sender = WebSender()
            self.sender.connect()
        else:
            self.mode = "unavailable"
            raise RuntimeError(
                "No sender available. Install Outlook Desktop or Playwright."
            )

    def send(self, to: str, subject: str, html: str) -> dict:
        return self.sender.send(to, subject, html)

    def login_web(self) -> str:
        if self.mode != "web":
            return "Desktop mode — no login needed"
        sender = self.sender
        if not sender.wait_for_login(timeout_seconds=120):
            return "Login timed out"
        return f"Logged in as {sender.get_user_email()}"

    def status(self) -> dict:
        if self.mode == "desktop":
            return {"mode": "desktop", "ready": True, "user": "Outlook Desktop"}
        if self.mode == "web":
            return {
                "mode": "web",
                "ready": self.sender.logged_in,
                "user": self.sender.get_user_email() if self.sender.logged_in else None
            }
        return {"mode": "unavailable", "ready": False, "user": None}

    def shutdown(self):
        if self.sender:
            self.sender.shutdown()
```

Run: `pytest tests/test_sender.py -v`
Expected: PASS (when Playwright installed)

---

### Task 5: Flask Server

**Files:**
- Create: `server.py`

**Interfaces:**
- Consumes: `renderer` (load_template, get_fields, substitute, render_mjml), `sender.SenderDispatcher`
- Produces: Flask app with all routes

- [ ] **Step 1: Write server.py**

```python
from flask import Flask, request, jsonify, render_template_string
import sys
import atexit

from renderer import load_template, get_fields, substitute, render_mjml
from sender import SenderDispatcher

app = Flask(__name__, static_folder="static", static_url_path="/static")

INDEX_HTML = None

def get_index_html():
    global INDEX_HTML
    if INDEX_HTML is None:
        from pathlib import Path
        path = Path(__file__).parent / "templates" / "index.html"
        INDEX_HTML = path.read_text(encoding="utf-8")
    return INDEX_HTML

dispatcher = SenderDispatcher()

def cleanup():
    dispatcher.shutdown()

atexit.register(cleanup)


@app.route("/")
def index():
    return render_template_string(get_index_html())


@app.route("/api/fields")
def api_fields():
    return jsonify(get_fields())


@app.route("/api/preview", methods=["POST"])
def api_preview():
    data = request.get_json()
    if not data or "fields" not in data:
        return jsonify({"error": "fields required"}), 400

    try:
        template = load_template()
        substituted = substitute(template, data["fields"])
        html = render_mjml(substituted)
        return jsonify({"html": html})
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/send", methods=["POST"])
def api_send():
    data = request.get_json()
    if not data:
        return jsonify({"error": "request body required"}), 400

    required = ["to", "subject", "fields"]
    for key in required:
        if key not in data:
            return jsonify({"error": f"'{key}' is required"}), 400

    try:
        template = load_template()
        substituted = substitute(template, data["fields"])
        html = render_mjml(substituted)

        result = dispatcher.send(data["to"], data["subject"], html)
        return jsonify(result)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/session/status")
def api_session_status():
    return jsonify(dispatcher.status())


@app.route("/api/session/login", methods=["POST"])
def api_session_login():
    msg = dispatcher.login_web()
    return jsonify({"message": msg})


@app.route("/api/session/logout", methods=["POST"])
def api_session_logout():
    dispatcher.shutdown()
    return jsonify({"message": "Session closed"})


if __name__ == "__main__":
    print(f"\n  Email Blast Tool — Mode: {dispatcher.mode}")
    print(f"  Open http://localhost:5000 in your browser\n")
    app.run(debug=True, port=5000)
```

---

### Task 6: Frontend

**Files:**
- Create: `templates/index.html`
- Create: `static/style.css`

- [ ] **Step 1: Write index.html**

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Email Blast Tool</title>
    <link rel="stylesheet" href="/static/style.css">
</head>
<body>
    <div id="app">
        <header>
            <h1>Email Blast Tool</h1>
            <div id="status-bar">
                <span id="status-mode">Checking...</span>
                <span id="status-user"></span>
                <button id="btn-login" style="display:none">Connect Outlook</button>
            </div>
        </header>

        <main>
            <section id="compose-section">
                <h2>Compose</h2>
                <div class="form-group">
                    <label for="input-to">To:</label>
                    <input type="email" id="input-to" placeholder="customer@example.com" required>
                </div>
                <div class="form-group">
                    <label for="input-subject">Subject:</label>
                    <input type="text" id="input-subject" placeholder="Email subject" required>
                </div>
                <div id="dynamic-fields"></div>

                <div class="actions">
                    <button id="btn-preview">Preview</button>
                    <button id="btn-send" disabled>Send</button>
                </div>
            </section>

            <section id="preview-section" style="display:none">
                <h2>Preview</h2>
                <div id="preview-status"></div>
                <iframe id="preview-frame" sandbox="allow-same-origin"></iframe>
            </section>

            <section id="results-section" style="display:none">
                <h2>Send Result</h2>
                <div id="results-list"></div>
            </section>
        </main>
    </div>

    <script>
        let fields = [];
        let mode = '';

        async function loadStatus() {
            const res = await fetch('/api/session/status');
            const data = await res.json();
            mode = data.mode;
            document.getElementById('status-mode').textContent = data.mode === 'desktop'
                ? 'Outlook Desktop (auto)'
                : 'Outlook Web';
            document.getElementById('status-user').textContent = data.user || '';

            if (data.mode === 'web' && !data.ready) {
                document.getElementById('btn-login').style.display = 'inline-block';
            }
            if (data.mode === 'desktop' || data.ready) {
                document.getElementById('btn-send').disabled = false;
            }
        }

        async function loadFields() {
            const res = await fetch('/api/fields');
            fields = await res.json();
            const container = document.getElementById('dynamic-fields');
            container.innerHTML = fields.map(f =>
                `<div class="form-group">
                    <label for="field-${f.name}">${f.label}:</label>
                    ${f.type === 'textarea'
                        ? `<textarea id="field-${f.name}" rows="4"></textarea>`
                        : `<input type="text" id="field-${f.name}">`
                    }
                </div>`
            ).join('');
        }

        function getFieldValues() {
            const values = {};
            fields.forEach(f => {
                const el = document.getElementById(`field-${f.name}`);
                values[f.name] = el ? el.value : '';
            });
            return values;
        }

        async function doPreview() {
            const values = getFieldValues();
            const res = await fetch('/api/preview', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ fields: values })
            });
            const data = await res.json();

            if (data.error) {
                document.getElementById('preview-status').textContent = 'Error: ' + data.error;
                return;
            }

            document.getElementById('preview-section').style.display = 'block';
            document.getElementById('preview-status').textContent = 'Preview ready';
            const frame = document.getElementById('preview-frame');
            frame.srcdoc = data.html;
        }

        async function doSend() {
            const to = document.getElementById('input-to').value;
            const subject = document.getElementById('input-subject').value;
            const values = getFieldValues();

            if (!to || !subject) {
                alert('To and Subject are required');
                return;
            }

            const res = await fetch('/api/send', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ to, subject, fields: values })
            });
            const data = await res.json();

            document.getElementById('results-section').style.display = 'block';
            document.getElementById('results-list').innerHTML = data.success
                ? `<div class="result success">Sent to ${to}</div>`
                : `<div class="result error">Failed: ${data.error}</div>`;
        }

        async function doLogin() {
            document.getElementById('btn-login').disabled = true;
            document.getElementById('btn-login').textContent = 'Waiting for login...';
            const res = await fetch('/api/session/login', { method: 'POST' });
            const data = await res.json();
            await loadStatus();
            document.getElementById('btn-login').textContent = 'Connect Outlook';
            document.getElementById('btn-login').disabled = false;
        }

        document.getElementById('btn-preview').addEventListener('click', doPreview);
        document.getElementById('btn-send').addEventListener('click', doSend);
        document.getElementById('btn-login').addEventListener('click', doLogin);

        loadStatus();
        loadFields();
    </script>
</body>
</html>
```

- [ ] **Step 2: Write style.css**

```css
* { margin: 0; padding: 0; box-sizing: border-box; }

body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: #f5f5f5;
    color: #333;
    line-height: 1.5;
}

#app {
    max-width: 800px;
    margin: 0 auto;
    padding: 20px;
}

header {
    background: #fff;
    padding: 16px 20px;
    border-radius: 8px;
    margin-bottom: 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    display: flex;
    justify-content: space-between;
    align-items: center;
}

header h1 { font-size: 20px; }

#status-bar {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 13px;
    color: #666;
}

#status-mode { font-weight: 600; color: #333; }

section {
    background: #fff;
    padding: 20px;
    border-radius: 8px;
    margin-bottom: 16px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}

section h2 { font-size: 16px; margin-bottom: 12px; }

.form-group { margin-bottom: 12px; }

.form-group label {
    display: block;
    font-size: 13px;
    font-weight: 600;
    margin-bottom: 4px;
    color: #555;
}

.form-group input,
.form-group textarea {
    width: 100%;
    padding: 8px 12px;
    border: 1px solid #ddd;
    border-radius: 6px;
    font-size: 14px;
    font-family: inherit;
}

.form-group textarea { resize: vertical; }

.actions { margin-top: 16px; display: flex; gap: 10px; }

button {
    padding: 8px 20px;
    border: none;
    border-radius: 6px;
    font-size: 14px;
    cursor: pointer;
    font-weight: 600;
}

#btn-preview { background: #e5e7eb; color: #333; }
#btn-send { background: #0078d4; color: #fff; }
#btn-send:disabled { background: #ccc; cursor: not-allowed; }
#btn-login { background: #0078d4; color: #fff; }

#preview-frame {
    width: 100%;
    height: 400px;
    border: 1px solid #ddd;
    border-radius: 6px;
}

#preview-status {
    font-size: 13px;
    color: #666;
    margin-bottom: 8px;
}

.result {
    padding: 10px 14px;
    border-radius: 6px;
    margin-bottom: 6px;
    font-size: 14px;
}

.result.success { background: #d4edda; color: #155724; }
.result.error { background: #f8d7da; color: #721c24; }
```

---

### Task 7: End-to-End Verification

**Files:**
- Modify: `templates/email.mjml` (user provides real template)

- [ ] **Step 1: Install dependencies**

Run: `pip install -r requirements.txt`
Run: `python -m playwright install chromium`

- [ ] **Step 2: Start server and verify**

Run: `python server.py`
Open: `http://localhost:5000`
Verify: Status bar shows mode, fields load, preview works

- [ ] **Step 3: Test desktop send (Windows + Outlook)**

Fill form → Preview → Send → Verify email arrives

- [ ] **Step 4: Test web send (Playwright)**

Click "Connect Outlook" → Sign in → Fill form → Preview → Send → Verify email arrives
