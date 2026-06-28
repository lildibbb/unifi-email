# Email Blast Tool — Design Spec

**Date:** 2026-06-27
**Status:** Approved

## Overview

A local CLI-driven web tool that lets team members send personalized emails via their company Microsoft 365 accounts — without any Azure app registration, IT admin involvement, or database. Two paths: COM automation (Outlook Desktop) for company laptops, and Playwright browser automation (Outlook Web) for personal laptops. MJML templates with dynamic fields are rendered to HTML and sent through the user's already-authenticated Outlook session.

## Constraints

- No database — data flows through memory, destroyed on shutdown
- No Azure AD app registration — bypasses IT entirely
- No SMTP — company enforces MFA, SMTP AUTH unavailable
- Must work on Windows (company laptops) and Mac (personal laptops)
- Must support sending to multiple recipients in one session
- MJML templates stored in project folder, with `{{placeholder}}` variables
- Single-page UI, no frontend framework

## Architecture

```
project/
├── server.py              # Flask app, all routes, startup orchestration
├── sender.py              # Sender dispatcher: DesktopSender | WebSender
├── renderer.py            # MJML → HTML via subprocess (npx mjml)
├── requirements.txt       # flask, playwright, pywin32 (Windows only)
├── static/
│   └── style.css          # Minimal styling
├── templates/
│   ├── index.html         # Single-page UI (vanilla JS)
│   ├── email.mjml         # MJML template with {{placeholders}}
│   └── fields.json        # Field definitions: [{name, label, type}]
└── docs/
    └── superpowers/
        └── specs/
            └── 2026-06-27-email-blast-tool-design.md
```

## Routes

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/` | GET | Serve the SPA |
| `/api/fields` | GET | Return template fields from `fields.json` |
| `/api/preview` | POST | Substitute placeholders, render MJML → HTML, return preview |
| `/api/send` | POST | Send to one recipient, returns success/error |
| `/api/session/status` | GET | Is sender connected? Which mode? |
| `/api/session/login` | POST | Trigger Playwright login flow (web mode) |
| `/api/session/logout` | POST | Gracefully close sender session |

## Components

### server.py

Flask app. On startup, detects available sender mode and initializes it. All data (field values, recipients, rendered HTML) flows through POST bodies — nothing stored server-side.

### sender.py — SenderDispatcher

```python
class SenderDispatcher:
    def detect_mode() -> str:
        # 1. Try COM (windows + outlook installed)
        # 2. Fallback to WEB (Playwright)
        pass

    def initialize(mode: str):
        # COM: nothing needed (Outlook already running)
        # WEB: launch Playwright Chromium at outlook.office.com
        pass

    def send(to: str, subject: str, html: str) -> dict:
        # Returns {success: bool, error: str | None}
        pass

    def shutdown():
        # Close Playwright browser if open
        pass
```

**DesktopSender (win32com):**
- Connects to running Outlook application via COM
- Creates `MailItem`, sets To/Subject/HTMLBody, calls `.Send()`
- Fully automatic — no login, no browser, zero user interaction
- Windows only

**WebSender (Playwright):**
- Launches single Chromium instance at startup
- User logs in to outlook.office.com once manually (MFA OK)
- After login, browser persists session — all sends reuse same session
- Each send: navigates compose pane, injects To/Subject via `page.fill()`, injects HTML body via `page.evaluate()` clipboard injection, clicks Send
- Works on Windows/Mac/Linux

### renderer.py

```python
def load_template(path: str) -> str:
    # Read .mjml file

def get_fields() -> list[dict]:
    # Read fields.json → [{name, label, type}]

def substitute(template: str, values: dict) -> str:
    # Replace {{name}} with values["name"]

def render_mjml(mjml: str) -> str:
    # Write to temp file
    # Run: npx mjml temp.mjml -o temp.html
    # Read and return HTML
```

### Frontend (index.html)

Single HTML file with vanilla JS. Four-step flow:

1. **Status bar** — shows connection mode and status
2. **Compose form** — To, Subject, dynamic fields (from `/api/fields`)
3. **Preview panel** — renders `/api/preview` response in iframe/div
4. **Send controls** — Send button + per-recipient status rows

## Data Flow

```
1. Page loads → GET /api/fields → renders input form
2. User fills fields → POST /api/preview {fields: {name: "...", ...}}
   → Backend: substitute template → render MJML → return HTML
   → Frontend: show in preview iframe
3. User clicks Send → POST /api/send {to, subject, fields, html}
   → Backend: render final HTML → sender.send(to, subject, html)
   → Return {success, error?}
4. Frontend shows ✓ or ✗ per recipient
```

## Sender Detection Logic

```
On startup:
  1. if platform == "win32" AND win32com can connect to Outlook → COM mode
  2. else → WEB mode (Playwright)
```

## UX — Web Mode Login Flow

1. Status: "Web mode — click Connect to sign in"
2. User clicks "Connect" → POST /api/session/login
3. Playwright opens Chromium → outlook.office.com
4. User signs in manually (MFA, password — whatever their company requires)
5. Playwright detects navigation to inbox → marks session as ready
6. Status updates: "Connected as user@tm.com.my"
7. User composes and sends — all emails through same session
8. On app shutdown or logout → browser closes, session destroyed

## Error Handling

- COM connection failure → auto-fallback to Web mode
- Playwright login timeout (120s) → show error, allow retry
- MJML render failure → return stderr to frontend, prevent send
- Send failure → return error to frontend, continue with next recipient
- Browser crash mid-session → auto-restart and prompt re-login

## Non-Goals

- No database persistence
- No user management or multi-tenancy
- No email tracking or analytics
- No template editor — templates edited as files
- No SMTP support (MFA blocks it)
- No Azure AD app registration
- No token extraction or storage (v2 consideration only)
