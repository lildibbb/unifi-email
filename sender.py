import sys
import platform
import time
import threading
import queue
import logging

log = logging.getLogger(__name__)


class DesktopSender:
    def __init__(self):
        self.outlook = None

    @staticmethod
    def is_available() -> bool:
        if platform.system() != "Windows":
            return False
        try:
            import subprocess
            result = subprocess.run(
                'tasklist /FI "IMAGENAME eq OUTLOOK.EXE" 2>&1',
                capture_output=True, text=True, shell=True
            )
            return "OUTLOOK.EXE" in result.stdout.upper()
        except Exception:
            return False

    def connect(self):
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        self.outlook = win32com.client.Dispatch("Outlook.Application")

    def send(self, to: str, subject: str, html: str) -> dict:
        try:
            import pythoncom
            pythoncom.CoInitialize()
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


def _is_logged_in_url(url: str) -> bool:
    if not url:
        return False
    url_lower = url.lower()
    if "login.microsoftonline.com" in url_lower:
        return False
    if "login.live.com" in url_lower:
        return False
    if "microsoftonline.com" in url_lower and "/sas/" in url_lower:
        return False
    if "outlook.office.com" in url_lower and "/mail/" in url_lower:
        return True
    if "outlook.live.com" in url_lower and "/mail/" in url_lower:
        return True
    if "outlook.office.com" in url_lower and "/owa/" in url_lower:
        return True
    if "outlook.office.com" in url_lower and "/calendar/" in url_lower:
        return True
    return False


def _is_logged_in_dom(page) -> bool:
    try:
        result = page.evaluate("""
            () => {
                if (document.querySelector('[data-test-id="user-display-name"]')) return true;
                if (document.querySelector('[aria-label="New mail"]')) return true;
                if (document.querySelector('[data-app-section="MailModule"]')) return true;
                const inboxLabel = document.querySelector('[title="Inbox"], [aria-label="Inbox"]');
                if (inboxLabel) return true;
                const composeIndicators = document.querySelector(
                    'input[aria-label="To"], [role="textbox"][contenteditable="true"], button[aria-label="Send"]'
                );
                if (composeIndicators) return true;
                if (window.location.href.includes('outlook.office.com') && !window.location.href.includes('login')) {
                    const signInBtn = document.querySelector('a[href*="login"], input[type="submit"][value*="Sign in"]');
                    if (!signInBtn) return true;
                }
                return null;
            }
        """)
        return result is True
    except Exception:
        return False


class WebSender:
    def __init__(self):
        self._browser_thread = None
        self._cmd_queue = queue.Queue()
        self._result_queue = queue.Queue()
        self.logged_in = False
        self._user_email = None
        self._login_error = None
        self._connecting = False

    @staticmethod
    def is_available() -> bool:
        try:
            import playwright.sync_api
            return True
        except ImportError:
            return False

    def connect(self):
        self._connecting = True
        self.logged_in = False
        self._user_email = None
        self._login_error = None
        if self._browser_thread and self._browser_thread.is_alive():
            self._cmd_queue.put(("reset", []))
        else:
            self._browser_thread = threading.Thread(target=self._run_browser, daemon=True)
            self._browser_thread.start()

    def _run_browser(self):
        try:
            from playwright.sync_api import sync_playwright
            pw = sync_playwright().start()
            browser = pw.chromium.launch(headless=False, args=["--no-sandbox", "--guest"])
            context = browser.new_context(no_viewport=True)
            page = context.new_page()
            page.goto("https://outlook.office.com/mail/", timeout=60000)

            login_start = time.time()
            last_url_check = 0

            while True:
                try:
                    cmd, args = self._cmd_queue.get(timeout=0.3)
                except queue.Empty:
                    cmd = None

                now = time.time()
                if now - last_url_check > 2:
                    last_url_check = now
                    was_logged_in = self.logged_in
                    is_logged_in = _is_logged_in_dom(page)

                    if is_logged_in:
                        if not was_logged_in:
                            self._user_email = self._extract_email(page)
                        self.logged_in = True
                        self._connecting = False

                    elif self._connecting:
                        if now - login_start > 300:
                            self._connecting = False
                            self._login_error = "Login timed out (5 min)"

                if cmd == "send":
                    try:
                        self._result_queue.put(self._do_send(page, *args))
                    except Exception as e:
                        self._result_queue.put({"success": False, "error": str(e)})
                elif cmd == "reset":
                    try:
                        page.goto("https://outlook.office.com/mail/", timeout=30000, wait_until="domcontentloaded")
                        page.wait_for_timeout(2000)
                        self.logged_in = False
                        self._user_email = None
                        self._connecting = True
                        self._result_queue.put(None)
                    except Exception as e:
                        self._result_queue.put({"success": False, "error": str(e)})
                elif cmd == "email":
                    self._result_queue.put(self._extract_email(page))
                elif cmd == "shutdown":
                    try:
                        context.close()
                        browser.close()
                        pw.stop()
                    except Exception:
                        pass
                    self._result_queue.put(None)
                    self.logged_in = False
                    return
                elif cmd is not None:
                    self._result_queue.put(None)

        except Exception as e:
            self._connecting = False
            self._login_error = str(e)

    def _extract_email(self, page) -> str:
        try:
            email = page.evaluate("""
                () => {
                    const el = document.querySelector('[data-test-id="user-display-name"]');
                    if (el) return el.getAttribute('title') || el.textContent || '';
                    const me = document.querySelector('div[data-testid="O365Header"]');
                    if (me) return me.textContent || '';
                    return '';
                }
            """)
            return email.strip() or "Connected"
        except Exception:
            return "Connected"

    def _do_send(self, page, to, subject, html):
        import os
        if not self.logged_in:
            return {"success": False, "error": "Not logged in — click Connect Outlook first"}

        # ── Step 1: Verify session — only reload if DOM check fails ─────────────
        if not _is_logged_in_dom(page):
            try:
                page.goto("https://outlook.office.com/mail/", timeout=15000, wait_until="domcontentloaded")
                page.wait_for_timeout(1500)
            except Exception:
                pass
            if not _is_logged_in_dom(page):
                self.logged_in = False
                self._connecting = False
                return {"success": False, "error": "Session expired — click Connect Outlook to sign in again"}

        # ── Step 2: Open compose ─────────────────────────────────────────────────
        compose_page = page
        try:
            with page.context.expect_page(timeout=5000) as new_page_info:
                page.goto(
                    "https://outlook.office.com/mail/deeplink/compose",
                    timeout=30000,
                    wait_until="domcontentloaded",
                )
            popup = new_page_info.value
            popup.wait_for_load_state("domcontentloaded")
            compose_page = popup
        except Exception:
            # No popup — compose opened in the same tab (most common)
            compose_page = page

        # ── Step 3: Wait for the compose form to be ready ───────────────────────
        TO_SELECTORS = [
            'input[aria-label="To"]',
            'div[aria-label="To"] input',
            'div[aria-label="To"]',
            '[role="combobox"][aria-label*="To"]',
            'input[type="text"]:visible',
        ]

        compose_loaded = False
        for sel in TO_SELECTORS:
            try:
                compose_page.wait_for_selector(sel, timeout=12000, state="visible")
                compose_loaded = True
                break
            except Exception:
                continue

        if not compose_loaded:
            try:
                snap = compose_page.evaluate("""
                    () => JSON.stringify({
                        url: window.location.href,
                        inputs: [...document.querySelectorAll('input,textarea,[contenteditable]')].map(e=>({
                            tag:e.tagName, id:e.id, role:e.getAttribute('role'),
                            aria:e.getAttribute('aria-label'), ph:e.placeholder||'',
                            ce:e.contentEditable
                        })),
                        buttons: [...document.querySelectorAll('button')].map(b=>({
                            text:b.textContent.trim().slice(0,40),
                            aria:b.getAttribute('aria-label'), title:b.title
                        }))
                    })
                """)
                log_path = os.path.join(os.path.dirname(__file__), "compose_debug.json")
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write(snap)
            except Exception:
                pass
            return {"success": False, "error": "Compose window did not load — check compose_debug.json for DOM snapshot"}

        # ── Step 4: Fill the To field ────────────────────────────────────────────
        to_filled = False

        for strategy in ["label", "role", "css"]:
            try:
                if strategy == "label":
                    loc = compose_page.get_by_label("To", exact=True)
                elif strategy == "role":
                    loc = compose_page.get_by_role("combobox", name="To")
                else:
                    loc = compose_page.locator(
                        'input[aria-label="To"], [role="combobox"][aria-label*="To"], '
                        'div[aria-label="To"] input'
                    ).first

                loc.wait_for(state="visible", timeout=5000)
                loc.click(timeout=3000)
                compose_page.wait_for_timeout(200)
                compose_page.keyboard.type(to, delay=20)
                compose_page.wait_for_timeout(400)
                compose_page.keyboard.press("Enter")
                compose_page.wait_for_timeout(300)
                to_filled = True
                break
            except Exception:
                continue

        if not to_filled:
            try:
                compose_page.get_by_text("To", exact=True).first.click(timeout=3000)
                compose_page.wait_for_timeout(200)
                compose_page.keyboard.type(to, delay=20)
                compose_page.wait_for_timeout(400)
                compose_page.keyboard.press("Enter")
                compose_page.wait_for_timeout(300)
                to_filled = True
            except Exception:
                pass

        if not to_filled:
            try:
                compose_page.keyboard.press("Tab")
                compose_page.wait_for_timeout(200)
                compose_page.keyboard.type(to, delay=20)
                compose_page.wait_for_timeout(400)
                compose_page.keyboard.press("Enter")
                compose_page.wait_for_timeout(300)
                to_filled = True
            except Exception:
                pass

        if not to_filled:
            try:
                snap = compose_page.evaluate("""
                    () => JSON.stringify([...document.querySelectorAll('input,textarea,[contenteditable],[role="combobox"],[role="textbox"]')].map(e=>({
                        tag:e.tagName, id:e.id, role:e.getAttribute('role'),
                        aria:e.getAttribute('aria-label'), ph:e.placeholder||'',
                        ce:e.contentEditable, vis:e.offsetParent!==null
                    })))
                """)
                log_path = os.path.join(os.path.dirname(__file__), "to_field_debug.json")
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write(snap)
            except Exception:
                pass
            return {"success": False, "error": "Could not fill To field — check to_field_debug.json for DOM snapshot"}

        # ── Step 5: Fill Subject ─────────────────────────────────────────────────
        subject_filled = False
        for strategy in ["label", "placeholder", "css"]:
            try:
                if strategy == "label":
                    loc = compose_page.get_by_label("Add a subject")
                elif strategy == "placeholder":
                    loc = compose_page.get_by_placeholder("Add a subject")
                else:
                    loc = compose_page.locator(
                        'input[placeholder="Add a subject"], input[aria-label="Subject"], '
                        '[aria-label="Add a subject"]'
                    ).first

                loc.wait_for(state="visible", timeout=5000)
                loc.click(timeout=3000)
                loc.fill(subject, timeout=5000)
                subject_filled = True
                break
            except Exception:
                continue

        if not subject_filled:
            return {"success": False, "error": "Could not fill Subject field"}

        # ── Step 6: Inject HTML body ─────────────────────────────────────────────
        # IMPORTANT: We pass `html` as a JS argument (not string interpolation)
        # so that base64 image data, backticks, $ signs etc. are safe.
        compose_page.wait_for_timeout(400)

        _INJECT_SCRIPT = """
            (htmlBody) => {
                const bodySelectors = [
                    '[aria-label="Message body"]',
                    '[aria-label*="body" i]',
                    '[data-testid="compose-editor"]',
                    '[contenteditable="true"][role="textbox"]',
                    '#editorBody',
                    '.ms-composer [contenteditable="true"]',
                    '[contenteditable="true"]',
                ];
                for (const sel of bodySelectors) {
                    const editors = [...document.querySelectorAll(sel)];
                    // Pick the LAST match — body is always after the To/CC fields
                    const el = editors[editors.length - 1];
                    if (el && el.contentEditable === 'true') {
                        el.focus();
                        // Try execCommand first (React-safe)
                        document.execCommand('selectAll', false, null);
                        const ok = document.execCommand('insertHTML', false, htmlBody);
                        if (!ok || el.innerHTML.length < 10) {
                            // Fall back to direct innerHTML
                            el.innerHTML = htmlBody;
                            el.dispatchEvent(new Event('input', {bubbles: true}));
                            el.dispatchEvent(new Event('change', {bubbles: true}));
                        }
                        return true;
                    }
                }
                return false;
            }
        """

        def _inject_body(frame_or_page, body_html: str) -> bool:
            """Inject HTML by passing it as a JS argument — safe for any content."""
            try:
                return bool(frame_or_page.evaluate(_INJECT_SCRIPT, body_html))
            except Exception:
                return False

        injected = _inject_body(compose_page, html)

        # Try child iframes if main frame failed
        if not injected:
            try:
                for frame in compose_page.frames:
                    if frame == compose_page.main_frame:
                        continue
                    if _inject_body(frame, html):
                        injected = True
                        break
            except Exception:
                pass

        if not injected:
            return {"success": False, "error": "Could not inject email body into Outlook compose window"}

        compose_page.wait_for_timeout(500)

        # ── Step 7: Click Send ───────────────────────────────────────────────────
        sent = False
        for strategy in ["label", "role", "css"]:
            try:
                if strategy == "label":
                    loc = compose_page.get_by_label("Send")
                elif strategy == "role":
                    loc = compose_page.get_by_role("button", name="Send")
                else:
                    loc = compose_page.locator(
                        'button[aria-label="Send"], button[title="Send"]'
                    ).first

                loc.wait_for(state="visible", timeout=5000)
                loc.click(timeout=5000)
                sent = True
                break
            except Exception:
                continue

        if not sent:
            return {"success": False, "error": "Could not click Send button — Outlook UI may have changed"}

        # Wait for compose to close after send
        try:
            compose_page.wait_for_url("**/mail/**", timeout=8000)
        except Exception:
            pass
        compose_page.wait_for_timeout(800)
        return {"success": True, "error": None}

    def send(self, to, subject, html):
        if not self._browser_thread or not self._browser_thread.is_alive():
            return {"success": False, "error": "Not connected — click Connect Outlook"}
        self._cmd_queue.put(("send", [to, subject, html]))
        try:
            return self._result_queue.get(timeout=75)
        except queue.Empty:
            return {"success": False, "error": "Send timed out — the operation took too long"}

    def shutdown(self):
        if self._browser_thread and self._browser_thread.is_alive():
            self._cmd_queue.put(("shutdown", []))
            self._browser_thread.join(timeout=5)
        self.logged_in = False
        self._connecting = False

    def __str__(self):
        return "Outlook Web (Playwright)"


class SenderDispatcher:
    def __init__(self):
        self.sender = None
        self.mode = None
        self._detect()

    def _detect(self):
        if WebSender.is_available():
            self.mode = "web"
            self.sender = WebSender()
        else:
            self.mode = "unavailable"
            raise RuntimeError(
                "No sender available. Install Playwright: "
                "pip install playwright && python -m playwright install chromium"
            )

    def connect(self):
        if self.mode == "web":
            self.sender.connect()

    def send(self, to: str, subject: str, html: str) -> dict:
        return self.sender.send(to, subject, html)

    def status(self) -> dict:
        if self.mode == "desktop":
            return {"mode": "desktop", "ready": True, "user": "Outlook Desktop"}
        if self.mode == "web":
            sender = self.sender
            return {
                "mode": "web",
                "ready": sender.logged_in,
                "connecting": sender._connecting,
                "user": sender._user_email if sender.logged_in else None,
                "error": sender._login_error,
            }
        return {"mode": "unavailable", "ready": False, "user": None}

    def shutdown(self):
        if self.sender:
            self.sender.shutdown()
