// background.js — Email Blast Outlook Connector (MV3)

var OUTLOOK_URL  = 'https://outlook.cloud.microsoft/mail/';
var OUTLOOK_URL2 = 'https://outlook.office.com/mail/';
var COMPOSE_URL  = 'https://outlook.cloud.microsoft/mail/deeplink/compose';
var COMPOSE_URL2 = 'https://outlook.office.com/mail/deeplink/compose';

console.log('[EmailBlast] background.js loaded');

// ── Storage ──────────────────────────────────────────────────────────────────
function getState() {
  return chrome.storage.local.get({ status: 'disconnected', user: null, outlookTabId: null, appTabId: null });
}
function setState(p) { return chrome.storage.local.set(p); }

// Wait for a tab to finish loading (event-driven, keeps SW alive)
function waitForTabLoad(tabId, timeoutMs) {
  return new Promise(function (resolve) {
    var done = false;
    function finish() { if (!done) { done = true; chrome.tabs.onUpdated.removeListener(listener); resolve(); } }
    function listener(id, changeInfo) { if (id === tabId && changeInfo.status === 'complete') finish(); }
    chrome.tabs.onUpdated.addListener(listener);
    setTimeout(finish, timeoutMs || 12000);
  });
}

// ── URL check ────────────────────────────────────────────────────────────────
function isLoggedInUrl(url) {
  var u = (url || '').toLowerCase();
  if (!u) return false;
  if (u.indexOf('login.microsoftonline.com') > -1) return false;
  if (u.indexOf('login.live.com')            > -1) return false;
  if (u.indexOf('microsoftonline.com')        > -1) return false;
  if (u.indexOf('/login')                    > -1) return false;
  if (u.indexOf('/signin')                   > -1) return false;
  var onOutlook = u.indexOf('outlook.office.com')    > -1 ||
                  u.indexOf('outlook.live.com')       > -1 ||
                  u.indexOf('outlook.cloud.microsoft') > -1;
  if (!onOutlook) return false;
  return u.indexOf('/mail') > -1 || u.indexOf('/owa') > -1 || u.indexOf('/calendar') > -1;
}

// ── Notify web page ──────────────────────────────────────────────────────────
function notifyPage(msg) {
  getState().then(function (s) {
    if (s.appTabId) chrome.tabs.sendMessage(s.appTabId, msg, function () { void chrome.runtime.lastError; });
  });
  chrome.tabs.query({ url: ['http://localhost/*', 'http://127.0.0.1/*'] }, function (tabs) {
    (tabs || []).forEach(function (t) {
      chrome.tabs.sendMessage(t.id, msg, function () { void chrome.runtime.lastError; });
    });
  });
}

// ── Mark connected ───────────────────────────────────────────────────────────
function markConnected() {
  return getState().then(function (s) {
    if (s.status !== 'connecting') return;
    console.log('[EmailBlast] marking connected');
    return setState({ status: 'connected', user: 'Connected' }).then(function () {
      notifyPage({ action: 'status', status: 'connected', ready: true, user: 'Connected' });
    });
  });
}

// ── onUpdated ────────────────────────────────────────────────────────────────
chrome.tabs.onUpdated.addListener(function (tabId, changeInfo, tab) {
  if (changeInfo.status !== 'complete') return;
  console.log('[EmailBlast] onUpdated tabId=' + tabId + ' url=' + tab.url);
  if (!isLoggedInUrl(tab.url)) return;
  getState().then(function (s) {
    if (s.status !== 'connecting') return;
    console.log('[EmailBlast] onUpdated: adopting tab ' + tabId + ' → marking connected');
    setState({ outlookTabId: tabId }).then(function () { markConnected(); });
  });
});

// ── onRemoved ────────────────────────────────────────────────────────────────
chrome.tabs.onRemoved.addListener(function (tabId) {
  getState().then(function (s) {
    if (s.outlookTabId !== tabId)    return;
    if (s.status === 'disconnected') return;
    setState({ status: 'disconnected', user: null, outlookTabId: null });
    notifyPage({ action: 'status', status: 'disconnected' });
  });
});

// ── Messages ─────────────────────────────────────────────────────────────────
chrome.runtime.onMessage.addListener(function (msg, sender, sendResponse) {
  console.log('[EmailBlast] message received: ' + JSON.stringify(msg && msg.action));
  if (!msg || !msg.action) { sendResponse({}); return false; }

  if (sender && sender.tab && sender.tab.id) {
    var u = (sender.tab.url || '').toLowerCase();
    if (u.indexOf('localhost') > -1 || u.indexOf('127.0.0.1') > -1) {
      setState({ appTabId: sender.tab.id });
    }
  }

  if (msg.action === 'connect') {
    handleConnect()
      .then(function () { sendResponse({ ok: true }); })
      .catch(function (e) { console.error('[EmailBlast] connect error', e); sendResponse({ ok: false, error: String(e) }); });
    return true;
  }
  if (msg.action === 'disconnect') {
    handleDisconnect().then(function () { sendResponse({ ok: true }); }).catch(function () { sendResponse({}); });
    return true;
  }
  if (msg.action === 'status') {
    getState().then(function (s) {
      console.log('[EmailBlast] status poll: ' + s.status + ' tabId=' + s.outlookTabId);
      if (s.status === 'connecting') {
        var ALL_OUTLOOK = ['https://outlook.office.com/*', 'https://outlook.live.com/*', 'https://outlook.cloud.microsoft/*'];
        if (s.outlookTabId) {
          chrome.tabs.get(s.outlookTabId, function (tab) {
            void chrome.runtime.lastError;
            console.log('[EmailBlast] tabs.get url=' + (tab && tab.url));
            if (tab && isLoggedInUrl(tab.url)) {
              setState({ status: 'connected', user: 'Connected' }).then(function () {
                notifyPage({ action: 'status', status: 'connected', ready: true, user: 'Connected' });
                sendResponse({ status: 'connected', user: 'Connected' });
              });
            } else if (!tab) {
              chrome.tabs.query({ url: ALL_OUTLOOK }, function (tabs) {
                void chrome.runtime.lastError;
                var found = null;
                for (var i = 0; tabs && i < tabs.length; i++) { if (isLoggedInUrl(tabs[i].url)) { found = tabs[i]; break; } }
                if (found) {
                  setState({ outlookTabId: found.id, status: 'connected', user: 'Connected' }).then(function () {
                    notifyPage({ action: 'status', status: 'connected', ready: true, user: 'Connected' });
                    sendResponse({ status: 'connected', user: 'Connected' });
                  });
                } else { sendResponse(s); }
              });
            } else { sendResponse(s); }
          });
        } else {
          chrome.tabs.query({ url: ALL_OUTLOOK }, function (tabs) {
            void chrome.runtime.lastError;
            var found = null;
            for (var i = 0; tabs && i < tabs.length; i++) { if (isLoggedInUrl(tabs[i].url)) { found = tabs[i]; break; } }
            if (found) {
              setState({ outlookTabId: found.id, status: 'connected', user: 'Connected' }).then(function () {
                notifyPage({ action: 'status', status: 'connected', ready: true, user: 'Connected' });
                sendResponse({ status: 'connected', user: 'Connected' });
              });
            } else { sendResponse(s); }
          });
        }
      } else {
        sendResponse(s);
      }
    }).catch(function () { sendResponse({ status: 'disconnected' }); });
    return true;
  }
  if (msg.action === 'send') {
    handleSend(msg.to, msg.subject, msg.html)
      .then(function (r) { sendResponse(r); })
      .catch(function (e) { sendResponse({ success: false, error: String(e) }); });
    return true;
  }
  sendResponse({});
  return false;
});

// ── handleConnect ─────────────────────────────────────────────────────────────
function handleConnect() {
  console.log('[EmailBlast] handleConnect start');
  return setState({ status: 'connecting', user: null, outlookTabId: null }).then(function () {
    notifyPage({ action: 'status', status: 'connecting', connecting: true });
    return new Promise(function (resolve, reject) {
      chrome.tabs.query({ url: ['https://outlook.office.com/*', 'https://outlook.live.com/*', 'https://outlook.cloud.microsoft/*'] }, function (tabs) {
        void chrome.runtime.lastError;
        console.log('[EmailBlast] found ' + (tabs ? tabs.length : 0) + ' existing outlook tabs');
        if (tabs && tabs.length > 0) {
          var t = tabs[0];
          var navUrl = (t.url || '').indexOf('cloud.microsoft') > -1 ? OUTLOOK_URL : OUTLOOK_URL2;
          console.log('[EmailBlast] updating existing tab ' + t.id + ' to ' + navUrl);
          chrome.tabs.update(t.id, { url: navUrl, active: true }, function (updated) {
            if (chrome.runtime.lastError) { reject(new Error(chrome.runtime.lastError.message)); return; }
            setState({ outlookTabId: updated.id }).then(resolve).catch(reject);
          });
        } else {
          console.log('[EmailBlast] creating new outlook tab');
          chrome.tabs.create({ url: OUTLOOK_URL, active: true }, function (tab) {
            if (chrome.runtime.lastError) { reject(new Error(chrome.runtime.lastError.message)); return; }
            console.log('[EmailBlast] tab created, tabId=' + tab.id);
            setState({ outlookTabId: tab.id }).then(resolve).catch(reject);
          });
        }
      });
    });
  });
}

// ── handleDisconnect ──────────────────────────────────────────────────────────
function handleDisconnect() {
  return getState().then(function (s) {
    var p = s.outlookTabId ? chrome.tabs.remove(s.outlookTabId).catch(function () {}) : Promise.resolve();
    return p.then(function () { return setState({ status: 'disconnected', user: null, outlookTabId: null }); })
      .then(function () { notifyPage({ action: 'status', status: 'disconnected' }); });
  });
}

// ── handleSend ────────────────────────────────────────────────────────────────
function handleSend(to, subject, html) {
  return getState().then(function (s) {
    if (s.status !== 'connected') return { success: false, error: 'Not connected — click Connect Outlook first' };
    if (!s.outlookTabId) return { success: false, error: 'No Outlook tab — click Connect again' };
    var tabId = s.outlookTabId;
    return new Promise(function (resolve) {
      chrome.tabs.get(tabId, function (tab) {
        void chrome.runtime.lastError;
        var baseComposeUrl = (tab && tab.url && tab.url.indexOf('cloud.microsoft') > -1) ? COMPOSE_URL : COMPOSE_URL2;
        
        // Pass to and subject as query parameters to let Outlook native code pre-populate them
        var composeUrl = baseComposeUrl + 
          '?to=' + encodeURIComponent(to) + 
          '&subject=' + encodeURIComponent(subject);

        console.log('[EmailBlast] navigating to compose: ' + composeUrl);
        chrome.tabs.update(tabId, { url: composeUrl, active: true }, function () {
          void chrome.runtime.lastError;
          waitForTabLoad(tabId, 12000).then(function () {
            console.log('[EmailBlast] compose tab loaded, injecting send script');
            tryComposeSend(tabId, to, subject, html).then(resolve);
          });
        });
      });
    });
  }).catch(function (e) { return { success: false, error: String(e) }; });
}

function tryComposeSend(tabId, to, subject, html) {
  console.log('[EmailBlast] tryComposeSend executing script on tabId=' + tabId);
  return chrome.scripting.executeScript({
    target: { tabId: tabId },
    func:   doComposeAndSend,
    args:   [{ to: to, subject: subject, html: html }]
  }).then(function (r) {
    var result = r && r[0] ? r[0].result : { success: false, error: 'No result from script' };
    console.log('[EmailBlast] tryComposeSend script result: ' + JSON.stringify(result));
    return result;
  }).catch(function (e) {
    console.error('[EmailBlast] tryComposeSend executeScript error: ', e);
    return { success: false, error: String(e) };
  });
}

// ── doComposeAndSend (injected into Outlook tab — MUST be async + self-contained)
// Uses await delay() NOT busy-wait — busy-waits block React rendering!

async function doComposeAndSend(params) {
  var to = params.to, subject = params.subject, html = params.html;

  function delay(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function findEl(sels) {
    for (var i = 0; i < sels.length; i++) {
      var el = document.querySelector(sels[i]);
      if (el && el.offsetParent !== null) return el;
    }
    return null;
  }
  function findBody() {
    var cands = document.querySelectorAll('[aria-label="Message body"],[contenteditable="true"][role="textbox"],[contenteditable="true"]');
    for (var c = cands.length - 1; c >= 0; c--) {
      if (cands[c].contentEditable === 'true') return cands[c];
    }
    return null;
  }

  // 1. Wait for either the To field OR the Body editor to appear (indicates compose is loaded)
  var bodyEl = null;
  var toEl = null;
  var TO_SELS = ['input[aria-label="To"]', 'div[aria-label="To"] input', '[role="combobox"][aria-label*="To"]', 'input[type="text"]:not([readonly]):not([disabled])'];
  
  for (var i = 0; i < 20; i++) {
    bodyEl = findBody();
    toEl = findEl(TO_SELS);
    if (bodyEl || toEl) break;
    await delay(1000);
  }
  
  if (!bodyEl && !toEl) {
    return { success: false, error: 'Compose page did not load within 20 seconds' };
  }

  // 2. Populate To field if not already set (fallback)
  var isToPopulated = false;
  if (toEl) {
    var container = toEl.closest('[role="combobox"]') || toEl.parentElement;
    if (container && container.textContent.toLowerCase().indexOf(to.toLowerCase()) > -1) {
      isToPopulated = true;
    }
  }

  if (!isToPopulated && toEl) {
    var _ns = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    toEl.focus(); toEl.click();
    if (toEl.tagName === 'INPUT' && _ns) _ns.call(toEl, to); else toEl.value = to;
    toEl.dispatchEvent(new Event('input',  { bubbles: true }));
    toEl.dispatchEvent(new Event('change', { bubbles: true }));
    toEl.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true }));
    await delay(2500); // wait for email resolution
  }

  // 3. Populate Subject field if not already set (fallback)
  var subEl = findEl(['input[placeholder="Add a subject"]', 'input[aria-label="Subject"]', '[aria-label="Add a subject"]']);
  var isSubjectPopulated = false;
  if (subEl && subEl.value && subEl.value.trim() !== '') {
    isSubjectPopulated = true;
  }

  if (!isSubjectPopulated && subEl) {
    var _ns = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    subEl.focus(); subEl.click();
    if (subEl.tagName === 'INPUT' && _ns) _ns.call(subEl, subject); else subEl.value = subject;
    subEl.dispatchEvent(new Event('input',  { bubbles: true }));
    subEl.dispatchEvent(new Event('change', { bubbles: true }));
  }

  // 4. Find and fill message body
  if (!bodyEl) {
    for (var b = 0; b < 10; b++) {
      bodyEl = findBody();
      if (bodyEl) break;
      await delay(500);
    }
  }
  if (!bodyEl) return { success: false, error: 'Could not find message body editor' };
  
  bodyEl.focus();
  await delay(200);
  document.execCommand('selectAll', false, null);
  if (!document.execCommand('insertHTML', false, html)) {
    bodyEl.innerHTML = html;
    bodyEl.dispatchEvent(new Event('input', { bubbles: true }));
  }
  await delay(1500);

  // 5. Find and click Send button
  var sendBtn = null;
  for (var sa = 0; sa < 10; sa++) {
    // A: Try selectors with case-insensitive matches for "send" or "hantar" (Malay)
    var SEND_SELS = [
      'button[aria-label*="send" i]',
      'button[title*="send" i]',
      'button[aria-label*="hantar" i]',
      'button[title*="hantar" i]',
      '[data-testid*="send" i]',
      '[data-testid*="hantar" i]'
    ];
    for (var j = 0; j < SEND_SELS.length; j++) {
      var btn = document.querySelector(SEND_SELS[j]);
      if (btn && !btn.disabled) { sendBtn = btn; break; }
    }
    // B: Fallback search in all buttons for text content
    if (!sendBtn) {
      var btns = document.querySelectorAll('button, [role="button"]');
      for (var bi = 0; bi < btns.length; bi++) {
        var bEl = btns[bi];
        var aria = (bEl.getAttribute('aria-label') || '').toLowerCase();
        var title = (bEl.getAttribute('title') || '').toLowerCase();
        var txt = (bEl.textContent || '').trim().toLowerCase();
        if (aria.indexOf('send') > -1 || aria.indexOf('hantar') > -1 ||
            title.indexOf('send') > -1 || title.indexOf('hantar') > -1 ||
            txt === 'send' || txt === 'hantar' || txt === 'senden' || txt === 'envoyer') {
          if (!bEl.disabled) { sendBtn = bEl; break; }
        }
      }
    }
    if (sendBtn) break;
    await delay(1000);
  }

  if (sendBtn) {
    sendBtn.click();
    return { success: true, error: null };
  } else {
    // C: Keyboard shortcut fallback (Ctrl + Enter)
    bodyEl.focus();
    await delay(200);
    bodyEl.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Enter',
      code: 'Enter',
      keyCode: 13,
      which: 13,
      ctrlKey: true,
      metaKey: true,
      bubbles: true,
      cancelable: true
    }));
    await delay(1000);
    return { success: true, error: null, warning: 'Used Keyboard Shortcut Fallback' };
  }
}

