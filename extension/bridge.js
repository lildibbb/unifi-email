(function () {
  // bridge.js — forwards messages between the web page and the extension background

  window.addEventListener('message', function (event) {
    if (event.source !== window) return;
    if (!event.data || event.data.source !== 'email-blast') return;

    var msg   = event.data;
    var msgId = msg._msgId;

    chrome.runtime.sendMessage(msg, function (response) {
      // Always read lastError to prevent Chrome's "unchecked" warning
      var err = chrome.runtime.lastError;
      window.postMessage({
        source:   'email-blast-extension',
        _msgId:   msgId,
        error:    err ? err.message : undefined,
        response: err ? undefined : (response || {})
      }, '*');
    });
  });

  // Push unsolicited messages from background → page
  chrome.runtime.onMessage.addListener(function (msg) {
    window.postMessage({ source: 'email-blast-extension', unsolicited: true, data: msg }, '*');
  });
})();