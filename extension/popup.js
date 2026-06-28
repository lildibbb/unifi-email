// popup.js - Extension toolbar popup logic
var APP_URL = 'http://localhost:5000/';
var _pollTimer = null;

function refreshUI(state) {
  var dot  = document.getElementById('dot');
  var txt  = document.getElementById('status-text');
  var btnC = document.getElementById('btn-connect');
  var btnD = document.getElementById('btn-disconnect');

  dot.className = 'dot';
  btnC.disabled = false;

  if (state.status === 'connected') {
    clearPoll();
    dot.classList.add('ok');
    txt.textContent = 'Connected' + (state.user ? ' \u2014 ' + state.user : '');
    txt.className   = 'ok';
    btnC.style.display = 'none';
    btnD.style.display = 'flex';
  } else if (state.status === 'connecting') {
    dot.classList.add('busy');
    txt.textContent = 'Signing in to Outlook...';
    txt.className   = '';
    btnC.disabled   = true;
    btnC.innerHTML  = '<span class="spin"></span> Connecting...';
    btnC.style.display = 'flex';
    btnD.style.display = 'none';
  } else {
    clearPoll();
    txt.textContent = state.error ? 'Error: ' + state.error : 'Disconnected';
    txt.className   = '';
    btnC.textContent   = 'Connect Outlook';
    btnC.style.display = 'flex';
    btnD.style.display = 'none';
  }
}

function loadState() {
  chrome.runtime.sendMessage({ action: 'status' }, function (resp) {
    if (chrome.runtime.lastError) return;
    refreshUI(resp || { status: 'disconnected' });
  });
}

// Poll storage every 2 s while connecting — survives service-worker restarts
function startPoll() {
  clearPoll();
  _pollTimer = setInterval(loadState, 2000);
}

function clearPoll() {
  if (_pollTimer) { clearInterval(_pollTimer); _pollTimer = null; }
}

function doConnect() {
  refreshUI({ status: 'connecting' });
  chrome.runtime.sendMessage({ action: 'connect' }, function () {
    // handleConnect() returns almost immediately now (no blocking poll).
    // Start a 2-second polling loop so the popup reflects storage changes
    // as soon as onUpdated fires and marks the session connected.
    startPoll();
  });
}

function doDisconnect() {
  chrome.runtime.sendMessage({ action: 'disconnect' }, function () {
    loadState();
  });
}

function openApp() {
  chrome.tabs.query({ url: ['http://localhost:5000/*', 'http://127.0.0.1:5000/*'] }, function (tabs) {
    if (tabs.length > 0) {
      chrome.tabs.update(tabs[0].id, { active: true });
    } else {
      chrome.tabs.create({ url: 'http://localhost:5000/' });
    }
    window.close();
  });
}

// Background pushes 'status' action when state changes — refresh immediately
chrome.runtime.onMessage.addListener(function (msg) {
  if (msg && msg.action === 'status') {
    loadState();
  }
});

loadState();