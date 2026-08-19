/**
 * auth-dialog.js — Mail Merge MSAL Authentication Dialog Logic
 * =============================================================
 * Extracted from auth-dialog.html to allow a strict Content Security Policy
 * (no 'unsafe-inline' in script-src). All auth, logging, and UI helpers live
 * here; auth-dialog.html contains only structural HTML + CSS.
 *
 * Load order within auth-dialog.html:
 *   1. Office.js (sync <script> in <head>)
 *   2. <script src="auth-dialog.js"> (this file — deferred below)
 *      • Defines helpers, starts logging, then dynamically injects the MSAL
 *        script tag (jsDelivr → alcdn fallback). When MSAL finishes loading,
 *        startAuth() is called to begin the Office.onReady → run() flow.
 */

// ── Early logging setup ───────────────────────────────────────────────────
// Grab the debug log container immediately — it exists in the DOM because
// auth-dialog.js is loaded at the bottom of <body> (after #debugLog).
var logEl = document.getElementById("debugLog");

/**
 * dbg(msg, cls) — appends a timestamped line to the debug log.
 * @param {string} msg - the log message to display
 * @param {string} [cls] - optional CSS class: "log-ok" (green) or "log-err" (red)
 */
function dbg(msg, cls) {
  var d = document.createElement("div");
  d.className = "log-line" + (cls ? " " + cls : "");
  d.textContent = new Date().toISOString().slice(11, 22) + " " + msg;
  logEl.appendChild(d);
  logEl.scrollTop = logEl.scrollHeight;
}

// ── Step progress data ────────────────────────────────────────────────────
var STEPS = [
  "Loading authentication…",
  "Preparing Office runtime…",
  "Initialising secure session…",
  "Checking for existing sign-in…",
  "Redirecting to Microsoft…"
];
var PROGRESS = [12, 28, 52, 72, 92];

/**
 * setStep(n, labelOverride) — advances the UI to step n.
 */
function setStep(n, labelOverride) {
  document.getElementById("stepLabel").textContent = labelOverride || STEPS[n] || "";
  document.getElementById("progressFill").style.width = PROGRESS[n] + "%";
  for (var i = 0; i < 5; i++) {
    var dot = document.getElementById("dot" + i);
    dot.className = "dot" + (i < n ? " done" : i === n ? " active" : "");
  }
}

/**
 * showError(msg) — transitions the UI to an error state.
 */
function showError(msg) {
  var rings = document.querySelectorAll(".pulse-ring");
  rings.forEach(function(r) { r.style.animationPlayState = "paused"; r.style.opacity = "0"; });
  document.getElementById("iconImg").classList.add("error-state");
  document.getElementById("progressFill").style.background = "#C4252A";
  document.getElementById("progressFill").style.boxShadow = "none";
  document.getElementById("progressFill").style.width = "100%";
  document.getElementById("stepLabel").textContent = "Sign-in failed";
  var box = document.getElementById("errorBox");
  box.textContent = msg;   // .textContent (not .innerHTML) — safe from XSS
  box.style.display = "block";
  dbg("ERROR: " + msg, "log-err");
}

// Log the first event — confirms the page loaded and this script ran.
dbg("Page loaded — loading MSAL...");
setStep(0);

// ── MSAL dynamic loader ───────────────────────────────────────────────────
/**
 * loadScript(src, onload, onerror) — dynamically injects a <script> element.
 * Used to load MSAL from CDN without inline event handler attributes, which
 * are blocked by a strict CSP that omits 'unsafe-inline'.
 */
function loadScript(src, onload, onerror) {
  var s = document.createElement("script");
  s.src = src;
  s.crossOrigin = "anonymous";
  s.onload  = onload;
  s.onerror = onerror;
  document.head.appendChild(s);
}

/**
 * loadMsalFallback() — attempts to load MSAL from Microsoft's alcdn after
 * jsDelivr fails (CDN down, network block, etc.).
 */
function loadMsalFallback() {
  dbg("jsDelivr MSAL FAILED — trying alcdn", "log-err");
  loadScript(
    "https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js",
    function() {
      dbg("alcdn fallback loaded OK", "log-ok");
      startAuth();
    },
    function() {
      dbg("BOTH MSAL CDNs failed!", "log-err");
      showError("Could not load the authentication library. Please check your network connection and try again.");
    }
  );
}

// ── Azure App Registration constants ─────────────────────────────────────
const CLIENT_ID    = "d06ae3cf-a7da-4264-b20e-ab8d70c06977";
const TENANT_ID    = "04ca7ab1-691d-45a2-928f-574b3e3300a0";

// Strip cache-bust query string so REDIRECT_URI matches the Azure app registration exactly.
const REDIRECT_URI = window.location.href.split('?')[0];

// Graph API scopes requested at sign-in.
const SCOPES = ["https://graph.microsoft.com/Mail.Send", "https://graph.microsoft.com/User.Read"];

// ── sendToParent ──────────────────────────────────────────────────────────
/**
 * sendToParent(data) — sends the auth result back to the taskpane via two
 * parallel channels: localStorage (for legacy Mac Outlook after a redirect
 * chain) and Office.context.ui.messageParent (for modern Outlook).
 *
 * @param {{ type: "token", token: string } | { type: "error", message: string }} data
 */
function sendToParent(data) {
  dbg("sendToParent → " + data.type);
  var payload = JSON.stringify(data);

  // Channel A: localStorage bridge ─────────────────────────────────────────
  try {
    localStorage.setItem('mm_auth_result', JSON.stringify(
      Object.assign({}, data, { ts: Date.now() })
    ));
    dbg("localStorage fallback written", "log-ok");
  } catch (lsErr) {
    dbg("localStorage write failed: " + lsErr.message);
  }

  // Channel B: messageParent with retry ────────────────────────────────────
  function tryMessageParent(attempts) {
    if (attempts <= 0) {
      dbg("messageParent exhausted — closing dialog via window.close()");
      showError("Sign-in complete — closing…");
      setTimeout(function() {
        try { window.close(); } catch(e) { dbg("window.close() failed: " + e.message); }
      }, 1200);
      return;
    }

    if (Office && Office.context && Office.context.ui && typeof Office.context.ui.messageParent === "function") {
      try {
        Office.context.ui.messageParent(payload);
        dbg("messageParent OK", "log-ok");
      } catch (e) {
        showError("Could not relay sign-in result: " + e.message);
      }
    } else {
      dbg("Office.context.ui not ready — retrying (" + attempts + ")...");
      setTimeout(function() { tryMessageParent(attempts - 1); }, 500);
    }
  }

  tryMessageParent(60);
}

// ── MSAL instance ─────────────────────────────────────────────────────────
// Office.initialize is a legacy hook — no-op prevents errors on older runtimes.
Office.initialize = function() {};

var msalInstance = null;

/**
 * createMsalInstance() — initialises the MSAL PublicClientApplication.
 * Returns true on success, false on failure.
 */
function createMsalInstance() {
  try {
    dbg("Creating MSAL instance...");
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId:    CLIENT_ID,
        authority:   "https://login.microsoftonline.com/" + TENANT_ID,
        redirectUri: REDIRECT_URI,
        navigateToLoginRequestUrl: true
      },
      cache: {
        cacheLocation:       "sessionStorage",
        storeAuthStateInCookie: false
      }
    });
    dbg("MSAL instance created OK", "log-ok");
    return true;
  } catch (e) {
    dbg("MSAL instance creation FAILED: " + e.message, "log-err");
    showError("Auth library init failed: " + e.message);
    return false;
  }
}

// ── Main auth flow ────────────────────────────────────────────────────────
/**
 * run() — the main async auth flow.
 * Steps: create MSAL → initialize → handleRedirectPromise → silent → popup → redirect.
 */
async function run() {
  dbg("run() started");
  if (!createMsalInstance()) return;

  try {
    setStep(2, "Initialising secure session…");
    dbg("calling initialize()...");
    await msalInstance.initialize();
    dbg("initialize() done", "log-ok");

    setStep(3, "Checking for existing sign-in…");
    dbg("calling handleRedirectPromise()...");

    var redirectResponse = await Promise.race([
      msalInstance.handleRedirectPromise(),
      new Promise(function(resolve) {
        setTimeout(function() {
          dbg("handleRedirectPromise timeout — proceeding");
          resolve(null);
        }, 6000);
      })
    ]);

    dbg("handleRedirectPromise → " + (redirectResponse ? "GOT RESPONSE" : "null"), "log-ok");

    if (redirectResponse && redirectResponse.accessToken) {
      setStep(4, "Signed in — returning to Mail Merge…");
      dbg("Sending token to parent", "log-ok");
      sendToParent({ type: "token", token: redirectResponse.accessToken });
      return;
    }

    // Silent token acquisition ──────────────────────────────────────────────
    var accounts = msalInstance.getAllAccounts();
    dbg("accounts in cache: " + accounts.length);

    if (accounts.length > 0) {
      try {
        setStep(3, "Restoring your session…");
        dbg("trying silent token...");
        var silent = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
        dbg("silent token OK", "log-ok");
        setStep(4, "Signed in — returning to Mail Merge…");
        sendToParent({ type: "token", token: silent.accessToken });
        return;
      } catch (se) {
        dbg("silent failed: " + (se.errorCode || se.message));
      }
    }

    // Interactive sign-in: popup first ─────────────────────────────────────
    setStep(4, "Opening Microsoft sign-in…");
    dbg("trying acquireTokenPopup...");
    try {
      var popupResponse = await msalInstance.acquireTokenPopup({ scopes: SCOPES });
      dbg("popup token OK", "log-ok");
      setStep(5, "Signed in — returning to Mail Merge…");
      sendToParent({ type: "token", token: popupResponse.accessToken });
      return;
    } catch (popupErr) {
      var pCode = popupErr.errorCode || "";
      dbg("popup error: " + pCode);

      if (pCode === "popup_window_error" || pCode === "empty_window_error") {
        dbg("Popup blocked — falling back to redirect...");
        await msalInstance.acquireTokenRedirect({ scopes: SCOPES });
        dbg("acquireTokenRedirect called (page navigating away)");
      } else {
        throw popupErr;
      }
    }

  } catch (e) {
    var code = e.errorCode || "";
    var msg  = e.errorMessage || e.message || "Unknown error";
    dbg("CATCH: " + code + " — " + msg, "log-err");

    if (code === "interaction_in_progress") {
      sessionStorage.clear();
      var m = "Sign-in interrupted. Please close this window and try again.";
      showError(m);
      setTimeout(function() { sendToParent({ type: "error", message: m }); }, 3000);
      return;
    }

    showError(msg);
    setTimeout(function() { sendToParent({ type: "error", message: msg }); }, 3000);
  }
}

// ── Entry point ────────────────────────────────────────────────────────────
/**
 * startAuth() — registers Office.onReady then calls run().
 * Also sets a 5-second safety-net timeout in case Office.onReady never fires.
 */
function startAuth() {
  setStep(1, "Preparing Office runtime…");
  dbg("registering Office.onReady...");

  Office.onReady(function() {
    dbg("Office.onReady fired", "log-ok");
    run();
  });

  setTimeout(function() {
    if (!msalInstance) {
      dbg("Office.onReady timeout — running directly");
      run();
    }
  }, 5000);
}

// ── Load MSAL dynamically (avoids inline onerror attribute, removing the need
//    for 'unsafe-inline' in CSP script-src). jsDelivr primary, alcdn fallback.
loadScript(
  "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js",
  function() {
    dbg("MSAL loaded OK from jsDelivr", "log-ok");
    startAuth();
  },
  loadMsalFallback
);
