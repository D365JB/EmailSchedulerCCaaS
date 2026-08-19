// Delegated authentication via MSAL (SPA + PKCE). No client secret.
// Strategy: reuse a cached account -> ssoSilent (using the agent's email as a hint) -> interactive popup.

var _msalApp = null;

function isConfigured() {
  return !!(CONFIG && CONFIG.clientId && CONFIG.clientId.indexOf('REPLACE') === -1 &&
            CONFIG.tenantId && CONFIG.tenantId.indexOf('REPLACE') === -1);
}

function computeRedirectUri() {
  if (CONFIG.redirectUri) return CONFIG.redirectUri;
  // Dynamics serves web resources through a cache-versioned path like
  // /{638...}/webresources/jmb_/index.html. That {token} changes on every publish and
  // can't be pre-registered in Entra, so rebuild a stable canonical URI that matches the
  // registered redirect (…/WebResources/<prefix>/blank.html), preserving that exact casing.
  var p = location.pathname;
  var idx = p.toLowerCase().indexOf('/webresources/');
  if (idx !== -1) {
    var after = p.substring(idx + '/webresources/'.length); // e.g. jmb_/index.html
    var folder = after.substring(0, after.lastIndexOf('/') + 1); // e.g. jmb_/
    return location.origin + '/WebResources/' + folder + 'blank.html';
  }
  var dir = p.substring(0, p.lastIndexOf('/') + 1);
  return location.origin + dir + 'blank.html';
}

function initAuth() {
  if (_msalApp) return _msalApp;
  _msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: CONFIG.clientId,
      authority: 'https://login.microsoftonline.com/' + CONFIG.tenantId,
      redirectUri: computeRedirectUri()
    },
    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
  });
  return _msalApp;
}

// Returns the model-driven app's Xrm object when the web resource is embedded in Dynamics.
function getXrm() {
  return (window.parent && window.parent.Xrm) ? window.parent.Xrm : (window.Xrm || null);
}

// When embedded in a model-driven app, read the signed-in agent's email for a seamless SSO hint.
async function getLoginHint() {
  try {
    var xrm = getXrm();
    if (!xrm || !xrm.Utility) return null;
    var uid = xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, '');
    var rec = await xrm.WebApi.retrieveRecord('systemuser', uid, '?$select=internalemailaddress');
    return (rec && rec.internalemailaddress) ? rec.internalemailaddress : null;
  } catch (e) {
    return null;
  }
}

async function ensureSignedInSilent() {
  var accounts = _msalApp.getAllAccounts();
  if (accounts && accounts.length) {
    _msalApp.setActiveAccount(accounts[0]);
    await getToken(CONFIG.scopes);
    return;
  }
  var hint = await getLoginHint();
  if (!hint) throw new Error('interaction_required');
  var res = await _msalApp.ssoSilent({ scopes: CONFIG.scopes, loginHint: hint });
  _msalApp.setActiveAccount(res.account);
}

async function interactiveSignIn() {
  var hint = await getLoginHint();
  var req = { scopes: CONFIG.scopes };
  if (hint) req.loginHint = hint;
  var res = await _msalApp.loginPopup(req);
  _msalApp.setActiveAccount(res.account);
}

async function getToken(scopes) {
  var account = _msalApp.getActiveAccount();
  if (!account) {
    var accts = _msalApp.getAllAccounts();
    if (accts && accts.length) { account = accts[0]; _msalApp.setActiveAccount(account); }
  }
  var request = { scopes: scopes || CONFIG.scopes, account: account };
  try {
    var res = await _msalApp.acquireTokenSilent(request);
    return res.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      var r2 = await _msalApp.acquireTokenPopup(request);
      return r2.accessToken;
    }
    throw e;
  }
}
