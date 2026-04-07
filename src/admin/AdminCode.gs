/**
 * NO~BULL BOOKS — ADMIN CONSOLE
 * ─────────────────────────────────────────────────────────────────────────────
 * Separate GAS deployment for registry management, client administration,
 * maintenance, and monitoring. Completely isolated from the client-facing Hub.
 *
 * Access: Owner only (edward@nobull.consulting)
 * Auth:   Request must include X-Admin-Secret header matching ADMIN_SECRET
 *         Script Property, OR be called directly from the GAS editor.
 *
 * API: POST /exec with JSON body { action, secret, ...params }
 * UI:  GET  /exec (admin dashboard — owner only)
 */

var ADMIN_VERSION = '1.0.0';

// ─────────────────────────────────────────────────────────────────────────────
// ENTRY POINTS
// ─────────────────────────────────────────────────────────────────────────────

function doGet(e) {
  // Verify caller is the owner
  var email = Session.getActiveUser().getEmail();
  var ownerEmail = 'edward@nobull.consulting';
  if (email.toLowerCase() !== ownerEmail.toLowerCase()) {
    return HtmlService.createHtmlOutput('<h1>Access denied</h1>')
      .setTitle('no~bull Admin');
  }

  // Handle registry ping from Hub or SetupService
  if (e.parameter.action === 'pingRegistry') {
    try {
      pingRegistry(e.parameter.sheetId, {
        email:       e.parameter.email       || '',
        companyName: e.parameter.companyName || '',
        version:     e.parameter.version     || ''
      });
    } catch(err) {
      Logger.log('pingRegistry error: ' + err);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Admin dashboard UI
  return HtmlService.createHtmlOutputFromFile('AdminDashboard')
    .setTitle('no~bull Admin Console')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents || '{}');
    var secret = PropertiesService.getScriptProperties().getProperty('ADMIN_SECRET') || '';

    // Validate shared secret
    if (!secret || body.secret !== secret) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, error: 'Unauthorised'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var result = _adminRoute(body.action, body);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    Logger.log('doPost error: ' + e.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false, error: e.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// ROUTER
// ─────────────────────────────────────────────────────────────────────────────

function _adminRoute(action, params) {
  switch(action) {
    // Registry
    case 'getAllRegistryClients':    return getAllRegistryClients(params);
    case 'registerClient':          return registerClient(params, params);
    case 'updateRegistryClient':    return updateRegistryClient(params.registryId, params, params);
    case 'deactivateRegistryClient':return deactivateRegistryClient(params.registryId, params.reason, params);
    case 'pingRegistry':            return pingRegistry(params.sheetId, params);
    case 'getInstanceMeta':         return getInstanceMeta(params);

    // Maintenance
    case 'runMaintenance':          return runMaintenance();
    case 'runManualBackup':         return runManualBackup();
    case 'verifyIntegrity':         return verifyIntegrity(params);
    case 'diagnoseSheets':          return diagnoseSheets(params);
    case 'getAdminStats':           return getAdminStats(params);

    // GDPR
    case 'eraseClient':             return eraseClient(params.clientId, params.retainFinancial);
    case 'exportClientData':        return exportClientData(params.clientId);

    default:
      return { success: false, error: 'Unknown action: ' + action };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// ADMIN DASHBOARD HTML
// ─────────────────────────────────────────────────────────────────────────────



// ─────────────────────────────────────────────────────────────────────────────
// SETUP HELPERS — run once from editor
// ─────────────────────────────────────────────────────────────────────────────


// ─────────────────────────────────────────────────────────────────────────────
// FRONTEND BRIDGE — called via google.script.run from AdminDashboard.html
// ─────────────────────────────────────────────────────────────────────────────

function handleAdminCall(action, params) {
  try {
    // Verify caller is owner
    var email = Session.getActiveUser().getEmail();
    if (email.toLowerCase() !== 'edward@nobull.consulting') {
      return { success: false, error: 'Unauthorised' };
    }
    return _adminRoute(action, params || {});
  } catch(e) {
    Logger.log('handleAdminCall error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function _setupAdminSecret() {
  var secret = Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      new Date().getTime().toString() + Math.random().toString())
  ).substring(0, 32);
  PropertiesService.getScriptProperties().setProperty('ADMIN_SECRET', secret);
  Logger.log('ADMIN_SECRET set to: ' + secret);
  Logger.log('Copy this to Hub Script Properties as ADMIN_SECRET');
}

function _checkAdminProps() {
  var props = PropertiesService.getScriptProperties().getProperties();
  Logger.log('ADMIN_SECRET: '    + (props['ADMIN_SECRET']    ? 'SET' : 'NOT SET'));
  Logger.log('REGISTRY_SHEET_ID: ' + (props['REGISTRY_SHEET_ID'] || 'NOT SET'));
}

function getAdminStats(params) {
  try {
    var r = getAllRegistryClients(params || {});
    var clients = r.clients || [];
    return {
      success:    true,
      total:      clients.length,
      active:     clients.filter(function(c){ return c.status === 'Active'; }).length,
      trial:      clients.filter(function(c){ return c.status === 'Trial'; }).length,
      suspended:  clients.filter(function(c){ return c.status === 'Suspended'; }).length
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getInstanceMeta(params) {
  try {
    var r = getAllRegistryClients(params || {});
    var clients = r.clients || [];
    var match = clients.filter(function(c){
      return c.sheetId === (params && params._sheetId);
    })[0];
    return { success: true, meta: match || {} };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}
