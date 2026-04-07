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
  return HtmlService.createHtmlOutput(_adminDashboardHtml())
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

function _adminDashboardHtml() {
  var clients = [];
  try {
    var r = getAllRegistryClients({});
    clients = r.clients || [];
  } catch(e) { Logger.log('Dashboard load error: ' + e); }

  var active    = clients.filter(function(c){ return c.status === 'Active'; }).length;
  var trial     = clients.filter(function(c){ return c.status === 'Trial'; }).length;
  var suspended = clients.filter(function(c){ return c.status === 'Suspended'; }).length;

  var rows = clients.map(function(c) {
    return '<tr>' +
      '<td>' + (c.companyName||'—') + '</td>' +
      '<td>' + (c.contactEmail||'—') + '</td>' +
      '<td>' + (c.plan||'—') + '</td>' +
      '<td><span style="padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:' +
        (c.status==='Active'?'#dcfce7;color:#166534':c.status==='Trial'?'#dbeafe;color:#1e40af':'#fee2e2;color:#991b1b') +
        '">' + (c.status||'—') + '</span></td>' +
      '<td>' + (c.lastSeen ? c.lastSeen.substring(0,10) : '—') + '</td>' +
      '<td><a href="' + (c.appLink||'#') + '" target="_blank" style="color:#0D7377">Open</a></td>' +
    '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<title>no~bull Admin Console</title>' +
    '<style>' +
    'body{font-family:-apple-system,sans-serif;background:#f8fafc;margin:0;padding:0}' +
    '.header{background:#14213D;color:#fff;padding:16px 32px;display:flex;align-items:center;gap:12px}' +
    '.header h1{font-size:18px;margin:0}.header span{color:#14A8AE}' +
    '.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;padding:24px 32px}' +
    '.stat{background:#fff;border-radius:8px;padding:20px;border:1px solid #e2e8f0}' +
    '.stat-val{font-size:32px;font-weight:700;color:#14213D}.stat-lbl{font-size:12px;color:#64748b;margin-top:4px}' +
    '.section{padding:0 32px 32px}' +
    'table{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;border:1px solid #e2e8f0}' +
    'th{background:#14213D;color:#fff;padding:10px 14px;text-align:left;font-size:12px}' +
    'td{padding:10px 14px;border-bottom:1px solid #f1f5f9;font-size:13px}' +
    'tr:hover td{background:#f8fafc}' +
    '</style></head><body>' +
    '<div class="header"><h1>🐂 no~bull <span>Admin Console</span></h1><span style="margin-left:auto;font-size:12px;opacity:0.6">v' + ADMIN_VERSION + '</span></div>' +
    '<div class="stats">' +
      '<div class="stat"><div class="stat-val">' + clients.length + '</div><div class="stat-lbl">Total clients</div></div>' +
      '<div class="stat"><div class="stat-val" style="color:#166534">' + active + '</div><div class="stat-lbl">Active</div></div>' +
      '<div class="stat"><div class="stat-val" style="color:#1e40af">' + trial + '</div><div class="stat-lbl">Trial</div></div>' +
      '<div class="stat"><div class="stat-val" style="color:#991b1b">' + suspended + '</div><div class="stat-lbl">Suspended</div></div>' +
    '</div>' +
    '<div class="section">' +
      '<table><thead><tr><th>Company</th><th>Email</th><th>Plan</th><th>Status</th><th>Last Seen</th><th></th></tr></thead>' +
      '<tbody>' + (rows || '<tr><td colspan="6" style="text-align:center;color:#94a3b8">No clients registered</td></tr>') + '</tbody>' +
      '</table>' +
    '</div>' +
    '</body></html>';
}

// ─────────────────────────────────────────────────────────────────────────────
// SETUP HELPERS — run once from editor
// ─────────────────────────────────────────────────────────────────────────────

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
