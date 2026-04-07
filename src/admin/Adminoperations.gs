/**
 * NO~BULL ADMIN — OPERATIONS
 * Health monitoring, audit log viewer, maintenance, communications.
 */

// ─────────────────────────────────────────────────────────────────────────────
// HEALTH MONITORING
// ─────────────────────────────────────────────────────────────────────────────

// Sheet names must exactly match SHEETS constants in Config.gs
var REQUIRED_SHEETS = [
  'Settings','Clients','Suppliers','Invoices','InvoiceLines',
  'Bills','BillLines','Transactions',
  'BankAccounts','BankTransactions',
  'ChartOfAccounts','Users','AuditLog',
  'VATReturns','InvoiceFiles','BillFiles',
  'CreditNotes','PurchaseOrders'
];

function checkClientHealth(params) {
  try {
    var sheetId = params && params.sheetId;
    if (!sheetId) return { success: false, message: 'No sheetId' };

    var ss      = SpreadsheetApp.openById(sheetId);
    var sheets  = ss.getSheets().map(function(s){ return s.getName(); });
    var missing = REQUIRED_SHEETS.filter(function(s){ return sheets.indexOf(s) === -1; });
    var status  = missing.length === 0 ? 'Healthy' : missing.length <= 2 ? 'Warning' : 'Error';

    // Check settings
    var settingsSheet = ss.getSheetByName('Settings');
    var hasSettings   = settingsSheet && settingsSheet.getLastRow() >= 2;

    // Check users
    var usersSheet = ss.getSheetByName('Users');
    var userCount  = usersSheet && usersSheet.getLastRow() > 1 ? usersSheet.getLastRow() - 1 : 0;

    // Check audit log size
    var auditSheet = ss.getSheetByName('AuditLog');
    var auditRows  = auditSheet ? auditSheet.getLastRow() - 1 : 0;

    return {
      success:     true,
      status:      status,
      sheetCount:  sheets.length,
      missingSheets: missing,
      hasSettings: hasSettings,
      userCount:   userCount,
      auditRows:   auditRows,
      checkedAt:   new Date().toISOString()
    };
  } catch(e) {
    return { success: false, status: 'Error', message: e.toString() };
  }
}

function checkAllClientsHealth(params) {
  try {
    var r       = getAllRegistryClients(params || {});
    var clients = r.clients || [];
    var results = clients.map(function(c) {
      if (!c.sheetId || c.status === 'Cancelled') {
        return { registryId: c.registryId, companyName: c.companyName, status: 'Skipped' };
      }
      try {
        var health = checkClientHealth({ sheetId: c.sheetId });
        return {
          registryId:    c.registryId,
          companyName:   c.companyName,
          sheetId:       c.sheetId,
          clientStatus:  c.status,
          healthStatus:  health.status,
          missingSheets: health.missingSheets || [],
          userCount:     health.userCount || 0,
          checkedAt:     health.checkedAt || ''
        };
      } catch(e) {
        return { registryId: c.registryId, companyName: c.companyName, healthStatus: 'Error', error: e.toString() };
      }
    });
    return { success: true, results: results };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// AUDIT LOG
// ─────────────────────────────────────────────────────────────────────────────

function getClientAuditLog(params) {
  try {
    var sheetId = params && params.sheetId;
    if (!sheetId) return { success: false, message: 'No sheetId' };
    var ss      = SpreadsheetApp.openById(sheetId);
    var sheet   = ss.getSheetByName('AuditLog');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, entries: [] };

    var limit    = params.limit ? parseInt(params.limit) : 50;
    var lastRow  = sheet.getLastRow();
    var startRow = Math.max(2, lastRow - limit + 1);
    var data     = sheet.getRange(startRow, 1, lastRow - startRow + 1, 7).getValues();

    var entries = [];
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      if (!row[0]) continue;
      entries.push({
        id:        row[0].toString(),
        timestamp: safeSerializeDate(row[1]),
        action:    row[2].toString(),
        entity:    row[3].toString(),
        entityId:  row[4].toString(),
        user:      row[5].toString(),
        detail:    row[6].toString()
      });
    }
    return { success: true, entries: entries, total: lastRow - 1 };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// COMMUNICATIONS
// ─────────────────────────────────────────────────────────────────────────────

function sendTrialExpiryReminders(params) {
  try {
    var r       = getAllRegistryClients(params || {});
    var clients = (r.clients || []).filter(function(c){ return c.status === 'Trial' && c.createdDate; });
    var sent    = 0;
    var skipped = 0;

    clients.forEach(function(c) {
      var daysElapsed  = Math.floor((new Date() - new Date(c.createdDate)) / 86400000);
      var daysLeft     = 14 - daysElapsed;

      // Send at day 11 (3 days left) and day 13 (1 day left)
      if (daysLeft === 3 || daysLeft === 1) {
        _sendTrialExpiryEmail(c, daysLeft);
        sent++;
      } else {
        skipped++;
      }
    });

    return { success: true, sent: sent, skipped: skipped };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function _sendTrialExpiryEmail(client, daysLeft) {
  var email = client.contactEmail;
  if (!email) return;

  var subject = daysLeft === 1
    ? 'Your no~bull books trial ends tomorrow'
    : 'Your no~bull books trial ends in ' + daysLeft + ' days';

  var appUrl = client.appLink || '';

  var body =
    'Hi,\n\n' +
    'Your no~bull books trial for ' + (client.companyName||'your account') + ' ' +
    (daysLeft === 1 ? 'ends tomorrow.' : 'ends in ' + daysLeft + ' days.') + '\n\n' +
    'To continue using no~bull books, please upgrade to a paid plan.\n\n' +
    'Access your account: ' + appUrl + '\n\n' +
    'Questions? Reply to this email.\n\n' +
    'Best regards,\n' +
    'Edward Jenkins\n' +
    'no~bull consulting\n' +
    'edward@nobull.consulting\n';

  var htmlBody =
    '<div style="font-family:-apple-system,sans-serif;max-width:540px;margin:0 auto">' +
    '<div style="background:#14213D;padding:20px 28px;border-radius:8px 8px 0 0">' +
      '<p style="color:#fff;font-size:18px;margin:0">🐂 <strong>no~bull</strong> <span style="color:#14A8AE">books</span></p>' +
    '</div>' +
    '<div style="background:#fff;border:1px solid #e2e8f0;border-top:none;padding:28px;border-radius:0 0 8px 8px">' +
      '<div style="background:#fef3c7;border:1px solid #f59e0b;border-radius:8px;padding:16px;margin-bottom:20px">' +
        '<p style="margin:0;color:#92400e;font-weight:600">⏰ Your trial ' + (daysLeft === 1 ? 'ends tomorrow' : 'ends in ' + daysLeft + ' days') + '</p>' +
      '</div>' +
      '<p style="color:#475569">Your no~bull books trial for <strong>' + (client.companyName||'your account') + '</strong> is coming to an end.</p>' +
      '<p style="color:#475569">To keep access to your invoices, bills, bank reconciliation and MTD VAT submissions, please upgrade to a paid plan.</p>' +
      '<a href="' + appUrl + '" style="display:inline-block;background:#0D7377;color:#fff;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:600;margin:16px 0">Continue with no~bull books →</a>' +
      '<hr style="border:none;border-top:1px solid #e2e8f0;margin:20px 0">' +
      '<p style="color:#94a3b8;font-size:12px">Questions? Reply to this email or contact <a href="mailto:edward@nobull.consulting" style="color:#0D7377">edward@nobull.consulting</a></p>' +
    '</div>' +
    '</div>';

  try {
    MailApp.sendEmail({ to: email, subject: subject, body: body, htmlBody: htmlBody });
    Logger.log('Trial expiry email sent to: ' + email + ' (' + daysLeft + ' days left)');
  } catch(e) {
    Logger.log('Failed to send trial expiry email to ' + email + ': ' + e);
  }
}

function sendBroadcast(params) {
  try {
    var subject  = params.subject;
    var bodyText = params.body;
    var htmlBody = params.htmlBody || bodyText;
    var filter   = params.filter || 'Active'; // 'Active', 'All', 'Trial'

    if (!subject || !bodyText) return { success: false, message: 'Subject and body required' };

    var r       = getAllRegistryClients(params || {});
    var clients = (r.clients || []).filter(function(c) {
      if (filter === 'All') return c.contactEmail;
      return c.status === filter && c.contactEmail;
    });

    var sent = 0;
    clients.forEach(function(c) {
      try {
        MailApp.sendEmail({ to: c.contactEmail, subject: subject, body: bodyText, htmlBody: htmlBody });
        sent++;
        Utilities.sleep(500); // Rate limiting
      } catch(e) {
        Logger.log('Broadcast failed for ' + c.contactEmail + ': ' + e);
      }
    });

    Logger.log('Broadcast sent to ' + sent + ' clients');
    return { success: true, sent: sent, total: clients.length };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// MAINTENANCE
// ─────────────────────────────────────────────────────────────────────────────

function runClientMaintenance(params) {
  try {
    var sheetId = params && params.sheetId;
    if (!sheetId) return { success: false, message: 'No sheetId' };

    var ss      = SpreadsheetApp.openById(sheetId);
    var actions = [];

    // Check and create missing sheets
    var AUDIT_HEADERS = ['AuditId','Timestamp','Action','Entity','EntityId','User','Detail'];
    REQUIRED_SHEETS.forEach(function(name) {
      if (!ss.getSheetByName(name)) {
        var newSheet = ss.insertSheet(name);
        if (name === 'AuditLog') {
          newSheet.getRange(1, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]);
        }
        actions.push('Created missing sheet: ' + name);
      } else if (name === 'AuditLog') {
        // Fix existing AuditLog with no headers
        var auditTab = ss.getSheetByName('AuditLog');
        var firstCell = auditTab.getLastRow() >= 1 ? auditTab.getRange(1,1).getValue().toString().trim() : '';
        if (!firstCell || firstCell !== 'AuditId') {
          if (auditTab.getLastRow() === 0) {
            auditTab.getRange(1, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]);
          } else {
            auditTab.insertRowBefore(1);
            auditTab.getRange(1, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]);
          }
          actions.push('Added headers to AuditLog');
        }
      }
    });

    // Trim oversized audit log (keep last 1000 rows)
    var auditSheet = ss.getSheetByName('AuditLog');
    if (auditSheet && auditSheet.getLastRow() > 1001) {
      var excess = auditSheet.getLastRow() - 1001;
      auditSheet.deleteRows(2, excess);
      actions.push('Trimmed audit log: removed ' + excess + ' old entries');
    }

    Logger.log('Maintenance complete for ' + sheetId + ': ' + actions.join(', '));
    return { success: true, actions: actions, message: actions.length ? actions.join('; ') : 'No action needed' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SCHEDULED TRIGGERS — run from Admin Console
// ─────────────────────────────────────────────────────────────────────────────

function installAdminTriggers() {
  // Remove existing
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });

  // Daily trial expiry check at 8am London
  ScriptApp.newTrigger('_dailyTrialCheck')
    .timeBased().atHour(8).everyDays(1).inTimezone('Europe/London').create();

  Logger.log('Admin triggers installed');
  return { success: true, message: 'Daily trial check trigger installed' };
}

function _dailyTrialCheck() {
  Logger.log('Daily trial check: ' + new Date().toISOString());
  sendTrialExpiryReminders({});
}