/**
 * NO~BULL BOOKS — CLIENT REGISTRY
 *
 * The registry is a single Google Sheet owned by the deployer (edward@nobull.consulting).
 * Its Spreadsheet ID is stored in Script Properties as REGISTRY_SHEET_ID.
 *
 * Each row represents one client instance:
 *   Col A  registryId       — unique registry record ID (REG_xxx)
 *   Col B  clientRef        — short human reference, e.g. "smith-consulting"
 *   Col C  companyName      — client's business name
 *   Col D  contactName      — primary contact person
 *   Col E  contactEmail     — primary contact email
 *   Col F  contactPhone     — contact phone
 *   Col G  sheetId          — Google Spreadsheet ID for this client's data
 *   Col H  sheetUrl         — full URL to the spreadsheet
 *   Col I  deployUrl        — no~bull books GAS /exec URL for this client
 *   Col J  plan             — e.g. "Solo", "Pro", "Accountant"
 *   Col K  status           — Active | Suspended | Cancelled | Trial
 *   Col L  createdDate      — ISO date
 *   Col M  lastSeen         — ISO datetime (updated on each login)
 *   Col N  lastSeenBy       — email of last user
 *   Col O  invoiceCount     — denormalised count (updated on ping)
 *   Col P  clientCount      — denormalised count
 *   Col Q  billCount        — denormalised count
 *   Col R  version          — app version string
 *   Col S  notes            — internal notes (billing, onboarding etc.)
 *   Col T  vatRegistered    — TRUE/FALSE
 *   Col U  vatNumber        — VAT reg number
 *   Col V  financialYearEnd — e.g. "31 March"
 *   Col W  country          — e.g. "UK"
 *   Col X  accountant       — accountant name / firm
 *   Col Y  accountantEmail  — accountant email
 *
 * SETUP: Run initRegistry() once from the Apps Script editor.
 * ─────────────────────────────────────────────────────────────────────────────
 */

var REGISTRY_VERSION = '1.0';
var REGISTRY_COLS = {
  REGISTRY_ID:    1,
  CLIENT_REF:     2,
  COMPANY_NAME:   3,
  CONTACT_NAME:   4,
  CONTACT_EMAIL:  5,
  CONTACT_PHONE:  6,
  SHEET_ID:       7,
  SHEET_URL:      8,
  DEPLOY_URL:     9,
  PLAN:           10,
  STATUS:         11,
  CREATED_DATE:   12,
  LAST_SEEN:      13,
  LAST_SEEN_BY:   14,
  INVOICE_COUNT:  15,
  CLIENT_COUNT:   16,
  BILL_COUNT:     17,
  VERSION:        18,
  NOTES:          19,
  VAT_REGISTERED: 20,
  VAT_NUMBER:     21,
  FY_END:         22,
  COUNTRY:        23,
  ACCOUNTANT:     24,
  ACCOUNTANT_EMAIL: 25
};

var REGISTRY_HEADERS = [
  'RegistryId', 'ClientRef', 'CompanyName', 'ContactName', 'ContactEmail',
  'ContactPhone', 'SheetId', 'SheetUrl', 'DeployUrl', 'Plan',
  'Status', 'CreatedDate', 'LastSeen', 'LastSeenBy',
  'InvoiceCount', 'ClientCount', 'BillCount', 'Version', 'Notes',
  'VATRegistered', 'VATNumber', 'FinancialYearEnd', 'Country',
  'Accountant', 'AccountantEmail'
];

// ─────────────────────────────────────────────────────────────────────────────
// REGISTRY ACCESS
// ─────────────────────────────────────────────────────────────────────────────

function _getRegistrySheet() {
  var id = PropertiesService.getScriptProperties().getProperty('REGISTRY_SHEET_ID');
  if (!id) throw new Error('REGISTRY_SHEET_ID not set. Run setRegistrySheetId("YOUR_SHEET_ID") first.');
  var ss    = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName('Registry');
  if (!sheet) throw new Error('Registry sheet not found in spreadsheet ' + id + '. Run initRegistry() first.');
  return sheet;
}

/**
 * initRegistry()
 * Run once from the Apps Script editor to create or repair the registry sheet.
 * Creates a new Google Sheet if REGISTRY_SHEET_ID is not set yet.
 */
function initRegistry() {
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty('REGISTRY_SHEET_ID');
  var ss;

  if (!id) {
    // Create a new sheet owned by the deployer
    ss = SpreadsheetApp.create('no~bull books — Client Registry');
    id = ss.getId();
    props.setProperty('REGISTRY_SHEET_ID', id);
    Logger.log('Created new registry sheet: ' + ss.getUrl());
  } else {
    ss = SpreadsheetApp.openById(id);
  }

  var sheet = ss.getSheetByName('Registry');
  if (!sheet) {
    sheet = ss.insertSheet('Registry');
    // Delete default Sheet1 if it exists
    var defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  }

  // Write or repair headers
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== 'RegistryId') {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, REGISTRY_HEADERS.length).setValues([REGISTRY_HEADERS]);
    sheet.getRange(1, 1, 1, REGISTRY_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#0f172a')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(3, 200);  // CompanyName
    sheet.setColumnWidth(8, 300);  // SheetUrl
    sheet.setColumnWidth(9, 300);  // DeployUrl
  }

  // Add a Dashboard sheet if not present
  var dashSheet = ss.getSheetByName('Dashboard');
  if (!dashSheet) {
    dashSheet = ss.insertSheet('Dashboard', 0);
    dashSheet.getRange('A1').setValue('no~bull books — Client Registry').setFontSize(16).setFontWeight('bold');
    dashSheet.getRange('A2').setValue('Last updated: ' + new Date().toLocaleString());
    dashSheet.getRange('A4').setValue('=COUNTA(Registry!A2:A)').setNumberFormat('0');
    dashSheet.getRange('B4').setValue('Total clients registered');
    dashSheet.getRange('A5').setFormula('=COUNTIF(Registry!K2:K,"Active")');
    dashSheet.getRange('B5').setValue('Active');
    dashSheet.getRange('A6').setFormula('=COUNTIF(Registry!K2:K,"Trial")');
    dashSheet.getRange('B6').setValue('Trial');
    dashSheet.getRange('A7').setFormula('=COUNTIF(Registry!K2:K,"Suspended")');
    dashSheet.getRange('B7').setValue('Suspended');
  }

  ss.setActiveSheet(ss.getSheetByName('Registry'));
  Logger.log('Registry initialised. Sheet URL: ' + ss.getUrl());
  Logger.log('REGISTRY_SHEET_ID = ' + id);
  return { success: true, sheetId: id, sheetUrl: ss.getUrl() };
}

/**
 * setRegistrySheetId(id)
 * Convenience function — run from editor to point to an existing registry sheet.
 */
function setRegistrySheetId(id) {
  PropertiesService.getScriptProperties().setProperty('REGISTRY_SHEET_ID', id);
  Logger.log('REGISTRY_SHEET_ID set to: ' + id);
}

// ─────────────────────────────────────────────────────────────────────────────
// CRUD OPERATIONS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * getAllRegistryClients(params)
 * Returns all registry entries. Called from the Admin Panel in the UI.
 */
function getAllRegistryClients(params) {
  try {
    var sheet = _getRegistrySheet();
    if (sheet.getLastRow() < 2) return { success: true, clients: [] };

    var data    = sheet.getDataRange().getValues();
    var clients = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      clients.push(_rowToClient(row));
    }

    return { success: true, clients: clients, total: clients.length };
  } catch(e) {
    Logger.log('getAllRegistryClients error: ' + e);
    return { success: false, message: e.toString(), clients: [] };
  }
}

/**
 * getRegistryClient(sheetId, params)
 * Look up a single client by their spreadsheet ID.
 */
function getRegistryClient(sheetId, params) {
  try {
    var sheet = _getRegistrySheet();
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][REGISTRY_COLS.SHEET_ID - 1] === sheetId) {
        return { success: true, client: _rowToClient(data[i]) };
      }
    }
    return { success: false, message: 'Client not found in registry.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * registerClient(clientData, params)
 * Add a new client to the registry. Called from Admin Panel or onboarding.
 * clientData: { companyName, contactName, contactEmail, contactPhone,
 *               sheetId, deployUrl, plan, notes, vatRegistered, vatNumber,
 *               fyEnd, country, accountant, accountantEmail, clientRef }
 */
function registerClient(clientData, params) {
  try {
    var sheet = _getRegistrySheet();

    // Check for duplicate sheetId
    if (clientData.sheetId) {
      var data = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [[]];
      for (var i = 1; i < data.length; i++) {
        if (data[i][REGISTRY_COLS.SHEET_ID - 1] === clientData.sheetId) {
          return { success: false, message: 'A client with this Sheet ID is already registered.' };
        }
      }
    }

    var registryId = generateId('REG');
    var now        = new Date().toISOString().split('T')[0];
    var sheetUrl   = clientData.sheetId
      ? 'https://docs.google.com/spreadsheets/d/' + clientData.sheetId
      : '';
    var clientRef  = clientData.clientRef ||
      (clientData.companyName || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

    sheet.appendRow([
      registryId,
      clientRef,
      clientData.companyName      || '',
      clientData.contactName      || '',
      clientData.contactEmail     || '',
      clientData.contactPhone     || '',
      clientData.sheetId          || '',
      sheetUrl,
      clientData.deployUrl        || ScriptApp.getService().getUrl(),
      clientData.plan             || 'Trial',
      clientData.status           || 'Trial',
      now,
      '',   // LastSeen — set on first login
      '',   // LastSeenBy
      0, 0, 0,   // invoice/client/bill counts
      REGISTRY_VERSION,
      clientData.notes            || '',
      clientData.vatRegistered    || false,
      clientData.vatNumber        || '',
      clientData.fyEnd            || '31 March',
      clientData.country          || 'UK',
      clientData.accountant       || '',
      clientData.accountantEmail  || ''
    ]);

    Logger.log('Registered client: ' + registryId + ' — ' + clientData.companyName);
    return {
      success:    true,
      registryId: registryId,
      clientRef:  clientRef,
      sheetUrl:   sheetUrl,
      message:    clientData.companyName + ' registered successfully.'
    };
  } catch(e) {
    Logger.log('registerClient error: ' + e);
    return { success: false, message: e.toString() };
  }
}

/**
 * updateRegistryClient(registryId, updates, params)
 * Update any fields for an existing registry entry.
 */
function updateRegistryClient(registryId, updates, params) {
  try {
    var sheet = _getRegistrySheet();
    var data  = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === registryId) {
        var row = i + 1;
        var map = {
          clientRef:       REGISTRY_COLS.CLIENT_REF,
          companyName:     REGISTRY_COLS.COMPANY_NAME,
          contactName:     REGISTRY_COLS.CONTACT_NAME,
          contactEmail:    REGISTRY_COLS.CONTACT_EMAIL,
          contactPhone:    REGISTRY_COLS.CONTACT_PHONE,
          sheetId:         REGISTRY_COLS.SHEET_ID,
          sheetUrl:        REGISTRY_COLS.SHEET_URL,
          deployUrl:       REGISTRY_COLS.DEPLOY_URL,
          plan:            REGISTRY_COLS.PLAN,
          status:          REGISTRY_COLS.STATUS,
          notes:           REGISTRY_COLS.NOTES,
          vatRegistered:   REGISTRY_COLS.VAT_REGISTERED,
          vatNumber:       REGISTRY_COLS.VAT_NUMBER,
          fyEnd:           REGISTRY_COLS.FY_END,
          country:         REGISTRY_COLS.COUNTRY,
          accountant:      REGISTRY_COLS.ACCOUNTANT,
          accountantEmail: REGISTRY_COLS.ACCOUNTANT_EMAIL
        };
        Object.keys(updates).forEach(function(key) {
          if (map[key]) sheet.getRange(row, map[key]).setValue(updates[key]);
        });
        // If sheetId updated, auto-update sheetUrl
        if (updates.sheetId) {
          sheet.getRange(row, REGISTRY_COLS.SHEET_URL).setValue(
            'https://docs.google.com/spreadsheets/d/' + updates.sheetId
          );
        }
        Logger.log('Updated registry client: ' + registryId);
        return { success: true, message: 'Registry entry updated.' };
      }
    }
    return { success: false, message: 'Registry entry not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * deactivateRegistryClient(registryId, reason, params)
 * Sets status to Cancelled and appends note.
 */
function deactivateRegistryClient(registryId, reason, params) {
  return updateRegistryClient(registryId, {
    status: 'Cancelled',
    notes:  '[Cancelled ' + new Date().toISOString().split('T')[0] + ': ' + (reason || '') + ']'
  }, params);
}

// ─────────────────────────────────────────────────────────────────────────────
// PING — called on each client session to keep registry fresh
// ─────────────────────────────────────────────────────────────────────────────

/**
 * pingRegistry(sheetId, meta)
 * Called automatically from getStartupData() on every client page load.
 * Updates LastSeen, counts, and version for the matching registry row.
 * Silently no-ops if registry is not configured.
 *
 * meta: { email, companyName, invoiceCount, clientCount, billCount, version }
 */
function pingRegistry(sheetId, meta) {
  try {
    var regId = PropertiesService.getScriptProperties().getProperty('REGISTRY_SHEET_ID');
    if (!regId) return; // registry not configured — skip silently

    var sheet = _getRegistrySheet();
    var data  = sheet.getDataRange().getValues();
    var now   = new Date().toISOString();
    var found = false;

    for (var i = 1; i < data.length; i++) {
      if (data[i][REGISTRY_COLS.SHEET_ID - 1] === sheetId) {
        var row = i + 1;
        sheet.getRange(row, REGISTRY_COLS.LAST_SEEN).setValue(now);
        if (meta.email)        sheet.getRange(row, REGISTRY_COLS.LAST_SEEN_BY).setValue(meta.email);
        if (meta.invoiceCount !== undefined) sheet.getRange(row, REGISTRY_COLS.INVOICE_COUNT).setValue(meta.invoiceCount);
        if (meta.clientCount  !== undefined) sheet.getRange(row, REGISTRY_COLS.CLIENT_COUNT).setValue(meta.clientCount);
        if (meta.billCount    !== undefined) sheet.getRange(row, REGISTRY_COLS.BILL_COUNT).setValue(meta.billCount);
        if (meta.version)      sheet.getRange(row, REGISTRY_COLS.VERSION).setValue(meta.version);
        if (meta.companyName && !data[i][REGISTRY_COLS.COMPANY_NAME - 1]) {
          sheet.getRange(row, REGISTRY_COLS.COMPANY_NAME).setValue(meta.companyName);
        }
        // Populate deployUrl if missing
        if (!data[i][REGISTRY_COLS.DEPLOY_URL - 1]) {
          sheet.getRange(row, REGISTRY_COLS.DEPLOY_URL).setValue(
            ScriptApp.getService().getUrl() + '?id=' + sheetId
          );
        }
        // Populate createdDate if missing
        if (!data[i][REGISTRY_COLS.CREATED_DATE - 1]) {
          sheet.getRange(row, REGISTRY_COLS.CREATED_DATE).setValue(new Date().toISOString().split('T')[0]);
        }
        found = true;
        break;
      }
    }

    // Auto-register if seen for the first time (new instance)
    if (!found && sheetId && meta.companyName) {
      registerClient({
        sheetId:      sheetId,
        companyName:  meta.companyName,
        contactEmail: meta.email || '',
        deployUrl:    ScriptApp.getService().getUrl() + '?id=' + sheetId,
        status:       'Trial',
        plan:         'Trial'
      });
    }
  } catch(e) {
    // Never let registry errors affect the client app
    Logger.log('pingRegistry error (non-fatal): ' + e.toString());
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// ADMIN PANEL — called from the UI
// ─────────────────────────────────────────────────────────────────────────────

/**
 * getRegistrySummary(params)
 * Returns summary stats + all clients for the Admin Panel.
 */
function getRegistrySummary(params) {
  try {
    var result = getAllRegistryClients(params);
    if (!result.success) return result;

    var clients  = result.clients;
    var active   = clients.filter(function(c){ return c.status === 'Active'; }).length;
    var trial    = clients.filter(function(c){ return c.status === 'Trial'; }).length;
    var totalInv = clients.reduce(function(s,c){ return s + (c.invoiceCount||0); }, 0);

    // Last-seen sort for "recently active" view
    clients.sort(function(a, b) {
      return (b.lastSeen || '') > (a.lastSeen || '') ? 1 : -1;
    });

    return {
      success:  true,
      summary: {
        total:           clients.length,
        active:          active,
        trial:           trial,
        totalInvoices:   totalInv,
        registrySheetId: PropertiesService.getScriptProperties().getProperty('REGISTRY_SHEET_ID') || ''
      },
      clients: clients
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// INTERNAL HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function _rowToClient(row) {
  return {
    registryId:     (row[0]  || '').toString(),
    clientRef:      (row[1]  || '').toString(),
    companyName:    (row[2]  || '').toString(),
    contactName:    (row[3]  || '').toString(),
    contactEmail:   (row[4]  || '').toString(),
    contactPhone:   (row[5]  || '').toString(),
    sheetId:        (row[6]  || '').toString(),
    sheetUrl:       (row[7]  || '').toString(),
    deployUrl:      (row[8]  || '').toString(),
    plan:           (row[9]  || 'Solo').toString(),
    status:         (row[10] || 'Active').toString(),
    createdDate:    safeSerializeDate(row[11]),
    lastSeen:       row[12] ? row[12].toString() : '',
    lastSeenBy:     (row[13] || '').toString(),
    invoiceCount:   parseInt(row[14]) || 0,
    clientCount:    parseInt(row[15]) || 0,
    billCount:      parseInt(row[16]) || 0,
    version:        (row[17] || '').toString(),
    notes:          (row[18] || '').toString(),
    vatRegistered:  row[19] === true || row[19] === 'TRUE',
    vatNumber:      (row[20] || '').toString(),
    fyEnd:          (row[21] || '31 March').toString(),
    country:        (row[22] || 'UK').toString(),
    accountant:     (row[23] || '').toString(),
    accountantEmail:(row[24] || '').toString(),
    // Computed fields
    sheetLink: row[6] ? 'https://docs.google.com/spreadsheets/d/' + row[6] : '',
    appLink:   row[8] ? row[8] : (row[6] ? 'https://script.google.com/a/macros/nobull.consulting/s/AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j/exec?id=' + row[6] : ''),
    trialEnd:  (function() {
      // Trial = 14 days from createdDate
      if (!row[11]) return '';
      var created = new Date(row[11]);
      if (isNaN(created.getTime())) return '';
      var trial = new Date(created.getTime() + 14 * 86400000);
      return safeSerializeDate(trial);
    })(),
    trialDaysLeft: (function() {
      if (!row[11]) return null;
      var created = new Date(row[11]);
      if (isNaN(created.getTime())) return null;
      var trial = new Date(created.getTime() + 14 * 86400000);
      var days = Math.ceil((trial - new Date()) / 86400000);
      return days;
    })()
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// EDITOR HELPERS — run these from the Apps Script editor
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Run from editor to manually add a client.
 */
function debug_registerClient() {
  var result = registerClient({
    companyName:   'Test Client Ltd',
    contactName:   'John Smith',
    contactEmail:  'john@testclient.com',
    sheetId:       'YOUR_CLIENT_SHEET_ID_HERE',
    plan:          'Solo',
    status:        'Trial',
    vatRegistered: false,
    country:       'UK',
    notes:         'Test registration'
  });
  Logger.log(JSON.stringify(result));
}

/**
 * Open the registry spreadsheet directly.
 */
function debug_openRegistry() {
  var id = PropertiesService.getScriptProperties().getProperty('REGISTRY_SHEET_ID');
  if (id) {
    Logger.log('Registry: https://docs.google.com/spreadsheets/d/' + id);
  } else {
    Logger.log('REGISTRY_SHEET_ID not set. Run initRegistry() first.');
  }
}
