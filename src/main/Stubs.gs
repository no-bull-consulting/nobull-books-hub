/**
 * NO~BULL BOOKS — STUBS & SUPPLEMENTARY FUNCTIONS
 * Functions routed in Api.gs that need concrete implementations.
 * These are functional stubs — expand each as needed.
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ─────────────────────────────────────────────────────────────────────────────
// FIXED ASSETS
// ─────────────────────────────────────────────────────────────────────────────

function getAllFixedAssets(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.FIXED_ASSETS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, assets: [] };

    var data   = sheet.getDataRange().getValues();
    var assets = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      assets.push({
        assetId:                row[0]  ? row[0].toString()  : '',
        name:                   row[1]  ? row[1].toString()  : '',
        category:               row[2]  ? row[2].toString()  : '',
        purchaseDate:           safeSerializeDate(row[3]),
        cost:                   parseFloat(row[4])  || 0,
        depreciationMethod:     row[5]  ? row[5].toString()  : 'StraightLine',
        usefulLifeYears:        parseFloat(row[6])  || 5,
        residualValue:          parseFloat(row[7])  || 0,
        accumulatedDepreciation:parseFloat(row[8])  || 0,
        netBookValue:           parseFloat(row[9])  || 0,
        status:                 row[10] ? row[10].toString() : 'Active',
        notes:                  row[11] ? row[11].toString() : ''
      });
    }
    return { success: true, assets: assets };
  } catch(e) {
    Logger.log('getAllFixedAssets error: ' + e.toString());
    return { success: false, message: e.toString(), assets: [] };
  }
}

function createFixedAsset(params) {
  try {
    _auth('coa.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.FIXED_ASSETS);
    if (!sheet) return { success: false, message: 'FixedAssets sheet not found.' };

    var id       = generateId('FA');
    var cost     = parseFloat(params.cost) || 0;
    var residual = parseFloat(params.residualValue) || 0;

    sheet.appendRow([
      id,
      params.name              || '',
      params.category          || 'Equipment',
      params.purchaseDate      || new Date().toISOString().split('T')[0],
      cost,
      params.depreciationMethod|| 'StraightLine',
      parseFloat(params.usefulLifeYears) || 5,
      residual,
      0,         // accumulatedDepreciation
      cost,      // netBookValue = cost on creation
      'Active',
      params.notes || ''
    ]);

    logAudit('CREATE', 'FixedAsset', id, { name: params.name });
    return { success: true, assetId: id, message: 'Fixed asset created.' };
  } catch(e) {
    Logger.log('createFixedAsset error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function updateFixedAsset(assetId, params) {
  try {
    _auth('coa.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.FIXED_ASSETS);
    if (!sheet) return { success: false, message: 'FixedAssets sheet not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === assetId) {
        var row = i + 1;
        if (params.name)              sheet.getRange(row, 2).setValue(params.name);
        if (params.category)          sheet.getRange(row, 3).setValue(params.category);
        if (params.purchaseDate)      sheet.getRange(row, 4).setValue(params.purchaseDate);
        if (params.cost !== undefined)sheet.getRange(row, 5).setValue(parseFloat(params.cost) || 0);
        if (params.notes !== undefined)sheet.getRange(row, 12).setValue(params.notes || '');
        logAudit('UPDATE', 'FixedAsset', assetId, { name: params.name });
        return { success: true, message: 'Asset updated.' };
      }
    }
    return { success: false, message: 'Asset not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function disposeFixedAsset(assetId, disposalDate, disposalProceeds, notes, params) {
  try {
    _auth('coa.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.FIXED_ASSETS);
    if (!sheet) return { success: false, message: 'FixedAssets sheet not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === assetId) {
        sheet.getRange(i + 1, 11).setValue('Disposed');
        sheet.getRange(i + 1, 12).setValue('Disposed ' + disposalDate + (notes ? ': ' + notes : ''));
        logAudit('DISPOSE', 'FixedAsset', assetId, { date: disposalDate, proceeds: disposalProceeds });
        return { success: true, message: 'Asset disposed.' };
      }
    }
    return { success: false, message: 'Asset not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function runDepreciationRun(periodEndDate, periodMonths, postToLedger, params) {
  try {
    _auth('coa.write', params);
    var db    = getDb(params || {});
    var sheet = db.getSheetByName(SHEETS.FIXED_ASSETS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, message: 'No assets to depreciate.' };

    var data      = sheet.getDataRange().getValues();
    var processed = 0;
    var totalDepr = 0;
    var months    = parseFloat(periodMonths) || 3;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] || row[10].toString() !== 'Active') continue;

      var cost          = parseFloat(row[4])  || 0;
      var residual      = parseFloat(row[7])  || 0;
      var usefulLife    = parseFloat(row[6])  || 5;
      var accumulated   = parseFloat(row[8])  || 0;
      var nbv           = parseFloat(row[9])  || cost;
      var method        = row[5] ? row[5].toString() : 'StraightLine';

      var periodDepr = 0;
      if (method === 'StraightLine') {
        var annualDepr = (cost - residual) / usefulLife;
        periodDepr     = annualDepr * (months / 12);
      } else if (method === 'ReducingBalance') {
        var rate       = 1 - Math.pow(residual / cost, 1 / usefulLife);
        periodDepr     = nbv * rate * (months / 12);
      }

      periodDepr    = Math.min(periodDepr, Math.max(0, nbv - residual));
      periodDepr    = Math.round(periodDepr * 100) / 100;
      accumulated  += periodDepr;
      nbv          -= periodDepr;
      totalDepr    += periodDepr;

      sheet.getRange(i + 1, 9).setValue(Math.round(accumulated * 100) / 100);
      sheet.getRange(i + 1, 10).setValue(Math.round(nbv * 100) / 100);
      if (nbv <= residual) sheet.getRange(i + 1, 11).setValue('FullyDepreciated');
      processed++;
    }

    // Log the depreciation run
    var runSheet = db.getSheetByName(SHEETS.DEPRECIATION_RUNS);
    if (runSheet) {
      runSheet.appendRow([
        generateId('DR'), periodEndDate, months, processed,
        Math.round(totalDepr * 100) / 100, postToLedger || false,
        new Date().toISOString().split('T')[0],
        Session.getActiveUser().getEmail()
      ]);
    }

    return {
      success:           true,
      assetsProcessed:   processed,
      totalDepreciation: Math.round(totalDepr * 100) / 100,
      message:           processed + ' asset(s) depreciated. Total: £' + (Math.round(totalDepr * 100) / 100).toFixed(2)
    };
  } catch(e) {
    Logger.log('runDepreciationRun error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getDepreciationSchedule(assetId, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.FIXED_ASSETS);
    if (!sheet) return { success: false, message: 'FixedAssets sheet not found.', schedule: [] };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === assetId) {
        var cost       = parseFloat(data[i][4]) || 0;
        var residual   = parseFloat(data[i][7]) || 0;
        var life       = parseFloat(data[i][6]) || 5;
        var method     = data[i][5] ? data[i][5].toString() : 'StraightLine';
        var startDate  = data[i][3] ? new Date(data[i][3]) : new Date();

        var schedule = [];
        var nbv      = cost;
        for (var y = 1; y <= life; y++) {
          var annualDepr = method === 'StraightLine'
            ? (cost - residual) / life
            : nbv * (1 - Math.pow(residual / cost, 1 / life));
          annualDepr = Math.min(annualDepr, Math.max(0, nbv - residual));
          nbv -= annualDepr;
          schedule.push({
            year:           y,
            yearEnd:        (startDate.getFullYear() + y) + '-03-31',
            depreciation:   Math.round(annualDepr * 100) / 100,
            accumulated:    Math.round((cost - nbv) * 100) / 100,
            netBookValue:   Math.round(nbv * 100) / 100
          });
          if (nbv <= residual) break;
        }
        return { success: true, schedule: schedule };
      }
    }
    return { success: false, message: 'Asset not found.', schedule: [] };
  } catch(e) {
    return { success: false, message: e.toString(), schedule: [] };
  }
}

function getDepreciationRuns(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.DEPRECIATION_RUNS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, runs: [] };
    var data = sheet.getDataRange().getValues();
    var runs = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      runs.push({
        runId:             data[i][0].toString(),
        periodEndDate:     safeSerializeDate(data[i][1]),
        periodMonths:      parseFloat(data[i][2]) || 3,
        assetsProcessed:   parseInt(data[i][3])   || 0,
        totalDepreciation: parseFloat(data[i][4]) || 0,
        postedToLedger:    data[i][5] === true,
        runDate:           safeSerializeDate(data[i][6]),
        runBy:             data[i][7] ? data[i][7].toString() : ''
      });
    }
    runs.sort(function(a, b) { return (b.runDate || '') > (a.runDate || '') ? 1 : -1; });
    return { success: true, runs: runs };
  } catch(e) {
    return { success: false, message: e.toString(), runs: [] };
  }
}

function initFixedAssetSheets(params) {
  return { success: true, message: 'Fixed asset sheets initialised by Initializer.gs.' };
}

// ─────────────────────────────────────────────────────────────────────────────
// FINANCIAL YEARS
// ─────────────────────────────────────────────────────────────────────────────

function getFinancialYears(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.FINANCIAL_YEARS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, years: [] };

    var data  = sheet.getDataRange().getValues();
    var years = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      years.push({
        yearId:     data[i][0].toString(),
        yearLabel:  data[i][1] ? data[i][1].toString() : '',
        startDate:  safeSerializeDate(data[i][2]),
        endDate:    safeSerializeDate(data[i][3]),
        status:     data[i][4] ? data[i][4].toString() : 'Open',
        closedDate: safeSerializeDate(data[i][5]),
        closedBy:   data[i][6] ? data[i][6].toString() : ''
      });
    }
    return { success: true, years: years };
  } catch(e) {
    Logger.log('getFinancialYears error: ' + e.toString());
    return { success: false, message: e.toString(), years: [] };
  }
}

function runPreCloseChecks(yearEndDate, params) {
  try {
    _auth('settings.write', params);
    var invs = (getAllInvoices(params).invoices || []);
    var bils = (getAllBills(params).bills || []);

    var unpaidInvs = invs.filter(function(i) {
      return i.status !== 'Paid' && i.status !== 'Void' && i.status !== 'Draft';
    });
    var unpaidBils = bils.filter(function(b) {
      return b.status !== 'Paid' && b.status !== 'Void';
    });

    var checks = [
      { name: 'Unpaid invoices',    passed: unpaidInvs.length === 0, detail: unpaidInvs.length + ' outstanding invoice(s)' },
      { name: 'Unpaid bills',       passed: unpaidBils.length === 0, detail: unpaidBils.length + ' outstanding bill(s)' },
      { name: 'Bank accounts exist',passed: (getBankAccounts(params).accounts || []).length > 0, detail: '' },
      { name: 'COA exists',         passed: (getAccounts({}, params).accounts || []).length > 0, detail: '' }
    ];

    var allPassed = checks.every(function(c) { return c.passed; });
    return { success: true, checks: checks, allPassed: allPassed };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getYearEndSummary(yearStart, yearEnd, params) {
  try {
    _auth('settings.write', params);
    var plResult = generateProfitLoss(yearStart, yearEnd, params);
    return {
      success:       true,
      yearStart:     yearStart,
      yearEnd:       yearEnd,
      totalRevenue:  plResult.success ? plResult.report.totalRevenue  : 0,
      totalExpenses: plResult.success ? plResult.report.totalExpenses : 0,
      netProfit:     plResult.success ? plResult.report.netProfit     : 0
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function closeFinancialYear(params) {
  try {
    _auth('settings.write', params);
    var db    = getDb(params || {});
    var sheet = db.getSheetByName(SHEETS.FINANCIAL_YEARS);
    if (!sheet) return { success: false, message: 'FinancialYears sheet not found.' };

    var yearId    = generateId('FY');
    var ctx       = _getCurrentUserContext(params);
    sheet.appendRow([
      yearId,
      params.label     || (params.yearStart + ' - ' + params.yearEnd),
      params.yearStart || '',
      params.yearEnd   || '',
      'Closed',
      new Date().toISOString().split('T')[0],
      ctx.email || ''
    ]);

    logAudit('CLOSE', 'FinancialYear', yearId, { label: params.label });
    return { success: true, yearId: yearId, message: 'Financial year closed.' };
  } catch(e) {
    Logger.log('closeFinancialYear error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function reopenFinancialYear(yearId, reason, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.FINANCIAL_YEARS);
    if (!sheet) return { success: false, message: 'FinancialYears sheet not found.' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === yearId) {
        sheet.getRange(i + 1, 5).setValue('Open');
        sheet.getRange(i + 1, 6).setValue('');
        logAudit('REOPEN', 'FinancialYear', yearId, { reason: reason });
        return { success: true, message: 'Financial year reopened.' };
      }
    }
    return { success: false, message: 'Financial year not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function initFinancialYearSheets(params) {
  return { success: true, message: 'Financial year sheets initialised by Initializer.gs.' };
}

// ─────────────────────────────────────────────────────────────────────────────
// RECURRING INVOICES
// ─────────────────────────────────────────────────────────────────────────────

function getAllRecurring(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.RECURRING_INVOICES);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, recurring: [] };

    var data      = sheet.getDataRange().getValues();
    var recurring = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var lines = [];
      try { lines = JSON.parse(row[6] || '[]'); } catch(e) { lines = []; }
      recurring.push({
        recurringId: row[0].toString(),
        clientId:    row[1] ? row[1].toString() : '',
        clientName:  row[2] ? row[2].toString() : '',
        frequency:   row[3] ? row[3].toString() : 'Monthly',
        nextRun:     safeSerializeDate(row[4]),
        lastRun:     '',
        lines:       lines,
        status:      row[7] ? row[7].toString() : 'Active',
        createdBy:   row[8] ? row[8].toString() : '',
        createdDate: safeSerializeDate(row[9]),
        total:       lines.reduce(function(s, l) {
          return s + ((parseFloat(l.qty)||1) * (parseFloat(l.unitPrice)||0) * (1 + (parseFloat(l.vatRate)||0)/100));
        }, 0)
      });
    }
    return { success: true, recurring: recurring };
  } catch(e) {
    Logger.log('getAllRecurring error: ' + e.toString());
    return { success: false, message: e.toString(), recurring: [] };
  }
}

function createRecurring(params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.RECURRING_INVOICES);
    if (!sheet) return { success: false, message: 'RecurringInvoices sheet not found.' };

    var ctx = _getCurrentUserContext(params);
    var id  = generateId('REC');
    sheet.appendRow([
      id,
      params.clientId    || '',
      params.clientName  || '',
      params.frequency   || 'Monthly',
      params.nextDate    || '',
      params.invoicePrefix || 'INV-',
      JSON.stringify(params.lines || []),
      'Active',
      ctx.email || '',
      new Date().toISOString().split('T')[0]
    ]);
    return { success: true, recurringId: id, message: 'Recurring schedule created.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function updateRecurring(recurringId, params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.RECURRING_INVOICES);
    if (!sheet) return { success: false, message: 'RecurringInvoices sheet not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === recurringId) {
        var row = i + 1;
        if (params.frequency)  sheet.getRange(row, 4).setValue(params.frequency);
        if (params.nextDate)   sheet.getRange(row, 5).setValue(params.nextDate);
        if (params.lines)      sheet.getRange(row, 7).setValue(JSON.stringify(params.lines));
        if (params.status)     sheet.getRange(row, 8).setValue(params.status);
        return { success: true, message: 'Recurring schedule updated.' };
      }
    }
    return { success: false, message: 'Recurring schedule not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteRecurring(recurringId, params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.RECURRING_INVOICES);
    if (!sheet) return { success: false, message: 'RecurringInvoices sheet not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === recurringId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Recurring schedule deleted.' };
      }
    }
    return { success: false, message: 'Not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function processRecurringInvoices(params) {
  try {
    var recResult = getAllRecurring(params);
    if (!recResult.success) return recResult;

    var today   = new Date(); today.setHours(0,0,0,0);
    var created = [];

    recResult.recurring.forEach(function(rec) {
      if (rec.status !== 'Active') return;
      if (!rec.nextRun) return;
      var nextRun = new Date(rec.nextRun); nextRun.setHours(0,0,0,0);
      if (nextRun > today) return;

      // Create invoice
      try {
        var invResult = createInvoice(rec.clientId, rec.lines, null, 'Auto-generated recurring invoice', params);
        if (invResult && invResult.success) {
          created.push(invResult.invoiceId);
          // Advance next run date
          var next = new Date(nextRun);
          if (rec.frequency === 'Weekly')        next.setDate(next.getDate() + 7);
          else if (rec.frequency === 'Monthly')  next.setMonth(next.getMonth() + 1);
          else if (rec.frequency === 'Quarterly')next.setMonth(next.getMonth() + 3);
          else if (rec.frequency === 'Annually') next.setFullYear(next.getFullYear() + 1);
          updateRecurring(rec.recurringId, { nextDate: next.toISOString().split('T')[0] }, params);
        }
      } catch(ie) {
        Logger.log('processRecurring: could not create invoice for ' + rec.recurringId + ': ' + ie);
      }
    });

    return { success: true, created: created, count: created.length,
             message: created.length + ' invoice(s) created.' };
  } catch(e) {
    Logger.log('processRecurringInvoices error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function installRecurringTrigger() {
  try {
    // Remove existing triggers first
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'processRecurringInvoices') {
        ScriptApp.deleteTrigger(t);
      }
    });
    ScriptApp.newTrigger('processRecurringInvoices')
      .timeBased().everyDays(1).atHour(6).create();
    return { success: true, message: 'Daily recurring trigger installed (runs at 06:00).' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// MISC STUBS (routed in Api.gs, safe defaults)
// ─────────────────────────────────────────────────────────────────────────────

function getSecurityStatus()           { return { success: true, status: 'OK' }; }
function getAuditLog(params)           { return { success: true, entries: [] }; }
function pingRegistry(event)           { /* no-op if REGISTRY_URL not set */ }
function getAllInstances()             { return { success: true, instances: [] }; }
function getInstanceMeta()            { return { success: true, meta: {} }; }
function getBackupStatus()            { return { success: true, hasTrigger: false, backupCount: 0 }; }
function installBackupTrigger()       { return { success: true, message: 'Backup trigger not yet implemented.' }; }
function removeBackupTrigger()        { return { success: true }; }
function runManualBackup()            { return { success: true, backupName: 'Manual backup not yet implemented.' }; }
function protectSensitiveSheets()     { return { success: true, message: 'Sheet protection not yet implemented.' }; }
function runSandboxValidation()       { return { success: true, summary: 'Sandbox validation not yet implemented.', results: [] }; }
function sandboxVATSubmitDryRun()     { return { success: true }; }
function verifyIntegrity(params)      { return { success: true }; }
function diagnoseSheets(params)       { return { success: true }; }
function rebuildAccountBalances(params){ return { success: true }; }
function cleanDuplicateTransactions(params){ return { success: true }; }
function verifySchemaIntegrity(params){ return { success: true }; }
function getIntegrityStatus(params)   { return { success: true }; }
function initializeSystem(params)     { return checkAndInitSheet(params); }
function createBackup(params)         { return { success: true }; }
function getITSAObligationsFromSheet(params){ return { success: true, obligations: [] }; }
function getITSASubmissions(params)   { return { success: true, submissions: [] }; }
function submitQuarterlyUpdate(nino, businessId, taxYear, quarter, income, params){ return { success: false, message: 'ITSA not yet configured.' }; }
function triggerAndGetCalculation(nino, taxYear, params){ return { success: false, message: 'ITSA not yet configured.' }; }
function eraseClient(clientId, retainFinancial){ return { success: false, message: 'GDPR erase not yet implemented.' }; }
function exportClientData(clientId)   { return { success: false, message: 'GDPR export not yet implemented.' }; }
function generateId(prefix) {
  return (prefix || 'ID') + '_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5).toUpperCase();
}
function logAudit(action, entity, id, detail) {
  try { Logger.log('[AUDIT] ' + action + ' | ' + entity + ' | ' + id + ' | ' + JSON.stringify(detail)); } catch(e) {}
}
function safeSerializeDate(val) {
  if (!val) return '';
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return val ? val.toString() : '';
    return d.toISOString().split('T')[0];
  } catch(e) { return ''; }
}
function _sendAlert(subject, body) {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) MailApp.sendEmail(email, 'no~bull books: ' + subject, body);
  } catch(e) { Logger.log('_sendAlert: ' + e); }
}
