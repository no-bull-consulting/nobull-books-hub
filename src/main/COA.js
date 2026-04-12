/**
 * NO~BULL BOOKS -- CHART OF ACCOUNTS
 * Account management, general ledger, trial balance
 *
 * KEY FIX: SPREADSHEET_ID removed everywhere; getDb(params) used throughout.
 * params threaded through every public function signature.
 * -----------------------------------------------------------------------------
 */

function getAccounts(filters, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.CHART_OF_ACCOUNTS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, accounts: [] };
    }

    var data     = sheet.getDataRange().getValues();
    var accounts = [];

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      var account = {
        accountCode:    data[i][0] ? data[i][0].toString().trim() : '',
        accountName:    data[i][1] ? data[i][1].toString().trim() : '',
        accountType:    data[i][2] ? data[i][2].toString().trim() : '',
        category:       data[i][3] ? data[i][3].toString().trim() : '',
        openingBalance: parseFloat(data[i][4]) || 0,
        currentBalance: parseFloat(data[i][5]) || 0,
        active:         data[i][6] === true || data[i][6] === 'TRUE' || data[i][6] === 'true',
        notes:          data[i][7] ? data[i][7].toString().trim() : ''
      };

      if (filters) {
        if (filters.accountType && account.accountType !== filters.accountType) continue;
        if (filters.category    && account.category    !== filters.category)    continue;
        if (filters.activeOnly  && !account.active)                             continue;
        if (filters.search) {
          var q = filters.search.toString().toLowerCase();
          if (account.accountName.toLowerCase().indexOf(q) < 0 &&
              account.accountCode.toLowerCase().indexOf(q) < 0) continue;
        }
      }

      accounts.push(account);
    }

    accounts.sort(function(a, b) { return a.accountCode.localeCompare(b.accountCode); });
    return { success: true, accounts: accounts };
  } catch(e) {
    Logger.log('getAccounts error: ' + e.toString());
    return { success: false, message: e.toString(), accounts: [] };
  }
}

function getAccountTypes(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.CHART_OF_ACCOUNTS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, types: [], categories: [] };
    }

    var data       = sheet.getDataRange().getValues();
    var typeSet    = {};
    var catSet     = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][2]) typeSet[data[i][2].toString().trim()] = true;
      if (data[i][3]) catSet[ data[i][3].toString().trim()] = true;
    }

    return {
      success:    true,
      types:      Object.keys(typeSet).sort(),
      categories: Object.keys(catSet).sort()
    };
  } catch(e) {
    Logger.log('getAccountTypes error: ' + e.toString());
    return { success: false, message: e.toString(), types: [], categories: [] };
  }
}

// -----------------------------------------------------------------------------
// GENERAL LEDGER
// -----------------------------------------------------------------------------

function getGeneralLedger(filters, params) {
  try {
    var ss = getDb(params || {});

    // Build account lookup map
    var coaSheet  = ss.getSheetByName(SHEETS.CHART_OF_ACCOUNTS);
    var accountMap = {};
    if (coaSheet && coaSheet.getLastRow() > 1) {
      var coaData = coaSheet.getDataRange().getValues();
      for (var i = 1; i < coaData.length; i++) {
        if (coaData[i][0]) {
          var code = coaData[i][0].toString().trim();
          accountMap[code] = {
            name:     coaData[i][1] ? coaData[i][1].toString().trim() : code,
            type:     coaData[i][2] ? coaData[i][2].toString().trim() : '',
            category: coaData[i][3] ? coaData[i][3].toString().trim() : ''
          };
        }
      }
    }

    // Read Transactions sheet
    var txnSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
    if (!txnSheet || txnSheet.getLastRow() < 2) {
      return { success: true, entries: [], totalEntries: 0, runningBalance: null, availableTypes: [] };
    }

    var data     = txnSheet.getDataRange().getValues();
    var entries  = [];
    var dateFrom = filters && filters.dateFrom ? new Date(filters.dateFrom) : null;
    var dateTo   = filters && filters.dateTo   ? new Date(filters.dateTo)   : null;
    if (dateTo) dateTo.setHours(23, 59, 59, 999);

    for (var j = 1; j < data.length; j++) {
      var row = data[j];
      if (!row[0]) continue;

      var txDate = row[1] ? new Date(row[1]) : null;
      if (dateFrom && txDate && txDate < dateFrom) continue;
      if (dateTo   && txDate && txDate > dateTo)   continue;

      var debitCode  = row[4] ? row[4].toString().trim() : '';
      var creditCode = row[5] ? row[5].toString().trim() : '';
      var txType     = row[2] ? row[2].toString().trim() : '';

      // Sanitise reference (guard against date objects stored in reference cell)
      var rawRef = row[3];
      var ref    = '';
      if (rawRef instanceof Date) {
        ref = safeSerializeDate(rawRef);
      } else if (rawRef) {
        var refStr = rawRef.toString().trim();
        ref = /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/.test(refStr)
          ? safeSerializeDate(new Date(refStr))
          : refStr;
      }

      var desc = row[7] ? row[7].toString().trim() : '';

      // Filters
      if (filters && filters.accountCode) {
        if (debitCode !== filters.accountCode && creditCode !== filters.accountCode) continue;
      }
      if (filters && filters.type && txType !== filters.type) continue;
      if (filters && filters.search) {
        var q = filters.search.toLowerCase();
        var dn = accountMap[debitCode]  ? accountMap[debitCode].name.toLowerCase()  : '';
        var cn = accountMap[creditCode] ? accountMap[creditCode].name.toLowerCase() : '';
        if (ref.toLowerCase().indexOf(q) < 0 && desc.toLowerCase().indexOf(q) < 0 &&
            dn.indexOf(q) < 0 && cn.indexOf(q) < 0 &&
            debitCode.toLowerCase().indexOf(q) < 0 && creditCode.toLowerCase().indexOf(q) < 0) continue;
      }

      var amount = parseFloat(row[6]) || 0;

      entries.push({
        transactionId: String(row[0]),
        date:          safeSerializeDate(row[1]),
        type:          txType,
        reference:     ref,
        debitCode:     debitCode,
        debitName:     accountMap[debitCode]  ? accountMap[debitCode].name  : debitCode,
        creditCode:    creditCode,
        creditName:    accountMap[creditCode] ? accountMap[creditCode].name : creditCode,
        amount:        amount,
        description:   desc,
        invoiceId:     String(row[8]  || ''),
        billId:        String(row[9]  || ''),
        reconciled:    row[10] === true || row[10] === 'TRUE'
      });
    }

    // Sort by date ascending, then by transactionId
    entries.sort(function(a, b) {
      var da = a.date || '', db = b.date || '';
      return da < db ? -1 : da > db ? 1 : a.transactionId.localeCompare(b.transactionId);
    });

    // Running balance when filtered to a single account
    var runningBalance = null;
    if (filters && filters.accountCode && accountMap[filters.accountCode]) {
      var acct          = accountMap[filters.accountCode];
      var balance       = 0;
      var debitIncreases = (acct.type === 'Asset' || acct.type === 'Expense');
      entries.forEach(function(e) {
        if (e.debitCode === filters.accountCode) {
          balance += debitIncreases ?  e.amount : -e.amount;
        } else {
          balance += debitIncreases ? -e.amount :  e.amount;
        }
        e.runningBalance = Math.round(balance * 100) / 100;
        e.side = e.debitCode === filters.accountCode ? 'DR' : 'CR';
      });
      runningBalance = Math.round(balance * 100) / 100;
    }

    // Distinct types for filter dropdown
    var typeSet2 = {};
    entries.forEach(function(e) { if (e.type) typeSet2[e.type] = true; });

    return {
      success:        true,
      entries:        entries,
      totalEntries:   entries.length,
      runningBalance: runningBalance,
      accountInfo:    filters && filters.accountCode ? (accountMap[filters.accountCode] || null) : null,
      availableTypes: Object.keys(typeSet2).sort()
    };
  } catch(e) {
    Logger.log('getGeneralLedger error: ' + e.toString());
    return { success: false, message: e.toString(), entries: [] };
  }
}

// -----------------------------------------------------------------------------
// TRIAL BALANCE
// -----------------------------------------------------------------------------

function getTrialBalance(dateFrom, dateTo, params) {
  try {
    var ss = getDb(params || {});

    // Build account map
    var coaSheet   = ss.getSheetByName(SHEETS.CHART_OF_ACCOUNTS);
    var accountMap = {};
    var accountOrder = [];

    if (coaSheet && coaSheet.getLastRow() > 1) {
      var coaData = coaSheet.getDataRange().getValues();
      for (var i = 1; i < coaData.length; i++) {
        if (coaData[i][0]) {
          var code = coaData[i][0].toString().trim();
          accountMap[code] = {
            code:           code,
            name:           coaData[i][1] ? coaData[i][1].toString().trim() : code,
            type:           coaData[i][2] ? coaData[i][2].toString().trim() : '',
            category:       coaData[i][3] ? coaData[i][3].toString().trim() : '',
            openingBalance: parseFloat(coaData[i][4]) || 0,
            totalDebits:    0,
            totalCredits:   0
          };
          accountOrder.push(code);
        }
      }
    }

    // Accumulate transaction debits/credits
    var txnSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
    if (txnSheet && txnSheet.getLastRow() > 1) {
      var data  = txnSheet.getDataRange().getValues();
      var dfrom = dateFrom ? new Date(dateFrom) : null;
      var dto   = dateTo   ? new Date(dateTo)   : null;
      if (dto) dto.setHours(23, 59, 59, 999);

      for (var j = 1; j < data.length; j++) {
        var row = data[j];
        if (!row[0]) continue;
        var txDate = row[1] ? new Date(row[1]) : null;
        if (dfrom && txDate && txDate < dfrom) continue;
        if (dto   && txDate && txDate > dto)   continue;

        var amount = parseFloat(row[6]) || 0;
        var dc     = row[4] ? row[4].toString().trim() : '';
        var cc     = row[5] ? row[5].toString().trim() : '';

        if (accountMap[dc]) accountMap[dc].totalDebits  += amount;
        if (accountMap[cc]) accountMap[cc].totalCredits += amount;
      }
    }

    // Build trial balance rows
    var rows         = [];
    var grandDebits  = 0;
    var grandCredits = 0;

    accountOrder.forEach(function(code) {
      var a              = accountMap[code];
      var debitIncreases = (a.type === 'Asset' || a.type === 'Expense');
      var balance        = a.openingBalance + (debitIncreases
        ? (a.totalDebits - a.totalCredits)
        : (a.totalCredits - a.totalDebits));
      balance = Math.round(balance * 100) / 100;

      var drBalance = 0, crBalance = 0;
      if      ( balance >= 0 &&  debitIncreases) drBalance =  balance;
      else if ( balance <  0 &&  debitIncreases) crBalance = -balance;
      else if ( balance >= 0 && !debitIncreases) crBalance =  balance;
      else                                        drBalance = -balance;

      grandDebits  += drBalance;
      grandCredits += crBalance;

      rows.push({
        code:         code,
        name:         a.name,
        type:         a.type,
        category:     a.category,
        totalDebits:  Math.round(a.totalDebits  * 100) / 100,
        totalCredits: Math.round(a.totalCredits * 100) / 100,
        balance:      balance,
        drBalance:    Math.round(drBalance  * 100) / 100,
        crBalance:    Math.round(crBalance  * 100) / 100
      });
    });

    return {
      success:      true,
      rows:         rows,
      grandDebits:  Math.round(grandDebits  * 100) / 100,
      grandCredits: Math.round(grandCredits * 100) / 100,
      balanced:     Math.abs(grandDebits - grandCredits) < 0.01
    };
  } catch(e) {
    Logger.log('getTrialBalance error: ' + e.toString());
    return { success: false, message: e.toString(), rows: [] };
  }
}

// -----------------------------------------------------------------------------
// ACCOUNT WRITE OPERATIONS
// -----------------------------------------------------------------------------

function createAccount(accountData, params) {
  try {
    _auth('coa.write', params);
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.CHART_OF_ACCOUNTS);
    if (!sheet) return { success: false, message: 'Chart of Accounts sheet not found.' };

    var data = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues() : [[]];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === accountData.accountCode.toString()) {
        return { success: false, message: 'Account code ' + accountData.accountCode + ' already exists.' };
      }
    }

    var opening = parseFloat(accountData.openingBalance) || 0;
    sheet.appendRow([
      accountData.accountCode,
      accountData.accountName,
      accountData.accountType,
      accountData.category        || '',
      opening,
      opening,
      accountData.active !== false,
      accountData.notes           || ''
    ]);

    var isBankAccount = accountData.accountType === 'Asset' &&
      (accountData.category === 'Bank Accounts' || accountData.isBankAccount === true);
    if (isBankAccount) _syncCOABankAccount(ss, accountData, 'create');

    logAudit('CREATE', 'Account', accountData.accountCode, { name: accountData.accountName });
    return {
      success:       true,
      message:       'Account ' + accountData.accountCode + ' -- ' + accountData.accountName + ' created.' +
                     (isBankAccount ? ' Bank account record also created.' : ''),
      isBankAccount: isBankAccount
    };
  } catch(e) {
    Logger.log('createAccount error: ' + e.toString());
    return { success: false, message: 'Error creating account: ' + e.toString() };
  }
}

function updateAccount(accountData, params) {
  try {
    _auth('coa.write', params);
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.CHART_OF_ACCOUNTS);
    if (!sheet) return { success: false, message: 'Chart of Accounts sheet not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === accountData.accountCode.toString()) {
        var row = i + 1;
        sheet.getRange(row, 2).setValue(accountData.accountName);
        sheet.getRange(row, 3).setValue(accountData.accountType);
        sheet.getRange(row, 4).setValue(accountData.category || '');
        sheet.getRange(row, 5).setValue(parseFloat(accountData.openingBalance) || 0);
        sheet.getRange(row, 7).setValue(accountData.active !== false);
        sheet.getRange(row, 8).setValue(accountData.notes || '');

        var isBankAccount = accountData.accountType === 'Asset' &&
          (accountData.category === 'Bank Accounts' || accountData.isBankAccount === true);
        if (isBankAccount) _syncCOABankAccount(ss, accountData, 'update');

        logAudit('UPDATE', 'Account', accountData.accountCode, { name: accountData.accountName });
        return {
          success:       true,
          message:       'Account ' + accountData.accountCode + ' updated.' +
                         (isBankAccount ? ' Bank account record also updated.' : ''),
          isBankAccount: isBankAccount
        };
      }
    }
    return { success: false, message: 'Account not found.' };
  } catch(e) {
    Logger.log('updateAccount error: ' + e.toString());
    return { success: false, message: 'Error updating account: ' + e.toString() };
  }
}

function deleteAccount(accountCode, params) {
  _auth('coa.write', params);
  var sheet = getDb(params).getSheetByName(SHEETS.CHART_OF_ACCOUNTS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === accountCode) {
      sheet.deleteRow(i + 1); // Hard delete for compliance cleanup
      logAudit('DELETE', 'Account', accountCode, {}, params);
      return { success: true };
    }
  }
}

// -----------------------------------------------------------------------------
// INTERNAL HELPERS
// -----------------------------------------------------------------------------

/**
 * _syncCOABankAccount
 * Keeps BankAccounts sheet in sync when a bank-type COA account is created/updated.
 * ss is already the correct spreadsheet -- no SPREADSHEET_ID needed.
 */
function _syncCOABankAccount(ss, accountData, action) {
  try {
    var bankSheet = ss.getSheetByName(SHEETS.BANK_ACCOUNTS);
    if (!bankSheet) return;

    var bankData    = bankSheet.getLastRow() > 1 ? bankSheet.getDataRange().getValues() : [[]];
    var code        = accountData.accountCode.toString();
    var existingRow = -1;

    for (var i = 1; i < bankData.length; i++) {
      var rowCode = bankData[i][10] ? bankData[i][10].toString().trim() : '';
      var rowName = bankData[i][1]  ? bankData[i][1].toString().trim().toLowerCase() : '';
      if (rowCode === code || rowName === (accountData.accountName || '').toLowerCase()) {
        existingRow = i + 1;
        break;
      }
    }

    if (action === 'create' && existingRow < 0) {
      var accountId = generateId('BA');
      bankSheet.appendRow([
        accountId,
        accountData.accountName,
        accountData.bankName        || '',
        accountData.bankAccountType || 'Current',
        accountData.sortCode        || '',
        accountData.accountNumber   || '',
        parseFloat(accountData.openingBalance) || 0,
        parseFloat(accountData.openingBalance) || 0,
        '',    // LastReconciled
        true,  // Active
        code   // NominalCode
      ]);
    } else if (existingRow > 0) {
      bankSheet.getRange(existingRow, 2).setValue(accountData.accountName);
      if (accountData.bankName)        bankSheet.getRange(existingRow, 3).setValue(accountData.bankName);
      if (accountData.bankAccountType) bankSheet.getRange(existingRow, 4).setValue(accountData.bankAccountType);
      if (accountData.sortCode)        bankSheet.getRange(existingRow, 5).setValue(accountData.sortCode);
      if (accountData.accountNumber)   bankSheet.getRange(existingRow, 6).setValue(accountData.accountNumber);
      bankSheet.getRange(existingRow, 10).setValue(accountData.active !== false);
      bankSheet.getRange(existingRow, 11).setValue(code);
    }
  } catch(e) {
    Logger.log('_syncCOABankAccount error: ' + e.toString());
  }
}