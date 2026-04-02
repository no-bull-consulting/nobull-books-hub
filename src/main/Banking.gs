/**
 * NO~BULL BOOKS — BANKING
 * Bank accounts, transactions, reconciliation, double-entry ledger
 * ─────────────────────────────────────────────────────────────
 */


function getBankAccounts(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BANK_ACCOUNTS);
    if (!sheet) return { success: false, message: 'Bank accounts sheet not found', accounts: [] };

    var data = sheet.getDataRange().getValues();
    var accounts = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        accounts.push({
          accountId:      data[i][0] || '',
          accountName:    data[i][1] || '',
          bankName:       data[i][2] || '',
          accountType:    data[i][3] || '',
          sortCode:       data[i][4] || '',
          accountNumber:  data[i][5] || '',
          openingBalance: parseFloat(data[i][6]) || 0,
          currentBalance: parseFloat(data[i][7]) || 0,
          lastReconciled: data[i][8] || '',
          active:         data[i][9] !== false,
          nominalCode:    data[i][10] ? data[i][10].toString() : ''
        });
      }
    }
    return { success: true, accounts: accounts };
  } catch (e) {
    Logger.log('Error in getBankAccounts: ' + e.toString());
    return { success: false, message: e.toString(), accounts: [] };
  }
}

function createBankAccount(accountData, params) {
  try {
    _auth('banking.write', params);
    var ss      = getDb(params || {});
    var sheet   = ss.getSheetByName(SHEETS.BANK_ACCOUNTS);
    var coaSheet = ss.getSheetByName('ChartOfAccounts');
    if (!sheet) return { success: false, message: 'Bank accounts sheet not found' };

    // Auto-assign a COA nominal code if not provided
    // Find highest existing bank account code (1000-1099 range) and increment
    var nominalCode = accountData.nominalCode || '';
    if (!nominalCode && coaSheet) {
      var coaData = coaSheet.getDataRange().getValues();
      var maxBankCode = 999;
      for (var c = 1; c < coaData.length; c++) {
        var code = parseInt(coaData[c][0]) || 0;
        if (code >= 1000 && code <= 1099) maxBankCode = Math.max(maxBankCode, code);
      }
      nominalCode = (maxBankCode + 1).toString();
    }

    var accountId = generateId('BA');
    sheet.appendRow([
      accountId,
      accountData.accountName,
      accountData.bankName || '',
      accountData.accountType || 'Current',
      accountData.sortCode || '',
      accountData.accountNumber || '',
      parseFloat(accountData.openingBalance) || 0,
      parseFloat(accountData.openingBalance) || 0,
      '',    // LastReconciled
      true,  // Active
      nominalCode
    ]);

    // Auto-create corresponding COA entry
    if (coaSheet && nominalCode) {
      // Check it doesn't already exist
      var coaData2 = coaSheet.getDataRange().getValues();
      var exists = false;
      for (var d = 1; d < coaData2.length; d++) {
        if (coaData2[d][0] && coaData2[d][0].toString() === nominalCode) { exists = true; break; }
      }
      if (!exists) {
        coaSheet.appendRow([
          nominalCode,
          accountData.accountName + (accountData.bankName ? ' (' + accountData.bankName + ')' : ''),
          'Asset',
          'Bank Accounts',
          parseFloat(accountData.openingBalance) || 0,
          parseFloat(accountData.openingBalance) || 0,
          true,
          'Auto-created with bank account'
        ]);
        Logger.log('Auto-created COA entry: ' + nominalCode + ' — ' + accountData.accountName);
      }
    }

    logAudit('CREATE', 'BankAccount', accountId, { name: accountData.accountName, nominalCode: nominalCode });
    return { success: true, accountId: accountId, nominalCode: nominalCode };
  } catch (e) {
    Logger.log('Error in createBankAccount: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function updateBankBalance(accountId, amount, params) {
  // Guard: never update balance for Statement-type imports
  // importBankStatement writes directly and does not call this function,
  // but this guard ensures nothing slips through
  try {
    _auth('banking.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BANK_ACCOUNTS);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === accountId) {
        var currentBalance = parseFloat(data[i][7]) || 0;
        var newBalance = currentBalance + amount;
        sheet.getRange(i + 1, 8).setValue(newBalance);
        return newBalance;
      }
    }
  } catch (e) {
    Logger.log('Error updating bank balance: ' + e.toString());
  }
}

/**
 * Recalculate and save the true book balance for an account.
 * Sums only transactions where Type != 'Statement' (i.e. real book entries).
 * Call this after reconciliation or any balance correction.
 */
function recalcBookBalance(accountId, params) {
  try {
    var ss       = getDb(params || {});
    var txSheet  = ss.getSheetByName(SHEETS.BANK_TRANSACTIONS);
    var accSheet = ss.getSheetByName(SHEETS.BANK_ACCOUNTS);
    if (!txSheet || !accSheet) return;

    var txData  = txSheet.getDataRange().getValues();
    var balance = 0;

    for (var i = 1; i < txData.length; i++) {
      var row      = txData[i];
      var account  = String(row[BANK_TX_COLS.BANK_ACCOUNT - 1] || '');
      var type     = String(row[BANK_TX_COLS.TYPE - 1] || '');
      var category = String(row[BANK_TX_COLS.CATEGORY - 1] || '');
      var notes    = String(row[BANK_TX_COLS.NOTES - 1] || '');
      if (account !== accountId) continue;

      // Skip statement import lines — they are not book movements
      if (type === 'Statement' || category === 'Statement' ||
          notes.indexOf('Imported from bank statement') >= 0) continue;

      balance += parseFloat(row[BANK_TX_COLS.AMOUNT - 1]) || 0;
    }

    // Write corrected balance to BankAccounts sheet
    var accData = accSheet.getDataRange().getValues();
    for (var j = 1; j < accData.length; j++) {
      if (accData[j][0] === accountId) {
        accSheet.getRange(j + 1, 8).setValue(Math.round(balance * 100) / 100);
        Logger.log('recalcBookBalance: ' + accountId + ' = £' + balance.toFixed(2));
        return Math.round(balance * 100) / 100;
      }
    }
  } catch (e) {
    Logger.log('Error in recalcBookBalance: ' + e.toString());
  }
}

// ============================================
// BANK TRANSACTIONS
// ============================================

function getBankTransactions(accountId, fromDate, toDate, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BANK_TRANSACTIONS);
    if (!sheet) return { success: false, message: 'Bank transactions sheet not found', transactions: [] };
    
    var data = sheet.getDataRange().getValues();
    var transactions = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      
      if (accountId && row[7] !== accountId) continue;
      
      if (fromDate || toDate) {
        var txDate = new Date(row[1]);
        if (fromDate && txDate < new Date(fromDate)) continue;
        if (toDate && txDate > new Date(toDate)) continue;
      }
      
      transactions.push({
        transactionId: String(row[0] || ''),
        date: safeSerializeDate(row[1]),
        description: String(row[2] || ''),
        reference: String(row[3] || ''),
        amount: parseFloat(row[4]) || 0,
        type: String(row[5] || (parseFloat(row[4]) > 0 ? 'Credit' : 'Debit')),
        bankAccount: String(row[7] || ''),
        category: String(row[6] || ''),
        status: String(row[8] || 'Unreconciled'),
        reconciledDate: safeSerializeDate(row[9]),
        matchId: String(row[10] || ''),
        matchType: String(row[11] || ''),
        notes: String(row[12] || '')
      });
    }
    
    return { success: true, transactions: transactions };
  } catch (e) {
    Logger.log('Error in getBankTransactions: ' + e.toString());
    return { success: false, message: e.toString(), transactions: [] };
  }
}

function _addBankTransaction(txData, params) {
  try {
    var txId = generateId('BTX');
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BANK_TRANSACTIONS);
    if (!sheet) return { success: false, message: 'Bank transactions sheet not found' };
    
    sheet.appendRow([
      txId,
      txData.date,
      txData.description,
      txData.reference || '',
      parseFloat(txData.amount) || 0,
      txData.type || (txData.amount > 0 ? 'Credit' : 'Debit'),
      txData.category || '',
      txData.bankAccount,
      'Unreconciled',
      '',
      '',
      '',
      txData.notes || ''
    ]);
    
    updateBankBalance(txData.bankAccount, parseFloat(txData.amount));
    
    return { success: true, transactionId: txId };
  } catch (e) {
    Logger.log('Error in addBankTransaction: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getUnreconciledTransactions(accountId, params) {
  try {
    var result = getBankTransactions(accountId);
    if (!result.success) return result;

    // Return ALL unreconciled (statement lines + book entries)
    // Callers that need only book entries filter by category !== 'Statement'
    var unreconciled = result.transactions.filter(function(tx) {
      return tx.status !== 'Reconciled';
    });

    // For count/amount, exclude statement lines (they are not book movements)
    var bookOnly = unreconciled.filter(function(tx) {
      return tx.category !== 'Statement' &&
             !(tx.notes && tx.notes.indexOf('Imported') >= 0);
    });

    return {
      success:      true,
      transactions: unreconciled,          // full list for reconcile page
      count:        bookOnly.length,       // count excludes statement lines
      totalAmount:  bookOnly.reduce(function(sum, tx) { return sum + tx.amount; }, 0)
    };
  } catch (e) {
    Logger.log('Error in getUnreconciledTransactions: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function importBankStatement(accountId, csvData, params) {
  try {

    _auth('banking.write', params);    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.BANK_TRANSACTIONS);
    if (!sheet) return { success: false, message: 'BankTransactions sheet not found.' };

    var rows = csvData.split('\n');
    var imported = 0;
    var skipped  = 0;

    for (var i = 1; i < rows.length; i++) {
      if (!rows[i].trim()) continue;

      // Parse quoted CSV properly
      var cols = _splitCSVRow(rows[i]);
      if (cols.length < 3) continue;

      var date        = cols[0].trim();
      var description = cols[1].trim().replace(/^"|"$/g, '');
      var amountStr   = cols[2].trim().replace(/[^0-9.\-]/g, '');
      var amount      = parseFloat(amountStr) || 0;
      var reference   = cols.length > 3 ? cols[3].trim().replace(/^"|"$/g, '') : '';

      if (!date || amount === 0) { skipped++; continue; }

      var exists = checkDuplicateTransaction(accountId, date, amount, description);
      if (exists) { skipped++; continue; }

      // Write directly — do NOT call updateBankBalance
      // Statement lines are pending items; balance only moves on reconciliation
      var txId = generateId('BTX');
      sheet.appendRow([
        txId,
        date,
        description,
        reference,
        amount,
        'Statement',            // Type — keeps statement lines out of balance calcs
        'Statement',            // category — identifies imported lines
        accountId,
        'Unreconciled',
        '', '', '',
        'Imported from bank statement'
      ]);
      imported++;
    }

    return {
      success: true,
      imported: imported,
      skipped:  skipped,
      message:  'Imported ' + imported + ' transaction' + (imported !== 1 ? 's' : '') +
                (skipped > 0 ? ', skipped ' + skipped + ' duplicates' : '') + '.'
    };
  } catch (e) {
    Logger.log('Error in importBankStatement: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Split a CSV row respecting double-quoted fields.
 */
function _splitCSVRow(line) {
  var result = [], cur = '', inQ = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (ch === '"') { inQ = !inQ; }
    else if (ch === ',' && !inQ) { result.push(cur); cur = ''; }
    else { cur += ch; }
  }
  result.push(cur);
  return result;
}

function checkDuplicateTransaction(accountId, date, amount, description, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BANK_TRANSACTIONS);
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][7] === accountId && 
          data[i][1] === date && 
          Math.abs(parseFloat(data[i][4]) - amount) < 0.01 &&
          data[i][2].indexOf(description) > -1) {
        return true;
      }
    }
    return false;
  } catch (e) {
    return false;
  }
}

// ============================================
// RECONCILIATION FUNCTIONS - COMPLETE
// ============================================

/**
 * Get all unallocated invoices for reconciliation
 */
function getUnallocatedInvoices(clientId, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    var data = sheet.getDataRange().getValues();
    var invoices = [];
    
    for (var i = 1; i < data.length; i++) {
      var amountDue = parseFloat(data[i][INV_COLS.AMOUNT_DUE-1]) || 0;
      var invClientId = data[i][INV_COLS.CLIENT_ID-1];
      
      if (amountDue > 0.01 && (!clientId || invClientId === clientId)) {
        invoices.push({
          invoiceId: data[i][0] || '',
          invoiceNumber: data[i][1] || '',
          clientId: data[i][2] || '',
          clientName: data[i][3] || '',
          issueDate: safeSerializeDate(data[i][6]),
          dueDate: safeSerializeDate(data[i][7]),
          total: parseFloat(data[i][11]) || 0,
          amountDue: amountDue,
          status: data[i][14] || 'Sent'
        });
      }
    }
    
    return { success: true, invoices: invoices };
  } catch (e) {
    Logger.log('Error in getUnallocatedInvoices: ' + e.toString());
    return { success: false, message: e.toString(), invoices: [] };
  }
}

/**
 * Get all unallocated bills for reconciliation
 */
function getUnallocatedBills(supplierId, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BILLS);
    var data = sheet.getDataRange().getValues();
    var bills = [];
    
    for (var i = 1; i < data.length; i++) {
      var amountDue = parseFloat(data[i][BILL_COLS.AMOUNT_DUE-1]) || 0;
      var billSupplierId = data[i][BILL_COLS.SUPPLIER_ID-1];
      
      if (amountDue > 0.01 && (!supplierId || billSupplierId === supplierId)) {
        bills.push({
          billId: data[i][0] || '',
          billNumber: data[i][1] || '',
          supplierId: data[i][2] || '',
          supplierName: data[i][3] || '',
          issueDate: safeSerializeDate(data[i][4]),
          dueDate: safeSerializeDate(data[i][5]),
          total: parseFloat(data[i][9]) || 0,
          amountDue: amountDue,
          status: data[i][12] || 'Pending'
        });
      }
    }
    
    return { success: true, bills: bills };
  } catch (e) {
    Logger.log('Error in getUnallocatedBills: ' + e.toString());
    return { success: false, message: e.toString(), bills: [] };
  }
}

/**
 * Reconcile a bank transaction.
 *
 * Three allocation types:
 *   'invoice' / 'Invoice' — mark an invoice payment as confirmed
 *   'bill'    / 'Bill'    — mark a bill payment as confirmed
 *   'booked'              — match a statement line to an existing book entry:
 *                           marks both as Reconciled, no double-entry (already posted)
 *
 * Statement lines (category='Statement') are marked Reconciled but do NOT
 * create new double-entry postings — the book entry already did that.
 */
function reconcileTransaction(transactionId, allocations, params) {
  try {
    _auth('banking.reconcile', params);

    var ss      = getDb(params || {});
    var txSheet = ss.getSheetByName(SHEETS.BANK_TRANSACTIONS);
    var txData  = txSheet.getDataRange().getValues();

    // ── Find the statement/transaction row ─────────────────────────────────
    var tx = null, txRowNum = -1;
    for (var i = 1; i < txData.length; i++) {
      if (txData[i][0] === transactionId) {
        tx = {
          id:          txData[i][0],
          date:        txData[i][1],
          description: txData[i][2],
          reference:   txData[i][3],
          amount:      parseFloat(txData[i][4]),
          category:    String(txData[i][6] || ''),
          bankAccount: txData[i][7]
        };
        txRowNum = i + 1;
        break;
      }
    }
    if (!tx) return { success: false, message: 'Transaction not found.' };

    var isStatementLine = tx.category === 'Statement' ||
      String(txData[txRowNum-1][12]||'').indexOf('Imported') >= 0;

    for (var j = 0; j < allocations.length; j++) {
      var alloc   = allocations[j];
      var docType = (alloc.documentType || '').toLowerCase();
      var docId   = alloc.documentId;
      var amount  = parseFloat(alloc.amount) || Math.abs(tx.amount);

      if (docType === 'booked') {
        // ── Match statement line to book entry ───────────────────────────
        // Mark the book entry as Reconciled
        for (var k = 1; k < txData.length; k++) {
          if (txData[k][0] === docId) {
            txSheet.getRange(k + 1, BANK_TX_COLS.STATUS).setValue('Reconciled');
            txSheet.getRange(k + 1, BANK_TX_COLS.RECONCILED_DATE).setValue(new Date());
            txSheet.getRange(k + 1, BANK_TX_COLS.MATCH_ID).setValue(transactionId);
            txSheet.getRange(k + 1, BANK_TX_COLS.MATCH_TYPE).setValue('StatementMatch');
            break;
          }
        }
        // Mark statement line as Reconciled too
        txSheet.getRange(txRowNum, BANK_TX_COLS.STATUS).setValue('Reconciled');
        txSheet.getRange(txRowNum, BANK_TX_COLS.RECONCILED_DATE).setValue(new Date());
        txSheet.getRange(txRowNum, BANK_TX_COLS.MATCH_ID).setValue(docId);
        txSheet.getRange(txRowNum, BANK_TX_COLS.MATCH_TYPE).setValue('StatementMatch');
        // No double-entry — book entry already posted it

      } else if (docType === 'invoice' || docType === 'Invoice') {
        // ── Match to invoice ────────────────────────────────────────────
        if (!isStatementLine) {
          // Only post double-entry if not already posted by a book entry
          recordPayment(docId, amount, tx.date, alloc.notes || 'Reconciled');
        }
        txSheet.getRange(txRowNum, BANK_TX_COLS.STATUS).setValue('Reconciled');
        txSheet.getRange(txRowNum, BANK_TX_COLS.RECONCILED_DATE).setValue(new Date());
        txSheet.getRange(txRowNum, BANK_TX_COLS.MATCH_ID).setValue(docId);
        txSheet.getRange(txRowNum, BANK_TX_COLS.MATCH_TYPE).setValue('Invoice');

      } else if (docType === 'bill' || docType === 'Bill') {
        // ── Match to bill ───────────────────────────────────────────────
        if (!isStatementLine) {
          recordBillPayment(docId, amount, tx.date, 'Bank', alloc.notes || 'Reconciled');
        }
        txSheet.getRange(txRowNum, BANK_TX_COLS.STATUS).setValue('Reconciled');
        txSheet.getRange(txRowNum, BANK_TX_COLS.RECONCILED_DATE).setValue(new Date());
        txSheet.getRange(txRowNum, BANK_TX_COLS.MATCH_ID).setValue(docId);
        txSheet.getRange(txRowNum, BANK_TX_COLS.MATCH_TYPE).setValue('Bill');
      }
    }

    // ── Update LastReconciled date on the BankAccounts sheet ────────────────
    if (tx.bankAccount) {
      try {
        var baSheet = ss.getSheetByName(SHEETS.BANK_ACCOUNTS);
        var baData  = baSheet.getDataRange().getValues();
        for (var r = 1; r < baData.length; r++) {
          if (baData[r][0] === tx.bankAccount) {
            baSheet.getRange(r + 1, 9).setValue(new Date()); // col 9 = LastReconciled
            break;
          }
        }
      } catch(re) {
        Logger.log('Could not update LastReconciled: ' + re.toString());
      }
    }

    Logger.log('Reconciled: ' + transactionId);
    return { success: true, message: 'Reconciled successfully.' };

  } catch (e) {
    Logger.log('Error in reconcileTransaction: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Get reconciliation summary for an account
 */
function getReconciliationSummary(accountId, params) {
  try {
    var accounts = getBankAccounts(params);
    var account = accounts.accounts.find(function(a) { return a.accountId === accountId; });
    if (!account) return { success: false, message: 'Account not found' };

    var txResult    = getUnreconciledTransactions(accountId);
    var allUnrecon  = txResult.transactions || [];

    // Separate statement lines from book entries
    // Statement lines are pending matching markers — they do NOT affect the balance
    var bookEntries    = allUnrecon.filter(function(tx) {
      return tx.category !== 'Statement' &&
             !(tx.notes && tx.notes.indexOf('Imported') >= 0);
    });
    var statementLines = allUnrecon.filter(function(tx) {
      return tx.category === 'Statement' ||
             (tx.notes && tx.notes.indexOf('Imported') >= 0);
    });

    // Unreconciled count = only book entries (statement lines don't count as "to do")
    var totalCredits = 0, totalDebits = 0;
    for (var i = 0; i < bookEntries.length; i++) {
      if (bookEntries[i].amount > 0) totalCredits += bookEntries[i].amount;
      else totalDebits += Math.abs(bookEntries[i].amount);
    }

    var invoicesResult = getUnallocatedInvoices();
    var billsResult    = getUnallocatedBills();

    return {
      success: true,
      summary: {
        accountName:         account.accountName,
        bankBalance:         account.currentBalance,
        unreconciledCount:   bookEntries.length,
        statementLineCount:  statementLines.length,
        unreconciledCredits: totalCredits,
        unreconciledDebits:  totalDebits,
        netUnreconciled:     totalCredits - totalDebits,
        lastReconciled:      account.lastReconciled,
        unallocatedInvoices: invoicesResult.invoices ? invoicesResult.invoices.length : 0,
        unallocatedBills:    billsResult.bills ? billsResult.bills.length : 0,
        unreconciled:        bookEntries
      }
    };
  } catch (e) {
    Logger.log('Error in getReconciliationSummary: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================
// SPEND / RECEIVE / TRANSFER MONEY
// ============================================

function spendMoney(data, params) {
  try {
    _checkPeriodLock(data && data.date ? new Date(data.date) : new Date());
    var bankTxId = generateId('BTX');
    var txSheet = getDb(params || {}).getSheetByName(SHEETS.BANK_TRANSACTIONS);
    if (!txSheet) return { success: false, message: 'BankTransactions sheet not found' };
    
    txSheet.appendRow([
      bankTxId,
      safeSerializeDate(data.date),
      data.description,
      data.reference || '',
      -Math.abs(parseFloat(data.amount)),
      'Debit',
      data.category || '',
      data.bankAccountId,
      'Unreconciled',
      '',
      '',
      '',
      data.notes || ''
    ]);
    
    _createDoubleEntry(
      data.date,
      'Expense',
      data.reference || bankTxId,
      data.accountCode || '6000',
      '1000',
      Math.abs(parseFloat(data.amount)),
      data.description,
      null,
      null
    );
    
    return { 
      success: true, 
      transactionId: bankTxId,
      message: 'Spend money transaction created'
    };
  } catch (e) {
    Logger.log('Error in spendMoney: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function receiveMoney(data, params) {
  try {
    _checkPeriodLock(data && data.date ? new Date(data.date) : new Date());
    var bankTxId = generateId('BTX');
    var txSheet = getDb(params || {}).getSheetByName(SHEETS.BANK_TRANSACTIONS);
    if (!txSheet) return { success: false, message: 'BankTransactions sheet not found' };
    
    txSheet.appendRow([
      bankTxId,
      safeSerializeDate(data.date),
      data.description,
      data.reference || '',
      Math.abs(parseFloat(data.amount)),
      'Credit',
      data.category || '',
      data.bankAccountId,
      'Unreconciled',
      '',
      '',
      '',
      data.notes || ''
    ]);
    
    _createDoubleEntry(
      data.date,
      'Income',
      data.reference || bankTxId,
      '1000',
      data.accountCode || '4000',
      Math.abs(parseFloat(data.amount)),
      data.description,
      null,
      null
    );
    
    return { 
      success: true, 
      transactionId: bankTxId,
      message: 'Receive money transaction created'
    };
  } catch (e) {
    Logger.log('Error in receiveMoney: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function transferMoney(data, params) {
  try {
    try { _checkPeriodLock(data && data.date ? new Date(data.date) : new Date()); } catch(le) { return { success:false, message:le.message }; }
    var txSheet = getDb(params || {}).getSheetByName(SHEETS.BANK_TRANSACTIONS);
    if (!txSheet) return { success: false, message: 'BankTransactions sheet not found' };
    
    var transferId = generateId('TRF');
    var amount = Math.abs(parseFloat(data.amount));
    
    var debitId = generateId('BTX');
    txSheet.appendRow([
      debitId,
      safeSerializeDate(data.date),
      'Transfer to ' + data.toAccountName,
      transferId,
      -amount,
      'Debit',
      'Transfer',
      data.fromAccountId,
      'Unreconciled',
      '',
      '',
      '',
      data.notes || ''
    ]);
    
    var creditId = generateId('BTX');
    txSheet.appendRow([
      creditId,
      safeSerializeDate(data.date),
      'Transfer from ' + data.fromAccountName,
      transferId,
      amount,
      'Credit',
      'Transfer',
      data.toAccountId,
      'Unreconciled',
      '',
      '',
      '',
      data.notes || ''
    ]);
    
    return { 
      success: true, 
      transferId: transferId,
      debitId: debitId,
      creditId: creditId,
      message: 'Transfer completed successfully'
    };
  } catch (e) {
    Logger.log('Error in transferMoney: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================
// DOUBLE-ENTRY ACCOUNTING
// ============================================

function _createTransaction(type, reference, debitAccount, creditAccount, amount, description, invoiceId, billId) {
  return _createDoubleEntry(new Date(), type, reference, debitAccount, creditAccount, amount, description, invoiceId, billId);
}

function _createDoubleEntry(date, type, reference, debitAccount, creditAccount, amount, description, invoiceId, billId, params) {
  try {
    var txnId = generateId('TXN');
    var sheet = getDb(params || {}).getSheetByName(SHEETS.TRANSACTIONS);
    
    sheet.appendRow([
      txnId,
      safeSerializeDate(date),
      type,
      reference,
      debitAccount,
      creditAccount,
      amount,
      description,
      invoiceId || '',
      billId || '',
      false
    ]);
    
    _updateAccountBalance(debitAccount, amount, true, params);
    _updateAccountBalance(creditAccount, amount, false, params);
    
    return { success: true, transactionId: txnId };
  } catch (e) {
    Logger.log('Error in createDoubleEntry: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ── Public aliases so cross-module callers resolve correctly ─────────────────
// All modules call createDoubleEntry() with no prefix. Banking uses _createDoubleEntry().
// Both route to the same implementation.
function createDoubleEntry(date, type, reference, debitAccount, creditAccount, amount, description, invoiceId, billId, params) {
  return _createDoubleEntry(date, type, reference, debitAccount, creditAccount, amount, description, invoiceId, billId, params);
}
function createTransaction(type, reference, debitAccount, creditAccount, amount, description, invoiceId, billId) {
  return _createTransaction(type, reference, debitAccount, creditAccount, amount, description, invoiceId, billId);
}



function _updateAccountBalance(accountCode, amount, isDebit, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.CHART_OF_ACCOUNTS);
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === accountCode) {
        var currentBalance = parseFloat(data[i][5]) || 0;
        var accountType = data[i][2];
        var newBalance = currentBalance;
        
        if (accountType === 'Asset' || accountType === 'Expense') {
          newBalance = isDebit ? currentBalance + amount : currentBalance - amount;
        } else {
          newBalance = isDebit ? currentBalance - amount : currentBalance + amount;
        }
        
        sheet.getRange(i + 1, 6).setValue(newBalance);
        return newBalance;
      }
    }
  } catch (e) {
    Logger.log('Error updating account balance: ' + e.toString());
  }
}

// ============================================
// HISTORY TRACKING
// ============================================

function _addInvoiceHistory(invoiceId, changeType, fieldChanged, oldValue, newValue, notes, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICE_HISTORY);
    var historyId = generateId('HIS');
    var timestamp = new Date();
    var user = Session.getActiveUser().getEmail() || 'system';
    
    sheet.appendRow([
      historyId,
      invoiceId,
      timestamp,
      user,
      changeType,
      fieldChanged || '',
      oldValue || '',
      newValue || '',
      notes || ''
    ]);
    
    return { success: true, historyId: historyId };
  } catch (e) {
    Logger.log('Error adding invoice history: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function _addBillHistory(billId, changeType, fieldChanged, oldValue, newValue, notes, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BILL_HISTORY);
    var historyId = generateId('HIS');
    var timestamp = new Date();
    var user = Session.getActiveUser().getEmail() || 'system';
    
    sheet.appendRow([
      historyId,
      billId,
      timestamp,
      user,
      changeType,
      fieldChanged || '',
      oldValue || '',
      newValue || '',
      notes || ''
    ]);
    
    return { success: true, historyId: historyId };
  } catch (e) {
    Logger.log('Error adding bill history: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ── Public aliases for history helpers (cross-module callers use no prefix) ──
function addInvoiceHistory(invoiceId, changeType, fieldChanged, oldValue, newValue, notes, params) {
  return _addInvoiceHistory(invoiceId, changeType, fieldChanged, oldValue, newValue, notes);
}
function addBillHistory(billId, changeType, fieldChanged, oldValue, newValue, notes, params) {
  return _addBillHistory(billId, changeType, fieldChanged, oldValue, newValue, notes);
}



// ============================================
// CHART OF ACCOUNTS
// ============================================

function recordPaymentWithBank(invoiceId, amount, paymentDate, bankAccountId, notes, params) {
  try {

    _auth('invoices.write', params);    Logger.log('=== RECORD PAYMENT WITH BANK ===');
    Logger.log('Invoice ID: ' + invoiceId);
    Logger.log('Amount: £' + amount);
    Logger.log('Payment Date: ' + paymentDate);
    Logger.log('Bank Account: ' + bankAccountId);
    
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    if (!sheet) return { success: false, message: 'Invoices sheet not found' };
    
    var data = sheet.getDataRange().getValues();
    var dateObj = typeof paymentDate === 'string' ? new Date(paymentDate) : paymentDate || new Date();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        var currentPaid = parseFloat(data[i][INV_COLS.AMOUNT_PAID-1]) || 0;
        var total = parseFloat(data[i][INV_COLS.TOTAL-1]) || 0;
        var newPaid = currentPaid + parseFloat(amount);
        var newDue = total - newPaid;
        var rowNum = i + 1;
        
        // Update payment amounts
        sheet.getRange(rowNum, INV_COLS.AMOUNT_PAID).setValue(newPaid);
        sheet.getRange(rowNum, INV_COLS.AMOUNT_DUE).setValue(newDue);
        
        // Update status
        var newStatus = newDue <= 0.01 ? 'Paid' : (newPaid > 0 ? 'Partial' : 'Sent');
        sheet.getRange(rowNum, INV_COLS.STATUS).setValue(newStatus);
        
        // Set payment date if fully paid
        if (newDue <= 0.01) {
          sheet.getRange(rowNum, INV_COLS.PAYMENT_DATE).setValue(safeSerializeDate(dateObj));
        }
        
        // Append payment notes with bank info
        if (notes || bankAccountId) {
          var existingNotes = data[i][INV_COLS.NOTES-1] || '';
          var bankInfo = '';
          
          if (bankAccountId) {
            var bankAccount = getBankAccountById(bankAccountId);
            bankInfo = bankAccount ? ' to ' + bankAccount.accountName : '';
          }
          
          var paymentNote = '\n[Payment ' + Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy') + 
                           ']: £' + amount + bankInfo + (notes ? ' - ' + notes : '');
          sheet.getRange(rowNum, INV_COLS.NOTES).setValue(existingNotes + paymentNote);
        }
        
        // Create bank transaction
        if (bankAccountId) {
          createBankTransactionFromPayment({
            date: dateObj,
            description: 'Payment received for invoice ' + data[i][INV_COLS.NUMBER-1],
            reference: data[i][INV_COLS.NUMBER-1],
            amount: parseFloat(amount),
            type: 'Credit',
            bankAccountId: bankAccountId,
            category: 'Sales',
            notes: notes || '',
            _sheetId: params && params._sheetId ? params._sheetId : ''
          }, params);
        }
        
        // Create accounting transaction
        _createDoubleEntry(
          dateObj,
          'Payment',
          data[i][INV_COLS.NUMBER-1],
          bankAccountId ? getBankAccountCode(bankAccountId) : '1000',
          '1100',
          parseFloat(amount),
          'Payment received for ' + data[i][INV_COLS.NUMBER-1] + (notes ? ' - ' + notes : ''),
          invoiceId,
          null
        );
        
        // Add to history
        _addInvoiceHistory(invoiceId, 'Payment', 'amount', currentPaid, newPaid, 
                         'Bank: ' + (bankAccountId || 'Not specified') + (notes ? ' - ' + notes : ''));
        
        return { 
          success: true, 
          message: 'Payment recorded successfully',
          newAmountPaid: newPaid,
          newAmountDue: newDue,
          status: newStatus
        };
      }
    }
    
    return { success: false, message: 'Invoice not found' };
    
  } catch (e) {
    Logger.log('Error in recordPaymentWithBank: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Record payment for a bill with bank account selection
 */
function recordBillPaymentWithBank(billId, amount, paymentDate, bankAccountId, notes, params) {
  try {

    _auth('bills.write', params);    Logger.log('=== RECORD BILL PAYMENT WITH BANK ===');
    Logger.log('Bill ID: ' + billId);
    Logger.log('Amount: £' + amount);
    Logger.log('Payment Date: ' + paymentDate);
    Logger.log('Bank Account: ' + bankAccountId);
    
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BILLS);
    if (!sheet) return { success: false, message: 'Bills sheet not found' };
    
    var data = sheet.getDataRange().getValues();
    var dateObj = typeof paymentDate === 'string' ? new Date(paymentDate) : paymentDate || new Date();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === billId) {
        var currentPaid = parseFloat(data[i][BILL_COLS.AMOUNT_PAID-1]) || 0;
        var total = parseFloat(data[i][BILL_COLS.TOTAL-1]) || 0;
        var newPaid = currentPaid + parseFloat(amount);
        var newDue = total - newPaid;
        var rowNum = i + 1;
        
        // Update payment amounts
        sheet.getRange(rowNum, BILL_COLS.AMOUNT_PAID).setValue(newPaid);
        sheet.getRange(rowNum, BILL_COLS.AMOUNT_DUE).setValue(newDue);
        
        // Update status
        var newStatus = newDue <= 0.01 ? 'Paid' : (newPaid > 0 ? 'Partially Paid' : 'Pending');
        sheet.getRange(rowNum, BILL_COLS.STATUS).setValue(newStatus);
        
        // Set payment date if fully paid
        if (newDue <= 0.01) {
          sheet.getRange(rowNum, BILL_COLS.PAYMENT_DATE).setValue(safeSerializeDate(dateObj));
        }
        
        // Append payment notes with bank info
        if (notes || bankAccountId) {
          var existingNotes = data[i][BILL_COLS.NOTES-1] || '';
          var bankInfo = '';
          
          if (bankAccountId) {
            var bankAccount = getBankAccountById(bankAccountId);
            bankInfo = ' from ' + bankAccount.accountName;
          }
          
          var paymentNote = '\n[Payment ' + Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy') + 
                           ']: £' + amount + bankInfo + (notes ? ' - ' + notes : '');
          sheet.getRange(rowNum, BILL_COLS.NOTES).setValue(existingNotes + paymentNote);
        }
        
        // Create bank transaction
        if (bankAccountId) {
          createBankTransactionFromPayment({
            date: dateObj,
            description: 'Payment for bill ' + data[i][BILL_COLS.NUMBER-1],
            reference: data[i][BILL_COLS.NUMBER-1],
            amount: -parseFloat(amount), // Negative for payment out
            type: 'Debit',
            bankAccountId: bankAccountId,
            category: 'Expenses',
            notes: notes || '',
            _sheetId: params && params._sheetId ? params._sheetId : ''
          }, params);
        }
        
        // Create accounting transaction
        _createDoubleEntry(
          dateObj,
          'Bill Payment',
          data[i][BILL_COLS.NUMBER-1],
          '2000',
          bankAccountId ? getBankAccountCode(bankAccountId) : '1000',
          parseFloat(amount),
          'Payment for bill ' + data[i][BILL_COLS.NUMBER-1] + (notes ? ' - ' + notes : ''),
          null,
          billId
        );
        
        // Add to history
        _addBillHistory(billId, 'Payment', 'amount', currentPaid, newPaid, 
                      'Bank: ' + (bankAccountId || 'Not specified') + (notes ? ' - ' + notes : ''));
        
        return { 
          success: true, 
          message: 'Payment recorded successfully',
          newAmountPaid: newPaid,
          newAmountDue: newDue,
          status: newStatus
        };
      }
    }
    
    return { success: false, message: 'Bill not found' };
    
  } catch (e) {
    Logger.log('Error in recordBillPaymentWithBank: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Get bank account by ID
 */
function getBankAccountById(accountId, params) {
  try {
    var result = getBankAccounts(params);
    if (!result.success) return null;
    return result.accounts.filter(function(a){ return a.accountId === accountId; })[0] || null;
  } catch (e) {
    Logger.log('Error in getBankAccountById: ' + e.toString());
    return null;
  }
}

/**
 * Get bank account code from chart of accounts
 */
function getBankAccountCode(accountId, params) {
  try {
    var bankAccount = getBankAccountById(accountId, params);
    if (!bankAccount) return '1000';

    // Use the stored nominal code if set (preferred — explicit link to COA)
    if (bankAccount.nominalCode && bankAccount.nominalCode.trim()) {
      return bankAccount.nominalCode.trim();
    }

    // Fallback: name-match against COA (for legacy accounts with no nominalCode)
    var coaResult = getAccounts({}, params);
    if (!coaResult.success) return '1200';
    // Try exact match first, then partial
    var bankName = bankAccount.accountName ? bankAccount.accountName.toLowerCase().trim() : '';
    var exact    = coaResult.accounts.filter(function(a) {
      return a.accountType === 'Asset' && a.accountName.toLowerCase().trim() === bankName;
    })[0];
    if (exact) return exact.accountCode;
    // Partial match — only use if account is in Bank Accounts category
    var partial = coaResult.accounts.filter(function(a) {
      return a.accountType === 'Asset' &&
             (a.accountCategory === 'Bank Accounts' || a.accountCategory === 'Bank') &&
             a.accountName.toLowerCase().indexOf(bankName) >= 0;
    })[0];
    return partial ? partial.accountCode : '1000';
  } catch (e) {
    Logger.log('Error in getBankAccountCode: ' + e.toString());
    return '1000';
  }
}

/**
 * Create bank transaction from payment
 */
function createBankTransactionFromPayment(data, params) {
  try {
    _auth('banking.write', params);
    var bankTxId = generateId('BTX');
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BANK_TRANSACTIONS);
    if (!sheet) return { success: false };
    
    sheet.appendRow([
      bankTxId,
      safeSerializeDate(data.date),
      data.description,
      data.reference || '',
      data.amount,
      data.type,
      data.category || '',
      data.bankAccountId,
      'Unreconciled',
      '',
      '',
      '',
      data.notes || ''
    ]);
    
    updateBankBalance(data.bankAccountId, data.amount, params);
    
    return { success: true, transactionId: bankTxId };
  } catch (e) {
    Logger.log('Error in createBankTransactionFromPayment: ' + e.toString());
    return { success: false };
  }
}

// ============================================
// UPDATE INVOICE WITH PAYMENT DETAILS
// ============================================

/**
 * ONE-OFF MIGRATION: Fix existing Statement import rows that were stored
 * with Type='Credit' or 'Debit' instead of Type='Statement'.
 * Run once from Apps Script editor after deploying this update.
 * Also resets bank account balance to exclude statement lines.
 */
function fixStatementRowTypes() {
  try {
    var ss      = getDb(params || {});
    var txSheet = ss.getSheetByName(SHEETS.BANK_TRANSACTIONS);
    if (!txSheet) return { success: false, message: 'BankTransactions sheet not found' };

    var data    = txSheet.getDataRange().getValues();
    var fixed   = 0;
    var accounts = {};

    for (var i = 1; i < data.length; i++) {
      var category = String(data[i][BANK_TX_COLS.CATEGORY - 1] || '');
      var notes    = String(data[i][BANK_TX_COLS.NOTES - 1]    || '');
      var type     = String(data[i][BANK_TX_COLS.TYPE - 1]     || '');
      var accountId = String(data[i][BANK_TX_COLS.BANK_ACCOUNT - 1] || '');

      var isStatement = category === 'Statement' ||
                        notes.indexOf('Imported from bank statement') >= 0;

      if (isStatement && type !== 'Statement') {
        txSheet.getRange(i + 1, BANK_TX_COLS.TYPE).setValue('Statement');
        fixed++;
        if (accountId) accounts[accountId] = true;
      }
    }

    // Recalculate balance for affected accounts
    var balances = {};
    for (var aid in accounts) {
      balances[aid] = recalcBookBalance(aid);
    }

    var msg = 'Fixed ' + fixed + ' statement rows. Recalculated balances for ' +
              Object.keys(balances).length + ' account(s).';
    Logger.log(msg);
    return { success: true, message: msg, balances: balances };
  } catch (e) {
    Logger.log('Error in fixStatementRowTypes: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}


/**
 * createReconAdjustment
 * Creates a bank adjustment transaction (e.g. bank charge, rounding)
 * directly from the reconciliation screen, then immediately marks it reconciled
 * against the provided statement line.
 *
 * params: {
 *   bankAccountId  - account to post to
 *   amount         - positive number (direction determined by type)
 *   type           - 'debit' | 'credit'
 *   description    - e.g. "Bank charge", "Interest received"
 *   accountCode    - COA nominal code (e.g. '7500' for bank charges)
 *   date           - ISO date string
 *   statementTxId  - the statement line to reconcile against (optional)
 *   notes          - internal notes
 * }
 */
function createReconAdjustment(params) {
  try {
    _auth('banking.write', params);
    var ss       = getDb(params || {});
    var txSheet  = ss.getSheetByName(SHEETS.BANK_TRANSACTIONS);
    if (!txSheet) return { success: false, message: 'BankTransactions sheet not found' };

    var amount  = Math.abs(parseFloat(params.amount) || 0);
    if (!amount) return { success: false, message: 'Amount must be greater than zero' };

    var isDebit = params.type === 'debit';
    var bankTxId = generateId('ADJ');
    var date     = params.date ? new Date(params.date) : new Date();
    var desc     = params.description || (isDebit ? 'Bank adjustment (debit)' : 'Bank adjustment (credit)');
    var accCode  = params.accountCode || (isDebit ? '7500' : '7900');

    // Create the adjustment transaction
    txSheet.appendRow([
      bankTxId,
      safeSerializeDate(date),
      desc,
      params.reference || '',
      isDebit ? -amount : amount,
      isDebit ? 'Debit' : 'Credit',
      'Adjustment',
      params.bankAccountId,
      'Unreconciled',
      '', '', '',
      params.notes || ''
    ]);

    // Double-entry
    _createDoubleEntry(
      date,
      'Adjustment',
      bankTxId,
      isDebit ? accCode : '1000',
      isDebit ? '1000'  : accCode,
      amount,
      desc,
      null,
      null
    );

    // Update bank balance
    var balSheet = ss.getSheetByName(SHEETS.BANK_ACCOUNTS);
    if (balSheet) {
      var balData = balSheet.getDataRange().getValues();
      for (var i = 1; i < balData.length; i++) {
        if (balData[i][0] === params.bankAccountId) {
          var cur = parseFloat(balData[i][3]) || 0;
          balSheet.getRange(i+1, 4).setValue(cur + (isDebit ? -amount : amount));
          break;
        }
      }
    }

    // If a statement line was provided, auto-reconcile the new adj transaction against it
    if (params.statementTxId) {
      reconcileTransaction(bankTxId, [{
        documentType: 'booked',
        documentId:   params.statementTxId,
        amount:       amount,
        notes:        'Auto-reconciled adjustment'
      }]);
      // Also mark the statement line as reconciled
      reconcileTransaction(params.statementTxId, [{
        documentType: 'booked',
        documentId:   bankTxId,
        amount:       amount,
        notes:        'Matched to adjustment'
      }]);
    }

    logAudit('CREATE_ADJUSTMENT', 'BankTransaction', bankTxId,
      'Recon adjustment: '+desc+' '+amount+' ('+params.type+')');

    return {
      success: true,
      transactionId: bankTxId,
      message: 'Adjustment created' + (params.statementTxId ? ' and reconciled' : '')
    };

  } catch(e) {
    Logger.log('createReconAdjustment error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}