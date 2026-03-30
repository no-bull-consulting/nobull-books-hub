/**
 * NO~BULL BOOKS — SETTINGS
 * Read/write settings, defaults, logo upload.
 *
 * SECURITY NOTE (S2-01, S2-03):
 * HMRC credentials (clientID, clientSecret, accessToken, tokenExpiry)
 * are stored in ScriptProperties — NOT in the spreadsheet.
 * Settings sheet columns 24-27 are preserved as empty placeholders
 * for backward compatibility but are never read or written.
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ── PropertiesService keys for HMRC credentials ──────────────────────────────
var HMRC_PROP_KEYS = {
  CLIENT_ID:     'hmrc_client_id',
  CLIENT_SECRET: 'hmrc_client_secret',
  ACCESS_TOKEN:  'hmrc_access_token',
  TOKEN_EXPIRY:  'hmrc_token_expiry'
};

function _getHMRCProps() {
  var props = PropertiesService.getScriptProperties();
  return {
    hmrcClientID:     props.getProperty(HMRC_PROP_KEYS.CLIENT_ID)     || '',
    hmrcClientSecret: props.getProperty(HMRC_PROP_KEYS.CLIENT_SECRET) || '',
    hmrcAccessToken:  props.getProperty(HMRC_PROP_KEYS.ACCESS_TOKEN)  || '',
    hmrcTokenExpiry:  props.getProperty(HMRC_PROP_KEYS.TOKEN_EXPIRY)  || ''
  };
}

function _setHMRCProps(settings) {
  var props = PropertiesService.getScriptProperties();
  if (settings.hmrcClientID     !== undefined) props.setProperty(HMRC_PROP_KEYS.CLIENT_ID,     settings.hmrcClientID     || '');
  if (settings.hmrcClientSecret !== undefined) props.setProperty(HMRC_PROP_KEYS.CLIENT_SECRET, settings.hmrcClientSecret || '');
  if (settings.hmrcAccessToken  !== undefined) props.setProperty(HMRC_PROP_KEYS.ACCESS_TOKEN,  settings.hmrcAccessToken  || '');
  if (settings.hmrcTokenExpiry  !== undefined) props.setProperty(HMRC_PROP_KEYS.TOKEN_EXPIRY,  settings.hmrcTokenExpiry  || '');
}

// ─────────────────────────────────────────────────────────────────────────────
// BANK TRANSACTIONS  (lightweight fetch — shares auth context with Settings)
// ─────────────────────────────────────────────────────────────────────────────

function fetchBankTransactions(accountId, fromDate, toDate) {
  try {
    var sheet = _ss().getSheetByName(SHEETS.BANK_TRANSACTIONS);
    if (!sheet) return { success: true, transactions: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, transactions: [] };
    var txns = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      if (accountId && row[BANK_TX_COLS.BANK_ACCOUNT - 1] !== accountId) continue;
      txns.push({
        txId:           row[0].toString(),
        date:           safeSerializeDate(row[1]),
        description:    row[2] ? row[2].toString() : '',
        reference:      row[3] ? row[3].toString() : '',
        amount:         parseFloat(row[4]) || 0,
        type:           row[5] ? row[5].toString() : '',
        category:       row[6] ? row[6].toString() : '',
        bankAccount:    row[7] ? row[7].toString() : '',
        status:         row[8] ? row[8].toString() : '',
        reconciledDate: safeSerializeDate(row[9]),
        matchId:        row[10] ? row[10].toString() : '',
        matchType:      row[11] ? row[11].toString() : '',
        notes:          row[12] ? row[12].toString() : ''
      });
    }
    return { success: true, transactions: txns };
  } catch(e) {
    return { success: false, message: e.toString(), transactions: [] };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// BANK ACCOUNTS
// ─────────────────────────────────────────────────────────────────────────────

function getBankAccounts(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BANK_ACCOUNTS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, accounts: [] };
    var data     = sheet.getDataRange().getValues();
    var accounts = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      accounts.push({
        accountId:      row[0].toString(),
        accountName:    row[1] ? row[1].toString() : '',
        bankName:       row[2] ? row[2].toString() : '',
        sortCode:       row[3] ? row[3].toString() : '',
        accountNumber:  row[4] ? row[4].toString() : '',
        currency:       row[5] ? row[5].toString() : 'GBP',
        openingBalance: parseFloat(row[6]) || 0,
        currentBalance: parseFloat(row[7]) || 0,
        active:         row[8] !== false && row[8] !== 'FALSE',
        notes:          row[9] ? row[9].toString() : ''
      });
    }
    return { success: true, accounts: accounts };
  } catch(e) {
    return { success: false, message: e.toString(), accounts: [] };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// VOID LOG
// ─────────────────────────────────────────────────────────────────────────────

function getVoidLog(params) {
  try {
    var inv  = getDb(params || {});
    var bill = inv; // same spreadsheet

    var invSheet  = inv.getSheetByName(SHEETS.INVOICES);
    var billSheet = bill.getSheetByName(SHEETS.BILLS);

    function readVoided(sheet, numCol, nameCol, isInvoice) {
      if (!sheet || sheet.getLastRow() < 2) return [];
      var data = sheet.getDataRange().getValues();
      var out  = [];
      for (var i = 1; i < data.length; i++) {
        var row    = data[i];
        var status = row[14] ? row[14].toString() : '';
        if (status !== 'Void' && status !== 'Voided') continue;
        var obj = {
          total:      parseFloat(row[isInvoice ? 11 : 9]) || 0,
          voidDate:   safeSerializeDate(row[isInvoice ? 19 : 16]),
          voidReason: row[isInvoice ? 20 : 17] ? row[isInvoice ? 20 : 17].toString() : '',
          voidedBy:   row[isInvoice ? 21 : 18] ? row[isInvoice ? 21 : 18].toString() : ''
        };
        if (isInvoice) {
          obj.invoiceNumber = row[1] ? row[1].toString() : '';
          obj.clientName    = row[3] ? row[3].toString() : '';
        } else {
          obj.billNumber    = row[1] ? row[1].toString() : '';
          obj.supplierName  = row[3] ? row[3].toString() : '';
        }
        out.push(obj);
      }
      return out;
    }

    return {
      success:        true,
      voidedInvoices: readVoided(invSheet,  1, 3, true),
      voidedBills:    readVoided(billSheet, 1, 3, false)
    };
  } catch(e) {
    return { success: false, message: e.toString(), voidedInvoices: [], voidedBills: [] };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// BAD DEBTS
// ─────────────────────────────────────────────────────────────────────────────

function getBadDebts(params) {
  try {
    _auth('invoices.read', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BAD_DEBTS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, badDebts: [] };
    var data = sheet.getDataRange().getValues();
    var bds  = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      bds.push({
        badDebtId:        row[0].toString(),
        invoiceId:        row[1] ? row[1].toString() : '',
        invoiceNumber:    row[2] ? row[2].toString() : '',
        clientId:         row[3] ? row[3].toString() : '',
        clientName:       row[4] ? row[4].toString() : '',
        writeOffDate:     safeSerializeDate(row[5]),
        amountWrittenOff: parseFloat(row[6]) || 0,
        vatElement:       parseFloat(row[7]) || 0,
        vatReclaimStatus: row[8] ? row[8].toString() : '',
        vatClaimDate:     safeSerializeDate(row[9]),
        reason:           row[10] ? row[10].toString() : '',
        writtenOffBy:     row[11] ? row[11].toString() : ''
      });
    }
    return { success: true, badDebts: bds };
  } catch(e) {
    return { success: false, message: e.toString(), badDebts: [] };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// CREDIT NOTES
// ─────────────────────────────────────────────────────────────────────────────

function getCreditNotes(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.CREDIT_NOTES);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, creditNotes: [] };
    var data = sheet.getDataRange().getValues();
    var cns  = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      cns.push({
        cnId:             row[0].toString(),
        cnNumber:         row[1] ? row[1].toString() : '',
        invoiceId:        row[2] ? row[2].toString() : '',
        clientId:         row[3] ? row[3].toString() : '',
        clientName:       row[4] ? row[4].toString() : '',
        issueDate:        safeSerializeDate(row[5]),
        subtotal:         parseFloat(row[6]) || 0,
        vat:              parseFloat(row[7]) || 0,
        total:            parseFloat(row[8]) || 0,
        status:           row[9]  ? row[9].toString()  : 'Draft',
        reason:           row[10] ? row[10].toString() : '',
        appliedDate:      safeSerializeDate(row[11]),
        appliedInvoiceId: row[12] ? row[12].toString() : ''
      });
    }
    return { success: true, creditNotes: cns };
  } catch(e) {
    return { success: false, message: e.toString(), creditNotes: [] };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// INVOICES  (read — write functions live in Invoices.gs)
// ─────────────────────────────────────────────────────────────────────────────

function getAllInvoices(params) {
  try {
    _auth('invoices.read', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, invoices: [] };
    var data = sheet.getDataRange().getValues();
    var invs = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[INV_COLS.ID - 1]) continue;
      invs.push({
        invoiceId:     row[INV_COLS.ID - 1].toString(),
        invoiceNumber: row[INV_COLS.NUMBER - 1]      ? row[INV_COLS.NUMBER - 1].toString()      : '',
        clientId:      row[INV_COLS.CLIENT_ID - 1]   ? row[INV_COLS.CLIENT_ID - 1].toString()   : '',
        clientName:    row[INV_COLS.CLIENT_NAME - 1] ? row[INV_COLS.CLIENT_NAME - 1].toString() : '',
        clientEmail:   row[INV_COLS.CLIENT_EMAIL - 1]? row[INV_COLS.CLIENT_EMAIL - 1].toString(): '',
        clientAddress: row[INV_COLS.CLIENT_ADDR - 1] ? row[INV_COLS.CLIENT_ADDR - 1].toString() : '',
        issueDate:     safeSerializeDate(row[INV_COLS.ISSUE_DATE - 1]),
        dueDate:       safeSerializeDate(row[INV_COLS.DUE_DATE - 1]),
        subtotal:      parseFloat(row[INV_COLS.SUBTOTAL - 1])    || 0,
        vatRate:       parseFloat(row[INV_COLS.VAT_RATE - 1])    || 0,
        vatTotal:      parseFloat(row[INV_COLS.VAT - 1])         || 0,
        total:         parseFloat(row[INV_COLS.TOTAL - 1])       || 0,
        amountPaid:    parseFloat(row[INV_COLS.AMOUNT_PAID - 1]) || 0,
        amountDue:     parseFloat(row[INV_COLS.AMOUNT_DUE - 1])  || 0,
        status:        row[INV_COLS.STATUS - 1]       ? row[INV_COLS.STATUS - 1].toString()       : 'Draft',
        paymentDate:   safeSerializeDate(row[INV_COLS.PAYMENT_DATE - 1]),
        notes:         row[INV_COLS.NOTES - 1]        ? row[INV_COLS.NOTES - 1].toString()        : '',
        pdfUrl:        row[INV_COLS.PDF_URL - 1]      ? row[INV_COLS.PDF_URL - 1].toString()      : '',
        bankAccount:   row[INV_COLS.BANK_ACCT - 1]   ? row[INV_COLS.BANK_ACCT - 1].toString()    : '',
        voidDate:      safeSerializeDate(row[INV_COLS.VOID_DATE - 1]),
        voidReason:    row[INV_COLS.VOID_REASON - 1] ? row[INV_COLS.VOID_REASON - 1].toString()  : '',
        voidedBy:      row[INV_COLS.VOIDED_BY - 1]   ? row[INV_COLS.VOIDED_BY - 1].toString()    : ''
      });
    }
    return { success: true, invoices: invs };
  } catch(e) {
    return { success: false, message: e.toString(), invoices: [] };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// BILLS  (read — write functions live in Bills.gs)
// ─────────────────────────────────────────────────────────────────────────────

function getAllBills(params) {
  try {
    _auth('bills.read', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BILLS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, bills: [] };
    var data  = sheet.getDataRange().getValues();
    var bills = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[BILL_COLS.ID - 1]) continue;
      bills.push({
        billId:        row[BILL_COLS.ID - 1].toString(),
        billNumber:    row[BILL_COLS.NUMBER - 1]       ? row[BILL_COLS.NUMBER - 1].toString()       : '',
        supplierId:    row[BILL_COLS.SUPPLIER_ID - 1]  ? row[BILL_COLS.SUPPLIER_ID - 1].toString()  : '',
        supplierName:  row[BILL_COLS.SUPPLIER_NAME - 1]? row[BILL_COLS.SUPPLIER_NAME - 1].toString(): '',
        issueDate:     safeSerializeDate(row[BILL_COLS.ISSUE_DATE - 1]),
        dueDate:       safeSerializeDate(row[BILL_COLS.DUE_DATE - 1]),
        subtotal:      parseFloat(row[BILL_COLS.SUBTOTAL - 1])    || 0,
        vatRate:       parseFloat(row[BILL_COLS.VAT_RATE - 1])    || 0,
        vatTotal:      parseFloat(row[BILL_COLS.VAT - 1])         || 0,
        total:         parseFloat(row[BILL_COLS.TOTAL - 1])       || 0,
        amountPaid:    parseFloat(row[BILL_COLS.AMOUNT_PAID - 1]) || 0,
        amountDue:     parseFloat(row[BILL_COLS.AMOUNT_DUE - 1])  || 0,
        status:        row[BILL_COLS.STATUS - 1]       ? row[BILL_COLS.STATUS - 1].toString()       : 'Pending',
        paymentDate:   safeSerializeDate(row[BILL_COLS.PAYMENT_DATE - 1]),
        notes:         row[BILL_COLS.NOTES - 1]        ? row[BILL_COLS.NOTES - 1].toString()        : '',
        reconciled:    row[BILL_COLS.RECONCILED - 1] === true || row[BILL_COLS.RECONCILED - 1] === 'TRUE',
        voidDate:      safeSerializeDate(row[BILL_COLS.VOID_DATE - 1]),
        voidReason:    row[BILL_COLS.VOID_REASON - 1]  ? row[BILL_COLS.VOID_REASON - 1].toString()  : '',
        voidedBy:      row[BILL_COLS.VOIDED_BY - 1]    ? row[BILL_COLS.VOIDED_BY - 1].toString()    : ''
      });
    }
    return { success: true, bills: bills };
  } catch(e) {
    return { success: false, message: e.toString(), bills: [] };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// PURCHASE ORDERS  (read)
// ─────────────────────────────────────────────────────────────────────────────

function getPurchaseOrders(statusFilter, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.PURCHASE_ORDERS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, purchaseOrders: [] };
    var data = sheet.getDataRange().getValues();
    var pos  = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var status = row[9] ? row[9].toString() : 'Draft';
      if (statusFilter && status !== statusFilter) continue;
      pos.push({
        poId:             row[0].toString(),
        poNumber:         row[1]  ? row[1].toString()  : '',
        supplierId:       row[2]  ? row[2].toString()  : '',
        supplierName:     row[3]  ? row[3].toString()  : '',
        issueDate:        safeSerializeDate(row[4]),
        expectedDelivery: safeSerializeDate(row[5]),
        subtotal:         parseFloat(row[6]) || 0,
        vat:              parseFloat(row[7]) || 0,
        total:            parseFloat(row[8]) || 0,
        status:           status,
        notes:            row[10] ? row[10].toString() : '',
        approvedBy:       row[11] ? row[11].toString() : '',
        billId:           row[12] ? row[12].toString() : ''
      });
    }
    return { success: true, purchaseOrders: pos };
  } catch(e) {
    return { success: false, message: e.toString(), purchaseOrders: [] };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SETTINGS  read / write
// ─────────────────────────────────────────────────────────────────────────────

function getSettings(params) {
  try {
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.SETTINGS);

    if (!sheet || sheet.getLastRow() < 2) {
      return getDefaultSettings();
    }

    var sheetCols = sheet.getLastColumn();
    var readCols  = Math.max(sheetCols, 1);
    var rawData   = sheet.getRange(2, 1, 1, readCols).getValues()[0];
    var data      = rawData;
    while (data.length < 47) data.push('');

    var hmrcProps = _getHMRCProps();

    return {
      companyName:          data[0]  ? data[0].toString()            : '',
      companyAddress:       data[1]  ? data[1].toString()            : '',
      companyPostcode:      data[2]  ? data[2].toString()            : '',
      companyPhone:         data[3]  ? data[3].toString()            : '',
      companyEmail:         data[4]  ? data[4].toString()            : '',
      vatRegNumber:         data[5]  ? data[5].toString()            : '',
      invoicePrefix:        data[6]  ? data[6].toString()            : 'INV-',
      nextInvoiceNumber:    data[7]  ? parseInt(data[7])  || 1       : 1,
      billPrefix:           data[8]  ? data[8].toString()            : 'BILL-',
      nextBillNumber:       data[9]  ? parseInt(data[9])  || 1       : 1,
      logoURL:              data[10] ? data[10].toString()           : '',
      bankName:             data[11] ? data[11].toString()           : '',
      accountName:          data[12] ? data[12].toString()           : '',
      sortCode:             data[13] ? data[13].toString()           : '',
      accountNumber:        data[14] ? data[14].toString()           : '',
      financialYearStart:   safeSerializeDate(data[15]) || '2026-04-01',
      financialYearEnd:     safeSerializeDate(data[16]) || '2027-03-31',
      currentFinancialYear: data[17] ? data[17].toString()           : '2026/27',
      vatRegistered:        data[18] === true || data[18] === 'TRUE' || data[18] === 'true',
      vatScheme:            data[19] ? data[19].toString()           : 'standard',
      vatRate:              data[20] ? parseFloat(data[20]) || 20    : 20,
      vatFrequency:         data[21] ? data[21].toString()           : 'quarterly',
      mtdEnabled:           data[22] === true || data[22] === 'TRUE' || data[22] === 'true',
      // cols 23-26: HMRC credentials — stored in PropertiesService, not sheet
      hmrcTestMode:         data[26] === true || data[26] === 'TRUE' || data[26] === 'true',
      hmrcNINO:             data[28] ? data[28].toString()           : '',
      mtdBusinessId:        data[29] ? data[29].toString()           : '',
      cnPrefix:             data[30] ? data[30].toString()           : 'CN-',
      nextCNNumber:         data[31] ? parseInt(data[31]) || 1       : 1,
      poPrefix:             data[32] ? data[32].toString()           : 'PO-',
      nextPONumber:         data[33] ? parseInt(data[33]) || 1       : 1,
      lockedBefore:         (data.length>34&&data[34]) ? data[34].toString()      : '',
      emailSubject:         (data.length>35&&data[35]) ? data[35].toString()      : '',
      emailBody:            (data.length>36&&data[36]) ? data[36].toString()      : '',
      paymentTerms:         (data.length>37&&data[37]) ? parseInt(data[37])||30   : 30,
      invoiceFooter:        (data.length>38&&data[38]) ? data[38].toString()      : '',
      templateAccentColor:  (data.length>39&&data[39]) ? data[39].toString()      : '#1a3c6b',
      templateLogoPosition: (data.length>40&&data[40]) ? data[40].toString()      : 'left',
      templateShowReference:(data.length>41&&data[41]) === true || data[41] === 'TRUE',
      templateFont:         (data.length>42&&data[42]) ? data[42].toString()      : 'sans',
      baseCurrency:         (data.length>43&&data[43]) ? data[43].toString()      : 'GBP',
      enabledCurrencies:    (data.length>44&&data[44]) ? data[44].toString().split(',') : ['GBP','EUR','USD'],
      businessStartDate:    safeSerializeDate(data[45]) || '',
      yearEndDay:           (data.length>46&&data[46])
        ? (function(v){ var d=safeSerializeDate(v); return d || (typeof v==='string'?v:'31 March'); })(data[46])
        : '31 March',
      // HMRC credentials from PropertiesService
      hmrcClientID:         hmrcProps.hmrcClientID,
      hmrcClientSecret:     hmrcProps.hmrcClientSecret,
      hmrcAccessToken:      hmrcProps.hmrcAccessToken,
      hmrcTokenExpiry:      hmrcProps.hmrcTokenExpiry
    };
  } catch(e) {
    Logger.log('getSettings error: ' + e.toString());
    return getDefaultSettings();
  }
}

function getDefaultSettings() {
  return {
    companyName: '', companyAddress: '', companyPostcode: '',
    companyPhone: '', companyEmail: '', vatRegNumber: '',
    invoicePrefix: 'INV-', nextInvoiceNumber: 1,
    billPrefix: 'BILL-', nextBillNumber: 1,
    logoURL: '', bankName: '', accountName: '', sortCode: '', accountNumber: '',
    financialYearStart: '2026-04-01', financialYearEnd: '2027-03-31',
    currentFinancialYear: '2026/27', vatRegistered: false,
    vatScheme: 'standard', vatRate: 20, vatFrequency: 'quarterly',
    mtdEnabled: false, hmrcTestMode: true,
    hmrcClientID: '', hmrcClientSecret: '', hmrcAccessToken: '', hmrcTokenExpiry: '',
    hmrcNINO: '', mtdBusinessId: '',
    cnPrefix: 'CN-', nextCNNumber: 1, poPrefix: 'PO-', nextPONumber: 1,
    paymentTerms: 30, baseCurrency: 'GBP',
    enabledCurrencies: ['GBP','EUR','USD'],
    templateAccentColor: '#1a3c6b', templateLogoPosition: 'left',
    templateShowReference: true, templateFont: 'sans',
    yearEndDay: '31 March'
  };
}

/**
 * updateSettings(settings)
 *
 * Saves all settings to the client's Settings sheet.
 * settings must contain _sheetId so getDb() can find the right spreadsheet.
 */
function updateSettings(settings) {
  try {
    var ss = getDb(settings);
    if (!ss) return { success: false, message: 'Could not open spreadsheet. Check _sheetId.' };

    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return { success: false, message: 'Settings sheet not found — run initial setup first.' };

    // Store HMRC credentials in PropertiesService (not the sheet)
    _setHMRCProps(settings);

    // Merge with existing so partial saves don't wipe untouched fields
    var existing = getSettings(settings);
    settings = Object.assign({}, existing, settings);

    // Key aliases from frontend
    if (settings.defaultPaymentTerms !== undefined) settings.paymentTerms    = settings.defaultPaymentTerms;
    if (settings.businessStartDate)                 settings.financialYearStart = settings.businessStartDate;

    var data = [
      settings.companyName          || '',
      settings.companyAddress       || '',
      settings.companyPostcode      || '',
      settings.companyPhone         || '',
      settings.companyEmail         || '',
      settings.vatRegNumber         || '',
      settings.invoicePrefix        || 'INV-',
      settings.nextInvoiceNumber    || 1,
      settings.billPrefix           || 'BILL-',
      settings.nextBillNumber       || 1,
      settings.logoURL              || '',
      settings.bankName             || '',
      settings.accountName          || '',
      settings.sortCode             || '',
      settings.accountNumber        || '',
      settings.financialYearStart   || '2026-04-01',
      settings.financialYearEnd     || '2027-03-31',
      settings.currentFinancialYear || '2026/27',
      settings.vatRegistered        || false,
      settings.vatScheme            || 'standard',
      settings.vatRate              || 20,
      settings.vatFrequency         || 'quarterly',
      settings.mtdEnabled           || false,
      '',   // col 24: hmrcClientID     — intentionally blank (PropertiesService)
      '',   // col 25: hmrcClientSecret — intentionally blank
      '',   // col 26: hmrcAccessToken  — intentionally blank
      settings.hmrcTestMode         || true,
      '',   // col 28: hmrcTokenExpiry  — intentionally blank
      settings.hmrcNINO             || '',
      settings.mtdBusinessId        || '',
      settings.cnPrefix             || 'CN-',
      settings.nextCNNumber         || 1,
      settings.poPrefix             || 'PO-',
      settings.nextPONumber         || 1,
      settings.lockedBefore         || '',
      settings.emailSubject         || '',
      settings.emailBody            || '',
      settings.paymentTerms         || 30,
      settings.invoiceFooter        || '',
      settings.templateAccentColor  || '#1a3c6b',
      settings.templateLogoPosition || 'left',
      settings.templateShowReference !== false,
      settings.templateFont         || 'sans',
      settings.baseCurrency         || 'GBP',
      (Array.isArray(settings.enabledCurrencies)
        ? settings.enabledCurrencies.join(',')
        : settings.enabledCurrencies || 'GBP,EUR,USD'),
      settings.businessStartDate    || '',
      settings.yearEndDay           || '31 March'
    ];

    // Ensure sheet has 47 columns
    try {
      var lastCol = sheet.getLastColumn();
      if (lastCol < 47) sheet.insertColumnsAfter(lastCol, 47 - lastCol);
    } catch(extErr) {
      Logger.log('Could not extend Settings sheet: ' + extErr);
    }

    sheet.getRange(2, 1, 1, 47).setValues([data]);
    return { success: true, message: 'Settings saved.' };

  } catch(e) {
    Logger.log('updateSettings ERROR: ' + e.toString());
    return { success: false, message: 'Error saving settings: ' + e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// LOGO UPLOAD
// ─────────────────────────────────────────────────────────────────────────────

/**
 * uploadLogo(params)
 * params: { base64Data, fileName, fileType, _sheetId }
 */
function uploadLogo(params) {
  try {
    _auth('settings.write', params);

    var base64Data = params.base64Data;
    var fileName   = params.fileName   || 'logo';
    var fileType   = params.fileType   || 'image/png';

    var decoded = Utilities.newBlob(Utilities.base64Decode(base64Data), fileType, fileName);
    var folder  = _getOrCreateLogoFolder();

    // Remove existing logos
    var existing = folder.getFiles();
    while (existing.hasNext()) { existing.next().setTrashed(true); }

    var file   = folder.createFile(decoded);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    var imgUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400';

    // Save URL back to settings — pass _sheetId through
    var current       = getSettings(params);
    current.logoURL   = imgUrl;
    current.logoDriveId = fileId;
    current._sheetId  = params._sheetId; // ensure correct sheet is targeted
    var saveResult    = updateSettings(current);

    if (!saveResult.success) {
      return { success: false, message: 'Logo saved to Drive but settings update failed: ' + saveResult.message };
    }

    logAudit('UPDATE', 'Settings', 'logoURL', { fileName: fileName, fileId: fileId });
    return { success: true, url: imgUrl, fileId: fileId, message: 'Logo uploaded.' };

  } catch(e) {
    Logger.log('uploadLogo error: ' + e.toString());
    return { success: false, message: 'Error uploading logo: ' + e.toString() };
  }
}

function _getOrCreateLogoFolder() {
  var folderName = 'no~bull books — Logos';
  var folders    = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// ─────────────────────────────────────────────────────────────────────────────
// MISC / DEV HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function getCurrentUser() {
  return {
    success: true,
    email: Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail()
  };
}

/**
 * migrateHMRCCredentialsToProperties()
 * One-off: moves HMRC creds from the sheet into ScriptProperties.
 * Run once from the Apps Script editor.
 */
function migrateHMRCCredentialsToProperties() {
  try {
    var sheet = _ss().getSheetByName(SHEETS.SETTINGS);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Settings sheet not found.' };
    var data        = sheet.getRange(2, 1, 1, 27).getValues()[0];
    var clientId    = data[23] ? data[23].toString() : '';
    var clientSecret= data[24] ? data[24].toString() : '';
    var accessToken = data[25] ? data[25].toString() : '';
    var tokenExpiry = data[27] ? data[27].toString() : '';
    var moved = 0;
    var props = PropertiesService.getScriptProperties();
    if (clientId)     { props.setProperty(HMRC_PROP_KEYS.CLIENT_ID,     clientId);     moved++; }
    if (clientSecret) { props.setProperty(HMRC_PROP_KEYS.CLIENT_SECRET, clientSecret); moved++; }
    if (accessToken)  { props.setProperty(HMRC_PROP_KEYS.ACCESS_TOKEN,  accessToken);  moved++; }
    if (tokenExpiry)  { props.setProperty(HMRC_PROP_KEYS.TOKEN_EXPIRY,  tokenExpiry);  moved++; }
    if (moved > 0) sheet.getRange(2, 24, 1, 4).clearContent();
    return { success: true, message: 'Migrated ' + moved + ' credential(s). Sheet columns cleared.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function diagSettings() {
  var ss      = _ss();
  var sheet   = ss.getSheetByName(SHEETS.SETTINGS);
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  Logger.log('Settings: ' + lastRow + ' rows × ' + lastCol + ' cols');
  var raw = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  Logger.log('Col 16 (financialYearStart): ' + raw[15]);
  Logger.log('Col 46 (businessStartDate):  ' + (raw[45] !== undefined ? raw[45] : 'MISSING'));
  Logger.log('Col 47 (yearEndDay):         ' + (raw[46] !== undefined ? raw[46] : 'MISSING'));
  var s = getSettings();
  Logger.log('getSettings.financialYearStart: ' + s.financialYearStart);
  Logger.log('getSettings.businessStartDate:  ' + s.businessStartDate);
}
