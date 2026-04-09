/**
 * NO~BULL BOOKS -- DATA IMPORT / MIGRATION
 * Handles CSV imports for contacts, invoices, bills and opening balances.
 * Each function validates, previews, then writes to the appropriate sheet.
 */

// -----------------------------------------------------------------------------
// CSV PARSER
// -----------------------------------------------------------------------------

function _parseCSV(csvText) {
  var lines  = csvText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  var result = [];
  lines.forEach(function(line) {
    if (!line.trim()) return;
    var row = [], cur = '', inQ = false;
    for (var i = 0; i < line.length; i++) {
      var ch = line[i];
      if (ch === '"') {
        if (inQ && line[i+1] === '"') { cur += '"'; i++; }
        else inQ = !inQ;
      } else if (ch === ',' && !inQ) {
        row.push(cur.trim()); cur = '';
      } else {
        cur += ch;
      }
    }
    row.push(cur.trim());
    result.push(row);
  });
  return result;
}

function _rowsToObjects(rows) {
  if (rows.length < 2) return [];
  var headers = rows[0].map(function(h){ return h.trim(); });
  return rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i){ obj[h] = (row[i] || '').trim(); });
    return obj;
  });
}

// -----------------------------------------------------------------------------
// IMPORT CONTACTS (Clients + Suppliers)
// -----------------------------------------------------------------------------

function importContacts(params) {
  try {
    _auth('clients.write', params);
    var csvText = params.csvData;
    if (!csvText) return { success: false, message: 'No CSV data provided.' };

    var rows    = _parseCSV(csvText);
    var records = _rowsToObjects(rows);
    if (!records.length) return { success: false, message: 'No records found in CSV.' };

    var ss          = getDb(params || {});
    var clientSheet = ss.getSheetByName(SHEETS.CLIENTS);
    var supSheet    = ss.getSheetByName(SHEETS.SUPPLIERS);
    var settings    = getSettings(params);

    var imported = 0, skipped = 0, errors = [];

    records.forEach(function(r, idx) {
      try {
        var type = (r['Type'] || r['type'] || r['ContactType'] || 'Client').toLowerCase();
        var name = r['Name'] || r['ContactName'] || r['CompanyName'] || r['name'] || '';
        var email = r['Email'] || r['EmailAddress'] || r['email'] || '';
        var phone = r['Phone'] || r['PhoneNumber'] || r['phone'] || '';
        var addr  = r['Address'] || r['Street'] || r['address'] || '';
        var city  = r['City'] || r['Town'] || r['city'] || '';
        var country = r['Country'] || r['country'] || 'UK';
        var vatNum  = r['VATNumber'] || r['TaxNumber'] || r['vatNumber'] || '';

        if (!name) { skipped++; return; }

        var id = generateId(type.indexOf('sup') >= 0 ? 'SUP' : 'CLI');

        if (type.indexOf('sup') >= 0) {
          if (supSheet) {
            supSheet.appendRow([
              id, name, email, phone,
              [addr, city].filter(Boolean).join(', '),
              '', country, vatNum, '', new Date(), 'Active', 'Imported'
            ]);
          }
        } else {
          if (clientSheet) {
            clientSheet.appendRow([
              id, name, email, phone,
              [addr, city].filter(Boolean).join(', '),
              '', country, vatNum, '', new Date(), 'Active', 'Imported'
            ]);
          }
        }
        imported++;
      } catch(rowErr) {
        errors.push('Row ' + (idx+2) + ': ' + rowErr.toString());
      }
    });

    logAudit('IMPORT', 'Contacts', 'CSV', { imported: imported, skipped: skipped }, params);
    return { success: true, imported: imported, skipped: skipped, errors: errors };
  } catch(e) {
    Logger.log('importContacts error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// -----------------------------------------------------------------------------
// IMPORT INVOICES (Outstanding)
// -----------------------------------------------------------------------------

function importInvoices(params) {
  try {
    _auth('invoices.write', params);
    var csvText = params.csvData;
    if (!csvText) return { success: false, message: 'No CSV data provided.' };

    var rows    = _parseCSV(csvText);
    var records = _rowsToObjects(rows);
    if (!records.length) return { success: false, message: 'No records found.' };

    var ss       = getDb(params || {});
    var sheet    = ss.getSheetByName(SHEETS.INVOICES);
    var settings = getSettings(params);
    if (!sheet) return { success: false, message: 'Invoices sheet not found.' };

    var imported = 0, skipped = 0, errors = [];

    records.forEach(function(r, idx) {
      try {
        var invNum    = r['InvoiceNumber'] || r['Reference'] || r['Invoice Number'] || '';
        var clientName = r['ClientName'] || r['ContactName'] || r['Customer'] || r['Contact'] || '';
        var issueDate  = r['IssueDate'] || r['Date'] || r['InvoiceDate'] || r['Invoice Date'] || '';
        var dueDate    = r['DueDate'] || r['Due Date'] || r['DueDate'] || '';
        var subtotal   = parseFloat(r['Subtotal'] || r['NetAmount'] || r['Net'] || r['Amount'] || 0);
        var vatAmount  = parseFloat(r['VATAmount'] || r['TaxAmount'] || r['VAT'] || r['Tax'] || 0);
        var total      = parseFloat(r['Total'] || r['GrossAmount'] || r['Gross'] || 0) || (subtotal + vatAmount);
        var amountDue  = parseFloat(r['AmountDue'] || r['Outstanding'] || r['Balance'] || total);
        var currency   = r['Currency'] || r['CurrencyCode'] || 'GBP';
        var status     = r['Status'] || 'Sent';

        if (!clientName && !invNum) { skipped++; return; }
        if (total <= 0 && subtotal <= 0) { skipped++; return; }

        var invId = generateId('INV');
        var paid  = total - amountDue;

        // Map status
        var mappedStatus = 'Sent';
        var s = status.toLowerCase();
        if (s.indexOf('paid') >= 0 || s.indexOf('complete') >= 0) mappedStatus = 'Paid';
        else if (s.indexOf('partial') >= 0) mappedStatus = 'Partial';
        else if (s.indexOf('overdue') >= 0) mappedStatus = 'Sent';
        else if (s.indexOf('draft') >= 0) mappedStatus = 'Draft';

        sheet.appendRow([
          invId,
          invNum || invId,
          '',                     // clientId (unknown)
          clientName,
          issueDate ? new Date(issueDate) : new Date(),
          dueDate   ? new Date(dueDate)   : '',
          '',                     // terms
          issueDate ? new Date(issueDate) : new Date(),
          subtotal || total,
          parseFloat(r['VATRate'] || 0),
          vatAmount,
          total,
          paid,
          amountDue,
          mappedStatus,
          '',                     // paymentDate
          r['Notes'] || r['Description'] || 'Imported from ' + (params.source || 'CSV'),
          false,                  // reconciled
          '', '', '',             // void fields
          currency,
          parseFloat(r['ExchangeRate'] || 1)
        ]);
        imported++;
      } catch(rowErr) {
        errors.push('Row ' + (idx+2) + ': ' + rowErr.toString());
      }
    });

    logAudit('IMPORT', 'Invoices', 'CSV', { imported: imported }, params);
    return { success: true, imported: imported, skipped: skipped, errors: errors };
  } catch(e) {
    Logger.log('importInvoices error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// -----------------------------------------------------------------------------
// IMPORT BILLS (Outstanding)
// -----------------------------------------------------------------------------

function importBills(params) {
  try {
    _auth('bills.write', params);
    var csvText = params.csvData;
    if (!csvText) return { success: false, message: 'No CSV data provided.' };

    var rows    = _parseCSV(csvText);
    var records = _rowsToObjects(rows);
    if (!records.length) return { success: false, message: 'No records found.' };

    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.BILLS);
    if (!sheet) return { success: false, message: 'Bills sheet not found.' };

    var imported = 0, skipped = 0, errors = [];

    records.forEach(function(r, idx) {
      try {
        var supplierName = r['SupplierName'] || r['ContactName'] || r['Vendor'] || r['Contact'] || '';
        var billNum      = r['BillNumber'] || r['Reference'] || r['Bill Number'] || r['InvoiceNumber'] || '';
        var issueDate    = r['IssueDate'] || r['Date'] || r['BillDate'] || '';
        var dueDate      = r['DueDate'] || r['Due Date'] || '';
        var subtotal     = parseFloat(r['Subtotal'] || r['NetAmount'] || r['Net'] || r['Amount'] || 0);
        var vatAmount    = parseFloat(r['VATAmount'] || r['TaxAmount'] || r['VAT'] || 0);
        var total        = parseFloat(r['Total'] || r['GrossAmount'] || 0) || (subtotal + vatAmount);
        var amountDue    = parseFloat(r['AmountDue'] || r['Outstanding'] || r['Balance'] || total);
        var currency     = r['Currency'] || 'GBP';

        if (!supplierName && !billNum) { skipped++; return; }
        if (total <= 0) { skipped++; return; }

        var billId = generateId('BILL');
        var paid   = total - amountDue;
        var status = paid >= total ? 'Paid' : paid > 0 ? 'Partial' : 'Approved';

        sheet.appendRow([
          billId,
          billNum || billId,
          '',                     // supplierId
          supplierName,
          issueDate ? new Date(issueDate) : new Date(),
          dueDate   ? new Date(dueDate)   : '',
          subtotal,
          parseFloat(r['VATRate'] || 0),
          vatAmount,
          total,
          paid,
          amountDue,
          status,
          paid >= total ? new Date() : '',
          r['Notes'] || 'Imported from ' + (params.source || 'CSV'),
          false,
          '', '', '',
          currency,
          parseFloat(r['ExchangeRate'] || 1)
        ]);
        imported++;
      } catch(rowErr) {
        errors.push('Row ' + (idx+2) + ': ' + rowErr.toString());
      }
    });

    logAudit('IMPORT', 'Bills', 'CSV', { imported: imported }, params);
    return { success: true, imported: imported, skipped: skipped, errors: errors };
  } catch(e) {
    Logger.log('importBills error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// -----------------------------------------------------------------------------
// IMPORT OPENING BALANCES
// -----------------------------------------------------------------------------

function importOpeningBalances(params) {
  try {
    _auth('settings.write', params);
    var csvText = params.csvData;
    var asAtDate = params.asAtDate || new Date().toISOString().split('T')[0];
    if (!csvText) return { success: false, message: 'No CSV data provided.' };

    var rows    = _parseCSV(csvText);
    var records = _rowsToObjects(rows);
    if (!records.length) return { success: false, message: 'No records found.' };

    var ss      = getDb(params || {});
    var txSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
    var coaSheet = ss.getSheetByName(SHEETS.CHART_OF_ACCOUNTS);
    if (!txSheet) return { success: false, message: 'Transactions sheet not found.' };

    // Build COA map for validation
    var coaMap = {};
    if (coaSheet && coaSheet.getLastRow() > 1) {
      var coaData = coaSheet.getDataRange().getValues();
      for (var i = 1; i < coaData.length; i++) {
        if (coaData[i][0]) coaMap[coaData[i][0].toString()] = {
          name: coaData[i][1] ? coaData[i][1].toString() : '',
          type: coaData[i][2] ? coaData[i][2].toString() : ''
        };
      }
    }

    var imported = 0, skipped = 0, errors = [], warnings = [];
    var obRef = 'OB-' + asAtDate;

    records.forEach(function(r, idx) {
      try {
        var accountCode = (r['AccountCode'] || r['Code'] || r['Account Code'] || '').toString().trim();
        var accountName = r['AccountName'] || r['Account'] || r['Name'] || '';
        var debit       = parseFloat(r['Debit']  || r['DR'] || 0) || 0;
        var credit      = parseFloat(r['Credit'] || r['CR'] || 0) || 0;

        if (!accountCode && !accountName) { skipped++; return; }
        if (debit === 0 && credit === 0)  { skipped++; return; }

        // Validate account code exists
        if (accountCode && !coaMap[accountCode]) {
          warnings.push('Row ' + (idx+2) + ': Account ' + accountCode + ' not in COA -- entry created anyway');
        }

        var amount = debit > 0 ? debit : credit;
        var txId   = generateId('OB');

        // Opening balance: debit = debit account, credit = retained earnings (3200)
        var debitAcc  = debit  > 0 ? accountCode : '3200';
        var creditAcc = credit > 0 ? accountCode : '3200';

        txSheet.appendRow([
          txId,
          safeSerializeDate(new Date(asAtDate)),
          'Opening Balance',
          obRef,
          debitAcc,
          creditAcc,
          amount,
          accountName || (coaMap[accountCode] ? coaMap[accountCode].name : accountCode),
          '', '',
          false
        ]);
        imported++;
      } catch(rowErr) {
        errors.push('Row ' + (idx+2) + ': ' + rowErr.toString());
      }
    });

    logAudit('IMPORT', 'OpeningBalances', obRef, { imported: imported, asAtDate: asAtDate }, params);
    return { success: true, imported: imported, skipped: skipped, errors: errors, warnings: warnings };
  } catch(e) {
    Logger.log('importOpeningBalances error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// -----------------------------------------------------------------------------
// PREVIEW (validate without writing)
// -----------------------------------------------------------------------------

function previewImport(params) {
  try {
    var type    = params.importType;
    var csvText = params.csvData;
    if (!csvText) return { success: false, message: 'No CSV data.' };

    var rows    = _parseCSV(csvText);
    if (rows.length < 2) return { success: false, message: 'CSV must have a header row and at least one data row.' };

    var headers = rows[0];
    var records = _rowsToObjects(rows);
    var preview = records.slice(0, 5); // First 5 rows for preview

    // Detect platform from headers
    var platform = 'Generic';
    var hStr = headers.join(',').toLowerCase();
    if (hStr.indexOf('xero') >= 0 || headers.indexOf('*ContactName') >= 0) platform = 'Xero';
    else if (hStr.indexOf('freeagent') >= 0) platform = 'FreeAgent';
    else if (hStr.indexOf('quickbooks') >= 0 || hStr.indexOf('intuit') >= 0) platform = 'QuickBooks';

    return {
      success:       true,
      headers:       headers,
      totalRows:     records.length,
      preview:       preview,
      platform:      platform,
      importType:    type
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// -----------------------------------------------------------------------------
// TEMPLATE DOWNLOAD (returns CSV template as text)
// -----------------------------------------------------------------------------

function getImportTemplate(params) {
  var type = params.templateType;
  var templates = {
    contacts: 'Type,Name,Email,Phone,Address,City,Country,VATNumber\nClient,Acme Ltd,info@acme.com,01234567890,123 High Street,London,UK,GB123456789\nSupplier,Office Supplies Co,orders@officesupplies.com,,,,,',
    invoices: 'InvoiceNumber,ClientName,IssueDate,DueDate,Subtotal,VATRate,VATAmount,Total,AmountDue,Currency,Status,Notes\nINV-001,Acme Ltd,2026-01-01,2026-01-31,1000.00,20,200.00,1200.00,1200.00,GBP,Sent,\nINV-002,Beta Corp,2026-01-15,2026-02-14,500.00,0,0,500.00,250.00,GBP,Partial,Partially paid',
    bills: 'BillNumber,SupplierName,IssueDate,DueDate,Subtotal,VATRate,VATAmount,Total,AmountDue,Currency,Notes\nBILL-001,Office Supplies Co,2026-01-05,2026-02-05,200.00,20,40.00,240.00,240.00,GBP,\nBILL-002,Broadband Provider,2026-01-01,2026-01-31,50.00,20,10.00,60.00,60.00,GBP,Monthly',
    openingBalances: 'AccountCode,AccountName,Debit,Credit\n1200,Bank Account,15000.00,\n1100,Accounts Receivable,5000.00,\n2100,Accounts Payable,,3000.00\n3200,Retained Earnings,,17000.00'
  };

  var csv = templates[type];
  if (!csv) return { success: false, message: 'Unknown template type: ' + type };
  return { success: true, csv: csv, filename: 'nobullbooks-' + type + '-template.csv' };
}
