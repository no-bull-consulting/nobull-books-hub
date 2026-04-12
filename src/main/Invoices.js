/**
 * NO~BULL BOOKS — INVOICES (HMRC COMPLIANT v2.1)
 * Invoice creation, updates, payment recording, PDF, history
 * ─────────────────────────────────────────────────────────────
 */

/**
 * _parseLocalDate(dateStr)
 * Parses a date string (YYYY-MM-DD) as LOCAL time to avoid UTC timezone shift.
 */
function _parseLocalDate(dateStr) {
  if (!dateStr) return new Date();
  if (dateStr instanceof Date) return dateStr;
  var s = dateStr.toString().substring(0, 10);
  var parts = s.split('-');
  if (parts.length === 3) {
    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  }
  return new Date(dateStr);
}

function getAllInvoices(params) {
  try {
    _auth('invoices.read', params); // Enforced Auth check
    var ss = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.INVOICES);
    
    if (!sheet) return { success: false, message: 'Invoices sheet not found', invoices: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, invoices: [] };
    
    var invoices = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      
      invoices.push({
        invoiceId:     String(row[0]),
        invoiceNumber: String(row[1]),
        clientId:      String(row[2]),
        clientName:    String(row[3]),
        clientEmail:   String(row[4]),
        clientAddress: String(row[5]),
        issueDate:     safeSerializeDate(row[6]),
        dueDate:       safeSerializeDate(row[7]),
        subtotal:      parseFloat(row[8]) || 0,
        vatRate:       parseFloat(row[9]) || 0,
        vat:           parseFloat(row[10]) || 0,
        total:         parseFloat(row[11]) || 0,
        amountPaid:    parseFloat(row[12]) || 0,
        amountDue:     parseFloat(row[13]) || 0,
        status:        String(row[14] || 'Draft'),
        paymentDate:   safeSerializeDate(row[15]),
        notes:         String(row[16] || ''),
        pdfUrl:        String(row[17] || ''),
        currency:      String(row[18] || 'GBP'),
        exchangeRate:  parseFloat(row[19]) || 1,
        baseTotal:     parseFloat(row[20]) || 0
      });
    }
    return { success: true, invoices: invoices };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function createInvoice(clientId, lines, dueDate, notes, currency, exchangeRate, issueDateParam, params) {
  try {
    _auth('invoices.write', params);
    
    // HMRC COMPLIANCE: Sanitization & Type Casting 
    var safeClientId = String(clientId).substring(0, 50);
    var safeNotes    = String(notes || '').substring(0, 1000);
    var safeCurrency = /^[A-Z]{3}$/.test(currency) ? currency : 'GBP';
    var safeRate     = parseFloat(exchangeRate) || 1.0;
    var issueDate    = issueDateParam ? _parseLocalDate(issueDateParam) : new Date();

    // HMRC COMPLIANCE: Period Lock check
    _checkPeriodLock(issueDate, params); 
    
    var settings = getSettings(params);
    var ss = getDb(params || {});
    var invSheet = ss.getSheetByName(SHEETS.INVOICES);
    
    // Fetch client details with safety
    var clients = getAllClients(params).clients || [];
    var client  = clients.find(function(c) { return c.clientId === safeClientId; });
    if (!client) throw new Error("Client not found");

    var invoiceNumber = settings.invoicePrefix + settings.nextInvoiceNumber.toString().padStart(4, '0');
    var invoiceId     = generateId('INV');
    
    var clientTermsDays = parseInt(client.paymentTerms) || parseInt(settings.defaultPaymentTerms) || 30;
    var calculatedDueDate = dueDate ? _parseLocalDate(dueDate) : new Date(issueDate.getTime() + clientTermsDays*24*60*60*1000);
    
    var subtotal = 0;
    var vatTotal = 0;
    var validatedLines = [];

    // HMRC COMPLIANCE: Line item validation
    lines.forEach(function(l) {
      var qty   = parseFloat(l.quantity) || 1;
      var price = parseFloat(l.unitPrice) || 0;
      var rate  = parseFloat(l.vatRate) || 0;
      var net   = Math.round(qty * price * 100) / 100;
      var vat   = Math.round(net * (rate / 100) * 100) / 100;
      
      subtotal += net;
      vatTotal += vat;
      
      validatedLines.push([
        generateId('LINE'), invoiceId, 
        String(l.description).substring(0, 255), 
        qty, price, rate, net + vat, 
        String(l.accountCode || '4000')
      ]);
    });
    
    var total = subtotal + vatTotal;
    
    invSheet.appendRow([
      invoiceId, invoiceNumber, client.clientId, client.clientName, client.email,
      String(client.address || '') + ' ' + String(client.postcode || ''),
      safeSerializeDate(issueDate), safeSerializeDate(calculatedDueDate),
      subtotal, validatedLines[0][5], vatTotal, total,
      0, total, 'Draft', '', safeNotes, '',
      safeCurrency, safeRate, Math.round(total * safeRate * 100) / 100
    ]);
    
    var linesSheet = ss.getSheetByName(SHEETS.INVOICE_LINES);
    if (linesSheet && validatedLines.length > 0) {
      linesSheet.getRange(linesSheet.getLastRow() + 1, 1, validatedLines.length, 8).setValues(validatedLines);
    }
    
    settings.nextInvoiceNumber++;
    updateSettings(settings, params);
    
    createDoubleEntry(issueDate, 'Invoice', invoiceNumber, ACCOUNTS.DEBTORS, validatedLines[0][7], total, 'Invoice to ' + client.clientName, invoiceId, null, params);
    
    return { success: true, invoiceId: invoiceId, invoiceNumber: invoiceNumber };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function approveInvoice(invoiceId, params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    var data  = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        if (data[i][14] !== 'Draft') return { success: false, message: 'Only Draft invoices can be approved' };
        
        // Enforce lock check on approval date
        _checkPeriodLock(data[i][6], params); 
        
        sheet.getRange(i+1, INV_COLS.STATUS).setValue('Approved');
        addInvoiceHistory(invoiceId, 'StatusChange', 'status', 'Draft', 'Approved', 'Approved as Tax Invoice', params);
        return { success: true, message: 'Invoice approved' };
      }
    }
    return { success: false, message: 'Invoice not found' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function voidInvoice(invoiceId, reason, params) {
  try {
    _auth('invoices.write', params);
    var ss = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.INVOICES);
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        var status = String(data[i][14]);
        if (status === 'Void' || status === 'Paid') throw new Error("Cannot void " + status + " invoice");
        
        // Reverse Ledger Entries
        createDoubleEntry(new Date(), 'VoidReversal', String(data[i][1]), '4000', ACCOUNTS.DEBTORS, parseFloat(data[i][11]), "Void: " + reason, invoiceId, null, params);
        
        sheet.getRange(i + 1, INV_COLS.STATUS).setValue('Void');
        sheet.getRange(i + 1, INV_COLS.NOTES).setValue(String(data[i][16] || '') + '\n[Voided: ' + reason + ']');
        
        logAudit('VOID', 'Invoice', invoiceId, { reason: reason }, params);
        return { success: true, message: 'Invoice voided' };
      }
    }
    return { success: false, message: 'Invoice not found' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function _checkPeriodLock(date, params) {
  var s = getSettings(params);
  if (!s || !s.lockedBefore) return;
  var lockDate = new Date(s.lockedBefore);
  var txDate = new Date(date);
  if (txDate <= lockDate) {
    throw new Error('HMRC Compliance: This period is locked (closed on ' + s.lockedBefore + '). Records cannot be created or modified.');
  }
}