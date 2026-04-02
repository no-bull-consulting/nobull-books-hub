/**
 * NO~BULL BOOKS — INVOICES
 * Invoice creation, updates, payment recording, PDF, history
 * ─────────────────────────────────────────────────────────────
 */

function getAllInvoices(params) {
  try {
    Logger.log('=== getAllInvoices(params) called ===');
    
    var ss = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.INVOICES);
    
    if (!sheet) {
      Logger.log('❌ Invoices sheet not found');
      return { success: false, message: 'Invoices sheet not found', invoices: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      return { success: true, invoices: [] };
    }
    
    var invoices = [];
    
    // Process each row (skip header row)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Skip empty rows
      if (!row[0] || row[0] === '') continue;
      
      try {
        // Handle the extra PDF URL column (index 17)
        var invoice = {
          invoiceId: row[0] ? row[0].toString() : '',
          invoiceNumber: row[1] ? row[1].toString() : '',
          clientId: row[2] ? row[2].toString() : '',
          clientName: row[3] ? row[3].toString() : '',
          clientEmail: row[4] ? row[4].toString() : '',
          clientAddress: row[5] ? row[5].toString() : '',
          issueDate: row[6] ? safeSerializeDate(row[6]) : '',
          dueDate: row[7] ? safeSerializeDate(row[7]) : '',
          subtotal: parseFloat(String(row[8]).replace(/[£,]/g, '')) || 0,
          vatRate: parseFloat(row[9]) || 0,
          vat: parseFloat(String(row[10]).replace(/[£,]/g, '')) || 0,
          total: parseFloat(String(row[11]).replace(/[£,]/g, '')) || 0,
          amountPaid: parseFloat(String(row[12]).replace(/[£,]/g, '')) || 0,
          amountDue: parseFloat(String(row[13]).replace(/[£,]/g, '')) || 0,
          status: row[14] ? row[14].toString() : 'Draft',
          paymentDate: row[15] ? safeSerializeDate(row[15]) : '',
          notes: row[16] ? row[16].toString() : '',
          pdfUrl: row[17] ? row[17].toString() : '', // Extra column
          // col 18=Currency, 19=ExchangeRate, 20=BaseTotal (written by createInvoice)
          bankAccountId: '',
          currency:     (function(v){ var s=v?v.toString().trim():''; return /^[A-Z]{3}$/.test(s)?s:'GBP'; })(row[18]),
          exchangeRate: (function(v){ var r=parseFloat(v); return (r&&r>0.001&&r<10000)?r:1; })(row[19]),
          baseTotal:    parseFloat(row[20]) || 0
        };
        
        invoices.push(invoice);
      } catch (e) {
        Logger.log('⚠️ Error processing row ' + i + ': ' + e.toString());
      }
    }
    
    Logger.log('✅ Successfully loaded ' + invoices.length + ' invoices');
    
    return { 
      success: true, 
      invoices: invoices,
      count: invoices.length
    };
    
  } catch (e) {
    Logger.log('❌ Error in getAllInvoices: ' + e.toString());
    return { success: false, message: e.toString(), invoices: [] };
  }
}

function getInvoiceById(invoiceId, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        return {
          invoiceId: data[i][0] || '',
          invoiceNumber: data[i][1] || '',
          clientId: data[i][2] || '',
          clientName: data[i][3] || '',
          clientEmail: data[i][4] || '',
          clientAddress: data[i][5] || '',
          issueDate: safeSerializeDate(data[i][6]),
          dueDate: safeSerializeDate(data[i][7]),
          subtotal: parseFloat(String(data[i][8]).replace(/[£,]/g, '')) || 0,
          vatRate: parseFloat(data[i][9]) || 0,
          vat: parseFloat(String(data[i][10]).replace(/[£,]/g, '')) || 0,
          total: parseFloat(String(data[i][11]).replace(/[£,]/g, '')) || 0,
          amountPaid: parseFloat(String(data[i][12]).replace(/[£,]/g, '')) || 0,
          amountDue: parseFloat(String(data[i][13]).replace(/[£,]/g, '')) || 0,
          status: data[i][14] || 'Draft',
          paymentDate: safeSerializeDate(data[i][15]),
          notes: data[i][16] || '',
          pdfUrl: data[i][17] || '',
          // col 18 = Currency (written by createInvoice)
          // col 19 = ExchangeRate
          // col 20 = BaseTotal
          // BankAccount is not stored in the Invoices sheet currently
          bankAccountId: '',
          currency:     (function(v) {
            var s = v ? v.toString().trim() : '';
            return /^[A-Z]{3}$/.test(s) ? s : 'GBP';
          })(data[i][18]),
          exchangeRate: (function(v) {
            var r = parseFloat(v);
            return (r && r > 0.001 && r < 10000) ? r : 1;
          })(data[i][19]),
          baseTotal:    parseFloat(data[i][20]) || 0
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log('Error in getInvoiceById: ' + e.toString());
    return null;
  }
}

function getInvoiceLines(invoiceId, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICE_LINES);
    var data = sheet.getDataRange().getValues();
    var lines = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === invoiceId) {
        lines.push({
          lineId: data[i][0],
          description: data[i][2],
          quantity: parseFloat(data[i][3]) || 0,
          unitPrice: parseFloat(data[i][4]) || 0,
          vatRate: parseFloat(data[i][5]) || 0,
          amount: parseFloat(data[i][6]) || 0,
          accountCode: data[i][7] || '4000'
        });
      }
    }
    return lines;
  } catch (e) {
    Logger.log('Error in getInvoiceLines: ' + e.toString());
    return [];
  }
}

function createInvoice(clientId, lines, dueDate, notes, currency, exchangeRate, params) {
  try {
    _auth('invoices.write', params);
    _checkPeriodLock(new Date(), params); // check today's date
    Logger.log('Creating invoice for client: ' + clientId);
    
    var settings = getSettings(params);
    var ss = getDb(params || {});
    var invSheet = ss.getSheetByName(SHEETS.INVOICES);
    var clientsSheet = ss.getSheetByName(SHEETS.CLIENTS);
    
    if (!invSheet) return { success: false, message: 'Invoices sheet not found' };
    if (!clientsSheet) return { success: false, message: 'Clients sheet not found' };
    
    var clientsData = clientsSheet.getDataRange().getValues();
    var client = null;
    
    for (var i = 1; i < clientsData.length; i++) {
      if (clientsData[i][0] === clientId) {
        client = {
          clientId:    clientsData[i][0],
          clientName:  clientsData[i][1],
          email:       clientsData[i][2],  // col 2 = Email
          phone:       clientsData[i][3],  // col 3 = Phone
          address:     clientsData[i][4],  // col 4 = Address
          postcode:    clientsData[i][5],  // col 5 = Postcode
          country:     clientsData[i][6]   // col 6 = Country
        };
        break;
      }
    }
    
    if (!client) return { success: false, message: 'Client not found' };
    
    // Currency
    var baseCurr  = settings.baseCurrency || 'GBP';
    var invCurr   = currency || baseCurr;
    var fxRate    = invCurr === baseCurr ? 1.0 : (parseFloat(exchangeRate)||1.0);
    if (invCurr !== baseCurr && (!exchangeRate || exchangeRate === 1)) {
      var rateResult = convertToBase(1, invCurr);
      fxRate = rateResult.success ? rateResult.rate : 1.0;
    }

    var invoiceNumber = settings.invoicePrefix + settings.nextInvoiceNumber.toString().padStart(4, '0');
    var invoiceId = generateId('INV');
    var issueDate = new Date();
    // Use client payment terms if no explicit due date provided
    var clientTermsDays = 30;
    for (var ci = 1; ci < clientsData.length; ci++) {
      if (clientsData[ci][0] === clientId) {
        clientTermsDays = parseInt(clientsData[ci][8]) || parseInt(settings.paymentTerms) || 30;
        break;
      }
    }
    var calculatedDueDate = dueDate ? new Date(dueDate) : new Date(issueDate.getTime() + clientTermsDays*24*60*60*1000);
    
    var subtotal = 0;
    var vatTotal = 0;
    
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      var amount = parseFloat(line.quantity) * parseFloat(line.unitPrice);
      var vat = amount * (parseFloat(line.vatRate) / 100);
      subtotal += amount;
      vatTotal += vat;
    }
    
    var total = subtotal + vatTotal;
    
    // Append row with 18 columns (including PDF URL at the end)
    invSheet.appendRow([
      invoiceId,                                    // A: InvoiceId
      invoiceNumber,                                // B: InvoiceNumber
      client.clientId,                              // C: ClientId
      client.clientName,                            // D: ClientName
      client.email,                                 // E: ClientEmail
      (client.address || '') + (client.postcode ? ', ' + client.postcode : '') + (client.country && client.country !== 'UK' ? ', ' + client.country : ''),  // F: ClientAddress
      safeSerializeDate(issueDate),             // G: IssueDate
      safeSerializeDate(calculatedDueDate),         // H: DueDate
      subtotal,                                     // I: Subtotal
      lines[0].vatRate || 0,                        // J: VATRate
      vatTotal,                                     // K: VAT
      total,                                        // L: Total
      0,                                            // M: AmountPaid
      total,                                        // N: AmountDue
      'Draft',                                      // O: Status
      '',                                           // P: PaymentDate
      notes || '',                                  // Q: Notes
      '',                                           // R: PDF URL (new column)
      invCurr,                                      // S: Currency
      fxRate,                                       // T: ExchangeRate (foreign per base, or 1 if same)
      Math.round(total * fxRate * 100) / 100        // U: BaseTotal (in base currency)
    ]);
    
    // Save invoice lines
    var linesSheet = ss.getSheetByName(SHEETS.INVOICE_LINES);
    if (linesSheet) {
      for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        var lineId = generateId('LINE');
        var amount = parseFloat(line.quantity) * parseFloat(line.unitPrice);
        
        linesSheet.appendRow([
          lineId,
          invoiceId,
          line.description,
          parseFloat(line.quantity),
          parseFloat(line.unitPrice),
          parseFloat(line.vatRate) || 0,
          amount,
          line.accountCode || '4000'
        ]);
      }
    }
    
    settings.nextInvoiceNumber++;
    updateSettings(settings, params);
    
    // Create transaction
    createDoubleEntry(
      issueDate,
      'Invoice',
      invoiceNumber,
      '1100',
      lines[0].accountCode || '4000',
      total,
      'Invoice to ' + client.clientName,
      invoiceId,
      null
    );
    
    return { success: true, invoiceId: invoiceId, invoiceNumber: invoiceNumber };
    
  } catch (e) {
    Logger.log('Error in createInvoice: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function updateInvoiceStatus(invoiceId, newStatus, params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        sheet.getRange(i + 1, INV_COLS.STATUS).setValue(newStatus);
        addInvoiceHistory(invoiceId, 'StatusChange', 'status', data[i][INV_COLS.STATUS-1], newStatus, '');
        return { success: true, message: 'Status updated' };
      }
    }
    return { success: false, message: 'Invoice not found' };
  } catch (e) {
    Logger.log('Error in updateInvoiceStatus: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * approveInvoice(invoiceId)
 * Transitions invoice from Draft (Pro-Forma) → Approved (Tax Invoice).
 * From this point the invoice is included in VAT returns.
 */
/**
 * editInvoice(invoiceId, updates)
 * Updates editable invoice fields based on current status rules:
 *   Draft     → all fields editable (client, dates, lines, reference, notes)
 *   Approved  → dueDate, reference, notes only
 *   Sent      → dueDate, notes only
 *   Paid/Void → read-only, returns error
 */
/**
 * deleteInvoice(invoiceId)
 * Hard-deletes a DRAFT invoice and its lines.
 * Only Draft invoices can be deleted — Approved+ must be Voided instead.
 */
function deleteInvoice(invoiceId, params) {
  try {
    _auth('invoices.write', params);
    var ss      = getDb(params || {});
    var sheet   = ss.getSheetByName(SHEETS.INVOICES);
    var data    = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        var status = (data[i][14]||'Draft').toString();
        if (status !== 'Draft') {
          return { success:false, message:'Only Draft invoices can be deleted. Use Void for '+status+' invoices.' };
        }
        sheet.deleteRow(i + 1);
        // Delete lines
        var lineSheet = ss.getSheetByName(SHEETS.INVOICE_LINES);
        if (lineSheet) {
          var lineData = lineSheet.getDataRange().getValues();
          for (var j = lineData.length - 1; j >= 1; j--) {
            if (lineData[j][1] === invoiceId) lineSheet.deleteRow(j + 1);
          }
        }
        logAudit('DELETE', 'Invoice', invoiceId, { invoiceNumber: data[i][1] });
        return { success:true, message:'Invoice deleted' };
      }
    }
    return { success:false, message:'Invoice not found' };
  } catch(e) {
    Logger.log('deleteInvoice error: ' + e);
    return { success:false, message:e.toString() };
  }
}

function editInvoice(invoiceId, updates, params) {
  try {
    _auth('invoices.write', params);
    // Prevent edits if the invoice date falls in a locked period
    try { _checkPeriodLock(updates.issueDate ? new Date(updates.issueDate) : new Date(), params); } catch(le) { return { success:false, message:le.message }; }
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.INVOICES);
    if (!sheet) return { success:false, message:'Invoices sheet not found' };

    var data = sheet.getDataRange().getValues();
    var rowNum = -1, row = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) { rowNum = i+1; row = data[i]; break; }
    }
    if (!row) return { success:false, message:'Invoice not found' };

    var status = (row[14]||'Draft').toString();
    var locked = ['Paid','Void','Voided','Bad Debt'];
    if (locked.indexOf(status) >= 0) {
      return { success:false, message:'Cannot edit a '+status+' invoice.' };
    }

    // What is allowed at each status
    var canEditDates    = status === 'Draft' || status === 'Approved' || status === 'Sent';
    var canEditClient   = status === 'Draft';
    var canEditLines    = status === 'Draft';
    var canEditRef      = status === 'Draft' || status === 'Approved';
    var canEditNotes    = true; // always (internal field)
    var canForceStatus  = status === 'Draft' || status === 'Approved' || status === 'Sent';

    var changes = [];

    // Client (Draft only)
    if (canEditClient && updates.clientId) {
      var clients = getAllClients(params).clients || [];
      var client  = clients.filter(function(c){ return c.clientId === updates.clientId; })[0];
      if (client) {
        sheet.getRange(rowNum, 3).setValue(client.clientId);
        sheet.getRange(rowNum, 4).setValue(client.clientName);
        changes.push('client');
      }
    }

    // Issue date (Draft only)
    if (canEditLines && updates.issueDate) {
      sheet.getRange(rowNum, 7).setValue(updates.issueDate);
      changes.push('issueDate');
    }

    // Due date
    if (canEditDates && updates.dueDate) {
      sheet.getRange(rowNum, 8).setValue(updates.dueDate);
      changes.push('dueDate');
    }

    // Reference (Draft + Approved)
    if (canEditRef && updates.reference !== undefined) {
      sheet.getRange(rowNum, 6).setValue(updates.reference || '');
      changes.push('reference');
    }

    // Notes (always)
    if (canEditNotes && updates.notes !== undefined) {
      sheet.getRange(rowNum, 17).setValue(updates.notes || '');
      changes.push('notes');
    }

    // Force status transition (e.g. revert Approved → Draft for correction)
    if (canForceStatus && updates.forceStatus) {
      var allowed = { 'Approved':['Draft'], 'Sent':['Approved','Draft'] };
      var permitted = allowed[status] || [];
      if (permitted.indexOf(updates.forceStatus) >= 0) {
        sheet.getRange(rowNum, 15).setValue(updates.forceStatus);
        changes.push('status:'+status+'→'+updates.forceStatus);
      } else {
        return { success:false, message:'Cannot revert from '+status+' to '+updates.forceStatus };
      }
    }

    // Line items (Draft only) — rebuild lines
    if (canEditLines && updates.lines && updates.lines.length) {
      var lineSheet = ss.getSheetByName(SHEETS.INVOICE_LINES);
      if (lineSheet) {
        // Delete existing lines for this invoice
        var lineData = lineSheet.getDataRange().getValues();
        var rowsToDelete = [];
        for (var j = lineData.length-1; j >= 1; j--) {
          if (lineData[j][1] === invoiceId) rowsToDelete.push(j+1);
        }
        rowsToDelete.forEach(function(r){ lineSheet.deleteRow(r); });

        // Re-insert new lines
        var subtotal = 0, vatTotal = 0;
        updates.lines.forEach(function(l) {
          var lineId  = generateId('LIN');
          var qty     = parseFloat(l.quantity)||1;
          var price   = parseFloat(l.unitPrice)||0;
          var vatRate = parseFloat(l.vatRate)||0;
          var net     = qty * price;
          var vat     = net * vatRate / 100;
          var gross   = net + vat;
          subtotal   += net;
          vatTotal   += vat;
          lineSheet.appendRow([lineId, invoiceId, l.description, qty, price, vatRate, gross,
                                l.accountCode || '4000']);
        });
        var newTotal = subtotal + vatTotal;
        sheet.getRange(rowNum,  9).setValue(subtotal);   // Subtotal
        sheet.getRange(rowNum, 10).setValue(vatTotal);   // VAT
        sheet.getRange(rowNum, 11).setValue(newTotal);   // Total
        sheet.getRange(rowNum, 12).setValue(0);          // AmountPaid reset
        sheet.getRange(rowNum, 14).setValue(newTotal);   // AmountDue
        changes.push('lines+totals');
      }
    }

    if (changes.length === 0) return { success:true, message:'No changes made' };

    addInvoiceHistory(invoiceId, 'Edit', changes.join(','), '', '', 'Edited: '+changes.join(', '));
    return { success:true, message:'Invoice updated: '+changes.join(', '), changes:changes };

  } catch(e) {
    Logger.log('editInvoice error: '+e);
    return { success:false, message:e.toString() };
  }
}

function approveInvoice(invoiceId, params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        var current = data[i][14] || 'Draft';
        if (current !== 'Draft') return { success:false, message:'Only Draft invoices can be approved (current: '+current+')' };
        try { _checkPeriodLock(data[i][6], params); } catch(lockErr) { return { success:false, message:lockErr.message }; }
        sheet.getRange(i+1, INV_COLS.STATUS).setValue('Approved');
        addInvoiceHistory(invoiceId,'StatusChange','status','Draft','Approved','Invoice approved — now a Tax Invoice');
        return { success:true, message:'Invoice approved' };
      }
    }
    return { success:false, message:'Invoice not found' };
  } catch(e) {
    Logger.log('approveInvoice error: '+e);
    return { success:false, message:e.toString() };
  }
}

/**
 * markInvoiceSent(invoiceId)
 * Transitions Approved → Sent.
 */
function markInvoiceSent(invoiceId, params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        var current = data[i][14] || 'Draft';
        if (current !== 'Approved') return { success:false, message:'Only Approved invoices can be marked Sent (current: '+current+')' };
        sheet.getRange(i+1, INV_COLS.STATUS).setValue('Sent');
        addInvoiceHistory(invoiceId,'StatusChange','status','Approved','Sent','Invoice marked as sent to client');
        return { success:true, message:'Invoice marked as sent' };
      }
    }
    return { success:false, message:'Invoice not found' };
  } catch(e) {
    return { success:false, message:e.toString() };
  }
}

function _recordPayment(invoiceId, amount, paymentDate, notes, params) {
  try {
    try { _checkPeriodLock(paymentDate ? new Date(paymentDate) : new Date(), params); } catch(le) { return { success:false, message:le.message }; }
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    var data = sheet.getDataRange().getValues();
    
    var dateObj = typeof paymentDate === 'string' ? new Date(paymentDate) : paymentDate || new Date();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        var currentPaid = parseFloat(data[i][INV_COLS.AMOUNT_PAID-1]) || 0;
        var total = parseFloat(data[i][INV_COLS.TOTAL-1]) || 0;
        var newPaid = currentPaid + parseFloat(amount);
        var newDue = total - newPaid;
        var rowNum = i + 1;
        
        sheet.getRange(rowNum, INV_COLS.AMOUNT_PAID).setValue(newPaid);
        sheet.getRange(rowNum, INV_COLS.AMOUNT_DUE).setValue(newDue);
        
        var newStatus = newDue <= 0.01 ? 'Paid' : (newPaid > 0 ? 'Partial' : 'Sent');
        sheet.getRange(rowNum, INV_COLS.STATUS).setValue(newStatus);
        
        if (newDue <= 0.01) {
          sheet.getRange(rowNum, INV_COLS.PAYMENT_DATE).setValue(safeSerializeDate(dateObj));
        }
        
        if (notes && notes.trim() !== '') {
          var existingNotes = data[i][INV_COLS.NOTES-1] || '';
          var paymentNote = '\n[Payment ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy') + ']: £' + amount + ' - ' + notes;
          sheet.getRange(rowNum, INV_COLS.NOTES).setValue(existingNotes + paymentNote);
        }
        
        createDoubleEntry(
          dateObj,
          'Payment',
          data[i][INV_COLS.NUMBER-1],
          '1000',
          '1100',
          parseFloat(amount),
          'Payment received for ' + data[i][INV_COLS.NUMBER-1] + (notes ? ' - ' + notes : ''),
          invoiceId,
          null
        );
        
        addInvoiceHistory(invoiceId, 'Payment', 'amount', currentPaid, newPaid, notes);
        
        return { 
          success: true, 
          message: 'Payment recorded',
          newAmountPaid: newPaid,
          newAmountDue: newDue,
          status: newStatus
        };
      }
    }
    return { success: false, message: 'Invoice not found' };
  } catch (e) {
    Logger.log('Error in recordPayment: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function updateInvoice(invoiceId, updates, params) {
  try {
    _auth('invoices.write', params);
    Logger.log('=== UPDATE INVOICE ===');
    Logger.log('Invoice ID: ' + invoiceId);
    Logger.log('Updates: ' + JSON.stringify(updates));
    
    var ss = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.INVOICES);
    
    if (!sheet) return { success: false, message: 'Invoices sheet not found' };
    
    var data = sheet.getDataRange().getValues();
    var rowNum = -1;
    var oldData = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        rowNum = i + 1;
        oldData = data[i];
        break;
      }
    }
    
    if (rowNum === -1) return { success: false, message: 'Invoice not found' };
    
    // Update fields
    if (updates.status !== undefined) {
      sheet.getRange(rowNum, 15).setValue(updates.status); // Status column
    }
    
    if (updates.paymentDate !== undefined) {
      sheet.getRange(rowNum, 16).setValue(updates.paymentDate); // PaymentDate column
    }
    
    if (updates.amountPaid !== undefined) {
      sheet.getRange(rowNum, 13).setValue(updates.amountPaid); // AmountPaid column
    }
    
    if (updates.amountDue !== undefined) {
      sheet.getRange(rowNum, 14).setValue(updates.amountDue); // AmountDue column
    }
    
    if (updates.notes !== undefined) {
      var existingNotes = oldData[16] || '';
      var newNotes = updates.notes;
      if (updates.bankAccountId || updates.paymentRef) {
        var bankAccount = getBankAccountById(updates.bankAccountId);
        var bankInfo = bankAccount ? bankAccount.accountName : updates.bankAccountId;
        var paymentNote = '\n[Payment Update ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + 
                         '] Bank: ' + bankInfo + 
                         (updates.paymentRef ? ' Ref: ' + updates.paymentRef : '') +
                         (updates.amountPaid ? ' Amount: £' + updates.amountPaid : '');
        newNotes = existingNotes + paymentNote;
      }
      sheet.getRange(rowNum, 17).setValue(newNotes); // Notes column
    }
    
    // Update issueDate if provided
    if (updates.issueDate !== undefined && updates.issueDate !== '') {
      sheet.getRange(rowNum, 7).setValue(safeSerializeDate(updates.issueDate));
    }
    
    // Update dueDate if provided
    if (updates.dueDate !== undefined && updates.dueDate !== '') {
      sheet.getRange(rowNum, 8).setValue(safeSerializeDate(updates.dueDate));
    }
    
    // Persist bankAccountId to column 19 if provided
    if (updates.bankAccountId) {
      sheet.getRange(rowNum, 19).setValue(updates.bankAccountId);
    }
    
    // Add to history
    addInvoiceHistory(invoiceId, 'Edited', 'payment', 
                     JSON.stringify({ paid: oldData[12], status: oldData[14] }),
                     JSON.stringify({ paid: updates.amountPaid, status: updates.status, bank: updates.bankAccountId }),
                     'Payment details updated');
    
    // Create bank transaction ONLY when status is newly changing to Paid
    // (prevents duplicate transactions when re-saving an already-paid invoice)
    var oldStatus = oldData[14] ? oldData[14].toString() : '';
    var newStatus = updates.status || '';
    var isNewPayment = (oldStatus !== 'Paid') && (newStatus === 'Paid');
    
    if (updates.bankAccountId && updates.amountPaid > 0 && isNewPayment) {
      createBankTransactionFromPayment({
        date: updates.paymentDate || new Date(),
        description: 'Payment received for invoice ' + (oldData[1] || invoiceId),
        reference: updates.paymentRef || oldData[1] || invoiceId,
        amount: updates.amountPaid,
        type: 'Credit',
        bankAccountId: updates.bankAccountId,
        category: 'Sales',
        notes: 'Payment recorded via invoice edit'
      });
    }
    
    return { success: true, message: 'Invoice updated successfully' };
    
  } catch (e) {
    Logger.log('Error updating invoice: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================
// UPDATE BILL WITH PAYMENT DETAILS
// ============================================

function getInvoiceHistory(invoiceId, limit, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICE_HISTORY);
    if (!sheet) return { success: true, history: [] };
    
    var data = sheet.getDataRange().getValues();
    var history = [];
    
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === invoiceId) {
        history.push({
          historyId: data[i][0],
          invoiceId: data[i][1],
          timestamp: data[i][2],
          user: data[i][3],
          changeType: data[i][4],
          fieldChanged: data[i][5],
          oldValue: data[i][6],
          newValue: data[i][7],
          notes: data[i][8]
        });
        
        if (limit && history.length >= limit) break;
      }
    }
    
    return { success: true, history: history };
    
  } catch (e) {
    Logger.log('Error getting invoice history: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function updateInvoiceWithPayment(invoiceId, paymentData, params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === invoiceId) {
        var rowNum = i + 1;
        var changes = [];
        
        if (paymentData.amountPaid !== undefined) {
          var oldPaid = data[i][INV_COLS.AMOUNT_PAID-1];
          sheet.getRange(rowNum, INV_COLS.AMOUNT_PAID).setValue(paymentData.amountPaid);
          changes.push({ field: 'amountPaid', old: oldPaid, new: paymentData.amountPaid });
        }
        
        if (paymentData.amountDue !== undefined) {
          sheet.getRange(rowNum, INV_COLS.AMOUNT_DUE).setValue(paymentData.amountDue);
        }
        
        if (paymentData.status !== undefined) {
          var oldStatus = data[i][INV_COLS.STATUS-1];
          sheet.getRange(rowNum, INV_COLS.STATUS).setValue(paymentData.status);
          changes.push({ field: 'status', old: oldStatus, new: paymentData.status });
        }
        
        if (paymentData.paymentDate !== undefined) {
          sheet.getRange(rowNum, INV_COLS.PAYMENT_DATE).setValue(safeSerializeDate(paymentData.paymentDate));
        }
        
        if (paymentData.bankAccountId || paymentData.paymentNotes) {
          var existingNotes = data[i][INV_COLS.NOTES-1] || '';
          var bankInfo = paymentData.bankAccountId ? 'Bank: ' + (function(){var _b=getBankAccountById(paymentData.bankAccountId);return _b?_b.accountName:paymentData.bankAccountId;}()) : '';
          var note = '\n[Payment Update ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + 
                    '] ' + bankInfo + (paymentData.paymentNotes ? ' - ' + paymentData.paymentNotes : '');
          sheet.getRange(rowNum, INV_COLS.NOTES).setValue(existingNotes + note);
        }
        
        addInvoiceHistory(invoiceId, 'PaymentUpdate', 'payment', '', JSON.stringify(paymentData), 
                         'Payment details updated');
        
        return { success: true, changes: changes };
      }
    }
    
    return { success: false, message: 'Invoice not found' };
  } catch (e) {
    Logger.log('Error in updateInvoiceWithPayment: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// Add this to your Code.gs file
function checkInvoicesSheet() {
  try {
    var ss = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.INVOICES);
    
    if (!sheet) {
      return { 
        success: false, 
        message: 'Invoices sheet not found',
        action: 'Run initializeSystem() to create it'
      };
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var expectedHeaders = [
      'InvoiceId', 'InvoiceNumber', 'ClientId', 'ClientName', 'ClientEmail', 
      'ClientAddress', 'IssueDate', 'DueDate', 'Subtotal', 'VATRate', 
      'VAT', 'Total', 'AmountPaid', 'AmountDue', 'Status', 'PaymentDate', 'Notes'
    ];
    
    var missingHeaders = [];
    for (var i = 0; i < expectedHeaders.length; i++) {
      if (headers[i] !== expectedHeaders[i]) {
        missingHeaders.push({ index: i, expected: expectedHeaders[i], actual: headers[i] || 'MISSING' });
      }
    }
    
    if (missingHeaders.length > 0) {
      return {
        success: false,
        message: 'Invoices sheet has incorrect headers',
        missingHeaders: missingHeaders,
        action: 'Run initializeSystem() to recreate the sheet'
      };
    }
    
    var rowCount = sheet.getLastRow() - 1; // Subtract header row
    var sampleData = [];
    if (rowCount > 0) {
      sampleData = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
    
    return {
      success: true,
      message: 'Invoices sheet exists and has correct structure',
      rowCount: rowCount,
      sampleData: sampleData
    };
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function diagnoseInvoiceColumns() {
  var ss = getDb(params || {});
  var sheet = ss.getSheetByName(SHEETS.INVOICES);
  
  if (!sheet) {
    Logger.log('❌ Invoices sheet not found');
    return;
  }
  
  var lastCol = sheet.getLastColumn();
  Logger.log('📊 Invoices sheet has ' + lastCol + ' columns');
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('Headers:');
  headers.forEach(function(header, index) {
    Logger.log('  Column ' + (index + 1) + ': "' + header + '"');
  });
  
  if (sheet.getLastRow() > 1) {
    var firstRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    Logger.log('\nSample first row:');
    firstRow.forEach(function(value, index) {
      Logger.log('  Col ' + (index + 1) + ' (' + headers[index] + '): "' + value + '"');
    });
  }
  
  return {
    columnCount: lastCol,
    headers: headers
  };
}


// ============================================
// SA100 / SOLE TRADER TAX (SA103)
// ============================================

// Capital Allowances sheet: AssetId, Description, PurchaseDate, Cost, 
//   AssetType (AIA/FYA/Main Pool/Special Rate), PrivateUsePercent, 
//   DisposalDate, DisposalValue, TaxYear, Notes

function generateInvoicePDF(invoiceId, params) {
  try {
    var invoice = getInvoiceById(invoiceId, params);
    if (!invoice) return { success: false, message: 'Invoice not found' };
    
    var lines = getInvoiceLines(invoiceId, params);
    var settings = getSettings(params);
    var html = generateInvoiceHTML(invoice, lines, settings);
    
    var blob = Utilities.newBlob(html, 'text/html', invoice.invoiceNumber + '.html');
    var pdf = blob.getAs('application/pdf');
    pdf.setName(invoice.invoiceNumber + '.pdf');
    
    var folder = getOrCreateFolder('Invoices');
    var files = folder.getFilesByName(invoice.invoiceNumber + '.pdf');
    while (files.hasNext()) {
      var oldFile = files.next();
      oldFile.setTrashed(true);
    }
    
    var file = folder.createFile(pdf);
    
    return { 
      success: true, 
      pdfUrl: file.getUrl(), 
      fileId: file.getId(),
      invoiceNumber: invoice.invoiceNumber 
    };
  } catch (e) {
    Logger.log('Error in generateInvoicePDF: ' + e.toString());
    return { success: false, message: 'Error generating PDF: ' + e.toString() };
  }
}

function generateInvoiceHTML(invoice, lines, settings, params) {
  try {
    var logoBase64 = '';
    // Extract Drive file ID from the stored logo URL and use DriveApp directly
    // This is more reliable than UrlFetchApp which struggles with Drive auth
    if (settings && settings.logoURL) {
      try {
        var logoFileId = null;
        // Parse file ID from thumbnail URL: https://drive.google.com/thumbnail?id=FILE_ID&sz=...
        var thumbMatch = settings.logoURL.match(/[?&]id=([a-zA-Z0-9_-]+)/);
        if (thumbMatch) { logoFileId = thumbMatch[1]; }
        // Parse file ID from standard Drive URL: https://drive.google.com/file/d/FILE_ID/...
        if (!logoFileId) {
          var driveMatch = settings.logoURL.match(/\/d\/([a-zA-Z0-9_-]+)/);
          if (driveMatch) { logoFileId = driveMatch[1]; }
        }
        if (logoFileId) {
          var logoFile = DriveApp.getFileById(logoFileId);
          var logoBlob = logoFile.getBlob();
          logoBase64 = 'data:' + logoBlob.getContentType() + ';base64,' + Utilities.base64Encode(logoBlob.getBytes());
        } else {
          // Fallback: try fetching URL directly
          var res = UrlFetchApp.fetch(settings.logoURL, { muteHttpExceptions:true });
          if (res.getResponseCode() === 200) {
            var b = res.getBlob();
            logoBase64 = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
          }
        }
      } catch(e) { Logger.log('Logo load error: ' + e); }
    }

    var tz         = Session.getScriptTimeZone();
    var issueStr   = invoice.issueDate ? Utilities.formatDate(new Date(invoice.issueDate), tz, 'dd MMM yyyy') : '—';
    var dueStr     = invoice.dueDate   ? Utilities.formatDate(new Date(invoice.dueDate),   tz, 'dd MMM yyyy') : '—';
    var subtotal   = parseFloat(invoice.subtotal)  || 0;
    var vatAmt     = parseFloat(invoice.vatAmount) || 0;
    var total      = parseFloat(invoice.total)     || 0;
    var amountDue  = parseFloat(invoice.amountDue) || 0;
    var isPaid     = invoice.status === 'Paid';
    var isDraft    = invoice.status === 'Draft';
    var isVATReg   = !!(settings && (settings.vatRegistered === true || settings.vatRegistered === 'true' || settings.vatRegistered === 'TRUE'));
    var baseCurr   = settings.baseCurrency || 'GBP';
    var invCurr    = invoice.currency || baseCurr;
    var isForeign  = invCurr !== baseCurr;
    var fxRate     = (function(r) {
      var n = parseFloat(r);
      return (n && n > 0.001 && n < 10000) ? n : 1;
    })(invoice.exchangeRate);
    var currSymbol = { GBP:'£', EUR:'€', USD:'$', CHF:'Fr', SEK:'kr', NOK:'kr', DKK:'kr', JPY:'¥', CAD:'$', AUD:'$' }[invCurr] || invCurr+' ';
    var baseSymbol = { GBP:'£', EUR:'€', USD:'$' }[baseCurr] || baseCurr+' ';
    var invoiceLabel = isDraft
      ? (isVATReg ? 'PROFORMA INVOICE' : 'INVOICE')
      : (isVATReg ? 'TAX INVOICE' : 'INVOICE');
    var accentColor  = isDraft ? '#64748b' : (settings.templateAccentColor || '#1a3c6b');
    var logoPos      = settings.templateLogoPosition || 'left';
    var templateFont = settings.templateFont==='serif' ? '"Georgia",serif' : settings.templateFont==='mono' ? '"Courier New",monospace' : '"Helvetica Neue",Arial,sans-serif';

    // Line items
    var hasVAT   = lines.some(function(l){ return parseFloat(l.vatRate||0) > 0; });
    var isVATReg2 = isVATReg; // already declared above
    var linesHTML = '';
    for (var i = 0; i < lines.length; i++) {
      var l = lines[i];
      var rowBg  = i % 2 === 0 ? '#ffffff' : '#f7f9fc';
      var vr     = parseFloat(l.vatRate) || 0;
      var net    = parseFloat(l.unitPrice||0) * parseFloat(l.quantity||1);
      var vatAmt2 = net * vr / 100;
      var gross  = net + vatAmt2;
      var vatLbl = vr > 0 ? vr + '%' : (l.vatRate === -1 ? 'Exempt' : 'Zero');
      linesHTML +=
        '<tr style="background:' + rowBg + '">' +
        '<td style="padding:10px 14px;border-bottom:1px solid #e8edf3">' + escapeHtml(l.description||'') + '</td>' +
        '<td style="padding:10px 14px;text-align:center;border-bottom:1px solid #e8edf3;color:#555">' + (parseFloat(l.quantity)||0) + '</td>' +
        '<td style="padding:10px 14px;text-align:right;border-bottom:1px solid #e8edf3;color:#555">'+currSymbol + parseFloat(l.unitPrice||0).toFixed(2) + '</td>' +
        (hasVAT && isVATReg ? '<td style="padding:10px 14px;text-align:center;border-bottom:1px solid #e8edf3;color:#555">' + vatLbl + '</td>' : '') +
        '<td style="padding:10px 14px;text-align:right;border-bottom:1px solid #e8edf3;font-weight:500">'+currSymbol + gross.toFixed(2) + '</td>' +
        '</tr>';
    }

    var html = '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">' +
      '<style>' +
      '*{box-sizing:border-box;margin:0;padding:0}' +
      'body{font-family:'+templateFont+';font-size:12px;color:#2d3748;background:#fff;padding:40px 50px}' +
      'h1{font-size:32px;font-weight:700;letter-spacing:-0.5px;color:' + accentColor + '}' +
      '.label{font-size:10px;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;color:#94a3b8;margin-bottom:3px}' +
      '.value{font-size:13px;font-weight:500;color:#2d3748}' +
      'table.lines{width:100%;border-collapse:collapse;margin:24px 0}' +
      'table.lines th{background:' + accentColor + ';color:#fff;padding:10px 14px;text-align:left;font-size:11px;letter-spacing:0.5px;font-weight:600;text-transform:uppercase}' +
      'table.lines th:nth-child(2){text-align:center}' +
      'table.lines th:nth-child(n+3){text-align:right}' +
      '.totals-table{margin-left:auto;border-collapse:collapse}' +
      '.totals-table td{padding:5px 10px;font-size:12px}' +
      '.totals-table td:last-child{text-align:right;min-width:90px}' +
      '.grand-total td{font-size:15px;font-weight:700;padding:10px 10px;border-top:2px solid ' + accentColor + ';color:' + accentColor + '}' +
      '.amount-due-row td{font-size:13px;font-weight:700;color:#e53e3e}' +
      '.paid-stamp{display:inline-block;border:3px solid #22c55e;color:#22c55e;font-size:22px;font-weight:800;letter-spacing:3px;padding:4px 16px;border-radius:4px;transform:rotate(-8deg);opacity:0.85}' +
      '.divider{border:none;border-top:1px solid #e8edf3;margin:28px 0}' +
      '</style></head><body>';

    // ── Header ──────────────────────────────────────────────────────────────
    html += '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:36px">';

    // Company block
    html += '<div style="max-width:260px">';
    if (logoBase64) {
      html += '<img src="' + logoBase64 + '" style="max-height:70px;max-width:200px;object-fit:contain;margin-bottom:12px;display:block" alt="logo">';
    }
    html += '<div style="font-size:14px;font-weight:700;color:' + accentColor + ';margin-bottom:4px">' + escapeHtml(settings.companyName||'') + '</div>';
    if (settings.companyAddress) html += '<div style="color:#555;line-height:1.5">' + escapeHtml(settings.companyAddress) + '</div>';
    if (settings.companyPostcode) html += '<div style="color:#555">' + escapeHtml(settings.companyPostcode) + '</div>';
    if (settings.companyPhone)   html += '<div style="color:#555">T: ' + escapeHtml(settings.companyPhone) + '</div>';
    if (settings.companyEmail)   html += '<div style="color:#555">' + escapeHtml(settings.companyEmail) + '</div>';
    if (settings.vatRegNumber)   html += '<div style="color:#94a3b8;font-size:11px;margin-top:4px">VAT No: ' + escapeHtml(settings.vatRegNumber) + '</div>';
    html += '</div>';

    // Invoice title + meta
    html += '<div style="text-align:right">' +
      '<h1>INVOICE</h1>' +
      '<div style="margin-top:10px">' +
      '<div class="label">Invoice Number</div><div class="value" style="font-size:15px;font-weight:700">' + escapeHtml(invoice.invoiceNumber||'') + '</div>' +
      '</div>' +
      '<div style="display:flex;gap:28px;margin-top:14px;justify-content:flex-end">' +
      '<div><div class="label">Issue Date</div><div class="value">' + issueStr + '</div></div>' +
      '<div><div class="label">Due Date</div><div class="value" style="color:#e53e3e">' + dueStr + '</div></div>' +
      '</div>';
    if (isPaid) {
      html += '<div style="margin-top:14px"><span class="paid-stamp">PAID</span></div>';
    } else if (isDraft && isVATReg) {
      html += '<div style="margin-top:14px"><span style="display:inline-block;border:2px dashed #94a3b8;color:#94a3b8;font-size:13px;font-weight:700;letter-spacing:2px;padding:3px 12px;border-radius:3px">NOT A VAT DOCUMENT</span></div>';
    }
    html += '</div>';
    html += '</div>';

    // ── Bill To ──────────────────────────────────────────────────────────────
    html += '<div style="display:flex;gap:48px;margin-bottom:28px">';
    html += '<div><div class="label">Bill To</div>' +
      '<div style="font-size:13px;font-weight:700;margin-top:4px">' + escapeHtml(invoice.clientName||'') + '</div>';
    if (invoice.clientAddress) html += '<div style="color:#555;line-height:1.6">' + escapeHtml(invoice.clientAddress) + '</div>';
    html += '</div>';
    if (invoice.reference) {
      html += '<div><div class="label">Your Reference</div><div class="value" style="margin-top:4px">' + escapeHtml(invoice.reference) + '</div></div>';
    }
    html += '</div>';

    // ── Line Items ────────────────────────────────────────────────────────────
    html += '<table class="lines">' +
      '<thead><tr>' +
      '<th>Description</th><th style="text-align:center">Qty</th>' +
      '<th style="text-align:right">Unit Price</th>' +
      (hasVAT && isVATReg ? '<th style="text-align:center">VAT</th>' : '') +
      '<th style="text-align:right">Amount</th>' +
      '</tr></thead><tbody>' + linesHTML + '</tbody></table>';

    // ── Totals ────────────────────────────────────────────────────────────────
    // Build VAT breakdown by rate
    var vatBreakdown = {};
    for (var vi = 0; vi < lines.length; vi++) {
      var vr = parseFloat(lines[vi].vatRate) || 0;
      var vn = parseFloat(lines[vi].unitPrice||0) * parseFloat(lines[vi].quantity||1);
      var va = vn * vr / 100;
      if (!vatBreakdown[vr]) vatBreakdown[vr] = 0;
      vatBreakdown[vr] += va;
    }
    html += '<div style="display:flex;justify-content:flex-end">' +
      '<table class="totals-table">';
    html += '<tr><td style="color:#555">Subtotal</td><td>'+currSymbol + subtotal.toFixed(2) + '</td></tr>';
    if (isVATReg) {
      // Always show VAT row on VAT-registered invoices (even if £0)
      var vatRates = Object.keys(vatBreakdown).sort();
      if (vatRates.length > 0) {
        vatRates.forEach(function(rate) {
          var va = vatBreakdown[rate];
          html += '<tr><td style="color:#555">VAT ' + rate + '%</td><td>'+currSymbol + va.toFixed(2) + '</td></tr>';
        });
      } else {
        // No lines yet or all zero — show blank VAT row
        html += '<tr><td style="color:#555">VAT</td><td>'+currSymbol+'0.00</td></tr>';
      }
    }
    html += '<tr class="grand-total"><td>Total</td><td>'+currSymbol + total.toFixed(2) + '</td></tr>';
    if (isForeign) {
      var baseTotal = Math.round(total * fxRate * 100) / 100;
      html += '<tr><td style="color:#94a3b8;font-size:11px">'+baseCurr+' equivalent (rate: '+fxRate+')</td><td style="color:#94a3b8;font-size:11px">'+baseSymbol+baseTotal.toFixed(2)+'</td></tr>';
    }
    if (!isPaid && amountDue < total) {
      html += '<tr><td style="color:#555;font-size:11px">Amount Paid</td><td style="color:#22c55e">'+currSymbol + (total - amountDue).toFixed(2) + '</td></tr>';
    }
    html += '<tr class="amount-due-row"><td>Amount Due</td><td>'+currSymbol + amountDue.toFixed(2) + '</td></tr>';
    html += '</table></div>';

    // ── Bank Details ──────────────────────────────────────────────────────────
    var hasBankDetails = settings && (settings.bankName || settings.accountName || settings.sortCode || settings.accountNumber);
    if (hasBankDetails) {
      html += '<hr class="divider">' +
        '<div style="display:flex;gap:48px">' +
        '<div><div class="label" style="margin-bottom:8px">Payment Details</div>';
      if (settings.bankName)      html += '<div style="line-height:1.8"><span style="color:#94a3b8;font-size:11px">Bank: </span>' + escapeHtml(settings.bankName) + '</div>';
      if (settings.accountName)   html += '<div style="line-height:1.8"><span style="color:#94a3b8;font-size:11px">Account Name: </span>' + escapeHtml(settings.accountName) + '</div>';
      if (settings.sortCode)      html += '<div style="line-height:1.8"><span style="color:#94a3b8;font-size:11px">Sort Code: </span>' + escapeHtml(settings.sortCode) + '</div>';
      if (settings.accountNumber) html += '<div style="line-height:1.8"><span style="color:#94a3b8;font-size:11px">Account No: </span>' + escapeHtml(settings.accountNumber) + '</div>';
      html += '</div></div>';
    }

    // ── Notes / Footer ────────────────────────────────────────────────────────
    // Notes intentionally excluded from PDF — internal use only

    if (settings.invoiceFooter) {
      html += '<hr class="divider"><div style="text-align:center;color:#94a3b8;font-size:11px">' + escapeHtml(settings.invoiceFooter) + '</div>';
    }

    html += '</body></html>';
    return html;
  } catch(e) {
    Logger.log('generateInvoiceHTML error: ' + e);
    return '<html><body><h1>Error</h1><p>' + e.toString() + '</p></body></html>';
  }
}


// ============================================
// EMAIL
// ============================================

function sendInvoiceEmail(invoiceId, overrides, params) {
  try {
    overrides = overrides || {};
    var invoice = getInvoiceById(invoiceId, params);
    var settings = getSettings(params);
    
    if (!invoice) return { success: false, message: 'Invoice not found' };
    var toEmail = overrides.to || invoice.clientEmail || '';
    if (!toEmail) return { success: false, message: 'No recipient email address' };
    
    var pdfResult = generateInvoicePDF(invoiceId, params);
    if (!pdfResult.success) return pdfResult;
    
    var pdfFile = DriveApp.getFileById(pdfResult.fileId);
    
    // ── Build payment details string ─────────────────────────────────────────
    var paymentDetails = '';
    if (settings.bankName) {
      paymentDetails = 'Bank: ' + (settings.bankName||'') +
        '\nAccount Name: ' + (settings.accountName||'') +
        '\nSort Code: '    + (settings.sortCode||'') +
        '\nAccount Number: '+ (settings.accountNumber||'');
    }

    // ── Apply email template from Settings, fall back to default ─────────────
    var defaultSubject = 'Invoice {{invoiceNumber}} from {{companyName}}';
    var defaultBody    =
      'Dear {{clientName}},\n\n' +
      'Please find attached invoice {{invoiceNumber}} for {{total}}, due on {{dueDate}}.\n\n' +
      (paymentDetails ? 'Payment Details\n' + paymentDetails + '\n\n' : '') +
      'If you have any questions, please do not hesitate to get in touch.\n\n' +
      'Kind regards,\n{{companyName}}';

    var subjectTpl = (settings.emailSubject && settings.emailSubject.trim())
      ? settings.emailSubject : defaultSubject;
    var bodyTpl    = (settings.emailBody && settings.emailBody.trim())
      ? settings.emailBody    : defaultBody;

    // ── Substitute template variables ─────────────────────────────────────────
    function applyTemplate(tpl) {
      return tpl
        .replace(/\{\{invoiceNumber\}\}/g, invoice.invoiceNumber || '')
        .replace(/\{\{clientName\}\}/g,    invoice.clientName    || '')
        .replace(/\{\{total\}\}/g,         '£' + (invoice.total||0).toFixed(2))
        .replace(/\{\{dueDate\}\}/g,       formatDate ? formatDate(invoice.dueDate) : (invoice.dueDate||''))
        .replace(/\{\{amountDue\}\}/g,     '£' + (invoice.amountDue||0).toFixed(2))
        .replace(/\{\{companyName\}\}/g,   settings.companyName  || '')
        .replace(/\{\{paymentDetails\}\}/g,paymentDetails);
    }

    var subject = applyTemplate(subjectTpl);
    var body    = applyTemplate(bodyTpl);
    
    var mailOpts = {
      name: settings.companyName || '',
      attachments: [pdfFile.getAs(MimeType.PDF)]
    };
    if (overrides.cc)  mailOpts.cc  = overrides.cc;
    if (overrides.bcc) mailOpts.bcc = overrides.bcc;
    // Use overridden subject/body if provided (user edited in send modal)
    if (overrides.subject) subject = overrides.subject;
    if (overrides.body)    body    = overrides.body;
    GmailApp.sendEmail(toEmail, subject, body, mailOpts);
    
    if (invoice.status === 'Draft') {
      updateInvoiceStatus(invoiceId, 'Sent');
    }
    
    addInvoiceHistory(invoiceId, 'EmailSent', '', '', '', 'Email sent to ' + invoice.clientEmail);
    
    return { success: true, message: 'Email sent successfully' };
  } catch (e) {
    Logger.log('Error sending email: ' + e.toString());
    return { success: false, message: 'Error sending email: ' + e.toString() };
  }
}


function recordPayment(invoiceId, amount, paymentDate, notes, params) {
  return _recordPayment(invoiceId, amount, paymentDate, notes, params);
}

/**
 * generateClientStatement(clientId, startDate, endDate)
 * Produces an HTML statement of all invoices for a client over a period.
 * Returns { success, fileId, fileUrl, pdfUrl }
 */
function generateClientStatement(clientId, startDate, endDate, params) {
  try {
    var settings = getSettings(params);
    var allInvs  = getAllInvoices(params);
    if (!allInvs.success) return allInvs;

    var start = startDate ? new Date(startDate) : new Date(0);
    var end   = endDate   ? new Date(endDate)   : new Date();
    end.setHours(23,59,59,999);

    var invs = allInvs.invoices.filter(function(i){
      if (i.clientId !== clientId) return false;
      if (i.status === 'Void' || i.status === 'Draft') return false;
      var d = i.issueDate ? new Date(i.issueDate) : null;
      return d && d >= start && d <= end;
    }).sort(function(a,b){ return new Date(a.issueDate)-new Date(b.issueDate); });

    if (!invs.length) return { success:false, message:'No invoices found for this client in the selected period.' };

    var client    = { clientName: invs[0].clientName, clientEmail: invs[0].clientEmail, clientAddress: invs[0].clientAddress };
    var tz        = Session.getScriptTimeZone();
    var fmtDate2  = function(d){ return d ? Utilities.formatDate(new Date(d),tz,'dd MMM yyyy') : '—'; };
    var fmtAmt2   = function(n){ return '£'+(parseFloat(n)||0).toFixed(2); };

    var totalBilled  = invs.reduce(function(s,i){ return s+(parseFloat(i.total)||0); }, 0);
    var totalPaid    = invs.reduce(function(s,i){ return s+((parseFloat(i.total)||0)-(parseFloat(i.amountDue)||0)); }, 0);
    var totalOutstanding = totalBilled - totalPaid;

    var rows = invs.map(function(i){
      var isPaid = i.status==='Paid';
      var statusCol = isPaid?'#22c55e': (i.status==='Overdue'?'#ef4444':'#1d4ed8');
      return '<tr style="border-bottom:1px solid #e8edf3">' +
        '<td style="padding:8px 10px">'+fmtDate2(i.issueDate)+'</td>' +
        '<td style="padding:8px 10px;font-weight:600;font-family:monospace">'+i.invoiceNumber+'</td>' +
        '<td style="padding:8px 10px">'+(i.reference||'—')+'</td>' +
        '<td style="padding:8px 10px;text-align:right">'+fmtAmt2(i.total)+'</td>' +
        '<td style="padding:8px 10px;text-align:right;color:#22c55e">'+fmtAmt2((parseFloat(i.total)||0)-(parseFloat(i.amountDue)||0))+'</td>' +
        '<td style="padding:8px 10px;text-align:right;color:'+(parseFloat(i.amountDue)>0?'#ef4444':'#22c55e')+'">'+fmtAmt2(i.amountDue)+'</td>' +
        '<td style="padding:8px 10px;text-align:center"><span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;background:'+(isPaid?'#f0fdf4':'#fef2f2')+';color:'+statusCol+'">'+i.status+'</span></td>' +
      '</tr>';
    }).join('');

    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
      '<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:"Helvetica Neue",Arial,sans-serif;font-size:12px;color:#2d3748;padding:36px 44px}' +
      'h1{font-size:26px;font-weight:700;color:#1a3c6b;margin-bottom:4px}' +
      'table{width:100%;border-collapse:collapse}th{background:#1a3c6b;color:#fff;padding:9px 10px;text-align:left;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.4px}' +
      'th:nth-child(n+4){text-align:right}.summary{display:flex;gap:0;margin-top:24px;border:1px solid #e8edf3;border-radius:6px;overflow:hidden}' +
      '.sum-cell{flex:1;padding:14px 16px;border-right:1px solid #e8edf3}.sum-cell:last-child{border-right:none}' +
      '.sum-label{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:#94a3b8;margin-bottom:5px}' +
      '.sum-val{font-size:20px;font-weight:700}</style></head><body>' +
      '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:28px">' +
        '<div><h1>Account Statement</h1>' +
          '<div style="font-size:13px;font-weight:700;color:#2d3748;margin-top:8px">'+escapeHtml(client.clientName)+'</div>' +
          (client.clientAddress?'<div style="color:#555">'+escapeHtml(client.clientAddress)+'</div>':'') +
          (client.clientEmail?'<div style="color:#555">'+escapeHtml(client.clientEmail)+'</div>':'') +
        '</div>' +
        '<div style="text-align:right">' +
          '<div style="font-size:13px;font-weight:700">'+escapeHtml(settings.companyName||'')+'</div>' +
          (settings.companyAddress?'<div style="color:#555">'+escapeHtml(settings.companyAddress)+'</div>':'') +
          (settings.companyEmail?'<div style="color:#555">'+escapeHtml(settings.companyEmail)+'</div>':'') +
          '<div style="margin-top:6px;font-size:11px;color:#94a3b8">Statement period: '+fmtDate2(startDate)+' – '+fmtDate2(endDate)+'</div>' +
          '<div style="font-size:11px;color:#94a3b8">Printed: '+fmtDate2(new Date().toISOString())+'</div>' +
        '</div>' +
      '</div>' +
      '<table><thead><tr><th>Date</th><th>Invoice</th><th>Reference</th><th>Amount</th><th>Paid</th><th>Balance</th><th>Status</th></tr></thead>' +
      '<tbody>'+rows+'</tbody></table>' +
      '<div class="summary">' +
        '<div class="sum-cell"><div class="sum-label">Total Billed</div><div class="sum-val">'+fmtAmt2(totalBilled)+'</div></div>' +
        '<div class="sum-cell"><div class="sum-label">Total Paid</div><div class="sum-val" style="color:#22c55e">'+fmtAmt2(totalPaid)+'</div></div>' +
        '<div class="sum-cell" style="background:'+(totalOutstanding>0?'#fef2f2':'#f0fdf4')+'"><div class="sum-label">Outstanding</div>' +
          '<div class="sum-val" style="color:'+(totalOutstanding>0?'#ef4444':'#22c55e')+'">'+fmtAmt2(totalOutstanding)+'</div></div>' +
      '</div>' +
      (settings.invoiceFooter?'<div style="margin-top:24px;text-align:center;color:#94a3b8;font-size:11px">'+escapeHtml(settings.invoiceFooter)+'</div>':'') +
      '</body></html>';

    var blob    = Utilities.newBlob(html,'text/html',client.clientName+'-statement.html');
    var pdf     = blob.getAs('application/pdf');
    pdf.setName(client.clientName+' Statement '+fmtDate2(startDate)+' to '+fmtDate2(endDate)+'.pdf');
    var folder  = getOrCreateFolder('Client Statements');
    var file    = folder.createFile(pdf);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { success:true, fileId:file.getId(), pdfUrl:file.getUrl(),
             clientName:client.clientName, invoiceCount:invs.length,
             totalBilled:totalBilled, totalOutstanding:totalOutstanding };
  } catch(e) {
    Logger.log('generateClientStatement error: '+e.toString());
    return { success:false, message:e.toString() };
  }
}
// ─────────────────────────────────────────────────────────────────────────────
// VOID INVOICE
// ─────────────────────────────────────────────────────────────────────────────

function voidInvoice(invoiceId, reason, params) {
  try {
    _auth('invoices.write', params);
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.INVOICES);
    if (!sheet) return { success: false, message: 'Invoices sheet not found' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === invoiceId) {
        var status = (data[i][14] || '').toString();
        if (status === 'Void' || status === 'Voided') {
          return { success: false, message: 'Invoice is already voided.' };
        }
        var rowNum = i + 1;
        sheet.getRange(rowNum, 15).setValue('Void');
        if (data[i][16] !== undefined) {
          var existNotes = data[i][16] ? data[i][16].toString() : '';
          sheet.getRange(rowNum, 17).setValue(existNotes + '\n[Voided: ' + reason + ']');
        }

        // Reverse double-entry if invoice was Approved/Sent/Paid
        if (['Approved','Sent','Paid','Partial'].indexOf(status) >= 0) {
          try {
            createDoubleEntry(
              new Date(), 'VoidReversal', data[i][1] ? data[i][1].toString() : invoiceId,
              data[i][INV_COLS ? (INV_COLS.ACCOUNT_CODE - 1) : 9] || '4000',
              '1100',
              parseFloat(data[i][11]) || 0,
              'Void reversal: ' + reason,
              invoiceId, null, params
            );
          } catch(te) { Logger.log('voidInvoice double-entry reversal: ' + te); }
        }

        // Write to void log sheet
        var voidSheet = ss.getSheetByName(SHEETS.VOID_LOG || 'VoidLog');
        var voidEntry = null;
        if (voidSheet) {
          var voidId = generateId('VD');
          var user   = Session.getActiveUser().getEmail() || 'system';
          voidSheet.appendRow([
            voidId, invoiceId, data[i][1], data[i][3], data[i][11],
            new Date().toISOString().split('T')[0], user, reason, 'Invoice'
          ]);
          voidEntry = {
            voidId: voidId, invoiceId: invoiceId,
            invoiceNumber: data[i][1], clientName: data[i][3],
            total: data[i][11], voidDate: new Date().toISOString().split('T')[0],
            voidedBy: user, voidReason: reason
          };
        }

        addInvoiceHistory(invoiceId, 'Voided', 'status', status, 'Void', reason, params);
        logAudit('VOID', 'Invoice', invoiceId, { reason: reason });
        return { success: true, message: 'Invoice voided.', voidEntry: voidEntry };
      }
    }
    return { success: false, message: 'Invoice not found.' };
  } catch(e) {
    Logger.log('voidInvoice error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// WRITE OFF AS BAD DEBT
// ─────────────────────────────────────────────────────────────────────────────

function writeOffInvoice(invoiceId, writeOffDate, reason, params) {
  try {
    _auth('invoices.write', params);
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.INVOICES);
    if (!sheet) return { success: false, message: 'Invoices sheet not found' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === invoiceId) {
        var status    = (data[i][14] || '').toString();
        var amountDue = parseFloat(data[i][13]) || 0;
        var total     = parseFloat(data[i][11]) || 0;
        var vatAmt    = parseFloat(data[i][10]) || 0;

        if (status === 'Draft' || status === 'Void') {
          return { success: false, message: 'Cannot write off a ' + status + ' invoice.' };
        }

        var rowNum = i + 1;
        sheet.getRange(rowNum, 14).setValue(0);         // AmountDue = 0
        sheet.getRange(rowNum, 15).setValue('Bad Debt'); // Status
        var existNotes = data[i][16] ? data[i][16].toString() : '';
        sheet.getRange(rowNum, 17).setValue(existNotes + '\n[Bad Debt: ' + writeOffDate + ' — ' + reason + ']');

        // Double-entry: Dr Bad Debts (7800), Cr Trade Debtors (1100)
        try {
          createDoubleEntry(
            new Date(writeOffDate), 'BadDebt',
            data[i][1] ? data[i][1].toString() : invoiceId,
            '7800', '1100', amountDue,
            'Bad debt write-off: ' + reason,
            invoiceId, null, params
          );
        } catch(te) { Logger.log('writeOffInvoice double-entry: ' + te); }

        // VAT element — eligible for bad debt relief after 6 months
        var isVATReg = false;
        try {
          var sett = getSettings(params);
          isVATReg = !!(sett.vatRegistered === true || sett.vatRegistered === 'true');
        } catch(se) {}

        // Record in BadDebts sheet
        var bdSheet = ss.getSheetByName(SHEETS.BAD_DEBTS || 'BadDebts');
        var badDebt = null;
        if (bdSheet) {
          var bdId   = generateId('BD');
          var user   = Session.getActiveUser().getEmail() || 'system';
          var vatElem = isVATReg ? Math.round(vatAmt * (amountDue / (total || 1)) * 100) / 100 : 0;
          bdSheet.appendRow([
            bdId, invoiceId, data[i][1], data[i][2], data[i][3],
            writeOffDate, amountDue, vatElem,
            isVATReg ? 'Eligible' : 'N/A',
            '', reason, user
          ]);
          badDebt = {
            badDebtId: bdId, invoiceId: invoiceId,
            invoiceNumber: data[i][1], clientId: data[i][2], clientName: data[i][3],
            writeOffDate: writeOffDate, amountWrittenOff: amountDue,
            vatElement: vatElem, vatReclaimStatus: isVATReg ? 'Eligible' : 'N/A',
            reason: reason, writtenOffBy: user
          };
        }

        addInvoiceHistory(invoiceId, 'BadDebt', 'status', status, 'Bad Debt', reason, params);
        logAudit('BAD_DEBT', 'Invoice', invoiceId, { date: writeOffDate, amount: amountDue });
        return { success: true, message: 'Invoice written off as bad debt.', badDebt: badDebt };
      }
    }
    return { success: false, message: 'Invoice not found.' };
  } catch(e) {
    Logger.log('writeOffInvoice error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// VOID BILL (mirrors voidInvoice for bills)
// ─────────────────────────────────────────────────────────────────────────────

function voidBill(billId, reason, params) {
  try {
    _auth('bills.write', params);
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.BILLS);
    if (!sheet) return { success: false, message: 'Bills sheet not found' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === billId) {
        var status = (data[i][12] || '').toString();
        if (status === 'Void' || status === 'Voided') {
          return { success: false, message: 'Bill is already voided.' };
        }

        var rowNum = i + 1;
        sheet.getRange(rowNum, 13).setValue('Void');

        // Reverse double-entry if approved
        if (['Approved','Paid','Partially Paid'].indexOf(status) >= 0) {
          try {
            createDoubleEntry(
              new Date(), 'BillVoidReversal', data[i][1] ? data[i][1].toString() : billId,
              '2100', data[i][7] || '5000',
              parseFloat(data[i][9]) || 0,
              'Bill void reversal: ' + reason,
              null, billId, params
            );
          } catch(te) { Logger.log('voidBill double-entry: ' + te); }
        }

        // Write to void log
        var voidSheet = ss.getSheetByName(SHEETS.VOID_LOG || 'VoidLog');
        var voidEntry = null;
        if (voidSheet) {
          var voidId = generateId('VD');
          var user   = Session.getActiveUser().getEmail() || 'system';
          voidSheet.appendRow([
            voidId, billId, data[i][1], data[i][3], data[i][9],
            new Date().toISOString().split('T')[0], user, reason, 'Bill'
          ]);
          voidEntry = {
            voidId: voidId, billId: billId,
            billNumber: data[i][1], supplierName: data[i][3],
            total: data[i][9], voidDate: new Date().toISOString().split('T')[0],
            voidedBy: user, voidReason: reason
          };
        }

        addBillHistory(billId, 'Voided', 'status', status, 'Void', reason, params);
        logAudit('VOID', 'Bill', billId, { reason: reason });
        return { success: true, message: 'Bill voided.', voidEntry: voidEntry };
      }
    }
    return { success: false, message: 'Bill not found.' };
  } catch(e) {
    Logger.log('voidBill error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER: escapeHtml (used by generateInvoiceHTML / generateClientStatement)
// ─────────────────────────────────────────────────────────────────────────────

function escapeHtml(str) {
  return (str || '').toString()
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&#39;');
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER: getOrCreateFolder — used by PDF generation
// ─────────────────────────────────────────────────────────────────────────────

function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER: _checkPeriodLock — prevent writes into closed financial periods
// Safe stub — update with real period lock logic if needed.
// ─────────────────────────────────────────────────────────────────────────────

function _checkPeriodLock(date, params) {
  try {
    var s = getSettings(params);
    if (!s || !s.lockedBefore) return; // no lock set
    var lockDate = new Date(s.lockedBefore);
    var txDate   = date instanceof Date ? date : new Date(date);
    if (!isNaN(lockDate) && !isNaN(txDate) && txDate < lockDate) {
      throw new Error('Period locked: transactions before ' + s.lockedBefore + ' cannot be modified.');
    }
  } catch(e) {
    if (e.message && e.message.indexOf('Period locked') >= 0) throw e;
    // Swallow any other errors (settings unavailable etc.)
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// VAT BAD DEBT RELIEF
// ─────────────────────────────────────────────────────────────────────────────

function markBadDebtVATClaimed(badDebtId, claimDate, params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BAD_DEBTS || 'BadDebts');
    if (!sheet) return { success: false, message: 'BadDebts sheet not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === badDebtId) {
        sheet.getRange(i + 1, 9).setValue('Claimed');
        sheet.getRange(i + 1, 10).setValue(claimDate || new Date().toISOString().split('T')[0]);
        logAudit('VAT_RELIEF_CLAIMED', 'BadDebt', badDebtId, { claimDate: claimDate });
        return { success: true, message: 'VAT relief marked as claimed.' };
      }
    }
    return { success: false, message: 'Bad debt record not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// DELETE CLIENT
// ─────────────────────────────────────────────────────────────────────────────

// deleteClient moved to Contacts.gs
