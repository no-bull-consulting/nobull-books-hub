/**
 * NO~BULL BOOKS — TRANSACTIONS
 * Bills, Credit Notes, Purchase Orders, File Attachments, Bad Debts,
 * VAT calculations, Exchange Rates, WhatsApp, Remittance.
 */

// ─────────────────────────────────────────────────────────────────────────────
// BILLS
// ─────────────────────────────────────────────────────────────────────────────

function createBill(supplierId, lines, issueDate, dueDate, notes, params) {
  try {
    _auth('bills.write', params);
    var ss          = getDb(params || {});
    var billSheet   = ss.getSheetByName(SHEETS.BILLS);
    var linesSheet  = ss.getSheetByName(SHEETS.BILL_LINES);
    var supSheet    = ss.getSheetByName(SHEETS.SUPPLIERS);
    var settings    = getSettings(params);

    if (!billSheet)  return { success: false, message: 'Bills sheet not found' };
    if (!linesSheet) return { success: false, message: 'BillLines sheet not found' };

    // Look up supplier
    var supplier = null;
    if (supSheet) {
      var supData = supSheet.getDataRange().getValues();
      for (var i = 1; i < supData.length; i++) {
        if (supData[i][0] && supData[i][0].toString() === supplierId) {
          supplier = { supplierId: supData[i][0], supplierName: supData[i][1] };
          break;
        }
      }
    }
    if (!supplier) return { success: false, message: 'Supplier not found: ' + supplierId };

    // Calculate totals
    var subtotal = 0, vatTotal = 0;
    for (var j = 0; j < lines.length; j++) {
      var net = parseFloat(lines[j].unitPrice || 0) * parseFloat(lines[j].quantity || 1);
      var vat = net * (parseFloat(lines[j].vatRate || 0) / 100);
      subtotal += net;
      vatTotal += vat;
    }
    var total = subtotal + vatTotal;

    // Generate bill number
    var billNumber = (settings.billPrefix || 'BILL-') + (settings.nextBillNumber || 1).toString().padStart(4, '0');
    var billId = generateId('BILL');
    var today  = issueDate ? new Date(issueDate) : new Date();
    var due    = dueDate   ? new Date(dueDate)   : new Date(today.getTime() + (parseInt(settings.paymentTerms) || 30) * 86400000);

    billSheet.appendRow([
      billId,
      billNumber,
      supplier.supplierId,
      supplier.supplierName,
      safeSerializeDate(today),
      safeSerializeDate(due),
      subtotal,
      lines[0] ? (parseFloat(lines[0].vatRate) || 0) : 0,
      vatTotal,
      total,
      0,        // AmountPaid
      total,    // AmountDue
      'Pending',
      '',       // PaymentDate
      notes || '',
      false,    // Reconciled
      '',       // VoidDate
      '',       // VoidReason
      ''        // VoidedBy
    ]);

    // Write bill lines
    for (var k = 0; k < lines.length; k++) {
      var l   = lines[k];
      var net2 = parseFloat(l.unitPrice || 0) * parseFloat(l.quantity || 1);
      var vat2 = net2 * (parseFloat(l.vatRate || 0) / 100);
      linesSheet.appendRow([
        generateId('BL'),
        billId,
        l.description || '',
        parseFloat(l.quantity) || 1,
        parseFloat(l.unitPrice) || 0,
        parseFloat(l.vatRate) || 0,
        net2 + vat2
      ]);
    }

    // Increment bill number in settings
    settings.nextBillNumber = (settings.nextBillNumber || 1) + 1;
    updateSettings(settings);

    logAudit('CREATE', 'Bill', billId, { billNumber: billNumber }, params);
    return { success: true, billId: billId, billNumber: billNumber };
  } catch(e) {
    Logger.log('createBill error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getBillLines(billId, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BILL_LINES);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, lines: [] };
    var data  = sheet.getDataRange().getValues();
    var lines = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === billId) {
        lines.push({
          lineId:      data[i][0] ? data[i][0].toString() : '',
          billId:      data[i][1] ? data[i][1].toString() : '',
          description: data[i][2] ? data[i][2].toString() : '',
          quantity:    parseFloat(data[i][3]) || 1,
          unitPrice:   parseFloat(data[i][4]) || 0,
          vatRate:     parseFloat(data[i][5]) || 0,
          lineTotal:   parseFloat(data[i][6]) || 0
        });
      }
    }
    return { success: true, lines: lines };
  } catch(e) {
    Logger.log('getBillLines error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function editBill(billId, updates, params) {
  try {
    _auth('bills.write', params);
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.BILLS);
    if (!sheet) return { success: false, message: 'Bills sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === billId) {
        var row = i + 1;
        if (updates.notes      !== undefined) sheet.getRange(row, 15).setValue(updates.notes);
        if (updates.dueDate    !== undefined) sheet.getRange(row, 6).setValue(safeSerializeDate(new Date(updates.dueDate)));
        if (updates.issueDate  !== undefined) sheet.getRange(row, 5).setValue(safeSerializeDate(new Date(updates.issueDate)));
        if (updates.supplierId !== undefined) sheet.getRange(row, 3).setValue(updates.supplierId);
        if (updates.currency   !== undefined) sheet.getRange(row, 17).setValue(updates.currency);
        if (updates.exchangeRate !== undefined) sheet.getRange(row, 18).setValue(parseFloat(updates.exchangeRate)||1);

        // Recalculate totals from lines if provided
        if (updates.lines && updates.lines.length) {
          var lineSheet = ss.getSheetByName(SHEETS.BILL_LINES);
          if (lineSheet) {
            // Delete existing lines for this bill
            var lineData = lineSheet.getDataRange().getValues();
            for (var k = lineData.length - 1; k >= 1; k--) {
              if (lineData[k][1] && lineData[k][1].toString() === billId) {
                lineSheet.deleteRow(k + 1);
              }
            }
            // Write new lines
            var sub = 0, vatAmt = 0;
            updates.lines.forEach(function(l, li) {
              var qty   = parseFloat(l.qty) || 1;
              var price = parseFloat(l.unitPrice) || 0;
              var vat   = parseFloat(l.vatRate) || 0;
              var lSub  = qty * price;
              var lVat  = lSub * vat / 100;
              var lTot  = lSub + lVat;
              sub    += lSub;
              vatAmt += lVat;
              lineSheet.appendRow([
                generateId('BL'), billId,
                l.description || '', qty, price, vat, lTot,
                l.accountCode || '5000'
              ]);
            });
            var total = sub + vatAmt;
            sheet.getRange(row, 7).setValue(sub);       // Subtotal
            sheet.getRange(row, 8).setValue(vatAmt);    // VAT amount
            sheet.getRange(row, 10).setValue(total);    // Total
            sheet.getRange(row, 12).setValue(total);    // AmountDue (reset)
          }
        }
        logAudit('UPDATE', 'Bill', billId, updates, params);
        return { success: true };
      }
    }
    return { success: false, message: 'Bill not found: ' + billId };
  } catch(e) {
    Logger.log('editBill error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function approveBill(billId, params) {
  try {
    _auth('bills.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BILLS);
    if (!sheet) return { success: false, message: 'Bills sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === billId) {
        sheet.getRange(i + 1, 13).setValue('Approved');
        logAudit('UPDATE', 'Bill', billId, { status: 'Approved' }, params);
        return { success: true };
      }
    }
    return { success: false, message: 'Bill not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteBill(billId, params) {
  try {
    _auth('bills.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BILLS);
    if (!sheet) return { success: false, message: 'Bills sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === billId) {
        var status    = data[i][12] ? data[i][12].toString() : '';
        var amountPaid = parseFloat(data[i][10]) || 0;
        // Only block delete if money has already been paid against this bill
        if (amountPaid > 0) {
          return { success: false, message: 'Cannot delete a bill with payments recorded. Void it instead.' };
        }
        // Block delete of Approved/Paid bills — must void first
        if (status === 'Approved' || status === 'Paid') {
          return { success: false, message: 'Cannot delete an '+status+' bill. Void it first.' };
        }
        sheet.deleteRow(i + 1);
        logAudit('DELETE', 'Bill', billId, {}, params);
        return { success: true };
      }
    }
    return { success: false, message: 'Bill not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function adjustBillBalance(billId, amount, params) {
  try {
    _auth('bills.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.BILLS);
    if (!sheet) return { success: false, message: 'Bills sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === billId) {
        var row = i + 1;
        var paid = parseFloat(data[i][10]) || 0;
        var total = parseFloat(data[i][9]) || 0;
        var newPaid = Math.min(paid + amount, total);
        var newDue  = Math.max(total - newPaid, 0);
        sheet.getRange(row, 11).setValue(newPaid);
        sheet.getRange(row, 12).setValue(newDue);
        if (newDue <= 0) {
          sheet.getRange(row, 13).setValue('Paid');
          sheet.getRange(row, 14).setValue(new Date());
        } else if (newPaid > 0) {
          sheet.getRange(row, 13).setValue('Partial');
        }
        return { success: true, amountPaid: newPaid, amountDue: newDue };
      }
    }
    return { success: false, message: 'Bill not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function adjustInvoiceBalance(invoiceId, amount, params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.INVOICES);
    if (!sheet) return { success: false, message: 'Invoices sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === invoiceId) {
        var row   = i + 1;
        var paid  = parseFloat(data[i][12]) || 0;
        var total = parseFloat(data[i][11]) || 0;
        var newPaid = Math.min(paid + amount, total);
        var newDue  = Math.max(total - newPaid, 0);
        sheet.getRange(row, 13).setValue(newPaid);
        sheet.getRange(row, 14).setValue(newDue);
        if (newDue <= 0) {
          sheet.getRange(row, 15).setValue('Paid');
          sheet.getRange(row, 16).setValue(new Date());
        } else if (newPaid > 0) {
          sheet.getRange(row, 15).setValue('Partial');
        }
        return { success: true, amountPaid: newPaid, amountDue: newDue };
      }
    }
    return { success: false, message: 'Invoice not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// CREDIT NOTES
// ─────────────────────────────────────────────────────────────────────────────

function createCreditNote(invoiceId, lines, reason, issueDate, params) {
  try {
    _auth('invoices.write', params);
    var ss       = getDb(params || {});
    var cnSheet  = ss.getSheetByName(SHEETS.CREDIT_NOTES);
    var cnlSheet = ss.getSheetByName(SHEETS.CREDIT_NOTE_LINES);
    var settings = getSettings(params);
    if (!cnSheet) return { success: false, message: 'CreditNotes sheet not found' };

    var invoice = getInvoiceById(invoiceId, params);
    if (!invoice) return { success: false, message: 'Invoice not found: ' + invoiceId };

    var subtotal = 0, vatTotal = 0;
    for (var j = 0; j < lines.length; j++) {
      var net = parseFloat(lines[j].unitPrice || 0) * parseFloat(lines[j].quantity || 1);
      var vat = net * (parseFloat(lines[j].vatRate || 0) / 100);
      subtotal += net; vatTotal += vat;
    }
    var total = subtotal + vatTotal;

    var cnNumber = (settings.cnPrefix || 'CN-') + (settings.nextCNNumber || 1).toString().padStart(4, '0');
    var cnId     = generateId('CN');
    var today    = issueDate ? new Date(issueDate) : new Date();

    cnSheet.appendRow([
      cnId, cnNumber, invoiceId,
      invoice.clientId, invoice.clientName,
      safeSerializeDate(today),
      subtotal, vatTotal, total,
      'Open', reason || '', '', ''
    ]);

    if (cnlSheet) {
      for (var k = 0; k < lines.length; k++) {
        var l = lines[k];
        var net2 = parseFloat(l.unitPrice || 0) * parseFloat(l.quantity || 1);
        cnlSheet.appendRow([
          generateId('CNL'), cnId,
          l.description || '', parseFloat(l.quantity) || 1,
          parseFloat(l.unitPrice) || 0, parseFloat(l.vatRate) || 0,
          net2 + net2 * (parseFloat(l.vatRate || 0) / 100)
        ]);
      }
    }

    settings.nextCNNumber = (settings.nextCNNumber || 1) + 1;
    updateSettings(settings);

    logAudit('CREATE', 'CreditNote', cnId, { cnNumber: cnNumber }, params);
    return { success: true, cnId: cnId, cnNumber: cnNumber };
  } catch(e) {
    Logger.log('createCreditNote error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function applyCreditNote(cnId, invoiceId, params) {
  try {
    _auth('invoices.write', params);
    var ss      = getDb(params || {});
    var cnSheet = ss.getSheetByName(SHEETS.CREDIT_NOTES);
    if (!cnSheet) return { success: false, message: 'CreditNotes sheet not found' };
    var data = cnSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === cnId) {
        var amount = parseFloat(data[i][8]) || 0;
        var row    = i + 1;
        cnSheet.getRange(row, 10).setValue('Applied');
        cnSheet.getRange(row, 12).setValue(new Date());
        cnSheet.getRange(row, 13).setValue(invoiceId);
        adjustInvoiceBalance(invoiceId, amount, params);
        logAudit('UPDATE', 'CreditNote', cnId, { applied: invoiceId }, params);
        return { success: true };
      }
    }
    return { success: false, message: 'Credit note not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function voidCreditNote(cnId, params) {
  try {
    _auth('invoices.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.CREDIT_NOTES);
    if (!sheet) return { success: false, message: 'CreditNotes sheet not found' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === cnId) {
        sheet.getRange(i + 1, 10).setValue('Void');
        logAudit('UPDATE', 'CreditNote', cnId, { status: 'Void' }, params);
        return { success: true };
      }
    }
    return { success: false, message: 'Credit note not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// PURCHASE ORDERS
// ─────────────────────────────────────────────────────────────────────────────

function createPurchaseOrder(supplierId, lines, expectedDelivery, notes, params) {
  try {
    _auth('purchaseorders.write', params);
    var ss       = getDb(params || {});
    var poSheet  = ss.getSheetByName(SHEETS.PURCHASE_ORDERS);
    var polSheet = ss.getSheetByName(SHEETS.PURCHASE_ORDER_LINES);
    var supSheet = ss.getSheetByName(SHEETS.SUPPLIERS);
    var settings = getSettings(params);
    if (!poSheet) return { success: false, message: 'PurchaseOrders sheet not found' };

    var supplier = null;
    if (supSheet) {
      var supData = supSheet.getDataRange().getValues();
      for (var i = 1; i < supData.length; i++) {
        if (supData[i][0] && supData[i][0].toString() === supplierId) {
          supplier = { supplierId: supData[i][0], supplierName: supData[i][1] };
          break;
        }
      }
    }
    if (!supplier) return { success: false, message: 'Supplier not found: ' + supplierId };

    var subtotal = 0, vatTotal = 0;
    for (var j = 0; j < lines.length; j++) {
      var net = parseFloat(lines[j].unitPrice || 0) * parseFloat(lines[j].quantity || 1);
      subtotal += net;
      vatTotal += net * (parseFloat(lines[j].vatRate || 0) / 100);
    }
    var total  = subtotal + vatTotal;
    var poNumber = (settings.poPrefix || 'PO-') + (settings.nextPONumber || 1).toString().padStart(4, '0');
    var poId   = generateId('PO');

    poSheet.appendRow([
      poId, poNumber,
      supplier.supplierId, supplier.supplierName,
      safeSerializeDate(new Date()),
      expectedDelivery ? safeSerializeDate(new Date(expectedDelivery)) : '',
      subtotal, vatTotal, total,
      'Draft', notes || '', '', ''
    ]);

    if (polSheet) {
      for (var k = 0; k < lines.length; k++) {
        var l = lines[k];
        var net2 = parseFloat(l.unitPrice || 0) * parseFloat(l.quantity || 1);
        polSheet.appendRow([
          generateId('POL'), poId,
          l.description || '', parseFloat(l.quantity) || 1,
          parseFloat(l.unitPrice) || 0, parseFloat(l.vatRate) || 0,
          net2 + net2 * (parseFloat(l.vatRate || 0) / 100)
        ]);
      }
    }

    settings.nextPONumber = (settings.nextPONumber || 1) + 1;
    updateSettings(settings);

    logAudit('CREATE', 'PurchaseOrder', poId, { poNumber: poNumber }, params);
    return { success: true, poId: poId, poNumber: poNumber };
  } catch(e) {
    Logger.log('createPurchaseOrder error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getPurchaseOrderLines(poId, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.PURCHASE_ORDER_LINES);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, lines: [] };
    var data = sheet.getDataRange().getValues();
    var lines = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === poId) {
        lines.push({
          lineId: data[i][0] ? data[i][0].toString() : '',
          poId: data[i][1] ? data[i][1].toString() : '',
          description: data[i][2] ? data[i][2].toString() : '',
          quantity: parseFloat(data[i][3]) || 1,
          unitPrice: parseFloat(data[i][4]) || 0,
          vatRate: parseFloat(data[i][5]) || 0,
          lineTotal: parseFloat(data[i][6]) || 0
        });
      }
    }
    return { success: true, lines: lines };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function _updatePOStatus(poId, status, params) {
  var sheet = getDb(params || {}).getSheetByName(SHEETS.PURCHASE_ORDERS);
  if (!sheet) return { success: false, message: 'PurchaseOrders sheet not found' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === poId) {
      sheet.getRange(i + 1, 10).setValue(status);
      logAudit('UPDATE', 'PurchaseOrder', poId, { status: status }, params);
      return { success: true };
    }
  }
  return { success: false, message: 'PO not found' };
}

function submitPurchaseOrderForApproval(poId, params) {
  _auth('purchaseorders.write', params);
  return _updatePOStatus(poId, 'Pending Approval', params);
}

function approvePurchaseOrder(poId, params) {
  _auth('purchaseorders.write', params);
  var sheet = getDb(params || {}).getSheetByName(SHEETS.PURCHASE_ORDERS);
  if (!sheet) return { success: false, message: 'PurchaseOrders sheet not found' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === poId) {
      var ctx = _getCurrentUserContext(params);
      sheet.getRange(i + 1, 10).setValue('Approved');
      sheet.getRange(i + 1, 12).setValue(ctx.email);
      logAudit('UPDATE', 'PurchaseOrder', poId, { status: 'Approved' }, params);
      return { success: true };
    }
  }
  return { success: false, message: 'PO not found' };
}

function cancelPurchaseOrder(poId, params) {
  _auth('purchaseorders.write', params);
  return _updatePOStatus(poId, 'Cancelled', params);
}

function markPurchaseOrderSent(poId, params) {
  _auth('purchaseorders.write', params);
  return _updatePOStatus(poId, 'Sent', params);
}

function markPurchaseOrderPartial(poId, params) {
  _auth('purchaseorders.write', params);
  return _updatePOStatus(poId, 'Partial', params);
}

function receivePurchaseOrder(poId, params) {
  _auth('purchaseorders.write', params);
  return _updatePOStatus(poId, 'Received', params);
}

// ─────────────────────────────────────────────────────────────────────────────
// FILE ATTACHMENTS — Invoices and Bills
// ─────────────────────────────────────────────────────────────────────────────

function _getFileSheet(type, params) {
  var sheetName = type === 'invoice' ? SHEETS.INVOICE_FILES : SHEETS.BILL_FILES;
  return getDb(params || {}).getSheetByName(sheetName);
}

function getInvoiceFiles(invoiceId, params) {
  try {
    var sheet = _getFileSheet('invoice', params);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, files: [] };
    var data = sheet.getDataRange().getValues();
    var files = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === invoiceId) {
        files.push({
          fileId: data[i][0] ? data[i][0].toString() : '',
          invoiceId: data[i][1] ? data[i][1].toString() : '',
          fileName: data[i][2] ? data[i][2].toString() : '',
          fileURL: data[i][3] ? data[i][3].toString() : '',
          fileUrl: data[i][3] ? data[i][3].toString() : '',  // alias for frontend
          fileType: data[i][4] ? data[i][4].toString() : '',
          fileSize: '',  // not stored, placeholder for UI
          uploadedBy: data[i][5] ? data[i][5].toString() : '',
          uploadedDate: data[i][6] ? safeSerializeDate(data[i][6]) : '',
          description: data[i][7] ? data[i][7].toString() : ''
        });
      }
    }
    return { success: true, files: files };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function uploadInvoiceFile(params) {
  try {
    _auth('invoices.write', params);
    return _uploadFile('invoice', params.invoiceId, params);
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteInvoiceFile(params) {
  try {
    _auth('invoices.write', params);
    return _deleteFile('invoice', params.fileId, params);
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getBillFiles(billId, params) {
  try {
    var sheet = _getFileSheet('bill', params);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, files: [] };
    var data = sheet.getDataRange().getValues();
    var files = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === billId) {
        files.push({
          fileId: data[i][0] ? data[i][0].toString() : '',
          billId: data[i][1] ? data[i][1].toString() : '',
          fileName: data[i][2] ? data[i][2].toString() : '',
          fileURL: data[i][3] ? data[i][3].toString() : '',
          fileUrl: data[i][3] ? data[i][3].toString() : '',  // alias for frontend
          fileType: data[i][4] ? data[i][4].toString() : '',
          fileSize: '',  // not stored, placeholder for UI
          uploadedBy: data[i][5] ? data[i][5].toString() : '',
          uploadedDate: data[i][6] ? safeSerializeDate(data[i][6]) : '',
          description: data[i][7] ? data[i][7].toString() : ''
        });
      }
    }
    return { success: true, files: files };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function uploadBillFile(params) {
  try {
    _auth('bills.write', params);
    return _uploadFile('bill', params.billId, params);
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteBillFile(params) {
  try {
    _auth('bills.write', params);
    return _deleteFile('bill', params.fileId, params);
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function _uploadFile(type, parentId, params) {
  if (!params.base64Data) return { success: false, message: 'No file data provided' };
  var fileName = params.fileName || 'attachment';
  var fileType = params.fileType || 'application/octet-stream';
  var blob     = Utilities.newBlob(Utilities.base64Decode(params.base64Data), fileType, fileName);
  var folder   = getOrCreateFolder('no~bull books — Attachments');
  var file     = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var fileId   = file.getId();
  var fileURL  = file.getUrl();
  var ctx      = _getCurrentUserContext(params);

  var sheetName = type === 'invoice' ? SHEETS.INVOICE_FILES : SHEETS.BILL_FILES;
  var sheet     = getDb(params || {}).getSheetByName(sheetName);
  if (!sheet) return { success: false, message: sheetName + ' sheet not found' };

  sheet.appendRow([
    generateId('FILE'), parentId, fileName, fileURL, fileType,
    ctx.email, new Date(), params.description || ''
  ]);

  return { success: true, fileId: fileId, fileURL: fileURL, fileName: fileName };
}

function _deleteFile(type, fileId, params) {
  var sheetName = type === 'invoice' ? SHEETS.INVOICE_FILES : SHEETS.BILL_FILES;
  var sheet = getDb(params || {}).getSheetByName(sheetName);
  if (!sheet) return { success: false, message: sheetName + ' sheet not found' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === fileId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'File not found' };
}

// ─────────────────────────────────────────────────────────────────────────────
// BAD DEBTS
// ─────────────────────────────────────────────────────────────────────────────

function writeOffBadDebt(invoiceId, reason, params) {
  try {
    _auth('invoices.write', params);
    var ss      = getDb(params || {});
    var invSheet = ss.getSheetByName(SHEETS.INVOICES);
    var bdSheet  = ss.getSheetByName(SHEETS.BAD_DEBTS);
    var settings = getSettings(params);

    if (!invSheet) return { success: false, message: 'Invoices sheet not found' };

    var invoice = getInvoiceById(invoiceId, params);
    if (!invoice) return { success: false, message: 'Invoice not found' };
    if (invoice.amountDue <= 0) return { success: false, message: 'Invoice has no outstanding balance' };

    // Mark invoice as Bad Debt
    var invData = invSheet.getDataRange().getValues();
    for (var i = 1; i < invData.length; i++) {
      if (invData[i][0] && invData[i][0].toString() === invoiceId) {
        invSheet.getRange(i + 1, 15).setValue('Bad Debt');
        break;
      }
    }

    // Record in BadDebts sheet
    if (bdSheet) {
      var vatElement = invoice.amountDue * (parseFloat(settings.vatRate) || 20) / (100 + (parseFloat(settings.vatRate) || 20));
      var ctx = _getCurrentUserContext(params);
      bdSheet.appendRow([
        generateId('BD'),
        invoiceId,
        invoice.invoiceNumber,
        invoice.clientId,
        invoice.clientName,
        new Date(),
        invoice.amountDue,
        settings.vatRegistered ? vatElement : 0,
        'Eligible',
        '', // VATClaimDate
        reason || 'Bad debt write-off',
        ctx.email
      ]);
    }

    logAudit('UPDATE', 'Invoice', invoiceId, { status: 'Bad Debt', reason: reason }, params);
    return { success: true, message: 'Invoice written off as bad debt' };
  } catch(e) {
    Logger.log('writeOffBadDebt error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function checkBadDebtVATEligibility(invoiceId, params) {
  try {
    var invoice  = getInvoiceById(invoiceId, params);
    if (!invoice) return { success: false, message: 'Invoice not found' };
    var settings = getSettings(params);
    var issueDate = new Date(invoice.issueDate);
    var now       = new Date();
    var daysSince = Math.floor((now - issueDate) / 86400000);
    var eligible  = daysSince >= 180 && settings.vatRegistered && invoice.amountDue > 0;
    return {
      success: true,
      eligible: eligible,
      daysSinceIssue: daysSince,
      minimumDays: 180,
      vatElement: eligible
        ? invoice.amountDue * (parseFloat(settings.vatRate) || 20) / (100 + (parseFloat(settings.vatRate) || 20))
        : 0,
      message: eligible
        ? 'Eligible for VAT bad debt relief'
        : 'Not yet eligible — must be 6+ months overdue and VAT registered'
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// VAT CALCULATION
// ─────────────────────────────────────────────────────────────────────────────

function calculateVATReturn(fromDate, toDate, params) {
  try {
    _auth('mtd.read', params);
    var ss       = getDb(params || {});
    var invSheet = ss.getSheetByName(SHEETS.INVOICES);
    var bilSheet = ss.getSheetByName(SHEETS.BILLS);
    var settings = getSettings(params);

    var from = new Date(fromDate);
    var to   = new Date(toDate);
    to.setHours(23, 59, 59);

    var salesVAT = 0, salesNet = 0;
    var purchVAT = 0, purchNet = 0;

    if (invSheet && invSheet.getLastRow() > 1) {
      var invData = invSheet.getDataRange().getValues();
      for (var i = 1; i < invData.length; i++) {
        var invDate = new Date(invData[i][6]);
        var status  = (invData[i][14] || '').toString();
        if (invDate >= from && invDate <= to && status !== 'Void' && status !== 'Draft') {
          salesNet += parseFloat(invData[i][8]) || 0;
          salesVAT += parseFloat(invData[i][10]) || 0;
        }
      }
    }

    if (bilSheet && bilSheet.getLastRow() > 1) {
      var bilData = bilSheet.getDataRange().getValues();
      for (var j = 1; j < bilData.length; j++) {
        var bilDate = new Date(bilData[j][4]);
        var bilStatus = (bilData[j][12] || '').toString();
        if (bilDate >= from && bilDate <= to && bilStatus !== 'Void') {
          purchNet += parseFloat(bilData[j][6]) || 0;
          purchVAT += parseFloat(bilData[j][8]) || 0;
        }
      }
    }

    var vatDue    = salesVAT - purchVAT;
    return {
      success: true,
      box1: Math.round(salesVAT * 100) / 100,    // VAT on sales
      box4: Math.round(purchVAT * 100) / 100,    // VAT on purchases
      box5: Math.round(vatDue * 100) / 100,      // Net VAT payable
      box6: Math.round(salesNet * 100) / 100,    // Total sales (ex VAT)
      box7: Math.round(purchNet * 100) / 100,    // Total purchases (ex VAT)
      box2: 0, box3: Math.round(salesVAT * 100) / 100,
      box8: 0, box9: 0,
      fromDate: fromDate, toDate: toDate
    };
  } catch(e) {
    Logger.log('calculateVATReturn error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// EXCHANGE RATES
// ─────────────────────────────────────────────────────────────────────────────

function getExchangeRates(params) {
  try {
    var settings  = getSettings(params);
    var base      = settings.baseCurrency || 'GBP';
    var enabled   = settings.enabledCurrencies || ['GBP', 'EUR', 'USD'];
    var rates     = {};
    rates[base]   = 1;

    // Fetch live rates from exchangerate-api (free tier)
    try {
      var url  = 'https://open.er-api.com/v6/latest/' + base;
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        var data = JSON.parse(resp.getContentText());
        if (data.rates) {
          for (var i = 0; i < enabled.length; i++) {
            if (data.rates[enabled[i]]) rates[enabled[i]] = data.rates[enabled[i]];
          }
        }
      }
    } catch(fetchErr) {
      Logger.log('Exchange rate fetch failed: ' + fetchErr);
    }

    return { success: true, base: base, rates: rates };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getAvailableRates(params) {
  return getExchangeRates(params);
}

// ─────────────────────────────────────────────────────────────────────────────
// REMITTANCE ADVICE
// ─────────────────────────────────────────────────────────────────────────────

function generateRemittanceAdvice(billIds, params) {
  try {
    _auth('bills.read', params);
    var settings = getSettings(params);
    var bills    = [];
    for (var i = 0; i < billIds.length; i++) {
      var sheet = getDb(params || {}).getSheetByName(SHEETS.BILLS);
      if (!sheet) continue;
      var data = sheet.getDataRange().getValues();
      for (var j = 1; j < data.length; j++) {
        if (data[j][0] && data[j][0].toString() === billIds[i]) {
          bills.push({
            billNumber:   data[j][1],
            supplierName: data[j][3],
            issueDate:    safeSerializeDate(data[j][4]),
            dueDate:      safeSerializeDate(data[j][5]),
            total:        parseFloat(data[j][9]) || 0,
            amountDue:    parseFloat(data[j][11]) || 0
          });
          break;
        }
      }
    }

    var total = bills.reduce(function(sum, b) { return sum + b.amountDue; }, 0);

    return {
      success: true,
      remittance: {
        companyName:  settings.companyName,
        companyEmail: settings.companyEmail,
        date:         new Date().toISOString(),
        bills:        bills,
        totalAmount:  total
      }
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SA103 / CAPITAL ALLOWANCES (stubs — full implementation in roadmap)
// ─────────────────────────────────────────────────────────────────────────────

function getSA103Data(params) {
  return { success: true, data: {}, message: 'SA103 module coming soon' };
}

function saveSATaxAdjustments(params) {
  return { success: true, message: 'SA103 adjustments saved (stub)' };
}

function getCapitalAllowances(params) {
  return { success: true, allowances: [], message: 'Capital allowances module coming soon' };
}

function saveCapitalAllowance(params) {
  return { success: true, message: 'Capital allowance saved (stub)' };
}

function deleteCapitalAllowance(params) {
  return { success: true, message: 'Capital allowance deleted (stub)' };
}

// ─────────────────────────────────────────────────────────────────────────────
// WHATSAPP (stubs — credentials set via GAS editor)
// ─────────────────────────────────────────────────────────────────────────────

function getWhatsAppLink(params) {
  return { success: false, message: 'WhatsApp not configured' };
}

function sendInvoiceWhatsApp(params) {
  return { success: false, message: 'WhatsApp delivery not yet configured. Set up credentials in Settings → WhatsApp.' };
}
