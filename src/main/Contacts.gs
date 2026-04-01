/**
 * NO~BULL BOOKS — CONTACTS
 * Client and Supplier CRUD operations.
 *
 * Sheet schemas (from Initializer.gs):
 *   Clients:   ClientId, Name, Email, Phone, Address, Postcode, Country,
 *              VATNumber, ContactName, Notes, CreatedDate, Active
 *   Suppliers: SupplierId, Name, Email, Phone, Address, Postcode, Country,
 *              VATNumber, ContactName, Notes, CreatedDate, Active
 */

// ─────────────────────────────────────────────────────────────────────────────
// CLIENTS
// ─────────────────────────────────────────────────────────────────────────────

function getAllClients(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.CLIENTS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, clients: [] };

    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
    var clients = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue; // skip empty rows
      clients.push({
        clientId:    r[0]  ? r[0].toString()  : '',
        name:        r[1]  ? r[1].toString()  : '',
        email:       r[2]  ? r[2].toString()  : '',
        phone:       r[3]  ? r[3].toString()  : '',
        address:     r[4]  ? r[4].toString()  : '',
        postcode:    r[5]  ? r[5].toString()  : '',
        country:     r[6]  ? r[6].toString()  : '',
        vatNumber:   r[7]  ? r[7].toString()  : '',
        contactName: r[8]  ? r[8].toString()  : '',
        notes:       r[9]  ? r[9].toString()  : '',
        createdDate: r[10] ? r[10].toString() : '',
        active:      r[11] !== false && r[11] !== 'FALSE' && r[11] !== ''
      });
    }
    return { success: true, clients: clients };
  } catch(e) {
    Logger.log('getAllClients error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function createClient(params) {
  try {
    _auth('clients.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.CLIENTS);
    if (!sheet) return { success: false, message: 'Clients sheet not found' };

    var clientId = generateId('CLI');
    var name = params.clientName || params.name || '';
    sheet.appendRow([
      clientId,
      name,
      params.email          || '',
      params.phone          || '',
      params.address        || '',
      params.postcode       || '',
      params.country        || 'UK',
      params.vatRegNumber   || params.vatNumber || '',
      params.contactName    || '',
      params.notes          || '',
      new Date(),
      true
    ]);

    logAudit('CREATE', 'Client', clientId, { name: name }, params);
    return { success: true, clientId: clientId, clientName: name };
  } catch(e) {
    Logger.log('createClient error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function updateClient(clientId, params) {
  try {
    _auth('clients.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.CLIENTS);
    if (!sheet) return { success: false, message: 'Clients sheet not found' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === clientId) {
        var row = i + 1;
        var name = params.clientName || params.name;
        if (name              !== undefined) sheet.getRange(row, 2).setValue(name);
        if (params.email       !== undefined) sheet.getRange(row, 3).setValue(params.email);
        if (params.phone       !== undefined) sheet.getRange(row, 4).setValue(params.phone);
        if (params.address     !== undefined) sheet.getRange(row, 5).setValue(params.address);
        if (params.postcode    !== undefined) sheet.getRange(row, 6).setValue(params.postcode);
        if (params.country     !== undefined) sheet.getRange(row, 7).setValue(params.country);
        var vat = params.vatRegNumber || params.vatNumber;
        if (vat               !== undefined) sheet.getRange(row, 8).setValue(vat);
        if (params.contactName !== undefined) sheet.getRange(row, 9).setValue(params.contactName);
        if (params.notes       !== undefined) sheet.getRange(row, 10).setValue(params.notes);
        if (params.active      !== undefined) sheet.getRange(row, 12).setValue(params.active);
        logAudit('UPDATE', 'Client', clientId, { name: name }, params);
        return { success: true, clientId: clientId };
      }
    }
    return { success: false, message: 'Client not found: ' + clientId };
  } catch(e) {
    Logger.log('updateClient error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function deleteClient(clientId, params) {
  try {
    _auth('clients.write', params);
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.CLIENTS);
    if (!sheet) return { success: false, message: 'Clients sheet not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === clientId) {
        // Safety check — refuse if client has outstanding invoices
        var invSheet = ss.getSheetByName(SHEETS.INVOICES);
        if (invSheet) {
          var invData = invSheet.getDataRange().getValues();
          for (var j = 1; j < invData.length; j++) {
            if (invData[j][2] && invData[j][2].toString() === clientId) {
              var status = (invData[j][14] || '').toString();
              var due    = parseFloat(invData[j][13]) || 0;
              if (due > 0 && status !== 'Void') {
                return { success: false, message: 'Cannot delete client with outstanding invoices (£' + due.toFixed(2) + ' due). Void or write off all invoices first.' };
              }
            }
          }
        }
        var clientName = data[i][1] ? data[i][1].toString() : clientId;
        sheet.deleteRow(i + 1);
        logAudit('DELETE', 'Client', clientId, { name: clientName }, params);
        return { success: true, message: 'Client "' + clientName + '" deleted.' };
      }
    }
    return { success: false, message: 'Client not found.' };
  } catch(e) {
    Logger.log('deleteClient error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SUPPLIERS
// ─────────────────────────────────────────────────────────────────────────────

function getAllSuppliers(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.SUPPLIERS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, suppliers: [] };

    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
    var suppliers = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      suppliers.push({
        supplierId:  r[0]  ? r[0].toString()  : '',
        name:        r[1]  ? r[1].toString()  : '',
        email:       r[2]  ? r[2].toString()  : '',
        phone:       r[3]  ? r[3].toString()  : '',
        address:     r[4]  ? r[4].toString()  : '',
        postcode:    r[5]  ? r[5].toString()  : '',
        country:     r[6]  ? r[6].toString()  : '',
        vatNumber:   r[7]  ? r[7].toString()  : '',
        contactName: r[8]  ? r[8].toString()  : '',
        notes:       r[9]  ? r[9].toString()  : '',
        createdDate: r[10] ? r[10].toString() : '',
        active:      r[11] !== false && r[11] !== 'FALSE' && r[11] !== ''
      });
    }
    return { success: true, suppliers: suppliers };
  } catch(e) {
    Logger.log('getAllSuppliers error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function createSupplier(params) {
  try {
    _auth('suppliers.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.SUPPLIERS);
    if (!sheet) return { success: false, message: 'Suppliers sheet not found' };

    var supplierId = generateId('SUP');
    var name = params.supplierName || params.name || '';
    sheet.appendRow([
      supplierId,
      name,
      params.email          || '',
      params.phone          || '',
      params.address        || '',
      params.postcode       || '',
      params.country        || 'UK',
      params.vatRegNumber   || params.vatNumber || '',
      params.contactName    || '',
      params.notes          || '',
      new Date(),
      true
    ]);

    logAudit('CREATE', 'Supplier', supplierId, { name: name }, params);
    return { success: true, supplierId: supplierId, supplierName: name };
  } catch(e) {
    Logger.log('createSupplier error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function updateSupplier(supplierId, params) {
  try {
    _auth('suppliers.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.SUPPLIERS);
    if (!sheet) return { success: false, message: 'Suppliers sheet not found' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === supplierId) {
        var row = i + 1;
        var name = params.supplierName || params.name;
        if (name               !== undefined) sheet.getRange(row, 2).setValue(name);
        if (params.email       !== undefined) sheet.getRange(row, 3).setValue(params.email);
        if (params.phone       !== undefined) sheet.getRange(row, 4).setValue(params.phone);
        if (params.address     !== undefined) sheet.getRange(row, 5).setValue(params.address);
        if (params.postcode    !== undefined) sheet.getRange(row, 6).setValue(params.postcode);
        if (params.country     !== undefined) sheet.getRange(row, 7).setValue(params.country);
        var vat = params.vatRegNumber || params.vatNumber;
        if (vat                !== undefined) sheet.getRange(row, 8).setValue(vat);
        if (params.contactName !== undefined) sheet.getRange(row, 9).setValue(params.contactName);
        if (params.notes       !== undefined) sheet.getRange(row, 10).setValue(params.notes);
        if (params.active      !== undefined) sheet.getRange(row, 12).setValue(params.active);
        logAudit('UPDATE', 'Supplier', supplierId, { name: name }, params);
        return { success: true, supplierId: supplierId };
      }
    }
    return { success: false, message: 'Supplier not found: ' + supplierId };
  } catch(e) {
    Logger.log('updateSupplier error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function deleteSupplier(supplierId, params) {
  try {
    _auth('suppliers.write', params);
    var sheet = getDb(params || {}).getSheetByName(SHEETS.SUPPLIERS);
    if (!sheet) return { success: false, message: 'Suppliers sheet not found' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === supplierId) {
        sheet.getRange(i + 1, 12).setValue(false);
        logAudit('DELETE', 'Supplier', supplierId, {}, params);
        return { success: true };
      }
    }
    return { success: false, message: 'Supplier not found: ' + supplierId };
  } catch(e) {
    Logger.log('deleteSupplier error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}