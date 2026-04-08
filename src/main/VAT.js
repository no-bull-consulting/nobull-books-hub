/**
 * NO~BULL BOOKS -- VAT
 * Local VAT return storage (save drafts, read history).
 * MTD submission is handled by the HMRC integration layer.
 * -----------------------------------------------------------------------------
 */

/**
 * getVATReturns(params)
 * Returns all saved VAT returns for the client spreadsheet.
 */
function getVATReturns(params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.VAT_RETURNS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, returns: [] };
    }

    var data    = sheet.getDataRange().getValues();
    var returns = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      returns.push({
        returnId:      row[0]  ? row[0].toString()  : '',
        periodStart:   safeSerializeDate(row[1]),
        periodEnd:     safeSerializeDate(row[2]),
        box1:          parseFloat(row[3])  || 0,
        box2:          parseFloat(row[4])  || 0,
        box3:          parseFloat(row[5])  || 0,
        box4:          parseFloat(row[6])  || 0,
        box5:          parseFloat(row[7])  || 0,
        box6:          parseFloat(row[8])  || 0,
        box7:          parseFloat(row[9])  || 0,
        box8:          parseFloat(row[10]) || 0,
        box9:          parseFloat(row[11]) || 0,
        status:        row[12] ? row[12].toString() : 'Draft',
        submittedDate: safeSerializeDate(row[13]),
        periodKey:     row[14] ? row[14].toString() : '',
        // Aliases for frontend display
        outputVAT:     parseFloat(row[3])  || 0,
        inputVAT:      parseFloat(row[6])  || 0,
        netVAT:        parseFloat(row[7])  || 0,
        totalSales:    parseFloat(row[8])  || 0,
        savedDate:     safeSerializeDate(row[13])
      });
    }

    // Most recent first
    returns.sort(function(a, b) {
      return (b.periodEnd || '') > (a.periodEnd || '') ? 1 : -1;
    });

    return { success: true, returns: returns };
  } catch(e) {
    Logger.log('getVATReturns error: ' + e.toString());
    return { success: false, message: e.toString(), returns: [] };
  }
}

/**
 * saveVATReturn(params)
 * Saves a calculated VAT return as a draft to the VATReturns sheet.
 * params should contain all box values plus periodStart/periodEnd/periodKey.
 */
function saveVATReturn(params) {
  try {
    _auth('reports.tax', params);

    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.VAT_RETURNS);
    if (!sheet) return { success: false, message: 'VATReturns sheet not found -- run initial setup.' };

    var d = params;

    // Check if a return for this period already exists -- update rather than duplicate
    var existingRow = -1;
    if (sheet.getLastRow() > 1) {
      var existing = sheet.getDataRange().getValues();
      for (var i = 1; i < existing.length; i++) {
        var rowStart = existing[i][1] ? safeSerializeDate(existing[i][1]) : '';
        var rowEnd   = existing[i][2] ? safeSerializeDate(existing[i][2]) : '';
        if (rowStart === (d.periodStart || '') && rowEnd === (d.periodEnd || '')) {
          existingRow = i + 1;
          break;
        }
      }
    }

    var returnId = existingRow > 0
      ? sheet.getRange(existingRow, 1).getValue().toString()
      : generateId('VAT');

    var row = [
      returnId,
      d.periodStart   || '',
      d.periodEnd     || '',
      parseFloat(d.box1 || d.outputVAT || 0),
      parseFloat(d.box2 || 0),
      parseFloat(d.box3 || (d.outputVAT || 0) + (d.box2 || 0)),
      parseFloat(d.box4 || d.inputVAT  || 0),
      parseFloat(d.box5 || d.netVAT    || 0),
      parseFloat(d.box6 || d.totalSales      || 0),
      parseFloat(d.box7 || d.totalPurchases  || 0),
      parseFloat(d.box8 || 0),
      parseFloat(d.box9 || 0),
      d.status    || 'Draft',
      d.submittedDate || '',
      d.periodKey || ''
    ];

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    logAudit('SAVE', 'VATReturn', returnId, { period: d.periodStart + ' to ' + d.periodEnd });
    return { success: true, returnId: returnId, message: 'VAT return saved.' };

  } catch(e) {
    Logger.log('saveVATReturn error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}