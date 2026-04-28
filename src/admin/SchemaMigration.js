/**
 * NO~BULL ADMIN — SCHEMA MIGRATION
 * Upgrades all client instances to the latest schema defined in Initializer.js.
 */

function migrateAllClientSchemas() {
  var registry = getAllRegistryClients({}).clients || [];
  var results = [];

  registry.forEach(function(client) {
    if (!client.sheetId || client.status === 'Cancelled') return;
    
    try {
      var upgradeLog = upgradeClientSchema(client.sheetId);
      results.push({
        company: client.companyName,
        status: 'Success',
        actions: upgradeLog
      });
    } catch (e) {
      results.push({
        company: client.companyName,
        status: 'Error',
        error: e.toString()
      });
    }
  });

  return { success: true, results: results };
}

/**
 * Upgrades a single client sheet.
 */
function upgradeClientSchema(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var actions = [];
  
  // 1. Ensure new sheets exist (e.g., FXRatesLog, AuditLog)
  var REQUIRED = {
    'FXRatesLog': ['LogId','Date','BaseCurrency','TargetCurrency','Rate','Source'],
    'AuditLog': ['AuditId','Timestamp','Action','Entity','EntityId','User','Detail'],
    'FixedAssets': ['AssetId','Name','Category','PurchaseDate','Cost','DepreciationMethod',
                    'UsefulLifeYears','ResidualValue','AccumulatedDepreciation','NetBookValue','Status','Notes']
  };

  Object.keys(REQUIRED).forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, REQUIRED[name].length).setValues([REQUIRED[name]]);
      sheet.setFrozenRows(1);
      actions.push('Created sheet: ' + name);
    }
  });

  // 2. Ensure standard column extensions (e.g., Currency support in Invoices)
  var invSheet = ss.getSheetByName('Invoices');
  if (invSheet && invSheet.getLastColumn() < 24) { // Schema v2.1 extension
    invSheet.insertColumnsAfter(invSheet.getLastColumn(), 24 - invSheet.getLastColumn());
    actions.push('Extended Invoices columns for FX support');
  }

  if (actions.length > 0) {
    Logger.log('Migrated ' + sheetId + ': ' + actions.join(', '));
  }
  return actions;
}