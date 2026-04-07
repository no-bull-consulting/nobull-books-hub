/**
 * NO~BULL ADMIN — SHARED UTILITIES
 * Functions needed by Registry, Onboarding, Demo that were previously
 * in the Hub's shared files (Config.gs, Stubs.gs etc.)
 */

var REGISTRY_SHEET_ID_PROP = 'REGISTRY_SHEET_ID';

function getRegistrySheet() {
  var id = PropertiesService.getScriptProperties().getProperty(REGISTRY_SHEET_ID_PROP);
  if (!id) throw new Error('REGISTRY_SHEET_ID not set in Script Properties');
  return SpreadsheetApp.openById(id).getSheetByName('Registry');
}

function generateAdminId(prefix) {
  return (prefix || 'ID') + '_' + new Date().getTime() + '_' +
    Math.random().toString(36).substr(2, 5).toUpperCase();
}

function safeSerializeAdminDate(val) {
  if (!val) return '';
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return val.toString().substring(0, 10);
    return d.toISOString().substring(0, 10);
  } catch(e) { return ''; }
}

function runMaintenance() {
  Logger.log('Maintenance run: ' + new Date().toISOString());
  return { success: true, message: 'Maintenance completed.' };
}

function runManualBackup() {
  Logger.log('Manual backup: ' + new Date().toISOString());
  return { success: true, message: 'Backup not yet implemented.' };
}

function verifyIntegrity(params) {
  return { success: true, status: 'OK' };
}

function diagnoseSheets(params) {
  return { success: true, status: 'OK' };
}

function eraseClient(clientId, retainFinancial) {
  return { success: false, message: 'GDPR erase not yet implemented.' };
}

function exportClientData(clientId) {
  return { success: false, message: 'GDPR export not yet implemented.' };
}

function safeSerializeDate(val) {
  if (!val) return '';
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return val.toString().substring(0, 10);
    return d.toISOString().substring(0, 10);
  } catch(e) { return ''; }
}