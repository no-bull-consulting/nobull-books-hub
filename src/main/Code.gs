/**
 * NO~BULL BOOKS — CENTRAL HUB ROUTER
 *
 * URL modes:
 *   ?id=SHEET_ID  → main accounting app for existing client
 *   (no params)   → landing page pointing to the setup microservice
 *
 * The setup flow (creating client spreadsheets) is handled by a
 * SEPARATE deployment: SetupService.gs (executeAs: USER_ACCESSING).
 * Its URL is stored in Script Properties as SETUP_SERVICE_URL.
 */
function doGet(e) {
  var sheetId = e.parameter.id;

  // ── Main app ────────────────────────────────────────────────────────────────
  if (sheetId) {
    var tmpl = HtmlService.createTemplateFromFile('Index');
    tmpl.clientSheetId = sheetId;
    tmpl.scriptUrl     = ScriptApp.getService().getUrl();

    return tmpl.evaluate()
      .setTitle('no~bull books')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── Landing page ────────────────────────────────────────────────────────────
  var setupUrl = PropertiesService.getScriptProperties()
    .getProperty('SETUP_SERVICE_URL') || '#';
  var appUrl   = ScriptApp.getService().getUrl();

  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<link rel="icon" href="data:image/svg+xml,<svg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 100 100\'><text y=\'.9em\' font-size=\'90\'>🐂</text></svg>">' +
    '<title>no~bull books</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;' +
      'background:#0f172a;min-height:100vh;display:flex;align-items:center;' +
      'justify-content:center;padding:24px}' +
    '.card{background:#fff;border-radius:16px;padding:48px 40px;max-width:440px;' +
      'width:100%;text-align:center;box-shadow:0 24px 64px rgba(0,0,0,.4)}' +
    '.logo{font-family:Georgia,serif;font-size:22px;color:#0f172a;margin-bottom:4px}' +
    '.logo span{color:#2563eb}' +
    'p{font-size:15px;color:#64748b;line-height:1.7;margin:16px 0 28px}' +
    '.btn-primary{display:inline-block;background:#2563eb;color:#fff;' +
      'text-decoration:none;padding:14px 32px;border-radius:8px;' +
      'font-weight:600;font-size:15px;width:100%;text-align:center}' +
    '.sub{font-size:12px;color:#94a3b8;margin-top:14px}' +
    '</style></head><body>' +
    '<div class="card">' +
      '<div style="font-size:52px;margin-bottom:16px">🐂</div>' +
      '<div class="logo">no~bull <span>books</span></div>' +
      '<p>Straightforward accounting for UK sole traders &amp; small businesses.<br>' +
        'Your data lives in your own Google Sheet.</p>' +
      '<a href="' + setupUrl + '" class="btn-primary">Get started free →</a>' +
      '<p class="sub">14-day free trial · No credit card required</p>' +
    '</div>' +
    '</body></html>'
  );
}

/**
 * getDb(params)
 * Returns the correct spreadsheet for the current request.
 * Always use this — never SpreadsheetApp.openById() directly in other files.
 */
function getDb(params) {
  if (params && params._sheetId) {
    return SpreadsheetApp.openById(params._sheetId);
  }
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch(e) {
    return SpreadsheetApp.openById(DEFAULT_SPREADSHEET_ID);
  }
}

/**
 * include(filename)
 * Helper to include HTML file content within GAS templates.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * setSetupServiceUrl(url)
 * Run once from the Apps Script editor after deploying SetupService.
 * Stores the setup microservice URL so the landing page can link to it.
 */
function setSetupServiceUrl(url) {
  PropertiesService.getScriptProperties().setProperty('SETUP_SERVICE_URL', url);
  Logger.log('SETUP_SERVICE_URL set to: ' + url);
}
