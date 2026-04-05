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

  // ── HMRC OAuth redirect ─────────────────────────────────────────────────────
  // HMRC redirects back here with ?code=XXX&state=YYY after OAuth sign-in.
  // We detect this and auto-redirect to the app with the code pre-filled.
  if (e.parameter.code && !sheetId) {
    var code   = e.parameter.code;
    var appUrl = 'https://script.google.com/a/macros/nobull.consulting/s/AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j/exec?id=1gIFwQUtbhGaM3HIHbFFaT7lIAU4BN3IksAOv1_uuUKg&hmrc_code=' + encodeURIComponent(code);
    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
      + '<meta http-equiv="refresh" content="0;url=' + appUrl + '">'
      + '<title>Connecting to HMRC...</title>'
      + '<style>body{font-family:-apple-system,sans-serif;background:#0f172a;min-height:100vh;display:flex;align-items:center;justify-content:center;color:#fff;text-align:center}</style>'
      + '</head><body>'
      + '<p>Connected! Returning to no~bull books...</p>'
      + '<p><a href="' + appUrl + '" style="color:#60a5fa">Click here if not redirected</a></p>'
      + '</body></html>';
    return HtmlService.createHtmlOutput(html)
      .setTitle('Connecting to HMRC...')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── Main app ────────────────────────────────────────────────────────────────
  if (sheetId) {
    var tmpl = HtmlService.createTemplateFromFile('Index');
    tmpl.clientSheetId = sheetId;
    tmpl.scriptUrl     = ScriptApp.getService().getUrl();
    tmpl.ownerEmail    = e.parameter.ownerEmail || '';

    return tmpl.evaluate()
      .setTitle('no~bull books')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── Auto-initialise sheet if first visit ────────────────────────────────────
  // If the sheet exists but hasn't been initialised (no Settings tab), init it now.
  // This ensures the app works even if the setup service redirect was interrupted.
  if (e.parameter.id) {
    try {
      var testSs    = getDb({ _sheetId: e.parameter.id });
      var testSheet = testSs ? testSs.getSheetByName('Settings') : null;
      if (!testSheet) {
        Logger.log('Auto-initialising sheet: ' + e.parameter.id);
        checkAndInitSheet({
          _sheetId:    e.parameter.id,
          _ownerEmail: e.parameter.ownerEmail || ''
        });
      }
    } catch(autoInitErr) {
      Logger.log('Auto-init error (non-fatal): ' + autoInitErr.toString());
    }
  }

  // ── Registry ping from SetupService ────────────────────────────────────────
  // Called by SetupService after creating a new client sheet
  if (e.parameter.action === 'pingRegistry') {
    try {
      pingRegistry(e.parameter.sheetId, {
        email:       e.parameter.email       || '',
        companyName: e.parameter.companyName || '',
        version:     APP_VERSION
      });
    } catch(pingErr) {
      Logger.log('doGet pingRegistry error: ' + pingErr.toString());
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── Landing page ────────────────────────────────────────────────────────────
  var setupUrl = PropertiesService.getScriptProperties()
    .getProperty('SETUP_SERVICE_URL') || '#';

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
      'font-weight:600;font-size:15px;width:100%;text-align:center;' +
      'cursor:pointer;border:none;font-family:inherit}' +
    '.btn-primary:hover{background:#1d4ed8}' +
    '.sub{font-size:12px;color:#94a3b8;margin-top:14px}' +
    '</style></head><body>' +
    '<div class="card">' +
      '<div style="font-size:52px;margin-bottom:16px">🐂</div>' +
      '<div class="logo">no~bull <span>books</span></div>' +
      '<p>Straightforward accounting for UK sole traders &amp; small businesses.<br>' +
        'Your data lives in your own Google Sheet.</p>' +
      '<button class="btn-primary" onclick="window.top.location.href=\'' + setupUrl + '\'">Get started free →</button>' +
      '<p class="sub">14-day free trial · No credit card required</p>' +
    '</div>' +
    '</body></html>'
  )
  .setTitle('no~bull books')
  .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
 */
function setSetupServiceUrl(url) {
  PropertiesService.getScriptProperties().setProperty('SETUP_SERVICE_URL', url);
  Logger.log('SETUP_SERVICE_URL set to: ' + url);
}

/**
 * Verify script properties are set correctly.
 * Run from editor to check configuration.
 */
function _checkProps() {
  var props = PropertiesService.getScriptProperties().getProperties();
  Logger.log('SETUP_SERVICE_URL: ' + props['SETUP_SERVICE_URL']);
  Logger.log('REGISTRY_SHEET_ID: ' + props['REGISTRY_SHEET_ID']);
  Logger.log('TRIAL_DAYS: '        + props['TRIAL_DAYS']);
  Logger.log('GEMINI_API_KEY: '    + (props['GEMINI_API_KEY'] ? 'set (' + props['GEMINI_API_KEY'].substring(0,8) + '...)' : 'NOT SET'));
}