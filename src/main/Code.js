/**
 * NO~BULL BOOKS -- CENTRAL HUB ROUTER
 *
 * URL modes:
 *   ?id=SHEET_ID  -> main accounting app for existing client
 *   (no params)   -> landing page pointing to the setup microservice
 *
 * The setup flow (creating client spreadsheets) is handled by a
 * SEPARATE deployment: SetupService.gs (executeAs: USER_ACCESSING).
 * Its URL is stored in Script Properties as SETUP_SERVICE_URL.
 */
function doGet(e) {
  var sheetId = e.parameter.id;

  // -- HMRC OAuth redirect --------------------------------------------------
  if (e.parameter.code && !sheetId) { return _hmrcRedirectPage(e.parameter.code); }

  // -- Main app ----------------------------------------------------------------
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

  // -- Auto-initialise sheet if first visit ------------------------------------
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

  // -- Registry ping from SetupService ----------------------------------------
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

  // -- Landing page ------------------------------------------------------------
  var setupUrl = PropertiesService.getScriptProperties()
    .getProperty('SETUP_SERVICE_URL') || '#';

  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<link rel="icon" href="data:image/svg+xml,<svg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 100 100\'><text y=\'.9em\' font-size=\'90\'>?</text></svg>">' +
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
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 80" style="height:44px;width:auto;margin:0 auto 16px;display:block">' +
      '<g transform="translate(0,80) scale(0.029326,-0.029091)" fill="#14213D" stroke="none">' +
      '<path d="M0 2730 c0 -28 45 -114 72 -140 39 -36 89 -52 163 -52 72 -1 146 21 246 72 42 22 50 23 219 16 293 -12 487 -43 678 -107 105 -36 158 -66 206 -120 87 -96 87 -133 -9 -457 -42 -140 -74 -257 -71 -259 2 -3 33 11 68 30 142 79 272 117 408 117 100 0 268 -21 335 -42 39 -12 38 -12 -63 -19 -244 -18 -457 -113 -649 -293 -165 -153 -288 -385 -314 -590 -34 -271 87 -514 312 -625 113 -56 206 -74 339 -68 332 17 643 209 832 516 l50 81 364 0 364 0 0 23 c0 74 67 239 138 344 168 244 528 421 815 400 45 -3 102 -12 127 -20 l45 -13 -80 -7 c-181 -14 -341 -79 -488 -198 -350 -282 -431 -745 -173 -987 104 -97 220 -142 366 -142 402 0 788 342 843 747 12 96 0 201 -32 273 -11 24 -17 48 -15 52 8 13 126 9 207 -6 112 -22 193 -63 276 -141 39 -38 79 -81 89 -97 10 -18 41 -41 80 -60 55 -27 68 -30 119 -24 31 4 76 15 100 26 42 19 153 99 153 111 0 3 -9 32 -20 65 -26 75 -63 215 -86 324 -25 119 -25 299 0 365 48 127 96 153 447 240 145 36 285 71 312 77 28 7 47 17 47 26 0 11 -62 13 -382 10 -488 -4 -456 6 -706 -240 -67 -67 -124 -119 -126 -117 -3 2 15 56 40 118 36 92 42 117 32 126 -7 7 -67 55 -134 107 -66 53 -188 153 -270 224 -155 133 -270 214 -364 256 -75 34 -107 44 -239 73 l-114 25 -2279 0 -2278 0 0 -20z m2312 -1320 c134 -37 215 -129 237 -272 30 -183 -76 -383 -270 -511 -198 -132 -456 -114 -571 38 -48 63 -48 74 0 36 144 -115 368 -83 545 80 110 100 178 267 156 384 -13 70 -27 104 -66 152 -32 39 -122 93 -157 93 -8 0 -17 4 -21 10 -7 12 89 5 147 -10z m2359 -205 c143 -49 211 -206 160 -373 -65 -214 -278 -366 -488 -349 -83 6 -158 43 -200 98 -41 53 -42 69 -3 35 37 -30 117 -56 176 -56 106 0 253 79 326 175 57 75 82 146 82 235 1 120 -49 193 -161 238 l-58 23 60 -5 c33 -3 81 -12 106 -21z"/>' +
      '</g></svg>' +
      '<div class="logo">no~bull <span>books</span></div>' +
      '<p>Straightforward accounting for UK sole traders &amp; small businesses.<br>' +
        'Your data lives in your own Google Sheet.</p>' +
      '<button class="btn-primary" onclick="window.top.location.href=\'' + setupUrl + '\'">Get started free -></button>' +
      '<p class="sub">14-day free trial ? No credit card required</p>' +
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
 * Always use this -- never SpreadsheetApp.openById() directly in other files.
 */

function _hmrcRedirectPage(code) {
  var base = 'https://script.google.com/a/macros/nobull.consulting/s/';
  var dep  = 'AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j';
  var sid  = '1gIFwQUtbhGaM3HIHbFFaT7lIAU4BN3IksAOv1_uuUKg';
  var url  = base + dep + '/exec?id=' + sid + '&hmrc_code=' + encodeURIComponent(code);
  return HtmlService.createHtmlOutput(
    'Connected! Returning to no~bull books. If not redirected: ' + url
  ).setTitle('Connecting').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ─────────────────────────────────────────────────────────────────────────────
// CONTACT FORM HANDLER
// ─────────────────────────────────────────────────────────────────────────────

function _handleContactForm(e) {
  try {
    var body = JSON.parse(e.postData ? e.postData.contents : '{}');
    var nl   = '\n';
    var text =
      'New enquiry from nobull.consulting contact form' + nl + nl +
      'Name:    ' + (body.fname||'') + ' ' + (body.lname||'') + nl +
      'Email:   ' + (body.email||'') + nl +
      'Company: ' + (body.company||'—') + nl +
      'Subject: ' + (body.subject||'—') + nl +
      'Source:  ' + (body.source||'—') + nl + nl +
      'Message:' + nl + (body.message||'') + nl;

    MailApp.sendEmail({
      to:      'edward@nobull.consulting',
      replyTo: body.email || '',
      subject: 'no~bull enquiry: ' + (body.subject || 'General'),
      body:    text
    });

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    Logger.log('Contact form error: ' + e.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: e.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


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
function doPost(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
    if (action === 'contact') return _handleContactForm(e);
    // Otherwise route to API handler
    var body = JSON.parse(e.postData ? e.postData.contents : '{}');
    return ContentService.createTextOutput(JSON.stringify(handleApiCall(body.action, body)))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

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