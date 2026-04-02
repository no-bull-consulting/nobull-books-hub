/**
 * NO~BULL BOOKS — SETUP MICROSERVICE
 *
 * This is a SEPARATE Google Apps Script project from the main hub.
 * Deploy with:
 *   executeAs: USER_ACCESSING
 *   access:    Anyone with a Google account
 *
 * It does exactly one thing: create a blank Google Sheet in the client's
 * own Drive, share it with the hub operator, ping the registry, then
 * redirect the client to the main no~bull books app with their Sheet ID.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * SETUP INSTRUCTIONS (one-off):
 *
 *  1. Create a new Apps Script project at script.google.com
 *  2. Paste this file as Code.gs
 *  3. Set MAIN_APP_URL and HUB_OPERATOR_EMAIL below
 *  4. Deploy as Web App:
 *       Execute as: User accessing the web app
 *       Who has access: Anyone with a Google account
 *  5. Copy the /exec URL — store this in the main hub via setSetupServiceUrl()
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ── Configuration ─────────────────────────────────────────────────────────────
// Replace with your actual main hub /exec URL
var MAIN_APP_URL = 'https://script.google.com/a/macros/nobull.consulting/s/AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j/exec';

// The Google account the main hub runs as — needs editor access to every client sheet
var HUB_OPERATOR_EMAIL = 'edward@nobull.consulting';

// ── Routes ────────────────────────────────────────────────────────────────────
function doGet(e) {
  var step = e.parameter.step || '1';

  if (step === '2') {
    return _createAndRedirect(e.parameter.name || '');
  }

  return _serveLandingPage();
}

// ─────────────────────────────────────────────────────────────────────────────
// STEP 1 — Landing page
// ─────────────────────────────────────────────────────────────────────────────

function _serveLandingPage() {
  var setupUrl = ScriptApp.getService().getUrl();

  var html =
    '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<link rel="icon" href="data:image/svg+xml,<svg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 100 100\'><text y=\'.9em\' font-size=\'90\'>🐂</text></svg>">' +
    '<title>Get started — no~bull books</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;' +
      'background:#0f172a;min-height:100vh;display:flex;align-items:center;' +
      'justify-content:center;padding:24px}' +
    '.card{background:#fff;border-radius:16px;max-width:480px;width:100%;' +
      'overflow:hidden;box-shadow:0 24px 64px rgba(0,0,0,.4)}' +
    '.hdr{background:#0f172a;padding:28px 40px;display:flex;align-items:center;gap:14px}' +
    '.hdr-logo{font-family:Georgia,serif;font-size:18px;color:#fff}' +
    '.hdr-logo span{color:#93c5fd}' +
    '.body{padding:36px 40px}' +
    'h1{font-size:22px;font-weight:700;color:#0f172a;margin-bottom:10px}' +
    '.sub{font-size:14px;color:#64748b;line-height:1.7;margin-bottom:24px}' +
    '.field{margin-bottom:18px}' +
    'label{display:block;font-size:11px;font-weight:700;letter-spacing:.6px;' +
      'text-transform:uppercase;color:#374151;margin-bottom:6px}' +
    'input{width:100%;padding:12px 14px;border:1.5px solid #e2e8f0;border-radius:8px;' +
      'font-size:15px;font-family:inherit;outline:none;color:#0f172a}' +
    'input:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.1)}' +
    '.btn{width:100%;padding:14px;background:#2563eb;color:#fff;border:none;' +
      'border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;font-family:inherit}' +
    '.btn:hover{background:#1d4ed8}' +
    '.btn:disabled{background:#93c5fd;cursor:not-allowed}' +
    '.note{font-size:12px;color:#94a3b8;margin-top:16px;text-align:center;line-height:1.6}' +
    '.trust{display:flex;gap:16px;justify-content:center;flex-wrap:wrap;' +
      'padding:18px 40px;border-top:1px solid #f1f5f9;background:#f8fafc}' +
    '.trust-item{font-size:12px;color:#64748b;display:flex;align-items:center;gap:5px}' +
    '</style></head><body>' +
    '<div class="card">' +
      '<div class="hdr">' +
        '<div style="font-size:30px">🐂</div>' +
        '<div class="hdr-logo">no~bull <span>books</span></div>' +
      '</div>' +
      '<div class="body">' +
        '<h1>Create your account</h1>' +
        '<p class="sub">Your data lives in your own Google Drive — we never store it. ' +
          'Click the button below and we\'ll set everything up for you.</p>' +
        '<div class="field">' +
          '<label for="name">Your business name</label>' +
          '<input type="text" id="name" name="name" ' +
            'placeholder="e.g. Acme Consulting Ltd" autocomplete="organization">' +
        '</div>' +
        '<button class="btn" type="button" id="submitBtn" onclick="handleSubmit()">Create my workspace →</button>' +
        '<p class="note">You\'ll be asked to sign in with Google if you aren\'t already.<br>' +
          'This is so we can create your spreadsheet in <strong>your</strong> Drive.</p>' +
      '</div>' +
      '<div class="trust">' +
        '<div class="trust-item">✓ 14-day free trial</div>' +
        '<div class="trust-item">✓ Your data, your Drive</div>' +
        '<div class="trust-item">✓ No credit card needed</div>' +
      '</div>' +
    '</div>' +
    '<script>' +
    'function handleSubmit() {' +
      'var name = document.getElementById("name").value.trim();' +
      'if (!name) { document.getElementById("name").focus(); return; }' +
      'var btn = document.getElementById("submitBtn");' +
      'btn.disabled = true;' +
      'btn.textContent = "Setting up your workspace…";' +
      'var url = "' + setupUrl + '?step=2&name=" + encodeURIComponent(name);' +
      'window.top.location.href = url;' +
    '}' +
    'document.getElementById("name").addEventListener("keydown", function(e) {' +
      'if (e.key === "Enter") handleSubmit();' +
    '});' +
    '</script>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Get started — no~bull books')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─────────────────────────────────────────────────────────────────────────────
// STEP 2 — Create sheet in client's Drive, ping registry, redirect to main app
// Runs as USER_ACCESSING — SpreadsheetApp.create() goes into the client's Drive
// New client is assigned Owner role when checkAndInitSheet() runs on first load
// ─────────────────────────────────────────────────────────────────────────────

function _createAndRedirect(companyName) {
  try {
    var name  = (companyName || '').trim() || 'My Business';

    // Create sheet in client's Drive (runs as USER_ACCESSING)
    var ss    = SpreadsheetApp.create('no~bull books — ' + name);

    // Share with hub operator so the main hub (running as edward) can access it
    ss.addEditor(HUB_OPERATOR_EMAIL);

    var id    = ss.getId();
    var email = Session.getActiveUser().getEmail();
    // Pass email to main app so Initializer can seed the correct Owner
    var appUrl = MAIN_APP_URL + '?id=' + id +
      (email ? '&ownerEmail=' + encodeURIComponent(email) : '');

    // Ping the main hub registry so the client appears in Admin Panel immediately
    try {
      var pingUrl = MAIN_APP_URL +
        '?action=pingRegistry' +
        '&sheetId=' + encodeURIComponent(id) +
        '&email='   + encodeURIComponent(email) +
        '&companyName=' + encodeURIComponent(name);
      UrlFetchApp.fetch(pingUrl, { muteHttpExceptions: true });
    } catch(pingErr) {
      Logger.log('Registry ping failed (non-fatal): ' + pingErr.toString());
    }

    // Redirect to main app — window.top breaks out of any iframe
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html>' +
      '<head>' +
        '<title>Redirecting…</title>' +
        '<meta http-equiv="refresh" content="0;url=' + appUrl + '">' +
      '</head>' +
      '<body style="font-family:sans-serif;display:flex;align-items:center;' +
        'justify-content:center;min-height:100vh;background:#0f172a">' +
        '<div style="text-align:center;color:#94a3b8">' +
          '<div style="font-size:48px;margin-bottom:16px">🐂</div>' +
          '<p>Opening your workspace…</p>' +
          '<p style="font-size:12px;margin-top:8px">' +
            '<a href="' + appUrl + '" style="color:#60a5fa">Click here if not redirected</a>' +
          '</p>' +
        '</div>' +
        '<script>' +
          'try { window.top.location.href = "' + appUrl + '"; } catch(e) {' +
          '  window.location.href = "' + appUrl + '"; }' +
        '<\/script>' +
      '</body></html>'
    )
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch(err) {
    Logger.log('SetupService error: ' + err.toString());
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><title>Setup error</title></head>' +
      '<body style="font-family:sans-serif;padding:48px;text-align:center">' +
        '<h2 style="color:#dc2626">Something went wrong</h2>' +
        '<p style="color:#64748b;margin:16px 0">' + err.message + '</p>' +
        '<a href="' + ScriptApp.getService().getUrl() + '" target="_top" ' +
          'style="background:#2563eb;color:#fff;padding:12px 24px;border-radius:8px;' +
          'text-decoration:none;font-weight:600">Try again</a>' +
      '</body></html>'
    )
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}