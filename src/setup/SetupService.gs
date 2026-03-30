/**
 * NO~BULL BOOKS — SETUP MICROSERVICE
 *
 * This is a SEPARATE Google Apps Script project from the main hub.
 * Deploy with:
 *   executeAs: USER_ACCESSING
 *   access:    ANYONE
 *
 * It does exactly one thing: create a blank Google Sheet in the client's
 * own Drive, then redirect them to the main no~bull books app with their
 * new Sheet ID in the URL.
 *
 * The client sees a standard Google sign-in (if not already signed in),
 * grants access once, and arrives at the setup wizard.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * SETUP INSTRUCTIONS (one-off):
 *
 *  1. Create a new Apps Script project at script.google.com
 *  2. Paste this file as Code.gs
 *  3. Set MAIN_APP_URL below to your hub /exec URL
 *  4. Deploy as Web App:
 *       Execute as: User accessing the web app
 *       Who has access: Anyone
 *  5. Copy the /exec URL — this is your setup URL
 *     (Optionally point a custom domain to it via a redirect service)
 *
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ── Configuration ─────────────────────────────────────────────────────────────
// Replace with your actual main hub /exec URL
var MAIN_APP_URL = 'https://script.google.com/macros/s/YOUR_MAIN_DEPLOYMENT_ID/exec';

// ── Routes ────────────────────────────────────────────────────────────────────
function doGet(e) {
  var step = e.parameter.step || '1';

  if (step === '2') {
    // Step 2: User has authorised — create the sheet and redirect
    return _createAndRedirect(e.parameter.name || '');
  }

  // Step 1: Show the setup landing page
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
        '<form id="f" method="GET" action="' + setupUrl + '">' +
          '<input type="hidden" name="step" value="2">' +
          '<div class="field">' +
            '<label for="name">Your business name</label>' +
            '<input type="text" id="name" name="name" ' +
              'placeholder="e.g. Acme Consulting Ltd" autocomplete="organization" autofocus>' +
          '</div>' +
          '<button class="btn" type="submit">Create my workspace →</button>' +
          '<p class="note">You\'ll be asked to sign in with Google if you aren\'t already.<br>' +
            'This is so we can create your spreadsheet in <strong>your</strong> Drive.</p>' +
        '</form>' +
      '</div>' +
      '<div class="trust">' +
        '<div class="trust-item">✓ 14-day free trial</div>' +
        '<div class="trust-item">✓ Your data, your Drive</div>' +
        '<div class="trust-item">✓ No credit card needed</div>' +
      '</div>' +
    '</div>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Get started — no~bull books')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─────────────────────────────────────────────────────────────────────────────
// STEP 2 — Create sheet in client's Drive, redirect to main app
// Runs as USER_ACCESSING — SpreadsheetApp.create() goes into the client's Drive
// ─────────────────────────────────────────────────────────────────────────────

function _createAndRedirect(companyName) {
  try {
    var name  = (companyName || '').trim() || 'My Business';
    var ss    = SpreadsheetApp.create('no~bull books — ' + name);
    var id    = ss.getId();
    var appUrl= MAIN_APP_URL + '?id=' + id;

    // Immediate redirect to the main app — client never sees a success page,
    // they just land straight in the setup wizard.
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head>' +
      '<meta http-equiv="refresh" content="0;url=' + appUrl + '">' +
      '<title>Redirecting…</title></head>' +
      '<body style="font-family:sans-serif;display:flex;align-items:center;' +
        'justify-content:center;min-height:100vh;background:#0f172a">' +
        '<div style="text-align:center;color:#94a3b8">' +
          '<div style="font-size:48px;margin-bottom:16px">🐂</div>' +
          '<p>Opening your workspace…</p>' +
          '<p style="font-size:12px;margin-top:8px">' +
            '<a href="' + appUrl + '" style="color:#60a5fa">Click here if not redirected</a>' +
          '</p>' +
        '</div>' +
      '</body></html>'
    );

  } catch(e) {
    // Something went wrong — show a friendly error with retry
    Logger.log('SetupService error: ' + e.toString());
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><title>Setup error</title></head>' +
      '<body style="font-family:sans-serif;padding:48px;text-align:center">' +
        '<h2 style="color:#dc2626">Something went wrong</h2>' +
        '<p style="color:#64748b;margin:16px 0">' + e.message + '</p>' +
        '<a href="' + ScriptApp.getService().getUrl() + '" ' +
          'style="background:#2563eb;color:#fff;padding:12px 24px;border-radius:8px;' +
          'text-decoration:none;font-weight:600">Try again</a>' +
      '</body></html>'
    );
  }
}
