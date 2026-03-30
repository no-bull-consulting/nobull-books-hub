/**
 * NO~BULL BOOKS — ONBOARDING & LICENCING
 *
 * Handles:
 *  - Trial expiry checks on every startup (_checkLicence)
 *  - Welcome email to new clients (sendWelcomeEmail)
 *  - Automated provisioning helper (provisionNewClient)
 *
 * Trial period: defined by TRIAL_DAYS in Script Properties (default 14).
 * After expiry the app shows a "contact us to activate" screen.
 * ─────────────────────────────────────────────────────────────────────────────
 */

var DEFAULT_TRIAL_DAYS = 14;

// ─────────────────────────────────────────────────────────────────────────────
// LICENCE CHECK — called on every getStartupData
// ─────────────────────────────────────────────────────────────────────────────

/**
 * _checkLicence(sheetId)
 * Returns a licence object. Called from getStartupData in Api.gs.
 * If the registry is not configured, or the client is not found,
 * the app runs in "unregistered trial" mode for DEFAULT_TRIAL_DAYS.
 *
 * Returns:
 *   { status, daysRemaining, trialExpired, plan, message }
 */
function _checkLicence(sheetId) {
  try {
    // If registry not configured → permissive (don't block early dev instances)
    var regId = PropertiesService.getScriptProperties().getProperty('REGISTRY_SHEET_ID');
    if (!regId || !sheetId) {
      return { status: 'Active', plan: 'Solo', trialExpired: false, daysRemaining: null, registered: false };
    }

    // Look up client in registry
    var result = getRegistryClient(sheetId);
    if (!result.success) {
      // Not registered — treat as unregistered trial
      return _unregisteredTrial(sheetId);
    }

    var client = result.client;
    var status = client.status || 'Trial';

    // Hard statuses
    if (status === 'Suspended') {
      return {
        status: 'Suspended', plan: client.plan, trialExpired: false,
        daysRemaining: 0, registered: true,
        message: 'Your account has been suspended. Please contact support.'
      };
    }
    if (status === 'Cancelled') {
      return {
        status: 'Cancelled', plan: client.plan, trialExpired: false,
        daysRemaining: 0, registered: true,
        message: 'Your account has been cancelled.'
      };
    }
    if (status === 'Active') {
      // Ping registry with latest counts (non-blocking)
      try { _pingRegistryAsync(sheetId, client.plan); } catch(e) {}
      return { status: 'Active', plan: client.plan, trialExpired: false, daysRemaining: null, registered: true };
    }

    // Trial — check expiry
    if (status === 'Trial') {
      var trialDays = parseInt(
        PropertiesService.getScriptProperties().getProperty('TRIAL_DAYS') || DEFAULT_TRIAL_DAYS
      );
      var created = client.createdDate ? new Date(client.createdDate) : new Date();
      var now     = new Date();
      var elapsed = Math.floor((now - created) / 86400000);
      var remaining = trialDays - elapsed;

      if (remaining <= 0) {
        return {
          status: 'Trial', plan: client.plan, trialExpired: true,
          daysRemaining: 0, registered: true,
          companyName: client.companyName,
          message: 'Your ' + trialDays + '-day trial has ended. Contact us to activate your account.'
        };
      }

      try { _pingRegistryAsync(sheetId, client.plan); } catch(e) {}
      return {
        status: 'Trial', plan: client.plan, trialExpired: false,
        daysRemaining: remaining, registered: true,
        message: remaining <= 3
          ? 'Trial ends in ' + remaining + ' day' + (remaining === 1 ? '' : 's') + '.'
          : null
      };
    }

    return { status: status, plan: client.plan, trialExpired: false, daysRemaining: null, registered: true };

  } catch(e) {
    Logger.log('_checkLicence error (non-fatal): ' + e.toString());
    // Never block app on licence errors
    return { status: 'Active', plan: 'Solo', trialExpired: false, daysRemaining: null, registered: false };
  }
}

function _unregisteredTrial(sheetId) {
  // No registry entry — check if there's a local "first seen" date in sheet properties
  try {
    var trialDays = parseInt(
      PropertiesService.getScriptProperties().getProperty('TRIAL_DAYS') || DEFAULT_TRIAL_DAYS
    );
    var ss       = SpreadsheetApp.openById(sheetId);
    var props    = ss.createDeveloperMetadataFinder()
      .withKey('nbb_first_seen').find();

    var firstSeen;
    if (props.length > 0) {
      firstSeen = new Date(props[0].getValue());
    } else {
      // First time ever — stamp it
      firstSeen = new Date();
      ss.addDeveloperMetadata('nbb_first_seen', firstSeen.toISOString());
    }

    var elapsed   = Math.floor((new Date() - firstSeen) / 86400000);
    var remaining = trialDays - elapsed;

    if (remaining <= 0) {
      return {
        status: 'Trial', plan: 'Solo', trialExpired: true,
        daysRemaining: 0, registered: false,
        message: 'Your ' + trialDays + '-day trial has ended. Contact us to get started.'
      };
    }

    return {
      status: 'Trial', plan: 'Solo', trialExpired: false,
      daysRemaining: remaining, registered: false,
      message: remaining <= 3 ? 'Trial ends in ' + remaining + ' day' + (remaining === 1 ? '' : 's') + '.' : null
    };
  } catch(e) {
    return { status: 'Active', plan: 'Solo', trialExpired: false, daysRemaining: null, registered: false };
  }
}

function _pingRegistryAsync(sheetId, plan) {
  // Lightweight ping — just last seen + counts
  try {
    pingRegistry(sheetId, {
      email:   Session.getActiveUser().getEmail(),
      version: '1.0'
    });
  } catch(e) { /* non-fatal */ }
}

// ─────────────────────────────────────────────────────────────────────────────
// WELCOME EMAIL
// ─────────────────────────────────────────────────────────────────────────────

/**
 * sendWelcomeEmail(params)
 * Sends a branded welcome email to a newly registered client.
 *
 * params: { toEmail, companyName, contactName, appUrl, plan, trialDays }
 *
 * Call from Admin Panel after registering a client, or automatically
 * from provisionNewClient().
 */
function sendWelcomeEmail(params) {
  try {
    var to          = params.toEmail || params.contactEmail;
    var contactName = params.contactName || 'there';
    var company     = params.companyName || 'your business';
    var appUrl      = params.appUrl || params.deployUrl || '';
    var plan        = params.plan || 'Solo';
    var trialDays   = parseInt(
      PropertiesService.getScriptProperties().getProperty('TRIAL_DAYS') || DEFAULT_TRIAL_DAYS
    );

    if (!to) return { success: false, message: 'No email address provided.' };

    var subject = 'Welcome to no\u2060~\u2060bull books \ud83d\udc02 — your account is ready';

    var html =
      '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<style>' +
      'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;' +
      'background:#f8fafc;margin:0;padding:0;color:#1e293b}' +
      '.wrap{max-width:560px;margin:40px auto;background:#fff;border-radius:12px;' +
      'box-shadow:0 4px 24px rgba(0,0,0,.08);overflow:hidden}' +
      '.header{background:#0f172a;padding:32px 40px;text-align:center}' +
      '.logo{font-family:Georgia,serif;font-size:22px;color:#fff;letter-spacing:.5px}' +
      '.logo span{color:#93c5fd}' +
      '.body{padding:36px 40px}' +
      'h1{font-size:22px;font-weight:700;color:#0f172a;margin:0 0 16px}' +
      'p{font-size:15px;line-height:1.7;color:#475569;margin:0 0 16px}' +
      '.btn{display:inline-block;background:#2563eb;color:#fff;text-decoration:none;' +
      'padding:14px 32px;border-radius:8px;font-weight:600;font-size:15px;margin:8px 0 24px}' +
      '.card{background:#f1f5f9;border-radius:8px;padding:20px 24px;margin:20px 0}' +
      '.card h3{font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;' +
      'color:#94a3b8;margin:0 0 12px}' +
      '.step{display:flex;gap:12px;margin-bottom:12px;align-items:flex-start}' +
      '.num{width:24px;height:24px;border-radius:50%;background:#2563eb;color:#fff;' +
      'font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0}' +
      '.step-text{font-size:14px;color:#334155;line-height:1.5}' +
      '.url-box{background:#0f172a;color:#93c5fd;border-radius:6px;padding:12px 16px;' +
      'font-family:monospace;font-size:13px;word-break:break-all;margin:12px 0}' +
      '.footer{background:#f8fafc;padding:24px 40px;text-align:center;' +
      'font-size:12px;color:#94a3b8;border-top:1px solid #e2e8f0}' +
      '.footer a{color:#64748b}' +
      '</style></head><body>' +
      '<div class="wrap">' +
        '<div class="header">' +
          '<div class="logo">no\u2060~\u2060bull <span>books</span> \ud83d\udc02</div>' +
        '</div>' +
        '<div class="body">' +
          '<h1>Welcome, ' + _escHtml(contactName) + '!</h1>' +
          '<p>Your <strong>no~bull books</strong> account for <strong>' + _escHtml(company) + '</strong> ' +
          'is ready. You have a <strong>' + trialDays + '-day free trial</strong> on the <strong>' + plan + '</strong> plan.</p>' +

          (appUrl
            ? '<div style="text-align:center"><a href="' + appUrl + '" class="btn">Open no~bull books \u2192</a></div>' +
              '<div class="url-box">' + _escHtml(appUrl) + '</div>' +
              '<p style="font-size:13px;color:#94a3b8">Bookmark this link — it\'s your unique account URL.</p>'
            : '') +

          '<div class="card">' +
            '<h3>Get started in 3 steps</h3>' +
            '<div class="step"><div class="num">1</div>' +
              '<div class="step-text"><strong>Complete your setup wizard</strong> — enter your company name, ' +
              'invoice settings and financial year dates.</div></div>' +
            '<div class="step"><div class="num">2</div>' +
              '<div class="step-text"><strong>Add your bank account</strong> — Banking \u2192 + Add Bank Account. ' +
              'This enables transactions and reconciliation.</div></div>' +
            '<div class="step"><div class="num">3</div>' +
              '<div class="step-text"><strong>Seed your Chart of Accounts</strong> — Admin Panel \u2192 ' +
              'Quick Actions \u2192 \ud83c\udf31 Seed UK COA. Installs 90 standard UK nominal accounts.</div></div>' +
          '</div>' +

          '<div class="card">' +
            '<h3>What\'s included</h3>' +
            '<p style="font-size:14px;margin:0">' +
            '\u2713 Invoicing with PDF generation and email<br>' +
            '\u2713 Bills and purchase orders<br>' +
            '\u2713 Bank account management and reconciliation<br>' +
            '\u2713 VAT returns and HMRC MTD connection<br>' +
            '\u2713 P&amp;L, Balance Sheet, Cash Flow reports<br>' +
            '\u2713 SA103 self-assessment calculator<br>' +
            '\u2713 Gemini AI financial assistant<br>' +
            '\u2713 All data in <em>your own</em> Google Sheet — no lock-in' +
            '</p>' +
          '</div>' +

          '<p>Questions? Reply to this email and we\'ll get back to you quickly.</p>' +
          '<p style="font-size:14px;color:#64748b">If you didn\'t request this account, you can safely ignore this email.</p>' +
        '</div>' +
        '<div class="footer">' +
          '<p>no~bull books by <a href="mailto:edward@nobull.consulting">no~bull consulting</a></p>' +
          '<p>Your data lives in your Google Drive. We never see it.</p>' +
        '</div>' +
      '</div>' +
      '</body></html>';

    var plainText =
      'Welcome to no~bull books, ' + contactName + '!\n\n' +
      'Your account for ' + company + ' is ready.\n\n' +
      (appUrl ? 'Your app link:\n' + appUrl + '\n\nBookmark this — it\'s your unique URL.\n\n' : '') +
      'GET STARTED:\n' +
      '1. Complete the setup wizard (company name, invoice settings, financial year)\n' +
      '2. Add your bank account (Banking > + Add Bank Account)\n' +
      '3. Seed your Chart of Accounts (Admin Panel > Seed UK COA)\n\n' +
      'You have a ' + trialDays + '-day free trial on the ' + plan + ' plan.\n\n' +
      'Questions? Reply to this email.\n\n' +
      'no~bull books by no~bull consulting\n' +
      'edward@nobull.consulting';

    GmailApp.sendEmail(to, subject, plainText, {
      htmlBody: html,
      name: 'no~bull books',
      replyTo: 'edward@nobull.consulting'
    });

    Logger.log('Welcome email sent to: ' + to);
    return { success: true, message: 'Welcome email sent to ' + to };
  } catch(e) {
    Logger.log('sendWelcomeEmail error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function _escHtml(s) {
  return (s || '').toString()
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ─────────────────────────────────────────────────────────────────────────────
// FULL PROVISIONING — one function to do everything
// ─────────────────────────────────────────────────────────────────────────────

/**
 * provisionNewClient(params)
 * Full onboarding in one call:
 *   1. Creates a blank Google Sheet for the client
 *   2. Registers them in the registry
 *   3. Sends the welcome email
 *
 * params: { companyName, contactName, contactEmail, contactPhone,
 *           plan, notes, vatRegistered, accountant, accountantEmail }
 *
 * Can be called from the Admin Panel or a future Stripe webhook.
 * Returns: { success, sheetId, sheetUrl, appUrl, registryId }
 */
function provisionNewClient(params) {
  try {
    var sheetId = (params.sheetId || '').trim();
    if (!params.companyName) return { success: false, message: 'Company name is required.' };

    Logger.log('=== Provisioning: ' + params.companyName + ' (sheet: ' + (sheetId||'not yet created') + ')');

    // Build the app URL
    var baseUrl = ScriptApp.getService().getUrl();
    var appUrl  = sheetId ? baseUrl + '?id=' + sheetId : '';
    var sheetUrl= sheetId ? 'https://docs.google.com/spreadsheets/d/' + sheetId : '';

    // Register in the registry
    var regResult = registerClient({
      companyName:    params.companyName,
      contactName:    params.contactName    || '',
      contactEmail:   params.contactEmail   || '',
      contactPhone:   params.contactPhone   || '',
      sheetId:        sheetId,
      deployUrl:      appUrl,
      plan:           params.plan           || 'Solo',
      status:         params.status         || 'Trial',
      vatRegistered:  params.vatRegistered  || false,
      vatNumber:      params.vatNumber      || '',
      country:        params.country        || 'UK',
      accountant:     params.accountant     || '',
      accountantEmail:params.accountantEmail|| '',
      notes:          params.notes          || ''
    });

    if (!regResult.success) {
      return { success: false, message: 'Registry error: ' + regResult.message };
    }

    // Send welcome email
    var emailResult = { success: false, message: 'No email — skipped' };
    if (params.contactEmail) {
      emailResult = sendWelcomeEmail({
        toEmail:     params.contactEmail,
        contactName: params.contactName || params.companyName,
        companyName: params.companyName,
        appUrl:      appUrl,
        plan:        params.plan || 'Solo'
      });
    }

    return {
      success:    true,
      sheetId:    sheetId,
      sheetUrl:   sheetUrl,
      appUrl:     appUrl,
      registryId: regResult.registryId,
      emailSent:  emailResult.success,
      message:    params.companyName + ' registered.' +
                  (emailResult.success ? ' Welcome email sent.' : '') +
                  (appUrl ? '' : ' No sheet ID — client will create their own via the setup link.')
    };
  } catch(e) {
    Logger.log('provisionNewClient error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// ADMIN API ROUTES (add to Api.gs)
// ─────────────────────────────────────────────────────────────────────────────
// case 'provisionNewClient':  _auth('settings.write', params); return provisionNewClient(params);
// case 'sendWelcomeEmail':    _auth('settings.write', params); return sendWelcomeEmail(params);
// case 'resendWelcomeEmail':  _auth('settings.write', params); return sendWelcomeEmail(params);

// ─────────────────────────────────────────────────────────────────────────────
// EDITOR HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Run from the Apps Script editor to set trial length.
 */
function setTrialDays(days) {
  PropertiesService.getScriptProperties().setProperty('TRIAL_DAYS', String(days || DEFAULT_TRIAL_DAYS));
  Logger.log('Trial period set to ' + days + ' days.');
}

/**
 * Manually activate a client (move from Trial → Active).
 * Run from editor: activateClient('REG_xxx_xxx')
 */
function activateClient(registryId) {
  var r = updateRegistryClient(registryId, { status: 'Active' });
  Logger.log(JSON.stringify(r));
}

/**
 * Test provisioning without sending email.
 */
function debug_provision() {
  var r = provisionNewClient({
    companyName:  'Test Company Ltd',
    contactName:  'Test User',
    contactEmail: '', // blank = skip email
    plan:         'Solo',
    notes:        'Debug provisioning test'
  });
  Logger.log(JSON.stringify(r));
}
