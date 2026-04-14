/**
 * NO~BULL BOOKS -- AUTH / RBAC
 * Permission engine, user context, user management.
 *
 * KEY CHANGE (hub model): every function now accepts optional `params`
 * so the correct client spreadsheet (_sheetId) is used for Users lookups.
 * -----------------------------------------------------------------------------
 */

/**
 * apiError(userMessage, internalError)
 * Standard safe error response. Logs full detail server-side, returns
 * a clean message to the client.
 */
function apiError(userMessage, internalError) {
  if (internalError) Logger.log('ERROR: ' + userMessage + ' | ' + internalError.toString());
  return { success: false, message: userMessage };
}

// -----------------------------------------------------------------------------
// PERMISSION ENGINE
// -----------------------------------------------------------------------------

function _canDoPermission(role, action) {
  var perms = ROLE_PERMISSIONS[role] || [];
  for (var i = 0; i < perms.length; i++) {
    if (perms[i] === '*') return true;
    if (perms[i] === action) return true;
    var p = perms[i];
    if (p.charAt(p.length - 1) === '*') {
      var prefix = p.slice(0, -2);
      if (action.indexOf(prefix + '.') === 0) return true;
    }
  }
  return false;
}

/**
 * _getCurrentUserContext(params)
 *
 * Returns { email, role, canDo(action) } for the calling user.
 *
 * Bootstrap rule: if the Users sheet does not exist, the caller is
 * treated as Owner (first-deploy only). If it exists but is empty,
 * access is denied to prevent lockout from accidental sheet deletion.
 *
 * @param {Object} [params] - Request params; must contain _sheetId for hub model.
 */
function _getCurrentUserContext(params) {
  var email = '';

  // -- Identity: read owner email stored in Settings sheet -------------------
  // Since the hub runs as USER_DEPLOYING (always edward), we cannot use
  // Session.getActiveUser() to identify the real client. Instead, the client's
  // email is stored in their Settings sheet during onboarding by SetupService,
  // and read here on every request.
  //
  // Security model: knowing the ?id=SHEET_ID URL = authorised access.
  // The sheet ID is a 44-character random string -- effectively a private key.
  // Priority 1: client-verified email from OTP session
  if (params && params._verifiedEmail) {
    email = params._verifiedEmail.trim().toLowerCase();
    Logger.log('Identity: using OTP-verified email: ' + email);
  }

  // Priority 2: ownerEmail from Settings sheet (owner / legacy access)
  if (!email && params && params._sheetId) {
    try {
      var ss       = getDb(params);
      var settings = ss.getSheetByName(SHEETS.SETTINGS);
      if (settings && settings.getLastRow() >= 2) {
        var sData = settings.getRange(2, 1, 1, settings.getLastColumn()).getValues()[0];
        var headers = settings.getRange(1, 1, 1, settings.getLastColumn()).getValues()[0];
        var ownerEmailCol = headers.indexOf('ownerEmail');
        if (ownerEmailCol >= 0 && sData[ownerEmailCol]) {
          email = sData[ownerEmailCol].toString().toLowerCase().trim();
        }
      }
    } catch(e) {
      Logger.log('Identity: could not read ownerEmail from Settings: ' + e.toString());
    }
  }

  // -- Fallback: GAS session (works for edward's own instances) --------------
  if (!email) {
    try {
      email = Session.getActiveUser().getEmail();
      if (!email) email = Session.getEffectiveUser().getEmail();
    } catch(e) {}
  }

  function makeCtx(em, role) {
    return {
      email: em,
      role:  role,
      canDo: role === null
        ? function() { return false; }
        : (function(r) { return function(action) { return _canDoPermission(r, action); }; })(role)
    };
  }

  if (!email) return makeCtx('', 'ReadOnly');

  // -- Superuser override ----------------------------------------------------
  // Check 1: the resolved identity email matches the superuser email
  if (typeof SUPERUSER_EMAIL !== 'undefined' &&
      email && email.toLowerCase() === SUPERUSER_EMAIL.toLowerCase()) {
    return { email: email, role: 'Superuser', canDo: function() { return true; } };
  }

  // Check 2: the GAS session itself is the superuser (bypass for any client instance)
  // This allows edward to open any client's ?id=SHEET_ID URL and get Superuser access
  // even though their Settings sheet has the client's email as ownerEmail.
  try {
    var sessionEmail = Session.getActiveUser().getEmail();
    if (!sessionEmail) sessionEmail = Session.getEffectiveUser().getEmail();
    if (sessionEmail && typeof SUPERUSER_EMAIL !== 'undefined' &&
        sessionEmail.toLowerCase() === SUPERUSER_EMAIL.toLowerCase()) {
      return {
        email: sessionEmail,
        role:  'Superuser',
        canDo: function() { return true; }
      };
    }
  } catch(e) {}

  try {
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.USERS);

    // First deploy -- no Users sheet yet -> grant Owner
    if (!sheet) return makeCtx(email, 'Owner');

    // Sheet exists but empty -- deny (security: accidental deletion)
    if (sheet.getLastRow() < 2) {
      Logger.log('SECURITY: Users sheet empty -- access denied for ' + email);
      return makeCtx(email, null);
    }

    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var rowEmail  = rows[i][0] ? rows[i][0].toString().toLowerCase().trim() : '';
      var rowActive = rows[i][4] !== false && rows[i][4] !== 'FALSE' && rows[i][4] !== '';
      if (rowEmail === email.toLowerCase() && rowActive) {
        return makeCtx(email, rows[i][1].toString() || 'ReadOnly');
      }
    }
    // Not registered
    return makeCtx(email, null);

  } catch(e) {
    Logger.log('_getCurrentUserContext error: ' + e.toString());
    return makeCtx(email, 'ReadOnly');
  }
}

/**
 * _auth(action, params)
 *
 * Permission gate. Call at the top of any write/sensitive function.
 * Throws if denied (caught by the router's try/catch).
 * Returns the context object for audit logging.
 *
 * @param {string} action  - e.g. 'invoices.write'
 * @param {Object} [params] - Request params containing _sheetId
 */
function _auth(action, params) {
  var ctx = _getCurrentUserContext(params);
  if (ctx.role === null) {
    throw new Error(
      'Access denied: your account (' + ctx.email + ') is not registered in this system. ' +
      'Ask the Owner to add you in Settings -> Users.'
    );
  }
  if (!ctx.canDo(action)) {
    throw new Error(
      'Permission denied: your role (' + ctx.role + ') cannot perform \'' + action + '\'. ' +
      'Contact the Owner if you need access.'
    );
  }
  return ctx;
}

// -----------------------------------------------------------------------------
// USER MANAGEMENT
// -----------------------------------------------------------------------------

/**
 * getCurrentUserWithRole(params)
 * Returns email, role, and flattened permissions for the frontend.
 */
function getCurrentUserWithRole(params) {
  try {
    var ctx = _getCurrentUserContext(params);
    try { if (ctx.email) pingRegistry('login'); } catch(pe) {}
    return {
      success: true,
      email:   ctx.email,
      role:    ctx.role || 'Unregistered',
      permissions: {
        canWrite:             ctx.canDo('invoices.write'),
        canManageUsers:       ctx.canDo('users.manage'),
        canViewUsers:         ctx.canDo('users.view'),
        canManageSettings:    ctx.canDo('settings.write'),
        canViewReports:       ctx.canDo('reports.read'),
        canViewTaxReports:    ctx.canDo('reports.tax'),
        canManageBanking:     ctx.canDo('banking.reconcile'),
        canSubmitMTD:         ctx.canDo('mtd.submit'),
        canManageCredentials: ctx.canDo('credentials.manage'),
        canRunMaintenance:    ctx.canDo('maintenance.run'),
        canViewCOA:           ctx.canDo('coa.read'),
        canManageCOA:         ctx.canDo('coa.write'),
        canManageSuppliers:   ctx.canDo('suppliers.write')
      }
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * getAllUsers(params)
 * Returns all registered users. Requires users.view permission.
 */
function getAllUsers(params) {
  try {
    _auth('users.view', params);
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.USERS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, users: [] };
    var rows  = sheet.getDataRange().getValues();
    var users = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      users.push({
        email:     rows[i][0].toString(),
        role:      rows[i][1].toString(),
        addedBy:   rows[i][2].toString(),
        addedDate: safeSerializeDate(rows[i][3]),
        active:    rows[i][4] !== false && rows[i][4] !== 'FALSE',
        notes:     rows[i][5] ? rows[i][5].toString() : ''
      });
    }
    return { success: true, users: users };
  } catch(e) {
    return { success: false, message: e.toString(), users: [] };
  }
}

/**
 * manageUser(action, email, role, notes, params)
 * Add, update role, or deactivate a user.
 * action: 'add' | 'update' | 'deactivate'
 * Requires users.manage (Owner only per ROLE_PERMISSIONS).
 */

// -----------------------------------------------------------------------------
// USER INVITATION EMAIL
// -----------------------------------------------------------------------------

/**
 * _sendUserInvitation(email, role, invitedBy, settings, params)
 * Sends a friendly welcome email to a newly added user.
 */

// -----------------------------------------------------------------------------
// EMAIL OTP AUTHENTICATION
// -----------------------------------------------------------------------------

/**
 * sendOTP(email, params)
 * Generates a 6-digit OTP, stores it in Script Properties with 10min expiry,
 * and emails it to the user.
 */
function sendOTP(email, params) {
  try {
    if (!email || !email.trim()) return { success: false, message: 'Email required' };
    email = email.trim().toLowerCase();

    // Generate 6-digit OTP
    var otp     = Math.floor(100000 + Math.random() * 900000).toString();
    var expiry  = new Date().getTime() + (10 * 60 * 1000); // 10 minutes
    var key     = 'OTP_' + email.replace(/[^a-z0-9]/g, '_');
    var payload = otp + ':' + expiry;

    PropertiesService.getScriptProperties().setProperty(key, payload);

    var sheetId     = params && params._sheetId ? params._sheetId : '';
    var settings    = {};
    try { settings = getSettings(params || {}); } catch(e) {}
    var companyName = settings.companyName || 'no~bull books';

    var nl   = '\n';
    var body = 'Your no~bull books verification code' + nl + nl +
      'Code: ' + otp + nl + nl +
      'This code expires in 10 minutes.' + nl +
      'If you did not request this, please ignore this email.' + nl + nl +
      'no~bull books -- ' + companyName;

    MailApp.sendEmail({
      to:      email,
      subject: otp + ' is your no~bull books verification code',
      body:    body
    });

    Logger.log('OTP sent to: ' + email);
    return { success: true, message: 'Verification code sent to ' + email };
  } catch(e) {
    Logger.log('sendOTP error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * verifyOTP(email, otp, params)
 * Verifies the OTP and checks the user is in the Users sheet.
 * Returns { success, role, email } on success.
 */
function verifyOTP(email, otp, params) {
  try {
    if (!email || !otp) return { success: false, message: 'Email and code required' };
    email = email.trim().toLowerCase();
    otp   = otp.toString().trim();

    var key     = 'OTP_' + email.replace(/[^a-z0-9]/g, '_');
    var stored  = PropertiesService.getScriptProperties().getProperty(key);

    if (!stored) return { success: false, message: 'No verification code found. Please request a new one.' };

    var parts  = stored.split(':');
    var storedOtp    = parts[0];
    var storedExpiry = parseInt(parts[1]);

    // Check expiry
    if (new Date().getTime() > storedExpiry) {
      PropertiesService.getScriptProperties().deleteProperty(key);
      return { success: false, message: 'Code has expired. Please request a new one.' };
    }

    // Check OTP
    if (otp !== storedOtp) {
      return { success: false, message: 'Incorrect code. Please try again.' };
    }

    // Valid -- delete OTP so it can't be reused
    PropertiesService.getScriptProperties().deleteProperty(key);

    // Check Users sheet
    var ss    = getDb(params || {});
    var sheet = ss ? ss.getSheetByName(SHEETS.USERS) : null;

    // Superuser bypass
    if (typeof SUPERUSER_EMAIL !== 'undefined' && email === SUPERUSER_EMAIL.toLowerCase()) {
      return { success: true, email: email, role: 'Superuser', verified: true };
    }

    if (!sheet || sheet.getLastRow() < 2) {
      // No users sheet or empty -- grant Owner to first user (bootstrap)
      return { success: true, email: email, role: 'Owner', verified: true };
    }

    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var rowEmail  = rows[i][0] ? rows[i][0].toString().toLowerCase().trim() : '';
      var rowActive = rows[i][4] !== false && rows[i][4] !== 'FALSE' && rows[i][4] !== '';
      if (rowEmail === email && rowActive) {
        return { success: true, email: email, role: rows[i][1].toString() || 'ReadOnly', verified: true };
      }
    }

    // Not in Users sheet
    return {
      success:     false,
      accessDenied: true,
      email:       email,
      message:     'Your account (' + email + ') is not registered for this no~bull books instance. Ask the owner to add you in Settings > Users.'
    };

  } catch(e) {
    Logger.log('verifyOTP error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}




function _sendUserInvitation(email, role, invitedBy, settings, params) {
  try {
    var companyName = (settings && settings.companyName) ? settings.companyName : 'your company';
    var sheetId     = params && params._sheetId ? params._sheetId : '';
    var appUrl      = 'https://script.google.com/a/macros/nobull.consulting/s/AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j/exec?id=' + sheetId;
    var nl          = '\n';
    var subject     = 'You have been invited to no~bull books - ' + companyName;
    var body =
      'Hi,' + nl + nl +
      'You have been invited to access no~bull books for ' + companyName + '.' + nl + nl +
      'Email:    ' + email + nl +
      'Role:     ' + role + nl +
      'Added by: ' + invitedBy + nl + nl +
      'HOW TO ACCESS' + nl +
      'no~bull books uses your Google account for secure login.' + nl + nl +
      '1. Open no~bull books: ' + appUrl + nl + nl +
      '2. Sign in with your Google account (' + email + ').' + nl +
      '   If not a Gmail address, create a Google account at accounts.google.com' + nl +
      '   and choose Use my current email address instead.' + nl + nl +
      'Questions? Contact ' + invitedBy + nl + nl +
      'Best regards,' + nl +
      'no~bull books' + nl +
      'nobull.consulting';
    MailApp.sendEmail({ to: email, subject: subject, body: body });
    Logger.log('Invitation sent to: ' + email);
  } catch(e) {
    Logger.log('_sendUserInvitation error: ' + e.toString());
  }
}


function manageUser(action, email, role, notes, params) {
  try {
    _auth('users.manage', params);
    if (!email || !email.trim()) throw new Error('Email is required');
    email = email.trim().toLowerCase();

    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.USERS);
    if (!sheet) throw new Error('Users sheet not found -- run Settings -> Initialise System first.');

    var ctx  = _getCurrentUserContext(params);
    var rows = sheet.getLastRow() >= 2 ? sheet.getDataRange().getValues() : [[]];

    var existingRow = -1;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] && rows[i][0].toString().toLowerCase().trim() === email) {
        existingRow = i + 1; // 1-based sheet row
        break;
      }
    }

    if (action === 'add') {
      if (existingRow > 0) {
        var wasInactive = rows[existingRow - 1][4] === false || rows[existingRow - 1][4] === 'FALSE';
        if (wasInactive) {
          sheet.getRange(existingRow, 2).setValue(role || 'ReadOnly');
          sheet.getRange(existingRow, 5).setValue(true);
          if (notes) sheet.getRange(existingRow, 6).setValue(notes);
          logAudit('REACTIVATE_USER', 'User', email, { role: role });
          return { success: true, message: 'User reactivated with role ' + role };
        }
        return { success: false, message: 'User ' + email + ' already exists.' };
      }
      var validRoles = ['Owner','Admin','Accountant','Staff','ReadOnly'];
      if (validRoles.indexOf(role) === -1) throw new Error('Invalid role: ' + role);
      sheet.appendRow([email, role, ctx.email, new Date(), true, notes || '']);
      logAudit('ADD_USER', 'User', email, { role: role, addedBy: ctx.email });
      // Send invitation email to new user
      var settings = getSettings(params || {});
      _sendUserInvitation(email, role, ctx.email, settings, params);
      // Alert the owner
      _sendAlert('User account added',
        'New user added to no~bull books.\nEmail: ' + email + '\nRole: ' + role + '\nAdded by: ' + ctx.email);
      return { success: true, message: 'User ' + email + ' added as ' + role + '. Invitation email sent.' };
    }

    if (existingRow < 0) return { success: false, message: 'User not found: ' + email };

    if (action === 'update') {
      var validRoles2 = ['Owner','Admin','Accountant','Staff','ReadOnly'];
      if (validRoles2.indexOf(role) === -1) throw new Error('Invalid role: ' + role);
      var oldRole = rows[existingRow - 1][1];
      sheet.getRange(existingRow, 2).setValue(role);
      if (notes !== undefined) sheet.getRange(existingRow, 6).setValue(notes || '');
      logAudit('UPDATE_USER_ROLE', 'User', email, { oldRole: oldRole, newRole: role });
      return { success: true, message: 'User ' + email + ' role updated to ' + role };
    }

    if (action === 'deactivate') {
      sheet.getRange(existingRow, 5).setValue(false);
      logAudit('DEACTIVATE_USER', 'User', email, { deactivatedBy: ctx.email });
      _sendAlert('User account deactivated',
        'User removed from no~bull books.\nEmail: ' + email + '\nDeactivated by: ' + ctx.email);
      return { success: true, message: 'User ' + email + ' deactivated.' };
    }

    return { success: false, message: 'Unknown action: ' + action };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getAuthUrl() {
  // This is a direct, hardcoded link including the ITSA scopes
  return "https://test-api.service.hmrc.gov.uk/oauth/authorize?client_id=ylNDYLn5yc3ri0Gb2Sj7iVgCRWH2&response_type=code&scope=read:vat%20write:vat%20read:self-assessment%20write:self-assessment&redirect_uri=https://script.google.com/a/macros/nobull.consulting/s/AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j/exec";
}