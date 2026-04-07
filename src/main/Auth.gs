/**
 * NO~BULL BOOKS — AUTH / RBAC
 * Permission engine, user context, user management.
 *
 * KEY CHANGE (hub model): every function now accepts optional `params`
 * so the correct client spreadsheet (_sheetId) is used for Users lookups.
 * ─────────────────────────────────────────────────────────────────────────────
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

// ─────────────────────────────────────────────────────────────────────────────
// PERMISSION ENGINE
// ─────────────────────────────────────────────────────────────────────────────

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

  // ── Identity: read owner email stored in Settings sheet ───────────────────
  // Since the hub runs as USER_DEPLOYING (always edward), we cannot use
  // Session.getActiveUser() to identify the real client. Instead, the client's
  // email is stored in their Settings sheet during onboarding by SetupService,
  // and read here on every request.
  //
  // Security model: knowing the ?id=SHEET_ID URL = authorised access.
  // The sheet ID is a 44-character random string — effectively a private key.
  if (params && params._sheetId) {
    try {
      var ss       = getDb(params);
      var settings = ss.getSheetByName(SHEETS.SETTINGS);
      if (settings && settings.getLastRow() >= 2) {
        var sData = settings.getRange(2, 1, 1, settings.getLastColumn()).getValues()[0];
        // ownerEmail is stored in the Settings sheet column named 'ownerEmail'
        // We look it up by header position
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

  // ── Fallback: GAS session (works for edward's own instances) ──────────────
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

  // ── Superuser override ────────────────────────────────────────────────────
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

    // First deploy — no Users sheet yet → grant Owner
    if (!sheet) return makeCtx(email, 'Owner');

    // Sheet exists but empty — deny (security: accidental deletion)
    if (sheet.getLastRow() < 2) {
      Logger.log('SECURITY: Users sheet empty — access denied for ' + email);
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
      'Ask the Owner to add you in Settings → Users.'
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

// ─────────────────────────────────────────────────────────────────────────────
// USER MANAGEMENT
// ─────────────────────────────────────────────────────────────────────────────

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

// ─────────────────────────────────────────────────────────────────────────────
// USER INVITATION EMAIL
// ─────────────────────────────────────────────────────────────────────────────

/**
 * _sendUserInvitation(email, role, invitedBy, settings, params)
 * Sends a friendly welcome email to a newly added user.
 */
function _sendUserInvitation(email, role, invitedBy, settings, params) {
  try {
    var companyName = (settings && settings.companyName) ? settings.companyName : 'your company';
    var sheetId     = params && params._sheetId ? params._sheetId : '';
    var appUrl      = 'https://script.google.com/a/macros/nobull.consulting/s/AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j/exec?id=' + sheetId;

    var subject = 'You have been invited to no~bull books — ' + companyName;

    var nl   = '\n';
    var body = 'Hi,' + nl + nl +
      'You have been invited to access no~bull books for ' + companyName + '.' + nl + nl +
      'Your details:' + nl +
      '  Email:    ' + email + nl +
      '  Role:     ' + role + nl +
      '  Added by: ' + invitedBy + nl + nl +
      'HOW TO ACCESS' + nl +
      'no~bull books uses your Google account for secure login.' + nl + nl +
      '1. Click the link below to open no~bull books:' + nl +
      '   ' + appUrl + nl + nl +
      '2. Sign in with your Google account (' + email + ').' + nl +
      '   If your email is not a Gmail address, you can create a free' + nl +
      '   Google account linked to it at:' + nl +
      '   https://accounts.google.com/signup/v2/createaccount' + nl +
      '   (Choose "Use my current email address instead" on the signup page)' + nl + nl +
      '3. Once signed in, you will have ' + role + ' access to the books.' + nl + nl +
      'If you have any questions, please contact ' + invitedBy + '.' + nl + nl +
      'Best regards,' + nl +
      'no~bull books' + nl +
      'nobull.consulting' + nl;

    var htmlBody =
      '<div style="font-family:-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif;max-width:560px;margin:0 auto">' +
      '<div style="background:#14213D;padding:24px 32px;border-radius:8px 8px 0 0">' +
        '<p style="color:#fff;font-size:20px;margin:0">🐂 <strong>no~bull</strong> <span style="color:#14A8AE">books</span></p>' +
      '</div>' +
      '<div style="background:#fff;border:1px solid #e2e8f0;border-top:none;padding:32px;border-radius:0 0 8px 8px">' +
        '<h2 style="color:#14213D;margin:0 0 8px">You have been invited!</h2>' +
        '<p style="color:#64748b;margin:0 0 24px">You now have access to <strong>' + companyName + '</strong> on no~bull books.</p>' +
        '<table style="width:100%;border-collapse:collapse;margin-bottom:24px">' +
          '<tr><td style="padding:10px 12px;background:#f8fafc;border:1px solid #e2e8f0;font-weight:600;color:#14213D;width:120px">Email</td>' +
              '<td style="padding:10px 12px;border:1px solid #e2e8f0;color:#475569">' + email + '</td></tr>' +
          '<tr><td style="padding:10px 12px;background:#f8fafc;border:1px solid #e2e8f0;font-weight:600;color:#14213D">Role</td>' +
              '<td style="padding:10px 12px;border:1px solid #e2e8f0;color:#475569">' + role + '</td></tr>' +
          '<tr><td style="padding:10px 12px;background:#f8fafc;border:1px solid #e2e8f0;font-weight:600;color:#14213D">Invited by</td>' +
              '<td style="padding:10px 12px;border:1px solid #e2e8f0;color:#475569">' + invitedBy + '</td></tr>' +
        '</table>' +
        '<a href="' + appUrl + '" style="display:inline-block;background:#0D7377;color:#fff;padding:14px 28px;border-radius:8px;text-decoration:none;font-weight:600;font-size:15px;margin-bottom:28px">Open no~bull books →</a>' +
        '<hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0">' +
        '<h3 style="color:#14213D;margin:0 0 12px;font-size:15px">How to sign in</h3>' +
        '<p style="color:#64748b;font-size:14px;line-height:1.7;margin:0 0 12px">' +
          'no~bull books uses your <strong>Google account</strong> for secure login. Simply click the button above and sign in with <strong>' + email + '</strong>.' +
        '</p>' +
        '<p style="color:#64748b;font-size:14px;line-height:1.7;margin:0 0 24px">' +
          'If your email is not a Gmail address, you can create a free Google account linked to it. On the Google sign-in page, click <em>"Create account"</em> and choose <em>"Use my current email address instead"</em>.' +
        '</p>' +
        '<p style="color:#94a3b8;font-size:12px;margin:0">This invitation was sent by ' + invitedBy + ' · <a href="https://nobull.consulting" style="color:#0D7377">nobull.consulting</a></p>' +
      '</div>' +
      '</div>';

    MailApp.sendEmail({
      to:       email,
      subject:  subject,
      body:     body,
      htmlBody: htmlBody
    });

    Logger.log('Invitation sent to: ' + email);
  } catch(e) {
    Logger.log('_sendUserInvitation error: ' + e.toString());
    // Non-fatal — user is still added even if email fails
  }
}

function manageUser(action, email, role, notes, params) {
  try {
    _auth('users.manage', params);
    if (!email || !email.trim()) throw new Error('Email is required');
    email = email.trim().toLowerCase();

    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(SHEETS.USERS);
    if (!sheet) throw new Error('Users sheet not found — run Settings → Initialise System first.');

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