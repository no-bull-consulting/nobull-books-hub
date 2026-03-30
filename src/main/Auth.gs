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
  try {
    email = Session.getActiveUser().getEmail();
    if (!email) email = Session.getEffectiveUser().getEmail();
  } catch(e) {}

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
  if (typeof SUPERUSER_EMAIL !== 'undefined' &&
      email.toLowerCase() === SUPERUSER_EMAIL.toLowerCase()) {
    return { email: email, role: 'Superuser', canDo: function() { return true; } };
  }

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
      _sendAlert('User account added',
        'New user added to no~bull books.\nEmail: ' + email + '\nRole: ' + role + '\nAdded by: ' + ctx.email);
      return { success: true, message: 'User ' + email + ' added as ' + role };
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
