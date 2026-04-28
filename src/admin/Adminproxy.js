/**
 * NO~BULL BOOKS HUB — ADMIN CONSOLE PROXY
 * ─────────────────────────────────────────────────────────────────────────────
 * Thin wrapper that forwards registry/admin calls from the Hub to the
 * separate Admin Console GAS deployment.
 *
 * The Admin Console URL and shared secret are stored in Script Properties:
 *   ADMIN_CONSOLE_URL  — the Admin Console /exec URL
 *   ADMIN_SECRET       — shared secret for authentication
 */

function _callAdminConsole(action, params) {
  try {
    var url    = PropertiesService.getScriptProperties().getProperty('ADMIN_CONSOLE_URL');
    var secret = PropertiesService.getScriptProperties().getProperty('ADMIN_SECRET');

    if (!url) {
      Logger.log('ADMIN_CONSOLE_URL not set — falling back to local registry');
      return null; // Caller handles fallback
    }

    var body = Object.assign({}, params || {}, { action: action, secret: secret });

    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    return JSON.parse(response.getContentText());
  } catch(e) {
    Logger.log('_callAdminConsole error: ' + e.toString());
    return null;
  }
}

/**
 * pingRegistryRemote — forwards registry ping to Admin Console
 * Falls back to local pingRegistry if Admin Console not configured
 */
function pingRegistryRemote(sheetId, data) {
  var r = _callAdminConsole('pingRegistry', Object.assign({ sheetId: sheetId }, data));
  if (r) return r;
  // Fallback to local (during transition period)
  try { return pingRegistry(sheetId, data); } catch(e) { return { success: false }; }
}