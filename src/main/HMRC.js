/**
 * NO~BULL BOOKS — HMRC MTD INTEGRATION
 * OAuth token management, VAT obligations, submission, liabilities, payments.
 *
 * HMRC credentials (clientID, clientSecret, accessToken, tokenExpiry) are
 * stored in Script Properties — never in the spreadsheet.
 *
 * MTD OAuth flow:
 *  1. User clicks Connect → getHMRCManualAuthUrl() → opens browser popup
 *  2. User signs in to HMRC, copies auth code from redirect URL
 *  3. User pastes code → exchangeHMRCCode(code) → stores access token
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ── Script Property keys (same as Settings.gs HMRC_PROP_KEYS) ────────────────
var _HMRC_KEYS = {
  CLIENT_ID:     'hmrc_client_id',
  CLIENT_SECRET: 'hmrc_client_secret',
  ACCESS_TOKEN:  'hmrc_access_token',
  TOKEN_EXPIRY:  'hmrc_token_expiry'
};

function _getHMRCToken() {
  var props = PropertiesService.getScriptProperties();
  return {
    clientId:     props.getProperty(_HMRC_KEYS.CLIENT_ID)     || '',
    clientSecret: props.getProperty(_HMRC_KEYS.CLIENT_SECRET) || '',
    accessToken:  props.getProperty(_HMRC_KEYS.ACCESS_TOKEN)  || '',
    tokenExpiry:  props.getProperty(_HMRC_KEYS.TOKEN_EXPIRY)  || ''
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// AUTH STATUS
// ─────────────────────────────────────────────────────────────────────────────

function getHMRCAuthStatus(params) {
  try {
    var t        = _getHMRCToken();
    var settings = getSettings(params || {});
    var testMode = settings.hmrcTestMode !== false;

    if (!t.accessToken) {
      return { success: true, connected: false, expired: true, testMode: testMode, expiresIn: 0 };
    }

    var now       = new Date();
    var expiry    = t.tokenExpiry ? new Date(t.tokenExpiry) : null;
    var expired   = !expiry || expiry <= now;
    var expiresIn = expiry ? Math.max(0, Math.round((expiry - now) / 60000)) : 0;

    // Extract VRN from stored token metadata if available
    var vrn = settings.vatRegNumber ? settings.vatRegNumber.replace(/[^0-9]/g, '') : '';

    return {
      success:   true,
      connected: !expired,
      expired:   expired,
      expiresIn: expiresIn,
      testMode:  testMode,
      vrn:       vrn
    };
  } catch(e) {
    Logger.log('getHMRCAuthStatus error: ' + e.toString());
    return { success: false, connected: false, expired: true, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// OAUTH — MANUAL FLOW (copy-paste code)
// ─────────────────────────────────────────────────────────────────────────────

function getHMRCManualAuthUrl(params) {
  try {
    var t        = _getHMRCToken();
    var settings = getSettings(params || {});
    var testMode = settings.hmrcTestMode !== false;

    if (!t.clientId) {
      return { success: false, message: 'HMRC Client ID not set — add it in Settings → HMRC/MTD.' };
    }

    var baseUrl   = testMode
      ? 'https://test-api.service.hmrc.gov.uk/oauth/authorize'
      : 'https://api.service.hmrc.gov.uk/oauth/authorize';
    var scriptUrl = 'https://script.google.com/a/macros/nobull.consulting/s/AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j/exec';
    var state     = Utilities.base64Encode(new Date().getTime().toString());

    var url = baseUrl +
      '?response_type=code' +
      '&client_id=' + encodeURIComponent(t.clientId) +
      '&scope=' + encodeURIComponent('read:vat write:vat') +
      '&redirect_uri=' + encodeURIComponent(scriptUrl) +
      '&state=' + state;

    return { success: true, url: url };
  } catch(e) {
    Logger.log('getHMRCManualAuthUrl error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function exchangeHMRCCode(code, params) {
  try {
    var t        = _getHMRCToken();
    var settings = getSettings(params || {});
    var testMode = settings.hmrcTestMode !== false;

    if (!t.clientId || !t.clientSecret) {
      return { success: false, message: 'HMRC Client ID and Secret must be set in Settings first.' };
    }

    var tokenUrl  = testMode
      ? 'https://test-api.service.hmrc.gov.uk/oauth/token'
      : 'https://api.service.hmrc.gov.uk/oauth/token';
    var scriptUrl = 'https://script.google.com/a/macros/nobull.consulting/s/AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j/exec';

    var payload = 'grant_type=authorization_code' +
      '&code=' + encodeURIComponent(code) +
      '&redirect_uri=' + encodeURIComponent(scriptUrl) +
      '&client_id=' + encodeURIComponent(t.clientId) +
      '&client_secret=' + encodeURIComponent(t.clientSecret);

    var response = UrlFetchApp.fetch(tokenUrl, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: payload,
      muteHttpExceptions: true
    });

    var json = JSON.parse(response.getContentText());
    if (json.error) throw new Error(json.error_description || json.error);

    var expiry = new Date(new Date().getTime() + ((json.expires_in || 14400) * 1000));
    var props  = PropertiesService.getScriptProperties();
    props.setProperty(_HMRC_KEYS.ACCESS_TOKEN, json.access_token || '');
    props.setProperty(_HMRC_KEYS.TOKEN_EXPIRY,  expiry.toISOString());

    return {
      success: true,
      message: 'Connected to HMRC MTD. Token expires: ' + expiry.toLocaleString()
    };
  } catch(e) {
    Logger.log('exchangeHMRCCode error: ' + e.toString());
    return { success: false, message: 'Token exchange failed: ' + e.toString() };
  }
}

function testHMRCConnection(params) {
  try {
    var t    = _getHMRCToken();
    var settings = getSettings(params || {});
    var testMode = settings.hmrcTestMode !== false;

    if (!t.accessToken) {
      return { success: false, message: 'Not connected — no access token stored.' };
    }

    var baseUrl = testMode
      ? 'https://test-api.service.hmrc.gov.uk'
      : 'https://api.service.hmrc.gov.uk';

    var response = UrlFetchApp.fetch(baseUrl + '/hello/user', {
      headers: { Authorization: 'Bearer ' + t.accessToken },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      return { success: true, message: 'Connection to HMRC MTD confirmed.' };
    }
    return { success: false, message: 'HMRC returned HTTP ' + response.getResponseCode() + '. Token may be expired.' };
  } catch(e) {
    Logger.log('testHMRCConnection error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// VAT MTD — OBLIGATIONS, SUBMISSION, LIABILITIES, PAYMENTS
// ─────────────────────────────────────────────────────────────────────────────

function getVATObligations(vrn, fromDate, toDate, params) {
  try {
    var t        = _getHMRCToken();
    var settings = getSettings(params);
    var testMode = settings.hmrcTestMode !== false;

    if (!t.accessToken) return { success: false, message: 'Not connected to HMRC MTD.' };
    if (!vrn)           return { success: false, message: 'VAT registration number is required.' };

    var baseUrl = testMode
      ? 'https://test-api.service.hmrc.gov.uk'
      : 'https://api.service.hmrc.gov.uk';

    var url = baseUrl + '/organisations/vat/' + vrn + '/obligations?from=' + fromDate + '&to=' + toDate;
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + t.accessToken,
        Accept: 'application/vnd.hmrc.1.0+json'
      },
      muteHttpExceptions: true
    });

    var json = JSON.parse(response.getContentText());
    if (json.code || json.message) throw new Error(json.message || json.code);

    return { success: true, obligations: json.obligations || [] };
  } catch(e) {
    Logger.log('getVATObligations error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function submitVATReturn(vrn, periodKey, params) {
  try {
    var t        = _getHMRCToken();
    var settings = getSettings(params);
    var testMode = settings.hmrcTestMode !== false;

    if (!t.accessToken) return { success: false, message: 'Not connected to HMRC MTD.' };
    if (!vrn || !periodKey) return { success: false, message: 'VRN and period key are required.' };

    var baseUrl = testMode
      ? 'https://test-api.service.hmrc.gov.uk'
      : 'https://api.service.hmrc.gov.uk';

    var body = {
      periodKey:             periodKey,
      vatDueSales:           parseFloat(params.box1) || 0,
      vatDueAcquisitions:    parseFloat(params.box2) || 0,
      totalVatDue:           parseFloat(params.box3) || 0,
      vatReclaimedCurrPeriod:parseFloat(params.box4) || 0,
      netVatDue:             Math.abs(parseFloat(params.box5) || 0),
      totalValueSalesExVAT:  Math.round(parseFloat(params.box6) || 0),
      totalValuePurchasesExVAT: Math.round(parseFloat(params.box7) || 0),
      totalValueGoodsSuppliedExVAT: Math.round(parseFloat(params.box8) || 0),
      totalAcquisitionsExVAT: Math.round(parseFloat(params.box9) || 0),
      finalised: true
    };

    var response = UrlFetchApp.fetch(baseUrl + '/organisations/vat/' + vrn + '/returns', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Bearer ' + t.accessToken,
        Accept: 'application/vnd.hmrc.1.0+json'
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    var json = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 201) {
      throw new Error(json.message || json.code || 'HTTP ' + response.getResponseCode());
    }

    // Mark as submitted in local VAT returns sheet
    params.status        = 'Submitted';
    params.submittedDate = new Date().toISOString().split('T')[0];
    params.periodKey     = periodKey;
    try { saveVATReturn(params); } catch(se) { Logger.log('saveVATReturn after submit: ' + se); }

    logAudit('SUBMIT', 'VATReturn', vrn, { periodKey: periodKey });
    return { success: true, message: 'VAT return submitted to HMRC.', reference: json.paymentIndicator || '' };
  } catch(e) {
    Logger.log('submitVATReturn error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getVATLiabilities(vrn, fromDate, toDate, params) {
  try {
    var t        = _getHMRCToken();
    var settings = getSettings(params);
    var testMode = settings.hmrcTestMode !== false;

    if (!t.accessToken) return { success: false, message: 'Not connected to HMRC MTD.' };
    if (!vrn)           return { success: false, message: 'VAT registration number required.' };

    var baseUrl  = testMode
      ? 'https://test-api.service.hmrc.gov.uk'
      : 'https://api.service.hmrc.gov.uk';

    var url = baseUrl + '/organisations/vat/' + vrn + '/liabilities?from=' + fromDate + '&to=' + toDate;
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + t.accessToken, Accept: 'application/vnd.hmrc.1.0+json' },
      muteHttpExceptions: true
    });

    var json = JSON.parse(response.getContentText());
    if (json.code || (response.getResponseCode() !== 200)) {
      throw new Error(json.message || 'HTTP ' + response.getResponseCode());
    }
    return { success: true, liabilities: json.liabilities || [] };
  } catch(e) {
    Logger.log('getVATLiabilities error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getVATPayments(vrn, fromDate, toDate, params) {
  try {
    var t        = _getHMRCToken();
    var settings = getSettings(params);
    var testMode = settings.hmrcTestMode !== false;

    if (!t.accessToken) return { success: false, message: 'Not connected to HMRC MTD.' };
    if (!vrn)           return { success: false, message: 'VAT registration number required.' };

    var baseUrl  = testMode
      ? 'https://test-api.service.hmrc.gov.uk'
      : 'https://api.service.hmrc.gov.uk';

    var url = baseUrl + '/organisations/vat/' + vrn + '/payments?from=' + fromDate + '&to=' + toDate;
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + t.accessToken, Accept: 'application/vnd.hmrc.1.0+json' },
      muteHttpExceptions: true
    });

    var json = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) {
      throw new Error(json.message || 'HTTP ' + response.getResponseCode());
    }
    return { success: true, payments: json.payments || [] };
  } catch(e) {
    Logger.log('getVATPayments error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}