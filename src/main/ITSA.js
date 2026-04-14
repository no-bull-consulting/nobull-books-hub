/**
 * NO~BULL BOOKS — ITSA QUARTERLY CALCULATOR
 * Aggregates income and expenses for Self Assessment Quarterly Updates.
 */
function calculateITSAQuarter(startDate, endDate, params) {
  try {
    const ss = getDb(params);
    const sheet = ss.getSheetByName('Transactions');
    const data = sheet.getDataRange().getValues();
    
    let itsaReport = {
      income: 0,
      expenses: {
        directCosts: 0,    // 5xxx codes
        overheads: 0,      // 7xxx codes
        other: 0           // Anything else
      },
      netProfit: 0,
      estimatedTax: 0      // 20% Basic Rate placeholder
    };

    const start = new Date(startDate).getTime();
    const end = new Date(endDate).getTime();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[1]) continue; 

      const txDate = new Date(row[1]).getTime(); // Column B

      if (txDate >= start && txDate <= end) {
        const amount = parseFloat(row[6]) || 0; // Column G
        const debitCode = row[4] ? row[4].toString() : ''; // Column E
        const creditCode = row[5] ? row[5].toString() : ''; // Column F

        // --- INCOME TRACKING ---
        if (creditCode.startsWith('4')) {
          itsaReport.income += amount;
        }

        // --- EXPENSE TRACKING ---
        if (debitCode.startsWith('5')) {
          itsaReport.expenses.directCosts += amount;
        } else if (debitCode.startsWith('7')) {
          itsaReport.expenses.overheads += amount;
        } else if (debitCode.startsWith('6') || debitCode.startsWith('8')) {
          itsaReport.expenses.other += amount;
        }
      }
    }

    itsaReport.netProfit = itsaReport.income - (itsaReport.expenses.directCosts + itsaReport.expenses.overheads + itsaReport.expenses.other);
    
    // Quick Tax Est: (Net Profit - Personal Allowance / 4) * 20%
    // Note: This is a rough estimation for the UI only.
    itsaReport.estimatedTax = Math.max(0, itsaReport.netProfit * 0.20);

    logAudit('ITSA_CALC_RUN', 'System', startDate + ' to ' + endDate, { profit: itsaReport.netProfit }, params);

    return { success: true, report: itsaReport };

  } catch (e) {
    logAudit('ITSA_CALC_ERROR', 'System', 'ITSA Engine', e.toString(), params);
    return { success: false, error: e.toString() };
  }
}

/**
 * Fetches ITSA Obligations from HMRC
 * Uses NINO and Business ID from Settings
 */
function getITSAObligations(params) {
  try {
    const settings = getSettings(params);
    const nino = settings.hmrcNINO; // From your screenshot
    const businessId = settings.mtdBusinessId; // From your screenshot
    
    if (!nino || !businessId) throw new Error("NINO and Business ID must be configured in Settings.");

    const t = _getHMRCToken();
    const testMode = settings.hmrcTestMode !== false;
    const baseUrl = testMode ? 'https://test-api.service.hmrc.gov.uk' : 'https://api.service.hmrc.gov.uk';

    // ITSA endpoint differs from VAT
    const url = `${baseUrl}/income-tax/nino/${nino}/business-id/${businessId}/obligations`;
    
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + t.accessToken,
        'Accept': 'application/vnd.hmrc.1.0+json'
      },
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    
    logAudit('ITSA_OBLIGATIONS_LOAD', 'HMRC', nino, { status: response.getResponseCode() }, params);
    
    return { success: true, obligations: json.obligations || [] };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Sends the Quarterly Update to HMRC
 */
function submitITSAUpdate(quarterData, params) {
  try {
    const settings = getSettings(params);
    const nino = settings.hmrcNINO;
    const t = _getHMRCToken();
    
    // HMRC Payload structure for ITSA Quarterly Updates
    const payload = {
      "incomes": [{ "incomeSourceType": "self-employment", "amount": quarterData.income }],
      "expenses": [
        { "expenseType": "consolidated-expenses", "amount": (quarterData.expenses.directCosts + quarterData.expenses.overheads) }
      ]
    };

    const baseUrl = settings.hmrcTestMode !== false ? 'https://test-api.service.hmrc.gov.uk' : 'https://api.service.hmrc.gov.uk';
    const url = `${baseUrl}/income-tax/nino/${nino}/periodic-updates`;

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + t.accessToken, 'Accept': 'application/vnd.hmrc.1.0+json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());
    
    // Log the submission to the Audit Log
    logAudit('ITSA_SUBMISSION', 'HMRC', nino, { status: response.getResponseCode(), ref: result.submissionId }, params);

    return { success: response.getResponseCode() === 201, data: result };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Test call to fetch ITSA Obligations
 */
function testITSAConnection() {
  const nino = "TS960860B"; // [cite: 81]
  const t = _getHMRCToken(); // Ensure this gets the fresh token
  
  if (!t.accessToken) {
    console.log("No token found. Please re-authorize via the app UI.");
    return;
  }

  // The specific URL from your documentation screenshot
  const url = `https://test-api.service.hmrc.gov.uk/obligations/details/${nino}/income-and-expenditure?status=O`;

  const options = {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + t.accessToken,
      'Accept': 'application/vnd.hmrc.3.0+json', // Matches your subscription
      'Gov-Client-Connection-Method': 'WEB_APP_VIA_SERVER'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  console.log("Status: " + response.getResponseCode());
  console.log("Data: " + response.getContentText());
}

function getITSAObligationsFromSheet(params) {
  var props = PropertiesService.getScriptProperties();
  var accessToken = props.getProperty('hmrc_access_token');
  
  // 1. Get Settings and extract NINO using your 'hmrcNINO' key
  var settings = getSettings(params); 
  // We check for 'hmrcNINO' (your label) and fallback to 'nino' just in case
  var nino = (settings.hmrcNINO || settings.nino || '').replace(/\s/g, '').toUpperCase(); 
  
  // CRITICAL: Check the variable 'nino', not 'hmrcNINO'
  if (!nino) return { success: false, message: "NINO missing in Settings (checked column 'hmrcNINO')" };

  // 2. Correct HMRC ITSA Obligations URL
  // The endpoint is /individuals/obligations/nino/{nino}
  var url = "https://test-api.service.hmrc.gov.uk/individuals/obligations/nino/" + nino;
  
  var options = {
    "method": "get",
    "muteHttpExceptions": true,
    "headers": {
      "Authorization": "Bearer " + accessToken,
      "Accept": "application/vnd.hmrc.1.0+json",
      "Gov-Test-Scenario": "QUARTERLY_MET" 
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseText = response.getContentText();
    var data = JSON.parse(responseText);
    
    // If HMRC returns a 404, it often means "No obligations found for this user"
    if (response.getResponseCode() === 404) {
      return { success: true, obligations: [], message: "No obligations found for this NINO." };
    }

    if (response.getResponseCode() !== 200) {
      return { success: false, message: "HMRC Error: " + (data.message || responseText) };
    }

    // 3. Return the data
    return { success: true, obligations: data.obligations || [] };
    
  } catch (e) {
    return { success: false, message: "Connection Error: " + e.toString() };
  }
}