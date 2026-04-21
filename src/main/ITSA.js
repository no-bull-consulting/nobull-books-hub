/**
 * NO~BULL BOOKS — ITSA ENGINE v3.0 (CLEANED)
 * Target: HMRC MTD for ITSA 2026/27 Mandate
 */

// --- CONFIGURATION ---
// REPLACE THIS WITH YOUR SHEET ID FROM THE URL
const TARGET_SHEET_ID = "1gIFwQUtbhGaM3HIHbFFaT7lIAU4BN3IksAOv1_uuUKg"; 

/**
 * BRIDGE: Called by Sidebar to process a quarterly submission.
 */
/**
 * BRIDGE: Now with "Undefined" protection
 */
function submitITSAQuarterlyUpdate(nino, businessId, taxYear, quarter, income, params) {
  // 1. Safety Check: If nino or businessId are missing, stop here
  if (!nino || !businessId) {
    return { 
      success: false, 
      message: "Missing NINO or Business ID. Please check your Settings and Sidebar fields." 
    };
  }

  var q = (quarter || "").toString().trim();
  var selectedPeriodKey = "";

  if (q.includes("Q1") || q === "1") selectedPeriodKey = "2026-04-06_2026-07-05";
  else if (q.includes("Q2") || q === "2") selectedPeriodKey = "2026-07-06_2026-10-05";
  else if (q.includes("Q3") || q === "3") selectedPeriodKey = "2026-10-06_2027-01-05";
  else if (q.includes("Q4") || q === "4") selectedPeriodKey = "2027-01-06_2027-04-05";

  if (!selectedPeriodKey) return { success: false, message: "Quarter not recognized: " + q };

  var quarterData = {
    income: parseFloat(income) || 0,
    periodKey: selectedPeriodKey 
  };
  
  params = params || {};
  params.hmrcNINO = nino;
  params.mtdBusinessId = businessId;

  return submitITSAUpdate(quarterData, params);
}

/**
 * ENGINE: Cleaned and Safe
 */
function submitITSAUpdate(quarterData, params) {
  try {
    // 2. The Fix for the "Replace" error: Use toString() and a fallback
    const nino = (params.hmrcNINO || "").toString().replace(/\s/g, '').toUpperCase();
    const id = (params.mtdBusinessId || "").toString().replace(/\s/g, ''); 
    
    if (nino.length < 8) throw new Error("Invalid NINO provided.");

    const t = _getHMRCToken();
    const url = `https://test-api.service.hmrc.gov.uk/income-tax/nino/${nino}/self-employment/${id}/periodic-updates?taxYear=2026-27`;

    const payload = {
      "periodKey": quarterData.periodKey,
      "income": { "turnover": quarterData.income },
      "expenses": { "consolidatedExpenses": 0 }
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      headers: { 
        'Authorization': 'Bearer ' + t.accessToken, 
        'Accept': 'application/vnd.hmrc.1.0+json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    return { 
      success: (response.getResponseCode() < 300), 
      status: response.getResponseCode(),
      detail: response.getContentText() 
    };
  } catch (e) { 
    return { success: false, message: "ITSA Engine Error: " + e.toString() }; 
  }
}

/**
 * DIAGNOSTIC: Checks if Uli is visible to the current token.
 */
function checkTaxpayerExistence() {
  try {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sheet = ss.getSheetByName('Settings');
    
    // Pull NINO from Column AC (Index 29), Row 2
    const nino = sheet.getRange(2, 29).getValue().toString().replace(/\s/g, '').toUpperCase();
    console.log("Checking NINO: " + nino);

    const t = _getHMRCToken();
    const url = `https://test-api.service.hmrc.gov.uk/individuals/details/nino/${nino}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 
        'Authorization': 'Bearer ' + t.accessToken, 
        'Accept': 'application/vnd.hmrc.2.0+json' // Use v2.0 for 2026 sandbox users
      },
      muteHttpExceptions: true
    });

    console.log("HMRC Identity Status: " + response.getResponseCode());
    console.log("HMRC Response: " + response.getContentText());

  } catch (e) { console.log("Diagnostic Error: " + e.toString()); }
}

/**
 * UTILITY: Clears old sessions.
 */
function clearTokens() {
  PropertiesService.getScriptProperties().deleteProperty('HMRC_ACCESS_TOKEN');
  console.log("Session cleared.");
}

/**
 * FETCH OBLIGATIONS: The actual function being called by Api.gs
 */
function getITSAObligationsFromSheet(params) {
  try {
    const nino = (params.nino || params.hmrcNINO || "").toString().replace(/\s/g, '').toUpperCase();
    
    if (!nino) {
      return { success: false, message: "NINO is required to fetch obligations." };
    }

    const t = _getHMRCToken();
    // Sandbox endpoint for ITSA obligations
    const url = `https://test-api.service.hmrc.gov.uk/obligations/nino/${nino}/ITSA-Business`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 
        'Authorization': 'Bearer ' + t.accessToken, 
        'Accept': 'application/vnd.hmrc.1.0+json'
      },
      muteHttpExceptions: true
    });

    const resBody = JSON.parse(response.getContentText());

    if (response.getResponseCode() === 200) {
      return { 
        success: true, 
        obligations: resBody.obligations || [] 
      };
    } else {
      return { 
        success: false, 
        message: "HMRC Error: " + (resBody.message || response.getContentText()) 
      };
    }
  } catch (e) {
    return { success: false, message: "System Error: " + e.toString() };
  }
}

/**
 * ENGINE: Fetches obligations from HMRC for a specific NINO.
 * This is the function that was "missing" or misnamed.
 */
function getITSAObligations(params) {
  try {
    // 1. Get NINO from params (sent by sidebar) or fallback to Settings sheet
    let nino = (params.nino || params.hmrcNINO || "").toString().replace(/\s/g, '').toUpperCase();
    
    if (!nino) {
      const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
      const sheet = ss.getSheetByName('Settings');
      nino = sheet.getRange(2, 29).getValue().toString().replace(/\s/g, '').toUpperCase();
    }

    if (!nino || nino.length < 5) {
      return { success: false, message: "Valid NINO is required to fetch obligations." };
    }

    const t = _getHMRCToken();
    
    // 2. Call the HMRC Obligations endpoint for ITSA
    const url = `https://test-api.service.hmrc.gov.uk/obligations/nino/${nino}/ITSA-Business`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 
        'Authorization': 'Bearer ' + t.accessToken, 
        'Accept': 'application/vnd.hmrc.1.0+json'
      },
      muteHttpExceptions: true
    });

    const resCode = response.getResponseCode();
    const resBody = JSON.parse(response.getContentText());

    if (resCode === 200) {
      // Map HMRC fields to your UI table fields
      const obligations = (resBody.obligations || []).map(ob => ({
        taxYear: ob.taxYear || "2026-27",
        quarter: ob.periodKey,
        periodStart: ob.start,
        periodEnd: ob.end,
        dueDate: ob.due,
        status: ob.status === "O" ? "Open" : "Fulfilled"
      }));

      return { success: true, obligations: obligations };
    } else {
      return { 
        success: false, 
        message: "HMRC Error (" + resCode + "): " + (resBody.message || "Unknown error") 
      };
    }
  } catch (e) {
    return { success: false, message: "System Error: " + e.toString() };
  }
}