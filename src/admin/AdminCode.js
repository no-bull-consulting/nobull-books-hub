/** * NO~BULL BOOKS — ADMIN CONSOLE (v1.0.4)
 * Consolidated and Fixed for Cross-Domain Identity
 */
const REGISTRY_ID = "13os7wkggdTpk_9l4sh7kr4dhzlCPh9E7FfQoWE6_o74";
const SPOKE_BASE_URL = "https://script.google.com/a/macros/nobull.consulting/s/AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j/exec";

function doGet(e) {
  const action = e.parameter.action;
  const userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  
  // 1. ADMIN DASHBOARD ROUTE
  // If YOU access the script directly without parameters, show the Dashboard UI
  if (!action && userEmail === "edward@nobull.consulting") {
    return HtmlService.createTemplateFromFile('AdminDashboard')
      .evaluate()
      .setTitle("no~bull | Admin Console")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 2. STATUS CHECK (For non-admins or diagnostic pings)
  if (!action) {
    return ContentService.createTextOutput("Admin API Active. Authenticated as: " + (userEmail || "Anonymous"));
  }

  // 3. IDENTITY FALLBACK (The NobleProg Spinner Fix)
  if (!userEmail) {
    if (action === "onboard" || action === "getInstances") {
      const breakoutUrl = ScriptApp.getService().getUrl() + "?action=onboard";
      return HtmlService.createHtmlOutput(
        "<div style='font-family:sans-serif;text-align:center;padding:50px;'>" +
        "<h3>Identity Verification Required</h3>" +
        "<a href='" + breakoutUrl + "' target='_top' style='background:#3b82f6;color:white;padding:12px 24px;text-decoration:none;border-radius:6px;font-weight:bold;'>Verify & Continue →</a>" +
        "</div>"
      );
    }
  }

  // 4. DISCOVERY HANDSHAKE (JSONP for the website dashboard)
  if (action === "getInstances") {
    try {
      const r = getAllRegistryClients({}); 
      const clients = r.clients || [];
      const userWorkspaces = clients.filter(function(c) {
        return c.contactEmail && c.contactEmail.toLowerCase().trim() === userEmail;
      }).map(function(c) {
        return {
          companyName: c.companyName,
          status:      c.status,
          appUrl:      c.AppUrl || c.appLink || c.sheetId 
        };
      });

      userWorkspaces.push({
        companyName: "+ Create New Books",
        status: "Action Required",
        appUrl: ScriptApp.getService().getUrl() + "?action=onboard"
      });

      return JSONP(e, { success: true, email: userEmail, instances: userWorkspaces });
    } catch (err) {
      return JSONP(e, { success: false, error: err.toString() });
    }
  }

  // 5. ONBOARDING FORM ROUTE
  if (action === "onboard") {
    const tmp = HtmlService.createTemplateFromFile('OnboardingUI');
    tmp.userEmail = userEmail; 
    return tmp.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function createWorkspace(bizName, manualEmail) {
  try {
    const userEmail = (manualEmail && manualEmail.length > 5) ? manualEmail : Session.getActiveUser().getEmail();
    if (!userEmail) throw new Error("No email detected.");

    const templateId = "1gIFwQUtbhGaM3HIHbFFaT7lIAU4BN3IksAOv1_uuUKg"; 
    const newFile = DriveApp.getFileById(templateId).makeCopy(bizName + " Books");
    const newFileId = newFile.getId();
    newFile.addEditor(userEmail);

    // FIX: Generate the Spoke URL with ID
    const finalAppUrl = SPOKE_BASE_URL + "?id=" + newFileId;

    const registrySheet = SpreadsheetApp.openById(REGISTRY_ID).getSheetByName("Registry");
    registrySheet.appendRow([
      "REG_" + new Date().getTime(), bizName, bizName, "User", userEmail, 
      "", newFileId, newFile.getUrl(), finalAppUrl, "Standard", "Trial", 
      new Date(), new Date(), userEmail, 0, 0, 0, "1.0.4", "Trial Created", 
      "No", "", "", "UK", "", "" 
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function JSONP(e, data) {
  const callback = e.parameter.callback;
  return ContentService.createTextOutput(callback + "(" + JSON.stringify(data) + ")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * handleAdminCall
 * The secure gateway for the Admin Dashboard.
 * Maps frontend requests to backend functions.
 */
function handleAdminCall(action, params) {
  // 1. SECURITY: Only allow Edward to run admin functions
  var authorizedUser = "edward@nobull.consulting";
  var currentUser = Session.getActiveUser().getEmail().toLowerCase().trim();

  if (currentUser !== authorizedUser) {
    throw new Error("Unauthorized: Admin access restricted.");
  }

  // 2. ROUTING: Map the 'action' string to actual JS functions
  try {
    switch (action) {
      case 'getAdminStats':
        return getAdminStats(params);
        
      case 'getRegistryClients':
        return getAllRegistryClients(params || {});
        
      case 'updateRegistryClient':
        // This maps to the logic in Registry.js
        return updateRegistryClient(params.registryId, params);
        
      case 'migrateSchemas':
        return migrateAllClientSchemas();

      default:
        return { success: false, message: "Unknown admin action: " + action };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}