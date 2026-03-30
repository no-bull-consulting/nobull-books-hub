# One-Time Setup Commands

These functions are run **once** from the GAS editor (main hub) after first deployment.  
Open the editor at [script.google.com](https://script.google.com), select the main hub project, and run each function from the editor toolbar.

---

## Required — run in this order

### 1. Initialise the registry

Creates the central client registry Google Sheet in edward@nobull.consulting's Drive and stores its ID in Script Properties.

```javascript
function _runOnce_initRegistry() {
  initRegistry();
}
```

After running, note the registry Sheet ID printed in the execution log.

---

### 2. Set the Setup Service URL

```javascript
function _runOnce_setSetupUrl() {
  setSetupServiceUrl('https://script.google.com/macros/s/YOUR_SETUP_SCRIPT_ID/exec');
}
```

Replace `YOUR_SETUP_SCRIPT_ID` with the actual setup microservice deployment ID.

---

### 3. Set trial duration

```javascript
function _runOnce_setTrialDays() {
  setTrialDays(14);
}
```

---

### 4. Set Gemini API key

```javascript
function _runOnce_setGemini() {
  setGeminiKey('AIzaSy...');
}
```

Obtain from [aistudio.google.com](https://aistudio.google.com) using your personal Gmail account.

---

### 5. Set default spreadsheet (editor testing only)

Optional — only needed for debugging in the GAS editor where no `?id=` URL parameter is available.

```javascript
function _runOnce_setDefaultSheet() {
  // Use your own test sheet ID
  PropertiesService.getScriptProperties().setProperty('DEFAULT_SPREADSHEET_ID', 'YOUR_TEST_SHEET_ID');
}
```

---

## Verify setup

Run this to confirm all properties are set:

```javascript
function _verify_scriptProperties() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var required = ['REGISTRY_SHEET_ID', 'SETUP_SERVICE_URL', 'TRIAL_DAYS', 'GEMINI_API_KEY'];
  required.forEach(function(k) {
    Logger.log(k + ': ' + (props[k] ? '✓ set (' + props[k].substring(0,20) + '...)' : '✗ MISSING'));
  });
}
```

---

## Client activation

When a client pays and is ready to activate their trial:

```javascript
function _activateClient() {
  // Find the registryId in the Admin Panel or registry sheet (format: REG_xxxxx)
  activateClient('REG_xxxxx');
}
```

Status changes to Active on the client's next page load. Takes under 60 seconds total.

---

## HMRC sandbox test (optional)

After connecting HMRC credentials in Settings, run a smoke test:

```javascript
function _hmrcSandboxTest() {
  var result = runSandboxValidation();
  Logger.log(JSON.stringify(result, null, 2));
}
```
