# Architecture

## Overview

no~bull books is a **central hub** multi-tenant accounting SaaS. One Google Apps Script project hosts all application code. Each client has an isolated Google Sheet in their own Google Drive.

---

## Two-deployment model

### Main Hub (`src/main/`)
- `executeAs: USER_DEPLOYING` — runs as edward@nobull.consulting for all clients
- Handles all accounting logic, HMRC MTD, Gemini AI, registry management
- Receives client's Sheet ID via URL `?id=SHEET_ID` query parameter
- `getDb(params)` opens the correct sheet on every call using `params._sheetId`

### Setup Microservice (`src/setup/`)
- `executeAs: USER_ACCESSING` — runs as the **client's** Google account
- Separate GAS project with `drive.file` scope
- `SpreadsheetApp.create()` creates a new sheet in the **client's** Drive
- Pings registry then redirects to main hub with `?id=NEW_SHEET_ID`

---

## Data flow

```
1. Client visits main hub (?no params)
      ↓
2. "Get started free" → Setup Microservice URL
      ↓
3. Client enters business name → Google OAuth
      ↓
4. SetupService.gs: SpreadsheetApp.create() in client's Drive
      ↓
5. pingRegistry() auto-registers new client
      ↓
6. HTTP redirect → MAIN_APP_URL?id=SHEET_ID
      ↓
7. boot() → api('getStartupData') → _checkLicence()
      ↓
8. Blank sheet → showLoader() → api('runInitialSetup')
      ↓
9. checkAndInitSheet(): 32 tabs + seed Owner user
      ↓
10. showInitWizard(): Welcome → Company → Invoicing → Financial Year → Done
      ↓
11. _launchApp(): nav restored → nbNav('dashboard')
```

---

## Sheet architecture (per client)

`checkAndInitSheet()` in `Initializer.gs` creates 32 sheets on first load:

| Category | Sheets |
|----------|--------|
| Core | Settings, Users, AuditLog |
| Sales | Invoices, InvoiceLines, Clients, CreditNotes, CreditNoteLines, RecurringInvoices |
| Purchases | Bills, BillLines, Suppliers, PurchaseOrders, POLines |
| Banking | BankAccounts, Transactions, StatementLines, Reconciliation |
| Accounting | ChartOfAccounts, JournalEntries, FixedAssets, DepreciationRuns |
| Tax | VATReturns, ITSASubmissions, ITSAObligations |
| Admin | BackupLog, VoidLog, BadDebts, FinancialYears |

---

## API layer

All frontend calls go through a single `handleApiCall(action, paramsJson)` function in `Api.gs`. The router dispatches to the appropriate module function.

**Frontend (`api()` in Index.html):**
```javascript
async function api(action, params) {
  params._sheetId = CLIENT_CONTEXT.sheetId;  // always injected
  return new Promise(function(resolve) {
    google.script.run
      .withSuccessHandler(resolve)
      .withFailureHandler(function(err) { resolve({ success: false, error: err.message }); })
      .handleApiCall(action, JSON.stringify(params));
  });
}
```

**Backend (`handleApiCall` in Api.gs):**
```javascript
function handleApiCall(action, paramsJson) {
  var params = JSON.parse(paramsJson);
  var ctx = _getCurrentUserContext(params);
  _auth(action, params);      // throws if not permitted
  return _route(action, params, ctx);
}
```

---

## Licence enforcement

`_checkLicence(sheetId)` in `Onboarding.gs` is called on every `getStartupData`.

| Registry status | UI behaviour |
|----------------|--------------|
| Active | Normal |
| Trial > 3 days | Normal |
| Trial ≤ 3 days | Yellow dismissible banner |
| Trial expired | Full-screen blocking page |
| Suspended | Full-screen blocking page |
| Registry error | Returns Active (fail-open) |

---

## Script Properties (main hub)

| Key | Set by | Purpose |
|-----|--------|---------|
| `GEMINI_API_KEY` | `setGeminiKey()` | Gemini 2.5 Flash |
| `REGISTRY_SHEET_ID` | `initRegistry()` | Central client registry |
| `SETUP_SERVICE_URL` | `setSetupServiceUrl()` | Setup microservice /exec URL |
| `TRIAL_DAYS` | `setTrialDays(14)` | Trial duration |
| `DEFAULT_SPREADSHEET_ID` | Manual | Editor testing fallback only |

---

## Security model

- **Data isolation** — each client's data is in their own spreadsheet; no path between clients
- **Credentials** — Gemini key, HMRC secrets stored in Script Properties only
- **Scope minimisation** — setup service: `drive.file` + `spreadsheets`; main hub: `spreadsheets`, `drive`, `external_request`, `gmail.send`
- **Role-based access** — 5 roles (Owner, Admin, Accountant, Staff, ReadOnly); checked on every API call via `_auth()`
- **Period locking** — `_checkPeriodLock()` rejects transactions before the lock date
