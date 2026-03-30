# no~bull books 🐂

**UK cloud accounting for sole traders and small businesses — built on Google Apps Script and Google Sheets.**

> Owned and operated by [no~bull consulting](mailto:edward@nobull.consulting).

---

## What it is

no~bull books is a multi-tenant SaaS accounting application. All code runs in a single central Google Apps Script project. Each client gets their own Google Sheet in their own Google Drive — no shared database, no data mixing, no servers to manage.

---

## Repository structure

```
nobull-books/
├── src/
│   ├── main/               ← Main hub (executeAs: USER_DEPLOYING)
│   │   ├── appsscript.json     Manifest — scopes, runtimes
│   │   ├── Code.gs             doGet(), getDb(), include()
│   │   ├── Config.gs           SHEETS, INV_COLS, ROLE_PERMISSIONS, _ss()
│   │   ├── Auth.gs             _getCurrentUserContext(), _auth(), manageUser()
│   │   ├── Api.gs              handleApiCall() — single API router (159 routes)
│   │   ├── Initializer.gs      checkAndInitSheet() — 32-tab schema + seed Owner
│   │   ├── Settings.gs         getSettings(), updateSettings(), bank accounts, …
│   │   ├── Invoices.gs         Full invoice/bill/credit note/PO lifecycle
│   │   ├── Banking.gs          Bank accounts, transactions, reconciliation
│   │   ├── COA.gs              Chart of Accounts CRUD + General Ledger
│   │   ├── COA_Seed.gs         seedUKChartOfAccounts() — 90 UK HMRC-aligned accounts
│   │   ├── VAT.gs              getVATReturns(), saveVATReturn()
│   │   ├── HMRC.gs             MTD OAuth, VAT obligations, submissions, ITSA
│   │   ├── Reports.gs          P&L, Balance Sheet, Trial Balance, Cash Flow, GL
│   │   ├── GeminiService.gs    Gemini 2.5 Flash AI assistant
│   │   ├── Registry.gs         Central client registry (25-column sheet)
│   │   ├── Onboarding.gs       _checkLicence(), provisionNewClient(), activation
│   │   ├── Stubs.gs            Fixed assets, recurring invoices, year-end, backups
│   │   ├── Index.html          App shell — nav, boot(), wizard, all page renderers
│   │   ├── Code2.html          Settings, users, reconciliation, banking modals
│   │   └── Code3.html          Invoice modals, VAT/MTD UI, year-end, fixed assets
│   │
│   └── setup/              ← Setup microservice (executeAs: USER_ACCESSING)
│       ├── appsscript.json     Manifest — drive.file scope
│       ├── SetupService.gs     Creates client sheet in client's Drive, redirects
│       └── Setup.html          Landing page — business name form
│
├── docs/
│   ├── ARCHITECTURE.md         System design and data flow
│   ├── DEPLOYMENT.md           Step-by-step deployment guide
│   └── ONE_TIME_SETUP.md       GAS editor commands to run once after first deploy
│
├── .github/
│   └── workflows/
│       ├── deploy-main.yml     Auto-deploy main hub on push to src/main/
│       └── deploy-setup.yml    Auto-deploy setup service on push to src/setup/
│
└── README.md
```

---

## Architecture summary

```
Client browser
    │
    ▼  ?id=SHEET_ID
Main Hub GAS (executeAs: USER_DEPLOYING)
    │                           │
    ├─ reads/writes ──────────▶ Client's Google Sheet (in client's Drive)
    │
    └─ reads/writes ──────────▶ Registry Sheet (edward's Drive)

Setup Microservice (executeAs: USER_ACCESSING)
    ├─ SpreadsheetApp.create() ▶ New sheet in CLIENT's Drive
    └─ HTTP redirect ──────────▶ Main Hub ?id=NEW_SHEET_ID
```

Key architectural rules — enforced throughout the codebase:

| Rule | Detail |
|------|--------|
| `getDb(params)` | Always use instead of `SpreadsheetApp.openById()` or `_ss()` |
| `params._sheetId` | Injected by `api()` in the frontend; threaded through every GAS call |
| No arrow functions | GAS V8 compat — all `function()` declarations |
| No `safeSerializeDateTime` | Use `safeSerializeDate` only |
| `executeAs: USER_DEPLOYING` | Main hub runs as edward — clients get no script access |
| `executeAs: USER_ACCESSING` | Setup service runs as client — creates sheet in their Drive |

---

## Deployment

See [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md) for full instructions.

**Quick path:**

1. Clone this repo
2. Set up GitHub Secrets: `CLASP_TOKEN`, `MAIN_SCRIPT_ID`, `SETUP_SCRIPT_ID`
3. Push to `main` — workflows deploy automatically
4. Run one-time setup commands from GAS editor (see [docs/ONE_TIME_SETUP.md](docs/ONE_TIME_SETUP.md))

---

## GitHub Secrets required

| Secret | Description |
|--------|-------------|
| `CLASP_TOKEN` | JSON content of `~/.clasprc.json` after `clasp login` |
| `MAIN_SCRIPT_ID` | Script ID of the main hub GAS project |
| `SETUP_SCRIPT_ID` | Script ID of the setup microservice GAS project |

---

## Licence

Proprietary — no~bull consulting. All rights reserved.  
Contact: edward@nobull.consulting
