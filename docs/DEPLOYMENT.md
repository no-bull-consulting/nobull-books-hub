# Deployment Guide

## Prerequisites

- [Node.js](https://nodejs.org/) 18+
- [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`
- Google account with Apps Script API enabled at [script.google.com/home/usersettings](https://script.google.com/home/usersettings)
- GitHub repository with Actions enabled

---

## Step 1 — Create the two GAS projects

### Main hub

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Name it `no~bull books — Main Hub`
3. Copy the **Script ID** from Project Settings (you'll need it as `MAIN_SCRIPT_ID`)

### Setup microservice

1. Create a second new project
2. Name it `no~bull books — Setup Service`
3. Copy the **Script ID** (you'll need it as `SETUP_SCRIPT_ID`)

---

## Step 2 — Authenticate clasp locally

```bash
clasp login
# Follow the browser auth flow
# This writes ~/.clasprc.json
cat ~/.clasprc.json   # copy the full JSON content
```

---

## Step 3 — Set GitHub Secrets

In your GitHub repo → Settings → Secrets and variables → Actions → New repository secret:

| Secret name | Value |
|-------------|-------|
| `CLASP_TOKEN` | Full JSON content of `~/.clasprc.json` |
| `MAIN_SCRIPT_ID` | Script ID from Step 1 (main hub) |
| `SETUP_SCRIPT_ID` | Script ID from Step 1 (setup service) |

---

## Step 4 — Initial push

The GitHub Actions workflows deploy automatically on push. For the very first deploy:

```bash
git add .
git commit -m "Initial commit — no~bull books"
git push origin main
```

Both workflows will run. Check the Actions tab to confirm they succeed.

---

## Step 5 — Deploy in GAS (first time only)

After clasp push, you need to create a web app deployment in the GAS editor:

**Main hub:**
1. Open the main hub project in the GAS editor
2. Deploy → New deployment
3. Type: **Web app**
4. Execute as: **Me (edward@nobull.consulting)**
5. Who has access: **Anyone**
6. Copy the `/exec` URL — this is your `MAIN_APP_URL`

**Setup microservice:**
1. Open the setup service project
2. Deploy → New deployment
3. Type: **Web app**
4. Execute as: **User accessing the web app**
5. Who has access: **Anyone**
6. Copy the `/exec` URL — this is your `SETUP_SERVICE_URL`

> **Important:** After the first GAS deployment, subsequent deploys via `clasp deploy` (run by the GitHub Action) update the existing deployment automatically. The `/exec` URLs never change.

---

## Step 6 — One-time setup commands

Run these from the GAS editor (main hub) **once** after first deployment. See [ONE_TIME_SETUP.md](ONE_TIME_SETUP.md) for details.

```javascript
initRegistry()
setSetupServiceUrl('https://script.google.com/macros/s/YOUR_SETUP_ID/exec')
setTrialDays(14)
setGeminiKey('AIza...')
// Optional: set a test sheet for editor debugging
// setDefaultSpreadsheetId('YOUR_SHEET_ID')
```

---

## Subsequent deployments

Push to `main` and both workflows run automatically. Changes to `src/main/**` trigger the main hub workflow; changes to `src/setup/**` trigger the setup service workflow.

Every deploy creates a new versioned deployment. All clients get the new version on their next page load — no per-client action required.

---

## Rollback

In the GAS editor: Deploy → Manage deployments → Edit → select a previous version → Deploy.

Takes effect immediately for all new page loads.
