# no~bull Admin Console

Separate GAS deployment for registry management, client administration,
maintenance and monitoring. Completely isolated from the client-facing Hub.

## Files

| File | Purpose |
|------|---------|
| `AdminCode.gs` | doGet/doPost entry points, router, dashboard UI |
| `AdminUtils.gs` | Shared utilities (date helpers, stubs) |
| `Registry.gs` | Client registry — all registry operations |
| `Onboarding.gs` | Client provisioning and setup |
| `Demo.gs` | Demo instance management and nightly reset |
| `appsscript.json` | GAS project config |

## Setup

### 1. Create GAS Project
1. Go to script.google.com → New project
2. Name it "no~bull Admin Console"
3. Copy all `.gs` files into the project
4. Update `appsscript.json`

### 2. Set Script Properties
Run `_setupAdminSecret()` from the editor — note the generated secret.

Then set:
- `REGISTRY_SHEET_ID` — the registry spreadsheet ID
- `ADMIN_SECRET` — the secret generated above (copy to Hub too)

### 3. Deploy
- Deploy as Web App
- Execute as: Me (edward@nobull.consulting)  
- Access: Only myself

### 4. Configure Hub
In the Hub GAS project Script Properties, set:
- `ADMIN_CONSOLE_URL` — the Admin Console /exec URL
- `ADMIN_SECRET` — same secret as above

### 5. Test
Run `_checkAdminProps()` to verify configuration.
Open the deployed URL to see the admin dashboard.

## Security
- Admin Console access restricted to owner Google account
- All API calls authenticated with shared ADMIN_SECRET
- Deployed as USER_DEPLOYING (owner's identity) with MYSELF access
- Hub falls back to local registry if Admin Console unreachable
