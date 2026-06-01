# Deploy the Reconciler to Streamlit Community Cloud

This guide takes a clean copy of the `invoice-reconciler/` folder and gets it
running at a public URL on Streamlit Community Cloud (free) in ~15 minutes.

## What the deployed app does

- Users open the URL.
- Each user uploads their copy of the four source files:
  - OTS QBO Revenue Audit Report (`audit_rev_ots.xlsx`)
  - OTS Bullhorn billable charges (`billable_charges.xlsx`)
  - CTS QBO Revenue Audit Report (`audit_rev_cts.xlsx`)
  - CTS Bullhorn billable charges (`billable_charges_CTS.xlsx`)
- The app reconciles them in the user's browser session and renders the dashboard.
- Nothing persists on the server — files vanish when the session ends.

This design means the app is safe to publish publicly (anyone with the URL can
load the page, but they can only see data they themselves uploaded).

---

## One-time setup

### 1. Create a GitHub account and repo

1. Sign up at <https://github.com> (any plan, free is fine).
2. Create a new repository: name it `oaktree-invoice-reconciler` or similar.
   - **Public** is fine (the code contains no sensitive data — see `.gitignore`).
   - If you'd rather it be Private, that works too. Streamlit Cloud can deploy from either.
3. On the repo's landing page, click **Code → "Set up in Desktop"** if you use GitHub Desktop, or grab the HTTPS clone URL for command-line git.

### 2. Push the code to the repo

From a PowerShell prompt:

```powershell
cd "C:\Users\pwil\OakTree Software, Inc d b a OakTree Staffing\Management - PBW Agents\OakTree Operations\jarvis-phil\invoice-reconciler"
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-username>/oaktree-invoice-reconciler.git
git push -u origin main
```

> Important: confirm `git status` shows **no .xlsx files** before pushing.
> The `.gitignore` is configured to keep data out, but verify before pushing.

### 3. Sign up for Streamlit Community Cloud

1. Go to <https://share.streamlit.io>.
2. Click **Sign up** and authenticate with GitHub.
3. Authorize Streamlit to read your repos.

### 4. Deploy the app

1. From the Streamlit Cloud dashboard, click **New app**.
2. Pick the repo you just created.
3. Branch: `main`.
4. Main file path: `app.py`.
5. App URL — Streamlit will suggest a slug. Pick something memorable like
   `oaktree-reconciler` so the URL is e.g. `https://oaktree-reconciler.streamlit.app`.
6. Click **Deploy**. First deploy takes ~3 minutes (installing dependencies).

When it's done, you'll see your live URL at the top of the app page.

### 5. Share the URL

Send the URL to the people who need access. They'll need:
- The four .xlsx files on their machine (OneDrive sync gives them this).
- A modern browser.

**First-time user instructions** (paste this in the email/Slack):

> Open the URL. You'll see two tabs at the top — OTS and CTS. Within each tab,
> drag-and-drop the two source files for that entity:
>
> - **OTS tab:** `audit_rev_ots.xlsx` (from `_shared/quickbooks-data/`) and
>   `billable_charges.xlsx` (from `_shared/bullhorn-data/`).
> - **CTS tab:** `audit_rev_cts.xlsx` and `billable_charges_CTS.xlsx`.
>
> Click **Run reconciliation** and the dashboard populates. You can filter
> by month, week (Mon–Sun), client, contractor — and drill into any row to
> see the underlying Bullhorn and QBO lines.

---

## Updating the deployed app

### Code changes
Push to the `main` branch and Streamlit Cloud auto-redeploys (~30 seconds).

```powershell
git add .
git commit -m "Describe the change"
git push
```

### Data refresh
The source xlsx files are bundled in `data/`. To push fresh data to the
cloud after Brenda updates the QBO and Bullhorn exports:

```powershell
.\refresh_data.ps1
```

This script copies the four files from `_shared/` (your OneDrive-synced
folder) into `data/`, commits, and pushes. It skips the commit if nothing
actually changed.

If you want fully hands-free refresh, schedule it via Windows Task
Scheduler — point it at this script and pick a cadence (daily, hourly,
etc.). The cloud app stays current as long as your machine is on and
OneDrive is syncing.

### Restarting / debugging
The deployed app shows a "Manage app" link in the bottom-right — useful for
viewing logs or restarting if something's stuck.

---

## Costs and limits

Streamlit Community Cloud is **free** for public apps. As of 2026:

- 1 GB RAM per app.
- App sleeps after 7 days of no traffic — wakes up on next visit (~30 sec).
- No viewer cap for public apps.

If you ever outgrow this (heavy use, want private app with many viewers, or
need persistent storage), upgrade paths:

- **Streamlit Teams** (~$250/mo) — private apps, SSO, more resources.
- **Azure App Service** — host on OakTree's M365 tenant with SSO.
- **Self-host on an internal Windows machine** — Dockerfile in this repo (TBD).

---

## Troubleshooting

**"Error parsing QBO file"** — the file's columns aren't matching. Check it
opens cleanly in Excel and that header row 5 has "Transaction date",
"Description", and "Amount" somewhere. The parser is header-driven so column
position doesn't matter, but the names need to be in `COLUMN_ALIASES` in
[parse_qbo.py](parse_qbo.py).

**"actionable exceptions: thousands"** — usually means the QBO file covers
only part of the date range the Bullhorn file covers. Pick a single month
from the filter and the noise drops out.

**App is slow on first load** — Streamlit Cloud cold-starts can take 20-30
seconds after sleep. Subsequent loads are instant.

**App sleeps too aggressively** — Phil can bump the app from the Streamlit
Cloud dashboard. Or set up an uptime monitor (e.g. UptimeRobot) to hit the
URL every few hours.

---

## What lives in the repo vs. on the user's machine

In the repo (visible at GitHub):
- All Python source (`app.py`, `reconcile.py`, parsers, etc.)
- `requirements.txt`, `.gitignore`
- `BUILD_PLAN.md`, `PARITY_REPORT.md`, this file
- The legacy VBA source for reference (`_inputs/Revenue_Audit_VBA_source.bas`)

Never in the repo (stays on user machine / OneDrive):
- Any .xlsx file (source data or snapshots)
- The legacy workbook itself
- The `output/` directory
- Any per-user secrets

---

## Locking down later

If you decide the public URL is too open and want to limit access:

1. In the Streamlit Cloud dashboard, open the app's **Settings**.
2. Set **Sharing** to **Private (invited viewers only)**.
3. Add the email addresses (Google/Microsoft account) of people who should see it.

Note: Streamlit Community Cloud's private tier has a viewer cap (~3 as of 2026).
For larger teams you'd pay for Streamlit Teams or move to a different host.
