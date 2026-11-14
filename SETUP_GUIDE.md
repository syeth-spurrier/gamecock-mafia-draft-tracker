# Gamecock Mafia Draft & Pick Tracker (v5.0 Beta) — Setup & Deployment Guide

This guide walks through deploying the **Gamecock Mafia Draft & Pick Tracker (v5.0 Beta)** on **GitHub Pages** paired with a **Google Sheet + Apps Script Web App** in your Google Drive for **true 2-way live sync** across all 10 managers and automatic spreadsheet backups.

---

## 1. Files in This Folder

- `index.html` — Production v5.0 Beta web app (upload to GitHub Pages).
- `league_data.json` — Initial seed database & shared `cloudSyncUrl` config (upload to GitHub Pages).
- `manifest.json` — Progressive Web App manifest for iPhone/Android "Add to Home Screen" (upload to GitHub Pages).
- `icon.svg` — Garnet & gold Gamecock Mafia crest icon (upload to GitHub Pages).
- `google_apps_script_sync.gs` — Google Apps Script backend to paste into your Google Sheet's `Extensions -> Apps Script`.
- `gamecock_mafia_draft_pick_tracker_v1.html` — Identical standalone copy of `index.html`.

*(Security note: Neither your league passphrase nor your Commissioner PIN is stored in plaintext anywhere in code—only salted SHA-256 hashes are stored.)*

---

## 2. Part A — Create Your Google Sheet 2-Way Live Sync Backend (5 Min)

1. Open Google Drive and create a new Google Sheet named **`Gamecock Mafia Live Draft Tracker DB`**.
2. Click **Extensions → Apps Script**.
3. Replace any starter code in `Code.gs` with the full contents of `google_apps_script_sync.gs` and click **Save**.
4. Click **Deploy → New deployment**:
   - **Select type**: `Web app`
   - **Execute as**: `Me`
   - **Who has access**: `Anyone` *(requests are verified against your salted passphrase & Admin PIN hashes inside the script)*
5. Click **Deploy**, authorize your Google account (`Advanced -> Go to project -> Allow`), and copy the resulting **Web App URL** ending in `/exec`.

---

## 3. Part B — Publish Free on GitHub Pages (5 Min)

1. Create a new **Public** repository on GitHub (e.g., `gamecock-mafia-draft`).
2. Upload **`index.html`**, **`league_data.json`**, **`manifest.json`**, and **`icon.svg`** to the root of the repository and commit.
3. Open **Settings → Pages**, set **Source** to **Deploy from a branch** (`main` / `/ (root)`), and click **Save**.
4. Your live tracker will publish at `https://<your-username>.github.io/<repo-name>/` in ~60 seconds.

---

## 4. Part C — Link Your Sheet Once so All 10 Managers Auto-Connect (2 Min)

1. Open your GitHub Pages URL and unlock with your **League Passphrase**.
2. Click **Commissioner Mode** and unlock with your **Commissioner PIN**.
3. In the bottom footer, click **Live Sheet / GitHub Sync**:
   - Paste your `/exec` Web App URL into **1. Google Sheet 2-Way Live Sync** and click **Save Live Sync URL** (this automatically seeds your Google Sheet and creates `LIVE_STATE_JSON`, `TRADES_BACKUP`, and `PICK_LEDGER_BACKUP` tabs).
   - *(Recommended)* Update your **Commissioner PIN** under section **4. Commissioner PIN**.
   - *(Optional)* Add a Discord/Slack webhook URL under section **2. Optional Group Chat Trade Alert Webhook**.
4. Click **Export league_data.json** in the footer and upload that updated `league_data.json` to your GitHub repo **one time**. From that point forward, all 10 managers connect to your live Google Sheet automatically!

---

## 5. How Live Beta Tracking Works for Your 10 Managers

- **Direct Final Trades**: All 10 managers can log trades directly via **Log Trade** (or simulate first in **Compare 2 Teams → Launch What-If Sandbox**).
- **Settled vs. Conditional Trades**:
  - Settled pick transfers take effect **immediately** across ledgers, comparisons, the 200-pick draft board, and your Google Sheet backup.
  - Any multi-pick pools or conditional round swaps are automatically tagged **`Pending Final`** and can only be finalized by the Commissioner in **Commissioner Mode**.
- **Install on Phone**: On iOS Safari (`Share -> Add to Home Screen`) or Android Chrome (`Menu -> Add to Home Screen`), managers get a full-screen **GCM Picks** app icon.
- **Future Migration**: Your entire league state lives in portable `league_data.json` + `.csv` / Google Sheet backups, so migrating to any custom domain or database later requires zero data rework.
