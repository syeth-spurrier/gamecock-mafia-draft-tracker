# Gamecock Mafia Draft, Keeper & League HQ (v7.0) — Setup & Security Guide

This guide documents the **Gamecock Mafia Draft, Keeper & League HQ App (v7.0)** hosted on **GitHub Pages** and backed by a **2-Way Live Sync Google Sheet** in Google Drive.

---

## 1. Architecture & Files

- `index.html` / `gamecock_mafia_draft_pick_tracker_v1.html` — Production v7.0 Mobile-First Native Dashboard & Year-Round Keeper League HQ Progressive Web App.
- `league_data.json` — Sanitized public seed database containing double-hashed credential verifiers (`leaguePassphraseVerifier`, `commishPinVerifier`), the passphrase-encrypted cloud sync endpoint (`encryptedCloudSyncUrl`), and multi-season ESPN League HQ data (`2018–2026`).
- `manifest.json` & `icon.svg` — iPhone/Android "Add to Home Screen" PWA manifest and garnet/gold crest icon.
- `google_apps_script_sync.gs` — Google Apps Script backend running inside your private Google Sheet (`LIVE_STATE_JSON`, `TRADES_BACKUP`, `PICK_LEDGER_BACKUP`, `ESPN_ROSTERS_BACKUP`, and `ESPN_HISTORY_BACKUP` tabs) with automatic multi-cell chunk overflow (`B2` + `D2:M2`) to support full multi-season ESPN history without hitting Google Sheets' 50,000-character cell limit.

---

## 2. v7.0 Year-Round ESPN League HQ & Keeper Engine

1. **Tab 1 — `My Team` (`📋 Draft Picks` + `🏈 Roster & Keepers`)**:
   - Toggle between your 20-round draft portfolio and your live ESPN franchise roster (Offense + IDP).
   - **12-Keeper Cap & Positional Cap Enforcement**: Enforces **12 Keepers max** (locking Rounds 1–12), **Max 4 RB**, **Max 4 WR**, and **Max 3 of any other single position** (`QB`, `TE`, `DL`, `LB`, `DB`, `K`, `D/ST`) with live meter pills and an explicit **`Confirm Keepers & Roster Tags`** button.
   - **1-Tap Roster Tags**: Tag players as **`🔒 Keep`**, **`🏷️ On Trade Block`**, or **`🛡️ Untouchable`**.
2. **Tab 2 — `Trade Center` (Mixed Player + Pick Trades)**:
   - The Active League Trade Block displays **both OTB Players and OTB Draft Picks** across all 10 teams.
   - The 2-Team Trade Machine and **Propose Trade** modal let managers package **both ESPN Roster Players and Draft Picks** in a single trade. Once the trade partner confirms the trade, traded players automatically move between rosters.
3. **Tab 3 — `Draft Board` (`🔒 Show Keepers (Rds 1–12)`)**:
   - Toggle the Keeper overlay on the 10×20 Draft Board to preview each franchise's locked keepers in Rounds 1–12 alongside traded picks.
4. **Tab 4 — `League HQ` (5 Sub-Views)**:
   - **`📊 Pick Capital`**: 10-Team Draft Pick Power Standings.
   - **`🏈 ESPN Standings`**: Season-by-season ESPN Standings (`2018–2026`) with W-L-T, PF, PA, Playoff Seed, Streak, All-Play W-L, Schedule Luck delta, plus a Commissioner **`Seed Draft Order from ESPN Standings`** button (inverse standings: 10th place → `1.01`, Champion → `1.10`).
   - **`🏆 All-Time Vault`**: Career Leaderboard for the 10 active managers (Championships 🏆, Toilet Bowls 🚽, Career W-L, Win %, Career PF, All-Play Win %) plus a dedicated **Former Managers / League Alumni (`Shane`, `Alex`)** table so past managers' seasons never skew current managers' career records.
   - **`⚔️ Rivalry Matrix`**: Interactive 10×10 Head-to-Head W-L matrix across all seasons.
   - **`📜 Record Book`**: All-Time Single-Week High Scores, Low Scores, Biggest Blowouts, & Closest Heartbreakers.

---

## 3. Syncing Private ESPN League Data (1-Click Browser Console Extractor)

Because your ESPN league is **Private**, external servers cannot query it without your session cookies. Instead of hunting down `espn_s2` or `SWID` cookies or installing Python:

1. Open the tracker and click the **`🏈 ESPN Sync`** button in the top header (or inside **Settings → ESPN Sync & Keeper Rules**).
2. Click **`Copy Extractor Script`**.
3. Log into [fantasy.espn.com](https://fantasy.espn.com) in Chrome, Safari, or Edge and open your Gamecock Mafia league page.
4. Press **`F12`** (or `Cmd+Option+J` on Mac / `Ctrl+Shift+J` on Windows) to open the **Browser Console**, paste the script, and press **Enter**.
5. The script automatically uses your active ESPN browser login to pull all historical seasons (`2018–present`), standings, weekly matchups, and live rosters, and copies the JSON directly to your clipboard.
6. Paste the JSON into the tracker's **ESPN Sync Modal** and click **`Import & Sync ESPN Data to Cloud`**.

---

## 4. Security & Zero-Leak Protections

1. **Encrypted Cloud Sync Endpoint (`encryptedCloudSyncUrl`)**: Your Google Apps Script `/exec` URL is encrypted at rest using a SHA-256 keystream derived from your league passphrase.
2. **Double-Hashed Credential Verifiers**: Public files store only double-hashed SHA-256 verifiers (`SHA256(verifier_salt + SHA256(salt + secret))`).
3. **Private Google Drive Storage**: Your Google Sheet (`Gamecock Mafia Live Draft Tracker DB`) stays 100% private to your Google account.
