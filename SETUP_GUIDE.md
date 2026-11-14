# Gamecock Mafia Draft & Pick Tracker (v5.0 Beta) — Setup & Security Guide

This guide documents the **Gamecock Mafia Draft & Pick Tracker (v5.0 Beta)** hosted on **GitHub Pages** and backed by a **2-Way Live Sync Google Sheet** in Google Drive.

---

## 1. Architecture & Files

- `index.html` / `gamecock_mafia_draft_pick_tracker_v1.html` — Production v5.0 Beta Progressive Web App.
- `league_data.json` — Sanitized public seed database containing double-hashed credential verifiers (`leaguePassphraseVerifier`, `commishPinVerifier`) and the passphrase-encrypted cloud sync endpoint (`encryptedCloudSyncUrl`).
- `manifest.json` & `icon.svg` — iPhone/Android "Add to Home Screen" PWA manifest and garnet/gold crest icon.
- `google_apps_script_sync.gs` — Google Apps Script backend running inside your private Google Sheet (`LIVE_STATE_JSON`, `TRADES_BACKUP`, `PICK_LEDGER_BACKUP` tabs).

---

## 2. Security & Zero-Leak Protections

1. **Encrypted Cloud Sync Endpoint (`encryptedCloudSyncUrl`)**: Your Google Apps Script `/exec` URL is encrypted at rest using a SHA-256 keystream derived from your league passphrase. It never appears in plaintext in `index.html` or `league_data.json`, and decrypts strictly in memory after a manager unlocks the league passphrase.
2. **Double-Hashed Credential Verifiers**: Public files store only double-hashed SHA-256 verifiers (`SHA256(verifier_salt + SHA256(salt + secret))`), preventing Pass-the-Hash replay from public repository files.
3. **Private Google Drive Storage**: Your Google Sheet (`Gamecock Mafia Live Draft Tracker DB`) stays 100% private to your Google account; league members never receive direct access to your Drive folder.
4. **Sanitized Public Exports**: Clicking **Export league_data.json** or **Copy Snapshot Link** automatically strips raw URLs, webhook URLs, and single-hashed wire tokens.
5. **Input & Spreadsheet Hardening**: All user strings pass through `escapeHtml()`, `safeId()`, `stripDangerousKeys()`, and spreadsheet formula-injection neutralization (`'=`, `'+`, `'-`, `'@`) across both `.csv` downloads and Google Sheet backup tabs.

---

## 3. Commissioner Quick Reference

- **Rotate Your Commissioner PIN**: Unlock **Commissioner Mode** → click **Live Sheet / GitHub Sync** in the footer → enter a new private PIN under **4. Commissioner PIN** and click **Update PIN**. Your private Google Sheet immediately enforces the new PIN across all devices.
- **Finalize Conditional Trades**: Standard managers can log settled trades directly (effective immediately). Any trade with multi-pick pools or conditional round swaps is forced to **`Pending Final`** until you resolve it in **Commissioner Mode → Conditional Trades**.
- **Roll Season Window**: After Draft Night, unlock **Commissioner Mode** and click **Roll Window →** to advance the active 2-year trading window (e.g., `2027 & 2028` → `2028 & 2029`).
