# Gamecock Mafia — Fantasy Football League Hub & Draft Tracker (`v7.3`)

Live mobile-first Progressive Web App (PWA) for the **Gamecock Mafia** 10-team keeper league (`2016–Present`), hosted on GitHub Pages with real-time 2-way Google Sheets Cloud Sync and ESPN Fantasy League integration.

- **Live App**: https://syeth-spurrier.github.io/gamecock-mafia-draft-tracker/

---

## 1. GitHub Repository & System Architecture Diagram

```mermaid
flowchart TB
  subgraph GH["GitHub Repository (syeth-spurrier/gamecock-mafia-draft-tracker)"]
    direction TB
    IDX["index.html & gamecock_mafia_draft_pick_tracker_v1.html<br/>(Single-File Mobile-First 100dvh PWA)"]
    SEED["league_data.json<br/>(Sanitized Initial Seed & 2016–Present Archive)"]
    GAS_FILE["google_apps_script_sync.gs<br/>(2-Way Cloud Sync + Weekly ESPN Auto-Sync)"]
    PWA["manifest.json + icon.svg + .nojekyll<br/>(Installable iOS/Android Home Screen App)"]
    DOCS["README.md + SETUP_GUIDE.md<br/>(Architecture & Admin Documentation)"]
  end

  subgraph UI["5-Tab Native Mobile & Desktop App Shell"]
    direction LR
    T1["1. My Team<br/>• Draft Picks (Rd 1–20)<br/>• Roster & 12 Keepers<br/>• Trophy Case & Nemesis"]
    T2["2. Trade Center<br/>• 2-Team Trade Machine<br/>• Multi-Year Pick Switcher<br/>• 1-Tap Send Proposal<br/>• Collapsible Trade Block"]
    T3["3. Draft Board<br/>• Clean Post-Keeper Rd 1–20<br/>• Sticky Team Headers<br/>• 'My Picks' 1-Tap Filter"]
    T4["4. League Hub<br/>• Capital Standings<br/>• ESPN Standings<br/>• All-Time Vault & Alumni<br/>• 10x10 Rivalry & Records"]
    T5["5. Activity<br/>• Partner Trade Confirmations<br/>• Pending Final Conditions<br/>• Emoji Reactions & Chat Share"]
  end

  subgraph CLOUD["Real-Time Multi-Device Sync (Google Workspace)"]
    direction TB
    WEBAPP["Google Apps Script Web App<br/>(LockService Atomic Write-Through + SHA-256 Auth)"]
    SHEET[("Google Sheet Database<br/>• LIVE_STATE_JSON<br/>• TRADES_BACKUP<br/>• PICK_LEDGER_BACKUP<br/>• ESPN_ROSTERS_BACKUP<br/>• ESPN_HISTORY_BACKUP")]
  end

  subgraph ESPN["Private ESPN Fantasy League (#30215)"]
    direction TB
    API["ESPN Fantasy v3 & leagueHistory API<br/>(2016–Present Standings, Rosters, H2H, IDP)"]
  end

  IDX --> UI
  SEED -.->|"Initial Fallback Seed"| IDX
  UI <-->|"45s Auto-Poll + Instant Tab-Focus Pull<br/>+ Atomic Action Push"| WEBAPP
  WEBAPP <-->|"Reads & Rebuilds 5 Tabs"| SHEET
  ESPN -->|"1-Click Browser Extractor (Admin)<br/>or Weekly Tuesday Apps Script Timer"| WEBAPP
```

---

## 2. Trade Proposal, Partner Validation & Multi-Device Sync Flow

```mermaid
sequenceDiagram
  autonumber
  actor TeamA as Proposer (e.g., Matt)
  participant AppA as Trade Center (Device A)
  participant GAS as Google Apps Script (LockService)
  participant Sheet as Google Sheet DB
  participant AppB as Activity Feed (Device B)
  actor TeamB as Trade Partner (e.g., Dylan)

  TeamA->>AppA: Taps "Propose Trade" -> Opens 2-Team Trade Machine
  TeamA->>AppA: Taps Players & Multi-Year Picks (2026/2027/2028) + "Send Trade Offer 🚀"
  AppA->>GAS: POST SUBMIT_TRADE_PROPOSAL (confirmedBy: [TeamA])
  GAS->>Sheet: Appends to pendingTrades in LIVE_STATE_JSON
  AppB->>GAS: 45s Auto-Poll / Tab Focus Pull
  GAS-->>AppB: Returns updated pendingTrades (Pulsing Notification Bell 🔔)
  TeamB->>AppB: Reviews KTC Balance Meter & Taps "Confirm Trade ✓"
  AppB->>GAS: POST CONFIRM_TRADE_PROPOSAL (confirmedBy: [TeamA, TeamB])
  alt Unconditional Trade
    GAS->>Sheet: Validates Ownership -> Moves Picks & Players -> Marks OFFICIAL
  else Has Conditional / Pool Terms
    GAS->>Sheet: Promotes to Activity with "PENDING FINAL" Badge (Awaits Commish Resolution)
  end
```

---

## 3. Repository File Map

| File | Role |
| :--- | :--- |
| `index.html` | Primary production entry point served by GitHub Pages (`v7.3`). |
| `gamecock_mafia_draft_pick_tracker_v1.html` | Master single-file application source (identical to `index.html`). |
| `google_apps_script_sync.gs` | Google Apps Script backend powering 2-way atomic sync, partner trade validation, and optional weekly Tuesday ESPN auto-sync (`setupWeeklyEspnSyncTrigger`). |
| `league_data.json` | Sanitized baseline seed containing 2016–Present league history, manager alias mapping (`Alex → Geltz`, `Shane → Boss`, `David → Mary`), Slot #10 Alumni lineage (`Scott → Alek → Josh`), and 2026–2028 pick ledgers. |
| `manifest.json` & `icon.svg` | Progressive Web App (PWA) manifest and vector icon for iOS/Android home-screen installation. |
| `SETUP_GUIDE.md` | Commissioner setup, ESPN sync, and security reference. |
