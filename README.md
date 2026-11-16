# Gamecock Mafia — Fantasy Football League Hub & Draft Tracker (`v11.0.0`)

Live mobile-first Progressive Web App (PWA) for the **Gamecock Mafia** 10-team keeper league (`2016–Present`), hosted on GitHub Pages with real-time 2-way Google Sheets Cloud Sync and ESPN Fantasy League integration.

- **Live App**: https://syeth-spurrier.github.io/gamecock-mafia-draft-tracker/

---

## 1. GitHub Repository & System Architecture Diagram

```mermaid
flowchart TB
  subgraph GH["GitHub Repository (syeth-spurrier/gamecock-mafia-draft-tracker)"]
    direction TB
    IDX["index.html & gamecock_mafia_draft_pick_tracker_v1.html<br/>(Single-File Mobile-First 100dvh PWA — v11.0.0)"]
    SEED["league_data.json<br/>(64 Canonical 2018–2027 Trades & 2016–Present Seed)"]
    GAS_FILE["google_apps_script_sync.gs<br/>(2-Way Cloud Sync + Weekly ESPN & FAAB Auto-Sync)"]
    PWA["manifest.json + sw.js + .nojekyll<br/>(Installable iOS/Android Home Screen App)"]
    DOCS["README.md + SETUP_GUIDE.md<br/>(Architecture & Admin Documentation)"]
  end

  subgraph UI["5-Tab Modern Navigation Shell"]
    direction LR
    T1["1. My Team War Room<br/>• 72h Pre-Draft Keeper Lock Hub<br/>• Draft Picks, Trade Offers & Keepers<br/>• Central Banker Dues Ledger ($150)"]
    T2["2. Trade Center<br/>• 2/3-Team Trade Machine & Fairness<br/>• Active-Era History & Waiting Queue<br/>• 1-Tap Jump to 2018–2027 Trade Vault"]
    T3["3. Draft Board<br/>• 14-Round Post-Keeper Live Draft<br/>• Round Tier Jump Bar & Snake Grid<br/>• Synced ESPN Drafted Player Overlay"]
    T4["4. Standings<br/>• ESPN Standings & Playoff/Bye Odds<br/>• 3-Pillar Power Rankings<br/>• Positional Top 10 & Free Agents"]
    T5["5. The Vault<br/>• 4-Stage VORP Player HOF & 3 Proof Chips<br/>• Automated FAAB & 6 Custom Badges<br/>• Trade Vault (64 Deals) & Time Capsules"]
  end

  subgraph CLOUD["Real-Time Multi-Device Sync (Google Workspace)"]
    direction TB
    WEBAPP["Google Apps Script Web App<br/>(LockService Atomic Write-Through + SHA-256 Auth)"]
    SHEET[("Google Sheet Database<br/>• LIVE_STATE_JSON (30-Snapshot Time Machine)<br/>• TRADES_BACKUP<br/>• PICK_LEDGER_BACKUP<br/>• ESPN_ROSTERS_BACKUP<br/>• ESPN_HISTORY_BACKUP")]
  end

  subgraph ESPN["Private ESPN Fantasy League (#30215)"]
    direction TB
    API["ESPN Fantasy v3 & leagueHistory API<br/>(2016–Present Standings, Rosters, mTransactions2 FAAB, mBoxscore)"]
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
  participant AppB as Notification Bell & Activity (Device B)
  actor TeamB as Trade Partner (e.g., Dylan)

  TeamA->>AppA: Taps Partner Avatar -> Stages Players & Picks + "Send Trade Offer 🚀"
  AppA->>GAS: POST PROPOSE_MEMBER_TRADE (confirmedBy: [TeamA])
  GAS->>Sheet: Appends to pendingTrades in LIVE_STATE_JSON
  AppB->>GAS: 45s Auto-Poll / Tab Focus Pull
  GAS-->>AppB: Returns updated pendingTrades (Pulsing Notification Bell 🔔)
  TeamB->>AppB: Reviews Trade Value Guide & Taps "Accept"
  AppB->>GAS: POST RESPOND_TRADE_PROPOSAL (ACCEPT)
  alt Unconditional Trade
    GAS->>Sheet: Validates Ownership -> Moves Picks & Players -> Marks OFFICIAL
  else Has Conditional / Pool Terms
    GAS->>Sheet: Promotes to Activity with "Waiting on Commish" Badge
  end
```

---

## 3. Repository File Map

| File | Role |
| :--- | :--- |
| `index.html` | Primary production entry point served by GitHub Pages (`v11.0.0`). |
| `gamecock_mafia_draft_pick_tracker_v1.html` | Master single-file application source (identical to `index.html`). |
| `google_apps_script_sync.gs` | Google Apps Script backend powering 2-way atomic sync, HOF 4-stage funnel (`maxClassSize: 3`), `mTransactions2` FAAB tracking, and weekly Tuesday ESPN auto-sync (`autoWeeklyEspnSync`). |
| `league_data.json` | Sanitized baseline seed containing all 64 canonical `2018–2027` Draft Sheet trades, locked `2018–2026` draft orders, manager alias mapping, and `2026–2028` pick ledgers (`v11.0.0`). |
| `manifest.json` & `sw.js` | Progressive Web App (PWA) manifest and Network-First Service Worker (`gcm-tracker-v11.0.0`). |
| `SETUP_GUIDE.md` | Commissioner setup, ESPN sync, and security reference. |
