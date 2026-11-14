/**
 * ============================================================================
 * GAMECOCK MAFIA DRAFT & PICK TRACKER — GOOGLE SHEETS 2-WAY LIVE SYNC BACKEND
 * ============================================================================
 * Attach this script to a Google Sheet in your Google Drive:
 *   1. Open a new Google Sheet (e.g., "Gamecock Mafia Live Draft Tracker DB")
 *   2. Click Extensions -> Apps Script
 *   3. Paste this entire file into Code.gs and click Save
 *   4. Click Deploy -> New deployment -> Select type: "Web app"
 *      - Execute as: "Me"
 *      - Who has access: "Anyone" (Access is protected by your salted league
 *        passphrase hash & Commissioner PIN hash inside this script!)
 *   5. Copy the Web App URL (ends in /exec) and paste it into your Tracker's
 *      Admin -> Live Sheet / GitHub Sync modal!
 *
 * Automatically maintains 3 tabs inside your Google Sheet:
 *   - LIVE_STATE_JSON     (Canonical JSON state + 30-version rollback history)
 *   - TRADES_BACKUP       (Human-readable spreadsheet table of every trade)
 *   - PICK_LEDGER_BACKUP  (Human-readable spreadsheet backup of traded picks & Pending Final items)
 * ============================================================================
 */

var DEFAULT_PASSPHRASE_HASH = "ed82a630acbc4c03fe026e616f3002f5d7ffd51a462cc9023efca21b90a1bd4e"; // Salted SHA-256 hash of default league passphrase
var DEFAULT_COMMISH_PIN_HASH = "9dc15277bc17bbea84f404584d0deff365108b37ccff158a6b2311ddcc5a6b8a"; // Salted SHA-256 hash of default Commissioner PIN

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureSheetsExist(ss) {
  var stateSheet = ss.getSheetByName("LIVE_STATE_JSON");
  if (!stateSheet) {
    stateSheet = ss.insertSheet("LIVE_STATE_JSON");
    stateSheet.getRange("A1:C1").setValues([["KEY", "VALUE", "UPDATED_ISO"]]).setFontWeight("bold");
    stateSheet.getRange("A2").setValue("CURRENT_STATE");
    stateSheet.getRange("A4:D4").setValues([["HISTORY_TIMESTAMP_MS", "UPDATED_ISO", "ACTOR_ACTION", "SNAPSHOT_JSON"]]).setFontWeight("bold");
    stateSheet.setColumnWidth(2, 500);
    stateSheet.setColumnWidth(4, 500);
  }

  var tradesSheet = ss.getSheetByName("TRADES_BACKUP");
  if (!tradesSheet) {
    tradesSheet = ss.insertSheet("TRADES_BACKUP");
  }

  var picksSheet = ss.getSheetByName("PICK_LEDGER_BACKUP");
  if (!picksSheet) {
    picksSheet = ss.insertSheet("PICK_LEDGER_BACKUP");
  }

  return {
    stateSheet: stateSheet,
    tradesSheet: tradesSheet,
    picksSheet: picksSheet
  };
}

function readStoredState(stateSheet) {
  var raw = stateSheet.getRange("B2").getValue();
  if (!raw || typeof raw !== "string") return null;
  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function getExpectedPassHash(storedState) {
  if (storedState && storedState.leaguePassphraseHash && /^[a-f0-9]{64}$/i.test(storedState.leaguePassphraseHash)) {
    return storedState.leaguePassphraseHash.toLowerCase();
  }
  return DEFAULT_PASSPHRASE_HASH;
}

function getExpectedCommishHash(storedState) {
  if (storedState && storedState.commishPinHash && /^[a-f0-9]{64}$/i.test(storedState.commishPinHash)) {
    return storedState.commishPinHash.toLowerCase();
  }
  return DEFAULT_COMMISH_PIN_HASH;
}

/**
 * GET endpoint: Returns canonical live JSON state to authorized league members.
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ensureSheetsExist(ss);
  var stored = readStoredState(sheets.stateSheet);

  var providedPass = (e && e.parameter && e.parameter.passHash) ? String(e.parameter.passHash).toLowerCase() : "";
  var expectedPass = getExpectedPassHash(stored);
  if (providedPass !== expectedPass) {
    return jsonResponse({ ok: false, error: "Unauthorized league passphrase." });
  }

  if (!stored) {
    return jsonResponse({ ok: true, empty: true, state: null });
  }
  return jsonResponse({ ok: true, empty: false, state: stored });
}

/**
 * POST endpoint: Uses text/plain JSON body so browsers send a simple CORS request
 * without failing OPTIONS preflight. Protected by LockService for concurrent trades.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (lockErr) {
    return jsonResponse({ ok: false, error: "Server busy saving another trade. Please try again in a moment." });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: "Missing request payload." });
    }

    var body = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ensureSheetsExist(ss);
    var stored = readStoredState(sheets.stateSheet);

    var providedPass = body && body.passHash ? String(body.passHash).toLowerCase() : "";
    var expectedPass = getExpectedPassHash(stored);
    if (providedPass !== expectedPass) {
      return jsonResponse({ ok: false, error: "Unauthorized league passphrase." });
    }

    var action = String(body.action || "");

    // ------------------------------------------------------------------------
    // ACTION 1: LEAGUE MEMBER LOGS A NEW OFFICIAL TRADE
    // Settled pick transfers are final immediately; any pools/conditionals are
    // strictly forced to Pending Final (only Commissioner can finalize them).
    // ------------------------------------------------------------------------
    if (action === "ADD_MEMBER_TRADE") {
      var baseState = stored || body.fallbackBaseState;
      if (!baseState || !Array.isArray(baseState.teams) || !Array.isArray(baseState.trades)) {
        return jsonResponse({ ok: false, error: "Initialize database from Commissioner Mode first." });
      }

      var tr = body.trade;
      if (!tr || !tr.headline || !tr.teamA || !tr.teamB) {
        return jsonResponse({ ok: false, error: "Invalid trade submission." });
      }

      var primaryYear = Number(tr.year) || Number(baseState.baseUpcomingYear) || 2027;
      var prefix = "T-" + primaryYear + "-";
      var maxNum = 0;
      for (var i = 0; i < baseState.trades.length; i++) {
        var existingId = String(baseState.trades[i].id || "");
        if (existingId.indexOf(prefix) === 0) {
          var n = parseInt(existingId.replace(prefix, ""), 10);
          if (!isNaN(n) && n > maxNum) maxNum = n;
        }
      }
      var assignedId = prefix + (maxNum + 1 < 10 ? "0" + (maxNum + 1) : String(maxNum + 1));

      // Enforce Pending Final on all pools & conditionals submitted by non-admins
      var cleanPools = Array.isArray(tr.pools) ? tr.pools.map(function(p) {
        p.resolvedOriginalOwner = null;
        return p;
      }) : [];

      var cleanConds = Array.isArray(tr.conditionals) ? tr.conditionals.map(function(c) {
        c.isResolved = false;
        c.resolvedRound = null;
        c.resolutionNote = "";
        return c;
      }) : [];

      var cleanTrade = {
        id: assignedId,
        year: primaryYear,
        date: String(tr.date || "").slice(0, 80),
        teamA: String(tr.teamA || ""),
        teamB: String(tr.teamB || ""),
        teamC: tr.teamC ? String(tr.teamC) : null,
        headline: String(tr.headline || "").slice(0, 400),
        notes: String(tr.notes || "").slice(0, 800),
        submittedBy: tr.submittedBy ? String(tr.submittedBy).slice(0, 60) : "",
        directTransfers: Array.isArray(tr.directTransfers) ? tr.directTransfers.map(function(dt) {
          return {
            year: Number(dt.year) || primaryYear,
            round: Math.min(20, Math.max(1, Number(dt.round) || 1)),
            originalOwner: String(dt.originalOwner || ""),
            fromOwner: String(dt.fromOwner || ""),
            toOwner: String(dt.toOwner || "")
          };
        }) : [],
        pools: cleanPools,
        conditionals: cleanConds
      };

      baseState.trades.push(cleanTrade);
      baseState.updatedAt = Date.now();

      saveAndMirrorToSheets(sheets, baseState, "MEMBER_TRADE:" + assignedId);
      return jsonResponse({ ok: true, assignedId: assignedId, state: baseState });
    }

    // ------------------------------------------------------------------------
    // ACTION 2: COMMISSIONER FULL STATE SYNC (ORDER LOCKS, RESOLUTIONS, EDITS)
    // ------------------------------------------------------------------------
    if (action === "ADMIN_FULL_SYNC") {
      var expectedCommish = getExpectedCommishHash(stored);
      var providedCommish = body.commishAuthHash ? String(body.commishAuthHash).toLowerCase() : "";

      if (providedCommish !== expectedCommish) {
        return jsonResponse({ ok: false, error: "Invalid Commissioner PIN authorization." });
      }

      var newState = body.state;
      if (!newState || !Array.isArray(newState.teams) || !Array.isArray(newState.trades)) {
        return jsonResponse({ ok: false, error: "Invalid state payload." });
      }

      newState.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, newState, body.reason || "ADMIN_SYNC");
      return jsonResponse({ ok: true, state: newState });
    }

    return jsonResponse({ ok: false, error: "Unknown sync action." });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Writes canonical JSON + rolling 30-snapshot history + human-readable Google Sheet tabs!
 */
function saveAndMirrorToSheets(sheets, state, actionLabel) {
  var nowIso = new Date().toISOString();
  var jsonStr = JSON.stringify(state);

  // 1. Update LIVE_STATE_JSON
  sheets.stateSheet.getRange("B2:C2").setValues([[jsonStr, nowIso]]);
  sheets.stateSheet.insertRowBefore(5);
  sheets.stateSheet.getRange("A5:D5").setValues([[state.updatedAt || Date.now(), nowIso, actionLabel, jsonStr]]);
  if (sheets.stateSheet.getLastRow() > 35) {
    sheets.stateSheet.deleteRows(36, sheets.stateSheet.getLastRow() - 35);
  }

  // Helper map of team IDs -> Names
  var teamMap = {};
  (state.teams || []).forEach(function(t) { teamMap[t.id] = t.name; });
  var tName = function(id) { return teamMap[id] || id || ""; };

  // 2. Rebuild TRADES_BACKUP tab
  sheets.tradesSheet.clearContents();
  var tradeRows = [[
    "Trade ID", "Primary Year", "Date / Sheet", "Logged By", "Managers Involved",
    "Headline", "Direct Pick Transfers", "Pools (Pending Final / Finalized)",
    "Conditional Swaps", "Notes"
  ]];

  (state.trades || []).forEach(function(tr) {
    var mgrs = [tr.teamA, tr.teamB, tr.teamC].filter(Boolean).map(tName).join(" <-> ");
    var directs = (tr.directTransfers || []).map(function(dt) {
      return dt.year + " Rd " + dt.round + " (" + tName(dt.originalOwner) + " -> " + tName(dt.toOwner) + ")";
    }).join("; ");
    var pools = (tr.pools || []).map(function(p) {
      var st = p.resolvedOriginalOwner ? ("Finalized -> " + tName(p.resolvedOriginalOwner)) : "PENDING FINAL";
      return p.year + " Rd " + p.round + " Pool (" + tName(p.poolHolder) + " owes 1 to " + tName(p.owedTo) + " | " + st + ")";
    }).join("; ");
    var conds = (tr.conditionals || []).map(function(c) {
      var st = c.isResolved ? (c.resolvedRound ? "Swapped Rd " + c.resolvedRound : "Settled") : "PENDING FINAL";
      return c.year + " Rd " + (c.roundsInvolved || []).join("/") + " (" + st + ": " + (c.summary || "") + ")";
    }).join("; ");

    tradeRows.push([
      tr.id || "",
      tr.year || "",
      tr.date || "",
      tr.submittedBy || "Commissioner",
      mgrs,
      tr.headline || "",
      directs,
      pools,
      conds,
      tr.notes || ""
    ]);
  });

  sheets.tradesSheet.getRange(1, 1, tradeRows.length, tradeRows[0].length).setValues(tradeRows);
  sheets.tradesSheet.getRange(1, 1, 1, tradeRows[0].length).setFontWeight("bold");

  // 3. Rebuild PICK_LEDGER_BACKUP tab (All traded picks & Pending Final conditions)
  sheets.picksSheet.clearContents();
  var pickRows = [[
    "Draft Year", "Round", "Type / Status", "Original Owner / Pool Holder",
    "Current Owner / Owed To", "Rule / Details", "Linked Trade ID"
  ]];

  (state.trades || []).forEach(function(tr) {
    (tr.directTransfers || []).forEach(function(dt) {
      pickRows.push([
        dt.year, dt.round, "SETTLED TRANSFER",
        tName(dt.originalOwner), tName(dt.toOwner),
        tr.headline || "", tr.id || ""
      ]);
    });
    (tr.pools || []).forEach(function(p) {
      pickRows.push([
        p.year, p.round, p.resolvedOriginalOwner ? "FINALIZED POOL" : "PENDING FINAL POOL",
        tName(p.poolHolder), tName(p.owedTo),
        p.ruleLabel || "", tr.id || ""
      ]);
    });
    (tr.conditionals || []).forEach(function(c) {
      pickRows.push([
        c.year, (c.roundsInvolved || []).join("/"), c.isResolved ? "FINALIZED SWAP" : "PENDING FINAL SWAP",
        tName(c.teamA) + " <-> " + tName(c.teamB), "Favors " + tName(c.favoredTeam),
        c.summary || "", tr.id || ""
      ]);
    });
  });

  sheets.picksSheet.getRange(1, 1, pickRows.length, pickRows[0].length).setValues(pickRows);
  sheets.picksSheet.getRange(1, 1, 1, pickRows[0].length).setFontWeight("bold");
}
