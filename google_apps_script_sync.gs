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
 * Automatically maintains 5 tabs inside your Google Sheet:
 *   - LIVE_STATE_JSON     (Canonical JSON state + 30-version rollback history, multi-cell chunk-safe)
 *   - TRADES_BACKUP       (Human-readable spreadsheet table of every trade)
 *   - PICK_LEDGER_BACKUP  (Human-readable spreadsheet backup of traded picks & Pending Final items)
 *   - ESPN_ROSTERS_BACKUP (Human-readable spreadsheet of synced ESPN rosters, keeper locks & OTB tags)
 *   - ESPN_HISTORY_BACKUP (Human-readable spreadsheet of all-time ESPN season standings & trophies)
 * ============================================================================
 */

var DEFAULT_PASSPHRASE_HASH = "ed82a630acbc4c03fe026e616f3002f5d7ffd51a462cc9023efca21b90a1bd4e"; // Salted SHA-256 hash of default league passphrase
var DEFAULT_COMMISH_PIN_HASH = "9dc15277bc17bbea84f404584d0deff365108b37ccff158a6b2311ddcc5a6b8a"; // Salted SHA-256 hash of default Commissioner PIN
var CELL_CHUNK_SIZE = 44000; // Safe margin below Google Sheets 50,000 char/cell limit

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

  var espnRostersSheet = ss.getSheetByName("ESPN_ROSTERS_BACKUP");
  if (!espnRostersSheet) {
    espnRostersSheet = ss.insertSheet("ESPN_ROSTERS_BACKUP");
  }

  var espnHistorySheet = ss.getSheetByName("ESPN_HISTORY_BACKUP");
  if (!espnHistorySheet) {
    espnHistorySheet = ss.insertSheet("ESPN_HISTORY_BACKUP");
  }

  return {
    stateSheet: stateSheet,
    tradesSheet: tradesSheet,
    picksSheet: picksSheet,
    espnRostersSheet: espnRostersSheet,
    espnHistorySheet: espnHistorySheet
  };
}

function readStoredState(stateSheet) {
  var raw = stateSheet.getRange("B2").getValue();
  if (!raw || typeof raw !== "string") return null;
  var joined = raw;
  // Support multi-cell overflow chunks in D2:Z2 when full ESPN history + rosters exceed 44,000 chars
  if (raw.length >= CELL_CHUNK_SIZE) {
    try {
      var extraRow = stateSheet.getRange(2, 4, 1, 20).getValues();
      if (extraRow && extraRow[0]) {
        for (var c = 0; c < extraRow[0].length; c++) {
          var part = extraRow[0][c];
          if (part && typeof part === "string") {
            joined += part;
          } else {
            break;
          }
        }
      }
    } catch (ignoreChunkErr) {}
  }
  try {
    return JSON.parse(joined);
  } catch (err) {
    try {
      return JSON.parse(raw);
    } catch (err2) {
      return null;
    }
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

function sanitizeStateForClient(stateObj) {
  if (!stateObj || typeof stateObj !== "object") return stateObj;
  var clone = JSON.parse(JSON.stringify(stateObj));
  delete clone.commishPinHash;
  delete clone.leaguePassphraseHash;
  delete clone.cloudSyncUrl;
  if (clone.teamProfiles && typeof clone.teamProfiles === "object") {
    Object.keys(clone.teamProfiles).forEach(function(tid) {
      if (clone.teamProfiles[tid]) {
        delete clone.teamProfiles[tid].teamPinHash;
      }
    });
  }
  return clone;
}

function safeSheetCell(val) {
  if (val === null || val === undefined) return "";
  if (typeof val === "number" || typeof val === "boolean") return val;
  var s = String(val);
  if (/^[=+\-@]/.test(s)) {
    return "'" + s;
  }
  return s;
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
  return jsonResponse({ ok: true, empty: false, state: sanitizeStateForClient(stored) });
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

    // Helper to compute current pick ownership for a given year in baseState
    function getCurrentPickOwnerInState(baseStateObj, yr, round, origOwner) {
      var owner = origOwner;
      var tradesList = Array.isArray(baseStateObj.trades) ? baseStateObj.trades : [];
      for (var ti = 0; ti < tradesList.length; ti++) {
        var tItem = tradesList[ti];
        var dts = Array.isArray(tItem.directTransfers) ? tItem.directTransfers : [];
        for (var di = 0; di < dts.length; di++) {
          var d = dts[di];
          if (Number(d.year) === Number(yr) && Number(d.round) === Number(round) && String(d.originalOwner) === String(origOwner)) {
            owner = String(d.toOwner);
          }
        }
        var pls = Array.isArray(tItem.pools) ? tItem.pools : [];
        for (var pi = 0; pi < pls.length; pi++) {
          var p = pls[pi];
          if (Number(p.year) === Number(yr) && Number(p.round) === Number(round) && p.resolvedOriginalOwner && String(p.resolvedOriginalOwner) === String(origOwner)) {
            owner = String(p.owedTo);
          }
        }
      }
      return owner;
    }

    function assignOfficialTradeId(baseStateObj, primaryYear) {
      var prefix = "T-" + primaryYear + "-";
      var maxNum = 0;
      var tradesList = Array.isArray(baseStateObj.trades) ? baseStateObj.trades : [];
      for (var i = 0; i < tradesList.length; i++) {
        var existingId = String(tradesList[i].id || "");
        if (existingId.indexOf(prefix) === 0) {
          var n = parseInt(existingId.replace(prefix, ""), 10);
          if (!isNaN(n) && n > maxNum) maxNum = n;
        }
      }
      return prefix + (maxNum + 1 < 10 ? "0" + (maxNum + 1) : String(maxNum + 1));
    }

    function sanitizeIncomingTradeObject(tr, primaryYear, assignedId) {
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

      var reqTeams = [String(tr.teamA || ""), String(tr.teamB || "")];
      if (tr.teamC) reqTeams.push(String(tr.teamC));

      var proposerId = String(tr.proposedByTeam || tr.proposerTeamId || tr.teamA || "");
      var confBy = Array.isArray(tr.confirmedBy)
        ? tr.confirmedBy.map(function(x) { return String(x || ""); }).filter(Boolean)
        : [proposerId];
      if (proposerId && confBy.indexOf(proposerId) === -1) confBy.push(proposerId);

      var cleanApprovals = {};
      if (tr.approvals && typeof tr.approvals === "object") {
        for (var ak in tr.approvals) {
          if (Object.prototype.hasOwnProperty.call(tr.approvals, ak) && tr.approvals[ak]) {
            cleanApprovals[String(ak)] = true;
          }
        }
      }
      for (var ci = 0; ci < confBy.length; ci++) {
        cleanApprovals[confBy[ci]] = true;
      }

      var cleanPlayers = Array.isArray(tr.playerTransfers) ? tr.playerTransfers.slice(0, 30).map(function(pt) {
        return {
          playerId: String(pt.playerId || "").slice(0, 40),
          name: String(pt.name || "").slice(0, 80),
          pos: String(pt.pos || "").slice(0, 12),
          nflTeam: String(pt.nflTeam || "").slice(0, 12),
          fromOwner: String(pt.fromOwner || "").slice(0, 30),
          toOwner: String(pt.toOwner || "").slice(0, 30)
        };
      }) : [];

      return {
        id: assignedId,
        year: primaryYear,
        date: String(tr.date || "").slice(0, 80),
        teamA: String(tr.teamA || ""),
        teamB: String(tr.teamB || ""),
        teamC: tr.teamC ? String(tr.teamC) : null,
        headline: String(tr.headline || "").slice(0, 400),
        notes: String(tr.notes || "").slice(0, 800),
        submittedBy: tr.submittedBy ? String(tr.submittedBy).slice(0, 60) : "",
        proposedByTeam: proposerId,
        proposerTeamId: proposerId,
        approvals: cleanApprovals,
        confirmedBy: Object.keys(cleanApprovals),
        requiredTeams: reqTeams,
        status: tr.status ? String(tr.status) : "OFFICIAL",
        createdAt: Number(tr.createdAt) || Date.now(),
        directTransfers: Array.isArray(tr.directTransfers) ? tr.directTransfers.map(function(dt) {
          return {
            year: Number(dt.year) || primaryYear,
            round: Math.min(20, Math.max(1, Number(dt.round) || 1)),
            originalOwner: String(dt.originalOwner || ""),
            fromOwner: String(dt.fromOwner || ""),
            toOwner: String(dt.toOwner || "")
          };
        }) : [],
        playerTransfers: cleanPlayers,
        pools: cleanPools,
        conditionals: cleanConds,
        reactions: (tr.reactions && typeof tr.reactions === "object") ? tr.reactions : {}
      };
    }

    function applyConfirmedPlayerTransfersToRosters(baseStateObj, tradeObj) {
      if (!baseStateObj || !baseStateObj.espnData || !baseStateObj.espnData.rosters || !tradeObj) return;
      var pts = Array.isArray(tradeObj.playerTransfers) ? tradeObj.playerTransfers : [];
      if (pts.length === 0) return;
      var rosters = baseStateObj.espnData.rosters;
      for (var i = 0; i < pts.length; i++) {
        var pt = pts[i];
        var fromList = Array.isArray(rosters[pt.fromOwner]) ? rosters[pt.fromOwner] : [];
        var toList = Array.isArray(rosters[pt.toOwner]) ? rosters[pt.toOwner] : [];
        var movedPlayer = null;
        for (var fi = 0; fi < fromList.length; fi++) {
          if (String(fromList[fi].id) === String(pt.playerId) || String(fromList[fi].name) === String(pt.name)) {
            movedPlayer = fromList.splice(fi, 1)[0];
            break;
          }
        }
        if (movedPlayer) {
          toList.push(movedPlayer);
          rosters[pt.fromOwner] = fromList;
          rosters[pt.toOwner] = toList;
        }
      }
    }

    // ------------------------------------------------------------------------
    // ACTION 1: LEAGUE MEMBER PROPOSES OR LOGS A TRADE
    // ------------------------------------------------------------------------
    if (action === "PROPOSE_MEMBER_TRADE" || action === "ADD_MEMBER_TRADE") {
      var baseState = stored || body.fallbackBaseState;
      if (!baseState || !Array.isArray(baseState.teams) || !Array.isArray(baseState.trades)) {
        return jsonResponse({ ok: false, error: "Initialize database from Commissioner Mode first." });
      }
      if (!Array.isArray(baseState.pendingTrades)) baseState.pendingTrades = [];

      var tr = body.trade || body.proposal;
      if (!tr || !tr.headline || !tr.teamA || !tr.teamB) {
        return jsonResponse({ ok: false, error: "Invalid trade submission." });
      }

      var primaryYear = Number(tr.year) || Number(baseState.baseUpcomingYear) || 2027;
      var dtsToCheck = Array.isArray(tr.directTransfers) ? tr.directTransfers : [];
      for (var dIdx = 0; dIdx < dtsToCheck.length; dIdx++) {
        var checkDt = dtsToCheck[dIdx];
        var currOwner = getCurrentPickOwnerInState(baseState, checkDt.year || primaryYear, checkDt.round, checkDt.originalOwner);
        if (currOwner !== String(checkDt.fromOwner)) {
          return jsonResponse({
            ok: false,
            error: "Pick Conflict: " + checkDt.fromOwner + " no longer owns " + (checkDt.year || primaryYear) + " Rd " + checkDt.round + " (" + checkDt.originalOwner + "'s Pick)."
          });
        }
      }

      var isProposal = (action === "PROPOSE_MEMBER_TRADE") || Boolean(body.requirePartnerConfirmation);
      if (isProposal) {
        var propId = String(tr.id || ("PROP-" + primaryYear + "-" + String(Date.now()).slice(-5)));
        var cleanProp = sanitizeIncomingTradeObject(tr, primaryYear, propId);
        cleanProp.status = "PENDING_CONFIRMATION";
        baseState.pendingTrades.push(cleanProp);
        baseState.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseState, "TRADE_PROPOSAL:" + propId);
        return jsonResponse({ ok: true, assignedId: propId, state: sanitizeStateForClient(baseState) });
      }

      var assignedId = assignOfficialTradeId(baseState, primaryYear);
      var cleanTrade = sanitizeIncomingTradeObject(tr, primaryYear, assignedId);
      baseState.trades.push(cleanTrade);
      applyConfirmedPlayerTransfersToRosters(baseState, cleanTrade);
      baseState.updatedAt = Date.now();

      saveAndMirrorToSheets(sheets, baseState, "MEMBER_TRADE:" + assignedId);
      return jsonResponse({ ok: true, assignedId: assignedId, state: sanitizeStateForClient(baseState) });
    }

    // ------------------------------------------------------------------------
    // ACTION 1A: PARTNER CONFIRMS, DECLINES, OR PROPOSER WITHDRAWS PROPOSAL
    // ------------------------------------------------------------------------
    if (action === "RESPOND_TRADE_PROPOSAL") {
      var baseStateResp = stored || body.fallbackBaseState;
      if (!baseStateResp || !Array.isArray(baseStateResp.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database from Commissioner Mode first." });
      }
      if (!Array.isArray(baseStateResp.pendingTrades)) baseStateResp.pendingTrades = [];
      if (!Array.isArray(baseStateResp.trades)) baseStateResp.trades = [];

      var propIdTarget = String(body.proposalId || "");
      var subAction = String(body.decision || body.responseAction || "ACCEPT"); // ACCEPT | DECLINE | WITHDRAW
      var actorTeamId = String(body.actorTeamId || body.responderTeamId || "");
      var replyNote = String(body.replyNote || "").slice(0, 240);

      var foundIdx = -1;
      for (var pIdx = 0; pIdx < baseStateResp.pendingTrades.length; pIdx++) {
        if (String(baseStateResp.pendingTrades[pIdx].id) === propIdTarget) {
          foundIdx = pIdx;
          break;
        }
      }
      if (foundIdx === -1) {
        return jsonResponse({ ok: false, error: "Trade proposal no longer exists or was already processed." });
      }

      var propObj = baseStateResp.pendingTrades[foundIdx];
      if (subAction === "DECLINE" || subAction === "WITHDRAW") {
        baseStateResp.pendingTrades.splice(foundIdx, 1);
        baseStateResp.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseStateResp, subAction + "_PROPOSAL:" + propIdTarget + (replyNote ? " (" + replyNote + ")" : ""));
        return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateResp) });
      }

      // ACCEPT flow
      if (!propObj.approvals || typeof propObj.approvals !== "object") propObj.approvals = {};
      var initPropTeam = propObj.proposedByTeam || propObj.proposerTeamId || propObj.teamA;
      if (initPropTeam) propObj.approvals[initPropTeam] = true;
      if (actorTeamId) propObj.approvals[actorTeamId] = true;

      var reqList = Array.isArray(propObj.requiredTeams) && propObj.requiredTeams.length >= 2
        ? propObj.requiredTeams
        : [propObj.teamA, propObj.teamB].concat(propObj.teamC ? [propObj.teamC] : []);

      if (body.commishAuthHash) {
        for (var rci = 0; rci < reqList.length; rci++) {
          propObj.approvals[reqList[rci]] = true;
        }
      }
      propObj.confirmedBy = Object.keys(propObj.approvals);

      var allConfirmed = reqList.every(function(reqTid) {
        return Boolean(propObj.approvals[reqTid]);
      });

      if (allConfirmed) {
        // Verify Double-Booked Pick Conflict Guard before finalizing!
        var propDts = Array.isArray(propObj.directTransfers) ? propObj.directTransfers : [];
        for (var vc = 0; vc < propDts.length; vc++) {
          var dtItem = propDts[vc];
          var liveOwner = getCurrentPickOwnerInState(baseStateResp, dtItem.year || propObj.year, dtItem.round, dtItem.originalOwner);
          if (liveOwner !== String(dtItem.fromOwner)) {
            baseStateResp.pendingTrades.splice(foundIdx, 1);
            baseStateResp.updatedAt = Date.now();
            saveAndMirrorToSheets(sheets, baseStateResp, "AUTO_VOID_CONFLICT:" + propIdTarget);
            return jsonResponse({
              ok: false,
              error: "Conflict Guard: Pick " + (dtItem.year || propObj.year) + " Rd " + dtItem.round + " (" + dtItem.originalOwner + ") was traded in another deal first. Proposal voided.",
              state: sanitizeStateForClient(baseStateResp)
            });
          }
        }

        baseStateResp.pendingTrades.splice(foundIdx, 1);
        var finalId = assignOfficialTradeId(baseStateResp, Number(propObj.year) || 2027);
        propObj.id = finalId;
        if (replyNote) {
          propObj.notes = (propObj.notes ? propObj.notes + " | " : "") + "Partner Note: " + replyNote;
        }
        baseStateResp.trades.push(propObj);
        applyConfirmedPlayerTransfersToRosters(baseStateResp, propObj);

        // Auto-prune any remaining open proposals whose directTransfers now conflict
        baseStateResp.pendingTrades = baseStateResp.pendingTrades.filter(function(otherProp) {
          var oDts = Array.isArray(otherProp.directTransfers) ? otherProp.directTransfers : [];
          for (var oi = 0; oi < oDts.length; oi++) {
            var od = oDts[oi];
            var nowOwner = getCurrentPickOwnerInState(baseStateResp, od.year || otherProp.year, od.round, od.originalOwner);
            if (nowOwner !== String(od.fromOwner)) return false;
          }
          return true;
        });

        baseStateResp.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseStateResp, "CONFIRMED_TRADE:" + finalId);
        return jsonResponse({ ok: true, finalizedTradeId: finalId, state: sanitizeStateForClient(baseStateResp) });
      } else {
        baseStateResp.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseStateResp, "PARTIAL_CONFIRM:" + propIdTarget + ":" + actorTeamId);
        return jsonResponse({ ok: true, partial: true, state: sanitizeStateForClient(baseStateResp) });
      }
    }

    // ------------------------------------------------------------------------
    // ACTION 1C: LEAGUE MEMBER REACTS TO A FINALIZED TRADE (🔥 🤝 🌮 💀)
    // ------------------------------------------------------------------------
    if (action === "REACT_TO_TRADE") {
      var baseStateReact = stored || body.fallbackBaseState;
      if (!baseStateReact || !Array.isArray(baseStateReact.trades)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      var targetTradeId = String(body.tradeId || "");
      var emoji = String(body.emoji || "");
      var reactorTeam = String(body.teamId || "");
      var allowedEmojis = ["🔥", "🤝", "🌮", "💀"];
      if (allowedEmojis.indexOf(emoji) === -1 || !reactorTeam) {
        return jsonResponse({ ok: false, error: "Invalid reaction." });
      }
      for (var rti = 0; rti < baseStateReact.trades.length; rti++) {
        if (String(baseStateReact.trades[rti].id) === targetTradeId) {
          var trObj = baseStateReact.trades[rti];
          if (!trObj.reactions || typeof trObj.reactions !== "object") trObj.reactions = {};
          if (!Array.isArray(trObj.reactions[emoji])) trObj.reactions[emoji] = [];
          var existingIdx = trObj.reactions[emoji].indexOf(reactorTeam);
          if (existingIdx !== -1) {
            trObj.reactions[emoji].splice(existingIdx, 1);
          } else {
            trObj.reactions[emoji].push(reactorTeam);
          }
          baseStateReact.updatedAt = Date.now();
          saveAndMirrorToSheets(sheets, baseStateReact, "TRADE_REACTION:" + targetTradeId);
          return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateReact) });
        }
      }
      return jsonResponse({ ok: false, error: "Trade not found." });
    }

    // ------------------------------------------------------------------------
    // ACTION 1B: FRANCHISE MANAGER UPDATES THEIR OWN TEAM WAR ROOM / KEEPERS / IDENTITY
    // ------------------------------------------------------------------------
    if (action === "UPDATE_TEAM_PROFILE") {
      var baseStateProf = stored || body.fallbackBaseState;
      if (!baseStateProf || !Array.isArray(baseStateProf.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database from Commissioner Mode first." });
      }
      var teamId = String(body.teamId || "");
      var validTeam = baseStateProf.teams.some(function(t) { return t.id === teamId; });
      if (!validTeam) {
        return jsonResponse({ ok: false, error: "Invalid franchise ID." });
      }
      if (!baseStateProf.teamProfiles || typeof baseStateProf.teamProfiles !== "object") {
        baseStateProf.teamProfiles = {};
      }
      var existingProf = baseStateProf.teamProfiles[teamId] || {};
      if (existingProf.teamPinHash && /^[a-f0-9]{64}$/i.test(existingProf.teamPinHash)) {
        var providedTeamPinHash = body.teamPinHash ? String(body.teamPinHash).toLowerCase() : "";
        if (providedTeamPinHash !== existingProf.teamPinHash.toLowerCase()) {
          return jsonResponse({ ok: false, error: "Personal Team PIN required to update this franchise." });
        }
      }
      var incoming = body.profile || {};
      var sanitizedTags = {};
      if (incoming.pickTags && typeof incoming.pickTags === "object") {
        Object.keys(incoming.pickTags).forEach(function(k) {
          var v = incoming.pickTags[k];
          if ((v === "ON_BLOCK" || v === "UNTOUCHABLE") && /^\d{4}-\d{1,2}-[a-z0-9_]+$/i.test(k)) {
            sanitizedTags[k] = v;
          }
        });
      }
      var sanitizedPlayerTags = {};
      if (incoming.playerTags && typeof incoming.playerTags === "object") {
        Object.keys(incoming.playerTags).slice(0, 60).forEach(function(pk) {
          var pv = incoming.playerTags[pk];
          if ((pv === "ON_BLOCK" || pv === "UNTOUCHABLE") && /^[A-Za-z0-9_.-]+$/.test(pk)) {
            sanitizedPlayerTags[pk] = pv;
          }
        });
      }
      var sanitizedKeepers = [];
      if (Array.isArray(incoming.keepers)) {
        incoming.keepers.slice(0, 20).forEach(function(kid) {
          var ks = String(kid || "").trim().slice(0, 40);
          if (ks && sanitizedKeepers.indexOf(ks) === -1) sanitizedKeepers.push(ks);
        });
      } else if (Array.isArray(existingProf.keepers)) {
        sanitizedKeepers = existingProf.keepers;
      }

      var sanitizedChips = [];
      if (Array.isArray(incoming.needsChips)) {
        incoming.needsChips.slice(0, 12).forEach(function(c) {
          var s = String(c || "").trim().slice(0, 32);
          if (s && sanitizedChips.indexOf(s) === -1) sanitizedChips.push(s);
        });
      }
      // Once upstateKittyCursed is true, a non-admin manager CANNOT clear it themselves!
      var wasCursed = Boolean(existingProf.upstateKittyCursed);
      var nowCursed = wasCursed || Boolean(incoming.upstateKittyCursed);

      if (!wasCursed && nowCursed) {
        if (!Array.isArray(baseStateProf.hallOfShame)) baseStateProf.hallOfShame = [];
        var shameTeamObj = baseStateProf.teams.filter(function(t) { return t.id === teamId; })[0];
        var shameTeamName = shameTeamObj ? shameTeamObj.name : teamId;
        baseStateProf.hallOfShame.unshift({
          id: "SHAME-" + Date.now(),
          teamId: teamId,
          teamName: shameTeamName,
          message: shameTeamName + " selected the 'Upstate Kitties' colorway and activated the Sparkly Pink 8==D Curse!",
          timestamp: Date.now(),
          dateStr: new Date().toLocaleDateString()
        });
        baseStateProf.hallOfShame = baseStateProf.hallOfShame.slice(0, 30);
      }

      baseStateProf.teamProfiles[teamId] = {
        tradeBlockNote: String(incoming.tradeBlockNote || "").slice(0, 280),
        lookingFor: String(incoming.lookingFor || "").slice(0, 180),
        needsChips: sanitizedChips,
        pickTags: sanitizedTags,
        playerTags: incoming.playerTags && typeof incoming.playerTags === "object" ? sanitizedPlayerTags : (existingProf.playerTags || {}),
        keepers: sanitizedKeepers,
        keepersConfirmedAt: Number(incoming.keepersConfirmedAt || existingProf.keepersConfirmedAt || 0),
        customTeamName: String(incoming.customTeamName || "").trim().slice(0, 24),
        avatarType: ["MONOGRAM", "EMOJI", "IMAGE"].indexOf(incoming.avatarType) !== -1 ? incoming.avatarType : (existingProf.avatarType || "MONOGRAM"),
        avatarText: String(incoming.avatarText || "").trim().slice(0, 4),
        avatarEmoji: String(incoming.avatarEmoji || "").trim().slice(0, 12),
        avatarBg: String(incoming.avatarBg || "").trim().slice(0, 24),
        avatarImageDataUrl: String(incoming.avatarImageDataUrl || "").trim().slice(0, 8000),
        upstateKittyCursed: nowCursed,
        upstateKittyCursedAt: nowCursed ? (Number(existingProf.upstateKittyCursedAt) || Date.now()) : 0,
        flairBadge: String(existingProf.flairBadge || ""),
        teamPinVerifier: (incoming.teamPinVerifier && /^[a-f0-9]{64}$/i.test(incoming.teamPinVerifier)) ? incoming.teamPinVerifier.toLowerCase() : "",
        teamPinHash: (incoming.teamPinHash && /^[a-f0-9]{64}$/i.test(incoming.teamPinHash)) ? incoming.teamPinHash.toLowerCase() : "",
        updatedAt: Date.now()
      };
      if (Array.isArray(incoming.roster)) {
        if (!baseStateProf.espnData || typeof baseStateProf.espnData !== "object") {
          baseStateProf.espnData = { rosters: {} };
        }
        if (!baseStateProf.espnData.rosters || typeof baseStateProf.espnData.rosters !== "object") {
          baseStateProf.espnData.rosters = {};
        }
        baseStateProf.espnData.rosters[teamId] = incoming.roster.slice(0, 40).map(function(rp, rIdx) {
          return {
            id: String(rp && rp.id ? rp.id : ("p-" + teamId + "-" + (rIdx + 1))).slice(0, 40),
            espnId: String(rp && rp.espnId ? rp.espnId : "").slice(0, 20),
            name: String(rp && rp.name ? rp.name : "").trim().slice(0, 60),
            pos: String(rp && rp.pos ? rp.pos : "FLEX").trim().toUpperCase().slice(0, 12),
            nflTeam: String(rp && rp.nflTeam ? rp.nflTeam : "NFL").trim().toUpperCase().slice(0, 8),
            slot: String(rp && rp.slot ? rp.slot : "BE").trim().toUpperCase().slice(0, 10),
            projPts: Number(rp && rp.projPts) || 0,
            totalPts: Number(rp && rp.totalPts) || 0
          };
        }).filter(function(rp) { return Boolean(rp.name); });
      }
      baseStateProf.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateProf, (!wasCursed && nowCursed) ? ("UPSTATE_KITTY_CURSE:" + teamId) : ("TEAM_WAR_ROOM:" + teamId));
      return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateProf) });
    }

    // ------------------------------------------------------------------------
    // ACTION 1D: 1-CLICK ESPN CONSOLE EXTRACTOR OR COMMISSIONER ESPN DATA SYNC
    // ------------------------------------------------------------------------
    if (action === "ESPN_SYNC_UPDATE") {
      var baseStateEspn = stored || body.fallbackBaseState;
      if (!baseStateEspn || !Array.isArray(baseStateEspn.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database from Commissioner Mode first." });
      }
      var incomingEspn = body.espnData;
      if (!incomingEspn || typeof incomingEspn !== "object") {
        return jsonResponse({ ok: false, error: "Missing espnData payload." });
      }
      var existingEspn = (baseStateEspn.espnData && typeof baseStateEspn.espnData === "object") ? baseStateEspn.espnData : {};
      baseStateEspn.espnData = {
        leagueId: String(incomingEspn.leagueId || existingEspn.leagueId || "").slice(0, 32),
        leagueName: String(incomingEspn.leagueName || existingEspn.leagueName || "Gamecock Mafia").slice(0, 80),
        currentSeason: Number(incomingEspn.currentSeason) || Number(existingEspn.currentSeason) || 2026,
        lastSyncedAt: Date.now(),
        keeperRules: (incomingEspn.keeperRules && typeof incomingEspn.keeperRules === "object")
          ? incomingEspn.keeperRules
          : (existingEspn.keeperRules || { maxKeepers: 12, keeperRounds: 12, posCaps: { RB: 4, WR: 4, OTHER: 3 } }),
        teamSlotMap: (incomingEspn.teamSlotMap && typeof incomingEspn.teamSlotMap === "object")
          ? incomingEspn.teamSlotMap
          : (existingEspn.teamSlotMap || {}),
        seasonOwnerOverrides: (incomingEspn.seasonOwnerOverrides && typeof incomingEspn.seasonOwnerOverrides === "object")
          ? incomingEspn.seasonOwnerOverrides
          : (existingEspn.seasonOwnerOverrides || {}),
        alumniManagers: Array.isArray(incomingEspn.alumniManagers)
          ? incomingEspn.alumniManagers
          : (existingEspn.alumniManagers || []),
        rosters: (incomingEspn.rosters && typeof incomingEspn.rosters === "object")
          ? incomingEspn.rosters
          : (existingEspn.rosters || {}),
        seasonStandings: (incomingEspn.seasonStandings && typeof incomingEspn.seasonStandings === "object")
          ? incomingEspn.seasonStandings
          : (existingEspn.seasonStandings || {}),
        headToHead: (incomingEspn.headToHead && typeof incomingEspn.headToHead === "object")
          ? incomingEspn.headToHead
          : (existingEspn.headToHead || {}),
        records: (incomingEspn.records && typeof incomingEspn.records === "object")
          ? incomingEspn.records
          : (existingEspn.records || {}),
        pastDraftPicks: (incomingEspn.pastDraftPicks && typeof incomingEspn.pastDraftPicks === "object")
          ? incomingEspn.pastDraftPicks
          : (existingEspn.pastDraftPicks || {}),
        rawTeams: Array.isArray(incomingEspn.rawTeams)
          ? incomingEspn.rawTeams.slice(0, 20)
          : (existingEspn.rawTeams || [])
      };
      baseStateEspn.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateEspn, "ESPN_SYNC_UPDATE:" + (baseStateEspn.espnData.leagueId || "LIVE"));
      return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateEspn) });
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

      if (!newState.commishPinHash && stored && stored.commishPinHash) {
        newState.commishPinHash = stored.commishPinHash;
      }
      if (!newState.leaguePassphraseHash && stored && stored.leaguePassphraseHash) {
        newState.leaguePassphraseHash = stored.leaguePassphraseHash;
      }
      delete newState.cloudSyncUrl;

      newState.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, newState, body.reason || "ADMIN_SYNC");
      return jsonResponse({ ok: true, state: sanitizeStateForClient(newState) });
    }

    return jsonResponse({ ok: false, error: "Unknown sync action." });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Writes canonical JSON (chunk-safe across columns if >44k chars) + rolling 30-snapshot history + human-readable Google Sheet tabs!
 */
function saveAndMirrorToSheets(sheets, state, actionLabel) {
  var nowIso = new Date().toISOString();
  var jsonStr = JSON.stringify(state);

  // 1. Update LIVE_STATE_JSON (Chunk-safe if jsonStr > 44,000 chars)
  var primaryChunk = jsonStr.slice(0, CELL_CHUNK_SIZE);
  sheets.stateSheet.getRange("B2:C2").setValues([[primaryChunk, nowIso]]);
  if (jsonStr.length > CELL_CHUNK_SIZE) {
    var overflowChunks = [];
    for (var pos = CELL_CHUNK_SIZE; pos < jsonStr.length; pos += CELL_CHUNK_SIZE) {
      overflowChunks.push(jsonStr.slice(pos, pos + CELL_CHUNK_SIZE));
    }
    while (overflowChunks.length < 10) overflowChunks.push("");
    sheets.stateSheet.getRange(2, 4, 1, overflowChunks.length).setValues([overflowChunks]);
  } else {
    try {
      sheets.stateSheet.getRange(2, 4, 1, 10).clearContent();
    } catch (ignoreClearErr) {}
  }

  sheets.stateSheet.insertRowBefore(5);
  sheets.stateSheet.getRange("A5:D5").setValues([[state.updatedAt || Date.now(), nowIso, safeSheetCell(actionLabel), primaryChunk]]);
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
    "Headline", "Direct Pick Transfers", "Player Transfers", "Pools (Pending Final / Finalized)",
    "Conditional Swaps", "Notes"
  ]];

  (state.trades || []).forEach(function(tr) {
    var mgrs = [tr.teamA, tr.teamB, tr.teamC].filter(Boolean).map(tName).join(" <-> ");
    var directs = (tr.directTransfers || []).map(function(dt) {
      return dt.year + " Rd " + dt.round + " (" + tName(dt.originalOwner) + " -> " + tName(dt.toOwner) + ")";
    }).join("; ");
    var players = (tr.playerTransfers || []).map(function(pt) {
      return pt.name + " (" + pt.pos + ", " + tName(pt.fromOwner) + " -> " + tName(pt.toOwner) + ")";
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
      safeSheetCell(tr.id || ""),
      tr.year || "",
      safeSheetCell(tr.date || ""),
      safeSheetCell(tr.submittedBy || "Commissioner"),
      safeSheetCell(mgrs),
      safeSheetCell(tr.headline || ""),
      safeSheetCell(directs),
      safeSheetCell(players),
      safeSheetCell(pools),
      safeSheetCell(conds),
      safeSheetCell(tr.notes || "")
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
        safeSheetCell(tName(dt.originalOwner)), safeSheetCell(tName(dt.toOwner)),
        safeSheetCell(tr.headline || ""), safeSheetCell(tr.id || "")
      ]);
    });
    (tr.pools || []).forEach(function(p) {
      pickRows.push([
        p.year, p.round, p.resolvedOriginalOwner ? "FINALIZED POOL" : "PENDING FINAL POOL",
        safeSheetCell(tName(p.poolHolder)), safeSheetCell(tName(p.owedTo)),
        safeSheetCell(p.ruleLabel || ""), safeSheetCell(tr.id || "")
      ]);
    });
    (tr.conditionals || []).forEach(function(c) {
      pickRows.push([
        c.year, safeSheetCell((c.roundsInvolved || []).join("/")), c.isResolved ? "FINALIZED SWAP" : "PENDING FINAL SWAP",
        safeSheetCell(tName(c.teamA) + " <-> " + tName(c.teamB)), safeSheetCell("Favors " + tName(c.favoredTeam)),
        safeSheetCell(c.summary || ""), safeSheetCell(tr.id || "")
      ]);
    });
  });

  sheets.picksSheet.getRange(1, 1, pickRows.length, pickRows[0].length).setValues(pickRows);
  sheets.picksSheet.getRange(1, 1, 1, pickRows[0].length).setFontWeight("bold");

  // 4. Rebuild ESPN_ROSTERS_BACKUP tab if espnData.rosters is present
  if (sheets.espnRostersSheet && state.espnData && state.espnData.rosters) {
    sheets.espnRostersSheet.clearContents();
    var rosterRows = [[
      "Franchise Manager", "Player ID", "Player Name", "Position", "NFL Team",
      "Lineup Slot", "Season Pts", "Keeper Locked?", "Trade Block Status"
    ]];
    var rostersMap = state.espnData.rosters || {};
    var profilesMap = state.teamProfiles || {};
    Object.keys(rostersMap).forEach(function(tid) {
      var pList = Array.isArray(rostersMap[tid]) ? rostersMap[tid] : [];
      var prof = profilesMap[tid] || {};
      var keepersArr = Array.isArray(prof.keepers) ? prof.keepers : [];
      var pTags = prof.playerTags || {};
      pList.forEach(function(pl) {
        var pid = String(pl.id || "");
        var isKeeper = keepersArr.indexOf(pid) !== -1 ? "LOCKED KEEPER" : "";
        var tagStatus = pTags[pid] || "";
        rosterRows.push([
          safeSheetCell(tName(tid)),
          safeSheetCell(pid),
          safeSheetCell(pl.name || ""),
          safeSheetCell(pl.pos || ""),
          safeSheetCell(pl.nflTeam || ""),
          safeSheetCell(pl.slot || "Bench"),
          Number(pl.pts) || 0,
          safeSheetCell(isKeeper),
          safeSheetCell(tagStatus)
        ]);
      });
    });
    sheets.espnRostersSheet.getRange(1, 1, rosterRows.length, rosterRows[0].length).setValues(rosterRows);
    sheets.espnRostersSheet.getRange(1, 1, 1, rosterRows[0].length).setFontWeight("bold");
  }

  // 5. Rebuild ESPN_HISTORY_BACKUP tab if espnData.seasonStandings is present
  if (sheets.espnHistorySheet && state.espnData && state.espnData.seasonStandings) {
    sheets.espnHistorySheet.clearContents();
    var histRows = [[
      "Season Year", "Final Rank", "Manager ID", "ESPN Team Name", "Wins", "Losses", "Ties",
      "Points For (PF)", "Points Against (PA)", "All-Play W", "All-Play L", "Champion?"
    ]];
    var standingsByYr = state.espnData.seasonStandings || {};
    Object.keys(standingsByYr).sort().reverse().forEach(function(yr) {
      var rows = Array.isArray(standingsByYr[yr]) ? standingsByYr[yr] : [];
      rows.forEach(function(r) {
        histRows.push([
          Number(yr) || yr,
          Number(r.finalStanding) || 0,
          safeSheetCell(r.managerId || ""),
          safeSheetCell(r.teamName || ""),
          Number(r.wins) || 0,
          Number(r.losses) || 0,
          Number(r.ties) || 0,
          Number(r.pointsFor) || 0,
          Number(r.pointsAgainst) || 0,
          Number(r.allPlayWins) || 0,
          Number(r.allPlayLosses) || 0,
          r.isChampion ? "CHAMPION 🏆" : ""
        ]);
      });
    });
    sheets.espnHistorySheet.getRange(1, 1, histRows.length, histRows[0].length).setValues(histRows);
    sheets.espnHistorySheet.getRange(1, 1, 1, histRows[0].length).setFontWeight("bold");
  }
}

/**
 * OPTIONAL AUTOMATED WEEKLY ESPN SYNC (HANDS-FREE SERVER-SIDE CRON)
 * To enable hands-free Tuesday morning ESPN roster & standings syncs:
 * 1. In Google Apps Script -> Project Settings -> Script Properties, add:
 *    - ESPN_LEAGUE_ID : "30215"
 *    - ESPN_S2        : "<your espn_s2 cookie value from fantasy.espn.com>"
 *    - ESPN_SWID      : "{<your SWID cookie value>}"
 * 2. Run `setupWeeklyEspnSyncTrigger()` once in the Apps Script editor.
 */
function setupWeeklyEspnSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "autoWeeklyEspnSync") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("autoWeeklyEspnSync")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(6)
    .create();
}

function autoWeeklyEspnSync() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var leagueId = props.getProperty("ESPN_LEAGUE_ID") || "30215";
    var espnS2 = props.getProperty("ESPN_S2") || "";
    var espnSwid = props.getProperty("ESPN_SWID") || "";

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ensureSheetsExist(ss);
    var currentState = readStoredState(sheets.stateSheet);
    if (!currentState || !currentState.espnData) return;

    var currentYear = new Date().getFullYear();
    var url = "https://lm-api-reads.fantasy.espn.com/apis/v3/games/ffl/seasons/" +
      currentYear + "/segments/0/leagues/" + leagueId +
      "?view=mTeam&view=mRoster&view=mStandings";

    var headers = {};
    if (espnS2 && espnSwid) {
      headers["Cookie"] = "espn_s2=" + espnS2 + "; SWID=" + espnSwid;
    }

    var resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: headers,
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return;

    var data = JSON.parse(resp.getContentText());
    if (!data || !Array.isArray(data.teams)) return;

    var slotMap = currentState.espnData.teamSlotMap || {};
    var posMap = { 1: "QB", 2: "RB", 3: "WR", 4: "TE", 5: "K", 9: "DT", 10: "DE", 11: "LB", 12: "CB", 13: "S", 14: "DB", 16: "D/ST" };
    var slotNameMap = { 0: "QB", 2: "RB", 4: "WR", 6: "TE", 7: "OP", 8: "DT", 9: "DE", 10: "LB", 11: "DL", 12: "CB", 13: "S", 14: "DB", 15: "DP", 16: "D/ST", 17: "K", 20: "BE", 21: "IR", 23: "FLEX" };
    var teamNflMap = { 0:"FA",1:"ATL",2:"BUF",3:"CHI",4:"CIN",5:"CLE",6:"DAL",7:"DEN",8:"DET",9:"GB",10:"TEN",11:"IND",12:"KC",13:"LV",14:"LAR",15:"MIA",16:"MIN",17:"NE",18:"NO",19:"NYG",20:"NYJ",21:"PHI",22:"ARI",23:"PIT",24:"LAC",25:"SF",26:"SEA",27:"TB",28:"WAS",29:"CAR",30:"JAX",33:"BAL",34:"HOU" };

    data.teams.forEach(function(t) {
      var slotKey = String(t.id);
      var mgrId = slotMap[slotKey];
      if (!mgrId) return;
      var entries = (t.roster && Array.isArray(t.roster.entries)) ? t.roster.entries : [];
      if (entries.length > 0) {
        currentState.espnData.rosters = currentState.espnData.rosters || {};
        currentState.espnData.rosters[mgrId] = entries.map(function(e) {
          var p = (e.playerPoolEntry && e.playerPoolEntry.player) || {};
          var pid = p.id || e.playerId;
          var pos = posMap[p.defaultPositionId] || "FLEX";
          var pts = Number((e.playerPoolEntry && e.playerPoolEntry.appliedStatTotal) || 0);
          return {
            id: "espn_" + pid,
            espnId: pid,
            name: p.fullName || "Unknown Player",
            pos: pos,
            nflTeam: teamNflMap[p.proTeamId] || "NFL",
            slot: slotNameMap[e.lineupSlotId] || "BE",
            totalPts: Math.round(pts * 10) / 10,
            acquiredVia: e.acquisitionType || "ROSTER"
          };
        });
      }
    });

    currentState.espnData.lastSyncedAt = Date.now();
    currentState.updatedAt = Date.now();
    saveAndMirrorToSheets(sheets, currentState, "WEEKLY_ESPN_AUTO_SYNC");
  } catch (err) {
    Logger.log("autoWeeklyEspnSync error: " + err);
  } finally {
    lock.releaseLock();
  }
}

