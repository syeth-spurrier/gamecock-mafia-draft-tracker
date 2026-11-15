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

    function doesTeamOwnPlayerInState(baseStateObj, ownerId, playerId, playerName) {
      if (!baseStateObj || !baseStateObj.espnData || !baseStateObj.espnData.rosters) return true;
      var rosters = baseStateObj.espnData.rosters;
      var populatedTeams = Object.keys(rosters).filter(function(k) {
        return Array.isArray(rosters[k]) && rosters[k].length > 0;
      });
      if (populatedTeams.length === 0) return true;
      var list = Array.isArray(rosters[ownerId]) ? rosters[ownerId] : [];
      var targetId = String(playerId || "").trim();
      var targetName = String(playerName || "").trim().toLowerCase();
      for (var i = 0; i < list.length; i++) {
        if (targetId && String(list[i].id) === targetId) return true;
        if (targetName && String(list[i].name || "").trim().toLowerCase() === targetName) return true;
      }
      return false;
    }

    function pruneConflictingPendingProposals(baseStateObj) {
      if (!baseStateObj || !Array.isArray(baseStateObj.pendingTrades)) return;
      baseStateObj.pendingTrades = baseStateObj.pendingTrades.filter(function(otherProp) {
        var oDts = Array.isArray(otherProp.directTransfers) ? otherProp.directTransfers : [];
        for (var oi = 0; oi < oDts.length; oi++) {
          var od = oDts[oi];
          var nowOwner = getCurrentPickOwnerInState(baseStateObj, od.year || otherProp.year, od.round, od.originalOwner);
          if (nowOwner !== String(od.fromOwner)) return false;
        }
        var oPts = Array.isArray(otherProp.playerTransfers) ? otherProp.playerTransfers : [];
        for (var pi = 0; pi < oPts.length; pi++) {
          var op = oPts[pi];
          if (!doesTeamOwnPlayerInState(baseStateObj, op.fromOwner, op.playerId, op.playerName || op.name)) {
            return false;
          }
        }
        return true;
      });
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
        var resolvedName = String(pt.playerName || pt.name || "").slice(0, 80);
        return {
          playerId: String(pt.playerId || "").slice(0, 40),
          name: resolvedName,
          playerName: resolvedName,
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
        supersededProposalId: tr.supersededProposalId ? String(tr.supersededProposalId).slice(0, 40) : "",
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
      var profiles = (baseStateObj.teamProfiles && typeof baseStateObj.teamProfiles === "object") ? baseStateObj.teamProfiles : {};
      for (var i = 0; i < pts.length; i++) {
        var pt = pts[i];
        var pName = String(pt.playerName || pt.name || "");
        var fromList = Array.isArray(rosters[pt.fromOwner]) ? rosters[pt.fromOwner] : [];
        var toList = Array.isArray(rosters[pt.toOwner]) ? rosters[pt.toOwner] : [];
        var movedPlayer = null;
        for (var fi = 0; fi < fromList.length; fi++) {
          if ((pt.playerId && String(fromList[fi].id) === String(pt.playerId)) ||
              (pName && String(fromList[fi].name).toLowerCase() === pName.toLowerCase())) {
            movedPlayer = fromList.splice(fi, 1)[0];
            break;
          }
        }
        // Never fabricate phantom "Traded Player" objects; only transfer verified roster entries
        if (movedPlayer) {
          movedPlayer.slot = "BE";
          var alreadyOnDest = toList.some(function(dp) { return String(dp.id) === String(movedPlayer.id); });
          if (!alreadyOnDest) toList.push(movedPlayer);
          rosters[pt.fromOwner] = fromList;
          rosters[pt.toOwner] = toList;
          // Strip locked keeper & trade-block tags from former owner
          var fromProf = profiles[pt.fromOwner];
          if (fromProf) {
            if (Array.isArray(fromProf.keepers)) {
              fromProf.keepers = fromProf.keepers.filter(function(kid) { return String(kid) !== String(movedPlayer.id); });
            }
            if (fromProf.playerTags && typeof fromProf.playerTags === "object") {
              delete fromProf.playerTags[movedPlayer.id];
            }
          }
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

      var minTradeableYr = Number(baseState.baseUpcomingYear) || 2026;
      var primaryYear = Number(tr.year) || minTradeableYr;
      var dtsToCheck = Array.isArray(tr.directTransfers) ? tr.directTransfers : [];
      for (var dIdx = 0; dIdx < dtsToCheck.length; dIdx++) {
        var checkDt = dtsToCheck[dIdx];
        var dtYr = Number(checkDt.year) || primaryYear;
        if (dtYr < minTradeableYr) {
          return jsonResponse({ ok: false, error: "Completed draft year " + dtYr + " is locked (view-only)." });
        }
        var currOwner = getCurrentPickOwnerInState(baseState, dtYr, checkDt.round, checkDt.originalOwner);
        if (currOwner !== String(checkDt.fromOwner)) {
          return jsonResponse({
            ok: false,
            error: "Pick Conflict: " + checkDt.fromOwner + " no longer owns " + dtYr + " Rd " + checkDt.round + " (" + checkDt.originalOwner + "'s Pick)."
          });
        }
      }

      var ptsToCheck = Array.isArray(tr.playerTransfers) ? tr.playerTransfers : [];
      for (var pIdxCheck = 0; pIdxCheck < ptsToCheck.length; pIdxCheck++) {
        var checkPt = ptsToCheck[pIdxCheck];
        var checkPtName = checkPt.playerName || checkPt.name || checkPt.playerId;
        if (!doesTeamOwnPlayerInState(baseState, checkPt.fromOwner, checkPt.playerId, checkPtName)) {
          return jsonResponse({
            ok: false,
            error: "Player Conflict: " + checkPt.fromOwner + " no longer owns " + checkPtName + "."
          });
        }
      }

      // If this submission supersedes an earlier proposal (Counter-Offer), remove the superseded proposal atomically
      var supersededId = String(tr.supersededProposalId || body.supersededProposalId || "").trim();
      if (supersededId) {
        baseState.pendingTrades = baseState.pendingTrades.filter(function(ep) {
          return String(ep.id) !== supersededId;
        });
      }

      var isProposal = (action === "PROPOSE_MEMBER_TRADE") || Boolean(body.requirePartnerConfirmation);
      if (isProposal) {
        var propId = String(tr.id || ("PROP-" + primaryYear + "-" + String(Date.now()).slice(-5)));
        var cleanProp = sanitizeIncomingTradeObject(tr, primaryYear, propId);
        cleanProp.status = "PENDING_CONFIRMATION";
        baseState.pendingTrades.push(cleanProp);
        baseState.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseState, (supersededId ? "COUNTER_PROPOSAL:" : "TRADE_PROPOSAL:") + propId);
        return jsonResponse({ ok: true, assignedId: propId, state: sanitizeStateForClient(baseState) });
      }

      var assignedId = assignOfficialTradeId(baseState, primaryYear);
      var cleanTrade = sanitizeIncomingTradeObject(tr, primaryYear, assignedId);
      baseState.trades.push(cleanTrade);
      applyConfirmedPlayerTransfersToRosters(baseState, cleanTrade);
      pruneConflictingPendingProposals(baseState);
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
        // Verify Double-Booked Pick & Player Conflict Guards before finalizing!
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

        var propPts = Array.isArray(propObj.playerTransfers) ? propObj.playerTransfers : [];
        for (var vpc = 0; vpc < propPts.length; vpc++) {
          var ptItem = propPts[vpc];
          var ptLabel = ptItem.playerName || ptItem.name || ptItem.playerId;
          if (!doesTeamOwnPlayerInState(baseStateResp, ptItem.fromOwner, ptItem.playerId, ptLabel)) {
            baseStateResp.pendingTrades.splice(foundIdx, 1);
            baseStateResp.updatedAt = Date.now();
            saveAndMirrorToSheets(sheets, baseStateResp, "AUTO_VOID_PLAYER_CONFLICT:" + propIdTarget);
            return jsonResponse({
              ok: false,
              error: "Conflict Guard: Player " + ptLabel + " is no longer on " + ptItem.fromOwner + "'s roster. Proposal voided.",
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

        // Auto-prune any remaining open proposals whose directTransfers or playerTransfers now conflict
        pruneConflictingPendingProposals(baseStateResp);

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

      // Enforce Champion Rename Decree lock on Phallic Feaster's customTeamName
      var isCommishCaller = Boolean(body.commishAuthHash && String(body.commishAuthHash).toLowerCase() === getExpectedCommishHash(stored));
      var feasterLocked = Boolean(existingProf.feasterRenamedByChamp);
      var resolvedCustomName = (feasterLocked && !isCommishCaller)
        ? String(existingProf.customTeamName || "").trim().slice(0, 24)
        : String(incoming.customTeamName || "").trim().slice(0, 24);

      baseStateProf.teamProfiles[teamId] = {
        tradeBlockNote: String(incoming.tradeBlockNote || "").slice(0, 280),
        lookingFor: String(incoming.lookingFor || "").slice(0, 180),
        needsChips: sanitizedChips,
        pickTags: sanitizedTags,
        playerTags: incoming.playerTags && typeof incoming.playerTags === "object" ? sanitizedPlayerTags : (existingProf.playerTags || {}),
        keepers: sanitizedKeepers,
        keepersConfirmedAt: Number(incoming.keepersConfirmedAt || existingProf.keepersConfirmedAt || 0),
        customTeamName: resolvedCustomName,
        feasterRenamedByChamp: feasterLocked,
        feasterRenamedBy: String(existingProf.feasterRenamedBy || "").slice(0, 30),
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
    // ACTION 1E: OVERALL CHAMPION RENAMES THE PHALLIC FEASTER (TOILET BOWL LOSER)
    // ------------------------------------------------------------------------
    if (action === "CHAMP_RENAME_FEASTER") {
      var baseStateChamp = stored || body.fallbackBaseState;
      if (!baseStateChamp || !Array.isArray(baseStateChamp.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      var targetTid = String(body.targetTeamId || "");
      var actorTid = String(body.actorTeamId || "");
      var newFeasterName = String(body.customTeamName || "").trim().slice(0, 24);
      var lockFlag = body.lockRename !== false;
      var validTarget = baseStateChamp.teams.some(function(t) { return t.id === targetTid; });
      if (!validTarget) {
        return jsonResponse({ ok: false, error: "Invalid target franchise." });
      }

      var isCommishDecree = Boolean(body.commishAuthHash && String(body.commishAuthHash).toLowerCase() === getExpectedCommishHash(stored));
      var isAuthorizedChamp = isCommishDecree;
      if (!isAuthorizedChamp && actorTid) {
        var edChamp = baseStateChamp.espnData || {};
        var seasonsMap = edChamp.seasons || {};
        var sYrs = Object.keys(seasonsMap).map(Number).filter(Boolean).sort(function(a, b) { return b - a; });
        for (var syi = 0; syi < sYrs.length; syi++) {
          var sEntry = seasonsMap[String(sYrs[syi])];
          if (sEntry && sEntry.champion) {
            if (String(sEntry.champion) === actorTid) isAuthorizedChamp = true;
            break;
          }
        }
        if (!isAuthorizedChamp && baseStateChamp.teamProfiles && baseStateChamp.teamProfiles[actorTid] && baseStateChamp.teamProfiles[actorTid].flairBadge === "CHAMP") {
          isAuthorizedChamp = true;
        }
      }
      if (!isAuthorizedChamp) {
        return jsonResponse({ ok: false, error: "Only the reigning Overall Champion or Commissioner can rename the Phallic Feaster." });
      }

      if (!baseStateChamp.teamProfiles || typeof baseStateChamp.teamProfiles !== "object") {
        baseStateChamp.teamProfiles = {};
      }
      var targetProf = baseStateChamp.teamProfiles[targetTid] || {};
      if (lockFlag && newFeasterName) {
        targetProf.customTeamName = newFeasterName;
        targetProf.feasterRenamedByChamp = true;
        targetProf.feasterRenamedBy = actorTid || "commish";
        if (!targetProf.flairBadge) targetProf.flairBadge = "FEASTER";
        if (!Array.isArray(baseStateChamp.hallOfShame)) baseStateChamp.hallOfShame = [];
        var fObj = baseStateChamp.teams.filter(function(t) { return t.id === targetTid; })[0];
        var cObj = baseStateChamp.teams.filter(function(t) { return t.id === actorTid; })[0];
        baseStateChamp.hallOfShame.unshift({
          id: "FEASTER-" + Date.now(),
          teamId: targetTid,
          teamName: fObj ? fObj.name : targetTid,
          message: (cObj ? cObj.name : "The Overall Champion") + " exercised the Champion's Decree and renamed Phallic Feaster " + (fObj ? fObj.name : targetTid) + " to \"" + newFeasterName + "\"!",
          timestamp: Date.now(),
          dateStr: new Date().toLocaleDateString()
        });
        baseStateChamp.hallOfShame = baseStateChamp.hallOfShame.slice(0, 30);
      } else {
        targetProf.feasterRenamedByChamp = false;
        targetProf.feasterRenamedBy = "";
      }
      targetProf.updatedAt = Date.now();
      baseStateChamp.teamProfiles[targetTid] = targetProf;
      baseStateChamp.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateChamp, "CHAMP_RENAME_FEASTER:" + targetTid);
      return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateChamp) });
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
        leagueScoring: (incomingEspn.leagueScoring && typeof incomingEspn.leagueScoring === "object")
          ? incomingEspn.leagueScoring
          : (existingEspn.leagueScoring || null),
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
        seasons: (incomingEspn.seasons && typeof incomingEspn.seasons === "object")
          ? incomingEspn.seasons
          : (existingEspn.seasons || {}),
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
      // Auto-void open pending proposals if a staged player was dropped/moved on ESPN
      pruneConflictingPendingProposals(baseStateEspn);
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

      var syncReason = String(body.reason || "ADMIN_SYNC");
      var isFullReset = syncReason.indexOf("RESET") === 0 || syncReason.indexOf("IMPORT") === 0;

      // Preserve concurrent manager teamProfile saves (e.g., keepers/avatar/PIN) if newer than incoming snapshot
      if (!isFullReset && stored && stored.teamProfiles && typeof stored.teamProfiles === "object") {
        if (!newState.teamProfiles || typeof newState.teamProfiles !== "object") newState.teamProfiles = {};
        Object.keys(stored.teamProfiles).forEach(function(tid) {
          var sProf = stored.teamProfiles[tid];
          var nProf = newState.teamProfiles[tid];
          if (!sProf) return;
          if (!nProf || (Number(sProf.updatedAt) || 0) > (Number(nProf.updatedAt) || 0)) {
            newState.teamProfiles[tid] = sProf;
          } else if (sProf.teamPinHash && !nProf.teamPinHash && nProf.teamPinVerifier === sProf.teamPinVerifier) {
            // Preserve server-only teamPinHash when client sanitized copy didn't change PIN
            nProf.teamPinHash = sProf.teamPinHash;
          }
        });
      }

      // Preserve concurrent pending trade proposals submitted while Commissioner modal was open
      if (!isFullReset && stored && Array.isArray(stored.pendingTrades) && stored.pendingTrades.length > 0) {
        if (!Array.isArray(newState.pendingTrades)) newState.pendingTrades = [];
        var baseSnapshotTs = Number(body.clientSnapshotUpdatedAt || newState.updatedAt || 0);
        var existingPendingIds = {};
        newState.pendingTrades.forEach(function(pt) { if (pt && pt.id) existingPendingIds[String(pt.id)] = true; });
        stored.pendingTrades.forEach(function(spt) {
          if (spt && spt.id && !existingPendingIds[String(spt.id)] && Number(spt.createdAt || 0) > baseSnapshotTs) {
            newState.pendingTrades.push(spt);
          }
        });
        pruneConflictingPendingProposals(newState);
      }

      if (!newState.commishPinHash && stored && stored.commishPinHash) {
        newState.commishPinHash = stored.commishPinHash;
      }
      if (!newState.leaguePassphraseHash && stored && stored.leaguePassphraseHash) {
        newState.leaguePassphraseHash = stored.leaguePassphraseHash;
      }
      delete newState.cloudSyncUrl;

      newState.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, newState, syncReason);
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

    var ed = currentState.espnData;
    var now = new Date();
    // NFL fantasy seasons run Sep–Jan; Jan/Feb belong to prior calendar year's NFL season unless ESPN already rolled over
    var targetYear = (now.getMonth() <= 1 && Number(ed.currentSeason)) ? Number(ed.currentSeason) : now.getFullYear();
    var url = "https://lm-api-reads.fantasy.espn.com/apis/v3/games/ffl/seasons/" +
      targetYear + "/segments/0/leagues/" + leagueId +
      "?view=mTeam&view=mRoster&view=mStandings&view=mMatchupScore&view=mSettings";

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

    var seasonYr = Number(data.seasonId) || targetYear;
    ed.currentSeason = seasonYr;
    if (data.settings && data.settings.name) {
      ed.leagueName = String(data.settings.name).slice(0, 80);
    }

    // 1. Parse League Scoring & Lineup Settings for Dynamic Trade Value Calibration
    if (data.settings) {
      var slotCounts = (data.settings.rosterSettings && data.settings.rosterSettings.lineupSlotCounts) || {};
      var sItems = (data.settings.scoringSettings && Array.isArray(data.settings.scoringSettings.scoringItems))
        ? data.settings.scoringSettings.scoringItems
        : [];
      var statPts = {};
      for (var si = 0; si < sItems.length; si++) {
        var itm = sItems[si];
        if (itm && itm.statId !== undefined) {
          statPts[String(itm.statId)] = Number(itm.pointsOverrides && itm.pointsOverrides["16"] !== undefined ? itm.pointsOverrides["16"] : itm.points) || 0;
        }
      }
      var qbCnt = Number(slotCounts["0"] || 1);
      var opCnt = Number(slotCounts["7"] || 0);
      var idpCnt = Number(slotCounts["8"] || 0) + Number(slotCounts["9"] || 0) + Number(slotCounts["10"] || 0) +
                   Number(slotCounts["11"] || 0) + Number(slotCounts["12"] || 0) + Number(slotCounts["13"] || 0) +
                   Number(slotCounts["14"] || 0) + Number(slotCounts["15"] || 0);
      ed.leagueScoring = {
        ppr: statPts["53"] !== undefined ? statPts["53"] : 1,
        pprPoints: statPts["53"] !== undefined ? statPts["53"] : 1,
        passTdPts: statPts["4"] !== undefined ? statPts["4"] : 4,
        passTdPoints: statPts["4"] !== undefined ? statPts["4"] : 4,
        qbSlots: qbCnt,
        qbCount: qbCnt,
        opSlots: opCnt,
        opCount: opCnt,
        idpSlots: idpCnt,
        idpSlotCount: idpCnt,
        tacklePts: statPts["109"] || statPts["108"] || 1.5,
        tacklePoints: statPts["109"] || statPts["108"] || 1.5,
        sackPts: statPts["99"] || 4,
        sackPoints: statPts["99"] || 4,
        syncedFromEspn: true,
        lineupSlots: slotCounts
      };
    }

    var slotMap = ed.teamSlotMap || {};
    var posMap = { 1: "QB", 2: "RB", 3: "WR", 4: "TE", 5: "K", 9: "DL", 10: "DL", 11: "LB", 12: "DB", 13: "DB", 14: "DB", 16: "D/ST" };
    var slotNameMap = { 0: "QB", 2: "RB", 4: "WR", 6: "TE", 7: "OP", 8: "DT", 9: "DE", 10: "LB", 11: "DL", 12: "CB", 13: "S", 14: "DB", 15: "DP", 16: "D/ST", 17: "K", 20: "BE", 21: "IR", 23: "FLEX" };
    var teamNflMap = { 0:"FA",1:"ATL",2:"BUF",3:"CHI",4:"CIN",5:"CLE",6:"DAL",7:"DEN",8:"DET",9:"GB",10:"TEN",11:"IND",12:"KC",13:"LV",14:"LAR",15:"MIA",16:"MIN",17:"NE",18:"NO",19:"NYG",20:"NYJ",21:"PHI",22:"ARI",23:"PIT",24:"LAC",25:"SF",26:"SEA",27:"TB",28:"WAS",29:"CAR",30:"JAX",33:"BAL",34:"HOU" };

    // 2. Sync Live Rosters & Positional Ranks
    ed.rosters = ed.rosters || {};
    var allLeaguePlayersByPos = {};
    data.teams.forEach(function(t) {
      var slotKey = String(t.id);
      var mgrId = slotMap[slotKey];
      if (!mgrId) return;
      var entries = (t.roster && Array.isArray(t.roster.entries)) ? t.roster.entries : [];
      if (entries.length > 0) {
        var mappedRoster = entries.map(function(e) {
          var p = (e.playerPoolEntry && e.playerPoolEntry.player) || {};
          var pid = p.id || e.playerId;
          var pos = posMap[p.defaultPositionId] || "FLEX";
          var pts = Math.round(Number((e.playerPoolEntry && e.playerPoolEntry.appliedStatTotal) || 0) * 10) / 10;
          var plObj = {
            id: "espn_" + pid,
            espnId: pid,
            name: p.fullName || "Unknown Player",
            pos: pos,
            nflTeam: teamNflMap[p.proTeamId] || "NFL",
            slot: slotNameMap[e.lineupSlotId] || "BE",
            pts: pts,
            totalPts: pts,
            posRank: 0,
            acquiredVia: e.acquisitionType || "ROSTER"
          };
          if (!allLeaguePlayersByPos[pos]) allLeaguePlayersByPos[pos] = [];
          allLeaguePlayersByPos[pos].push(plObj);
          return plObj;
        });
        ed.rosters[mgrId] = mappedRoster;
      }
    });
    Object.keys(allLeaguePlayersByPos).forEach(function(posKey) {
      allLeaguePlayersByPos[posKey].sort(function(a, b) { return (b.totalPts || 0) - (a.totalPts || 0); });
      allLeaguePlayersByPos[posKey].forEach(function(pl, idx) { pl.posRank = idx + 1; });
    });

    // 3. Sync Weekly Matchups (All-Play, H2H Rivalries, Single-Week Records & Toilet Bowl)
    var regWeeks = Number(data.settings && data.settings.scheduleSettings && data.settings.scheduleSettings.matchupPeriodCount) || 14;
    var playoffTeamCnt = Number(data.settings && data.settings.scheduleSettings && data.settings.scheduleSettings.playoffTeamCount) || 5;
    var currentSp = Number(data.scoringPeriodId || (data.status && data.status.currentMatchupPeriod) || 0);
    var finalSp = Number((data.status && data.status.finalScoringPeriod) || 17);

    var allPlayBySlot = {};
    var toiletStatsBySlot = {};
    data.teams.forEach(function(t) {
      allPlayBySlot[String(t.id)] = { w: 0, l: 0 };
      toiletStatsBySlot[String(t.id)] = { wins: 0, pts: 0 };
    });

    if (!ed.headToHead || typeof ed.headToHead !== "object") ed.headToHead = {};
    if (!ed.syncedMatchupKeys || typeof ed.syncedMatchupKeys !== "object") ed.syncedMatchupKeys = {};
    if (!ed.records || typeof ed.records !== "object") ed.records = {};

    var schedule = Array.isArray(data.schedule) ? data.schedule : [];
    var scoresByWeek = {};
    schedule.forEach(function(m) {
      if (!m || !m.home || !m.away) return;
      var wk = Number(m.matchupPeriodId) || 0;
      var hScore = Number(m.home.totalPoints || 0);
      var aScore = Number(m.away.totalPoints || 0);
      if (wk <= 0 || (hScore === 0 && aScore === 0)) return;
      if (m.winner === "UNDECIDED") return;

      var hSlot = String(m.home.teamId);
      var aSlot = String(m.away.teamId);
      var hMgr = slotMap[hSlot] || "";
      var aMgr = slotMap[aSlot] || "";

      if (wk <= regWeeks) {
        if (!scoresByWeek[wk]) scoresByWeek[wk] = [];
        scoresByWeek[wk].push({ slot: hSlot, score: hScore });
        scoresByWeek[wk].push({ slot: aSlot, score: aScore });
      } else {
        // Postseason Consolation / Toilet Bowl tracking for non-playoff seeds
        var tier = String(m.playoffTierType || "");
        if (tier.indexOf("LOSERS") !== -1 || tier.indexOf("CONSOLATION") !== -1) {
          if (toiletStatsBySlot[hSlot]) {
            toiletStatsBySlot[hSlot].pts += hScore;
            if (hScore > aScore) toiletStatsBySlot[hSlot].wins += 1;
          }
          if (toiletStatsBySlot[aSlot]) {
            toiletStatsBySlot[aSlot].pts += aScore;
            if (aScore > hScore) toiletStatsBySlot[aSlot].wins += 1;
          }
        }
      }

      // Deduplicated Head-to-Head & Single-Week Record Updates
      var mKey = seasonYr + "-W" + wk + "-" + hSlot + "-" + aSlot;
      if (hMgr && aMgr && hMgr !== aMgr && !ed.syncedMatchupKeys[mKey]) {
        ed.syncedMatchupKeys[mKey] = 1;
        var pairKey = [hMgr, aMgr].sort().join("__");
        var rec = ed.headToHead[pairKey] || { a: [hMgr, aMgr].sort()[0], b: [hMgr, aMgr].sort()[1], aWins: 0, bWins: 0, ties: 0, aPts: 0, bPts: 0 };
        if (hMgr === rec.a) {
          rec.aPts = Math.round((Number(rec.aPts || 0) + hScore) * 10) / 10;
          rec.bPts = Math.round((Number(rec.bPts || 0) + aScore) * 10) / 10;
          if (hScore > aScore) rec.aWins += 1;
          else if (aScore > hScore) rec.bWins += 1;
          else rec.ties += 1;
        } else {
          rec.bPts = Math.round((Number(rec.bPts || 0) + hScore) * 10) / 10;
          rec.aPts = Math.round((Number(rec.aPts || 0) + aScore) * 10) / 10;
          if (hScore > aScore) rec.bWins += 1;
          else if (aScore > hScore) rec.aWins += 1;
          else rec.ties += 1;
        }
        ed.headToHead[pairKey] = rec;

        var checkScorePairs = [{ mgr: hMgr, sc: hScore, opp: aMgr, oppSc: aScore }, { mgr: aMgr, sc: aScore, opp: hMgr, oppSc: hScore }];
        checkScorePairs.forEach(function(sp) {
          if (!ed.records.highestWeekScore || sp.sc > Number(ed.records.highestWeekScore.score || 0)) {
            ed.records.highestWeekScore = { managerId: sp.mgr, score: Math.round(sp.sc * 100) / 100, year: seasonYr, week: wk, opponentId: sp.opp };
          }
          if (sp.sc > 0 && (!ed.records.lowestWeekScore || !ed.records.lowestWeekScore.score || sp.sc < Number(ed.records.lowestWeekScore.score))) {
            ed.records.lowestWeekScore = { managerId: sp.mgr, score: Math.round(sp.sc * 100) / 100, year: seasonYr, week: wk, opponentId: sp.opp };
          }
        });
        var margin = Math.round(Math.abs(hScore - aScore) * 100) / 100;
        if (margin > 0) {
          var winMgr = hScore > aScore ? hMgr : aMgr;
          var loseMgr = hScore > aScore ? aMgr : hMgr;
          if (!ed.records.biggestBlowout || margin > Number(ed.records.biggestBlowout.margin || 0)) {
            ed.records.biggestBlowout = { winnerId: winMgr, loserId: loseMgr, margin: margin, year: seasonYr, week: wk };
          }
          if (!ed.records.closestWin || !ed.records.closestWin.margin || margin < Number(ed.records.closestWin.margin)) {
            ed.records.closestWin = { winnerId: winMgr, loserId: loseMgr, margin: margin, year: seasonYr, week: wk };
          }
        }
      }
    });

    Object.keys(scoresByWeek).forEach(function(wkKey) {
      var list = scoresByWeek[wkKey];
      for (var i = 0; i < list.length; i++) {
        for (var j = 0; j < list.length; j++) {
          if (i === j) continue;
          if (list[i].score > list[j].score) allPlayBySlot[list[i].slot].w += 1;
          else if (list[i].score < list[j].score) allPlayBySlot[list[i].slot].l += 1;
        }
      }
    });

    // 4. Build Live Season Standings & Official Completed-Season Gate
    var maxGamesPlayed = 0;
    var rawTeamRows = data.teams.map(function(t) {
      var slotKey = String(t.id);
      var mgrId = slotMap[slotKey] || "";
      var recOverall = (t.record && t.record.overall) || {};
      var w = Number(recOverall.wins || 0);
      var l = Number(recOverall.losses || 0);
      var ties = Number(recOverall.ties || 0);
      if (w + l + ties > maxGamesPlayed) maxGamesPlayed = w + l + ties;
      var pf = Math.round(Number(recOverall.pointsFor || 0) * 10) / 10;
      var pa = Math.round(Number(recOverall.pointsAgainst || 0) * 10) / 10;
      var seed = Number(t.playoffSeed || 0);
      var finalRank = Number(t.rankCalculatedFinal || 0);
      var ap = allPlayBySlot[slotKey] || { w: 0, l: 0 };
      var tb = toiletStatsBySlot[slotKey] || { wins: 0, pts: 0 };
      var tNameFull = ((t.location || "") + " " + (t.nickname || "")).trim() || (t.name || ("Slot " + slotKey));
      return {
        espnTeamId: Number(t.id),
        managerId: mgrId,
        currentManagerId: mgrId,
        teamName: tNameFull,
        wins: w,
        losses: l,
        ties: ties,
        pf: pf,
        pa: pa,
        pointsFor: pf,
        pointsAgainst: pa,
        playoffSeed: seed,
        rank: finalRank || seed || 0,
        finalStanding: finalRank || seed || 0,
        allPlayWins: ap.w,
        allPlayLosses: ap.l,
        toiletWins: tb.wins,
        toiletPoints: Math.round(tb.pts * 10) / 10,
        isChampion: false
      };
    });

    rawTeamRows.sort(function(a, b) {
      var ra = a.rank || 99;
      var rb = b.rank || 99;
      if (ra !== rb) return ra - rb;
      if (b.wins !== a.wins) return b.wins - a.wins;
      return b.pf - a.pf;
    });
    rawTeamRows.forEach(function(r, idx) {
      if (!r.rank) r.rank = idx + 1;
      if (!r.finalStanding) r.finalStanding = r.rank;
    });

    var hasFinalRanks = data.teams.some(function(t) { return Number(t.rankCalculatedFinal || 0) === 1; });
    var isOfficiallyCompleted = Boolean(hasFinalRanks && currentSp >= finalSp && maxGamesPlayed >= 13);

    var champId = "";
    var runnerUpId = "";
    var toiletWinnerId = "";
    var feasterId = "";
    if (isOfficiallyCompleted && rawTeamRows.length >= 2) {
      champId = rawTeamRows[0].managerId || "";
      runnerUpId = rawTeamRows[1].managerId || "";
      rawTeamRows[0].isChampion = true;
      var nonPlayoff = rawTeamRows.filter(function(r) {
        return r.playoffSeed > playoffTeamCnt || (!r.playoffSeed && r.rank > playoffTeamCnt);
      });
      if (nonPlayoff.length >= 2) {
        nonPlayoff.sort(function(a, b) {
          if (b.toiletWins !== a.toiletWins) return b.toiletWins - a.toiletWins;
          if (b.toiletPoints !== a.toiletPoints) return b.toiletPoints - a.toiletPoints;
          return a.rank - b.rank;
        });
        toiletWinnerId = nonPlayoff[0].managerId || "";
        feasterId = nonPlayoff[nonPlayoff.length - 1].managerId || "";
      }
    }

    if (!ed.seasons || typeof ed.seasons !== "object") ed.seasons = {};
    if (!ed.seasonStandings || typeof ed.seasonStandings !== "object") ed.seasonStandings = {};
    ed.seasons[String(seasonYr)] = {
      year: seasonYr,
      seasonCompleted: isOfficiallyCompleted,
      playoffTeamCount: playoffTeamCnt,
      champion: champId,
      runnerUp: runnerUpId,
      toiletBowlWinner: toiletWinnerId,
      phallicFeaster: feasterId,
      standings: rawTeamRows
    };
    ed.seasonStandings[String(seasonYr)] = rawTeamRows;

    // 5. Auto-void open trade proposals if a staged player left fromOwner's roster on waivers/ESPN
    if (Array.isArray(currentState.pendingTrades) && currentState.pendingTrades.length > 0) {
      currentState.pendingTrades = currentState.pendingTrades.filter(function(prop) {
        var pts = Array.isArray(prop.playerTransfers) ? prop.playerTransfers : [];
        for (var i = 0; i < pts.length; i++) {
          var pt = pts[i];
          var ownerRoster = Array.isArray(ed.rosters[pt.fromOwner]) ? ed.rosters[pt.fromOwner] : [];
          var pLabel = String(pt.playerName || pt.name || "").trim().toLowerCase();
          var stillOwned = ownerRoster.some(function(rp) {
            return (pt.playerId && String(rp.id) === String(pt.playerId)) ||
                   (pLabel && String(rp.name || "").trim().toLowerCase() === pLabel);
          });
          if (!stillOwned) return false;
        }
        return true;
      });
    }

    ed.lastSyncedAt = Date.now();
    currentState.updatedAt = Date.now();
    saveAndMirrorToSheets(sheets, currentState, "WEEKLY_ESPN_AUTO_SYNC");
  } catch (err) {
    Logger.log("autoWeeklyEspnSync error: " + err);
  } finally {
    lock.releaseLock();
  }
}

