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
        pools: cleanPools,
        conditionals: cleanConds,
        reactions: (tr.reactions && typeof tr.reactions === "object") ? tr.reactions : {}
      };
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
    // ACTION 1B: FRANCHISE MANAGER UPDATES THEIR OWN TEAM WAR ROOM / IDENTITY
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
      baseStateProf.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateProf, (!wasCursed && nowCursed) ? ("UPSTATE_KITTY_CURSE:" + teamId) : ("TEAM_WAR_ROOM:" + teamId));
      return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateProf) });
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
 * Writes canonical JSON + rolling 30-snapshot history + human-readable Google Sheet tabs!
 */
function saveAndMirrorToSheets(sheets, state, actionLabel) {
  var nowIso = new Date().toISOString();
  var jsonStr = JSON.stringify(state);

  // 1. Update LIVE_STATE_JSON
  sheets.stateSheet.getRange("B2:C2").setValues([[jsonStr, nowIso]]);
  sheets.stateSheet.insertRowBefore(5);
  sheets.stateSheet.getRange("A5:D5").setValues([[state.updatedAt || Date.now(), nowIso, safeSheetCell(actionLabel), jsonStr]]);
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
      safeSheetCell(tr.id || ""),
      tr.year || "",
      safeSheetCell(tr.date || ""),
      safeSheetCell(tr.submittedBy || "Commissioner"),
      safeSheetCell(mgrs),
      safeSheetCell(tr.headline || ""),
      safeSheetCell(directs),
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
}
