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
 *   6. (v13.0.0 Native Gemini AI) In Apps Script -> Project Settings -> Script Properties,
 *      add property `GEMINI_API_KEY` with your Google AI Studio Gemini API key.
 *      The key stays strictly server-side and is never returned to the browser.
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

function readRowChunkedJson(stateSheet, rowNum, primaryColNum, overflowStartColNum, maxOverflowCols) {
  var raw = stateSheet.getRange(rowNum, primaryColNum).getValue();
  if (!raw || typeof raw !== "string") return null;
  var joined = raw;
  if (raw.length >= CELL_CHUNK_SIZE) {
    try {
      var sheetMaxCols = stateSheet.getMaxColumns();
      var availCols = Math.max(0, sheetMaxCols - overflowStartColNum + 1);
      var colsToRead = Math.min(maxOverflowCols || 120, availCols);
      if (colsToRead > 0) {
        var extraRow = stateSheet.getRange(rowNum, overflowStartColNum, 1, colsToRead).getValues();
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

function readStoredState(stateSheet) {
  // Primary chunk in B2 (col 2), overflow chunks across D2+ (up to 120 cols ~ 5.3MB)
  return readRowChunkedJson(stateSheet, 2, 2, 4, 120);
}

function readSnapshotStateAtRow(stateSheet, rowNum) {
  // Snapshot rows 5..35 store primary chunk in D (col 4) and overflow chunks across E+ (up to 120 cols ~ 5.3MB)
  return readRowChunkedJson(stateSheet, rowNum, 4, 5, 120);
}

function listSnapshotsFromSheet(stateSheet) {
  var lastRow = stateSheet.getLastRow();
  if (lastRow < 5) return [];
  var numRows = Math.min(30, lastRow - 4);
  var metaVals = stateSheet.getRange(5, 1, numRows, 4).getValues();
  var out = [];
  for (var i = 0; i < metaVals.length; i++) {
    var r = metaVals[i];
    var ts = Number(r[0]) || 0;
    var iso = String(r[1] || "");
    var act = String(r[2] || "");
    var snapPreview = String(r[3] || "");
    if (!ts && !iso && !snapPreview) continue;
    out.push({
      rowNumber: 5 + i,
      index: i,
      timestamp: ts,
      updatedIso: iso,
      actionLabel: act,
      sizeApprox: snapPreview.length
    });
  }
  return out;
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

function isAuthorizedCommishOrCoCommish(storedState, body) {
  var provided = body && body.commishAuthHash ? String(body.commishAuthHash).toLowerCase() : "";
  var actorTid = body && (body.actorTeamId || body.coCommishTeamId) ? String(body.actorTeamId || body.coCommishTeamId).toLowerCase().trim() : "";
  if (!provided || !/^[a-f0-9]{64}$/i.test(provided)) {
    return { authorized: false, isMasterCommish: false, isCoCommish: false, actorTeamId: actorTid };
  }
  if (provided === getExpectedCommishHash(storedState)) {
    return { authorized: true, isMasterCommish: true, isCoCommish: false, actorTeamId: actorTid || "syeth" };
  }
  var coList = (storedState && Array.isArray(storedState.coCommissioners)) ? storedState.coCommissioners : [];
  var coHashes = (storedState && storedState.coCommishPinHashes && typeof storedState.coCommishPinHashes === "object") ? storedState.coCommishPinHashes : {};
  if (actorTid && coList.indexOf(actorTid) !== -1 && coHashes[actorTid] && String(coHashes[actorTid]).toLowerCase() === provided) {
    return { authorized: true, isMasterCommish: false, isCoCommish: true, actorTeamId: actorTid };
  }
  // Fallback scan if actorTeamId wasn't explicitly passed on a legacy payload
  for (var i = 0; i < coList.length; i++) {
    var cid = String(coList[i] || "").toLowerCase();
    if (cid && coHashes[cid] && String(coHashes[cid]).toLowerCase() === provided) {
      return { authorized: true, isMasterCommish: false, isCoCommish: true, actorTeamId: cid };
    }
  }
  return { authorized: false, isMasterCommish: false, isCoCommish: false, actorTeamId: actorTid };
}

function normalizeStrategyModeTag(rawTag) {
  var up = String(rawTag || "").trim().toUpperCase();
  if (up === "CONTENDING" || up === "ALL_IN") return "CONTENDING";
  if (up === "HOLDING_SERVE" || up === "BALANCED") return "HOLDING_SERVE";
  if (up === "REBUILDING" || up === "RETOOLING") return "REBUILDING";
  return "";
}

function sanitizeStateForClient(stateObj) {
  if (!stateObj || typeof stateObj !== "object") return stateObj;
  var clone = JSON.parse(JSON.stringify(stateObj));
  delete clone.commishPinHash;
  delete clone.coCommishPinHashes;
  delete clone.leaguePassphraseHash;
  delete clone.cloudSyncUrl;
  delete clone.oneSignalRestKey;
  delete clone.geminiApiKey;
  delete clone.GEMINI_API_KEY;
  if (clone.settings && typeof clone.settings === "object") {
    delete clone.settings.geminiApiKey;
    delete clone.settings.GEMINI_API_KEY;
  }
  if (clone.pushConfig && typeof clone.pushConfig === "object") {
    delete clone.pushConfig.oneSignalRestKey;
    delete clone.pushConfig.restApiKey;
  }
  if (clone.teamProfiles && typeof clone.teamProfiles === "object") {
    Object.keys(clone.teamProfiles).forEach(function(tid) {
      if (clone.teamProfiles[tid]) {
        delete clone.teamProfiles[tid].teamPinHash;
      }
    });
  }
  clone.hasOneSignalRestKey = Boolean(
    (stateObj.oneSignalRestKey && String(stateObj.oneSignalRestKey).trim()) ||
    (stateObj.pushConfig && stateObj.pushConfig.oneSignalRestKey && String(stateObj.pushConfig.oneSignalRestKey).trim())
  );
  var hasGemKey = false;
  try {
    var sProps = PropertiesService.getScriptProperties();
    hasGemKey = Boolean(sProps && sProps.getProperty("GEMINI_API_KEY") && String(sProps.getProperty("GEMINI_API_KEY")).trim());
  } catch (ignoreGemPropErr) {}
  clone.hasGeminiApiKey = hasGemKey;
  return clone;
}

/**
 * Dispatches server-side Web Push via OneSignal REST API.
 * Enforces 10:00 PM – 8:00 AM EST quiet hours on non-urgent broadcasts while delivering live trade offers immediately.
 */
function dispatchOneSignalPush(stateObj, opts) {
  try {
    if (!stateObj || !opts) return { sent: false, reason: "NO_OPTS" };
    var props = null;
    try { props = PropertiesService.getScriptProperties(); } catch (ignorePropErr) {}
    var appId = String(
      (stateObj.pushConfig && stateObj.pushConfig.oneSignalAppId) ||
      stateObj.oneSignalAppId ||
      (props && props.getProperty("ONESIGNAL_APP_ID")) ||
      ""
    ).trim();
    var restKey = String(
      (stateObj.pushConfig && stateObj.pushConfig.oneSignalRestKey) ||
      stateObj.oneSignalRestKey ||
      (props && props.getProperty("ONESIGNAL_REST_KEY")) ||
      ""
    ).trim();
    if (!appId || !restKey) {
      return { sent: false, reason: "MISSING_ONESIGNAL_KEYS" };
    }

    var category = String(opts.category || "BROADCAST").toUpperCase(); // DIRECT_TRADE | CONDITIONAL | BROADCAST | HOF | KEEPER | CUSTOM
    var isUrgentDirect = Boolean(opts.urgent || category === "DIRECT_TRADE" || category === "CONDITIONAL");
    var targetTeamIds = Array.isArray(opts.targetTeamIds) ? opts.targetTeamIds : null;

    // Collect opted-in player/subscription IDs from teamProfiles[tid].pushPrefs
    var profiles = (stateObj.teamProfiles && typeof stateObj.teamProfiles === "object") ? stateObj.teamProfiles : {};
    var subIds = [];
    Object.keys(profiles).forEach(function(tid) {
      if (targetTeamIds && targetTeamIds.indexOf(tid) === -1) return;
      var pf = profiles[tid] && profiles[tid].pushPrefs;
      if (!pf || !pf.enabled || !pf.subscriptionId) return;
      if (category === "DIRECT_TRADE" && pf.directTradeAlerts === false) return;
      if (category === "CONDITIONAL" && pf.conditionalAlerts === false) return;
      if (category === "BROADCAST" && pf.tradeBroadcasts === false) return;
      if (category === "HOF" && pf.hofAlerts === false) return;
      if (category === "KEEPER" && pf.keeperAlerts === false) return;
      var sid = String(pf.subscriptionId).trim();
      if (sid && subIds.indexOf(sid) === -1) subIds.push(sid);
    });

    // Check EST Quiet Hours (10:00 PM – 8:00 AM America/New_York) for non-urgent broadcasts
    var sendAfterIso = null;
    if (!isUrgentDirect) {
      try {
        var estHourStr = Utilities.formatDate(new Date(), "America/New_York", "H");
        var estHour = parseInt(estHourStr, 10);
        if (!isNaN(estHour) && (estHour >= 22 || estHour < 8)) {
          var delayHours = estHour >= 22 ? (24 - estHour + 8) : (8 - estHour);
          var releaseDate = new Date(Date.now() + delayHours * 3600 * 1000);
          releaseDate.setMinutes(0, 0, 0);
          sendAfterIso = releaseDate.toUTCString();
        }
      } catch (tzErr) {}
    }

    var payload = {
      app_id: appId,
      headings: { en: String(opts.title || "Gamecock Mafia Tracker").slice(0, 100) },
      contents: { en: String(opts.message || "").slice(0, 260) },
      url: String(opts.url || "https://syeth-spurrier.github.io/gamecock-mafia-draft-tracker/")
    };
    if (subIds.length > 0) {
      payload.include_subscription_ids = subIds;
    } else if (!targetTeamIds) {
      payload.included_segments = ["Subscribed Users"];
    } else {
      return { sent: false, reason: "NO_SUBSCRIBED_TARGETS" };
    }
    if (sendAfterIso) {
      payload.send_after = sendAfterIso;
    }

    var resp = UrlFetchApp.fetch("https://onesignal.com/api/v1/notifications", {
      method: "post",
      contentType: "application/json; charset=utf-8",
      headers: {
        Authorization: "Basic " + restKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return {
      sent: resp.getResponseCode() >= 200 && resp.getResponseCode() < 300,
      statusCode: resp.getResponseCode(),
      queuedForMorning: Boolean(sendAfterIso),
      sendAfter: sendAfterIso || null
    };
  } catch (pushErr) {
    return { sent: false, error: String(pushErr) };
  }
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

function appendServerActivityEvent(stateObj, evtType, title, detail, meta) {
  if (!stateObj || typeof stateObj !== "object") return;
  if (!Array.isArray(stateObj.activityLog)) stateObj.activityLog = [];
  var nowMs = Date.now();
  var entry = {
    id: "ACT-" + evtType + "-" + nowMs + "-" + Math.floor(Math.random() * 1000),
    type: String(evtType || "LEAGUE_EVENT").slice(0, 40),
    title: String(title || "").slice(0, 180),
    detail: String(detail || "").slice(0, 360),
    year: Number((meta && meta.year) || stateObj.activeYear || stateObj.baseUpcomingYear || 2026),
    teamId: meta && meta.teamId ? String(meta.teamId).slice(0, 30) : "",
    actorTeamId: meta && meta.actorTeamId ? String(meta.actorTeamId).slice(0, 30) : "",
    timestamp: nowMs,
    dateStr: new Date(nowMs).toLocaleDateString()
  };
  stateObj.activityLog.unshift(entry);
  if (stateObj.activityLog.length > 100) {
    stateObj.activityLog = stateObj.activityLog.slice(0, 100);
  }
}

/* ============================================================================
 * WAVE 4 (v13.0.0) NATIVE GEMINI AI ENGINE — TONE PRESETS, ANTI-SLOP & CACHE
 * ============================================================================ */
var GEMINI_MODEL_NAME = "gemini-2.5-flash";
var AI_CACHE_MAX_BYTES = 35000;
var AI_MEMBER_HOURLY_LIMIT = 25;

var AI_TONE_PRESETS = {
  DRY_BEAT: {
    name: "Dry Beat Writer",
    directive: "Write like a deadpan, veteran beat reporter covering a ruthless 10-man keeper league as if it were the real NFL wire. Deliver dry, raised-eyebrow sarcasm buried inside crisp, matter-of-fact journalism. No caps-lock, no cheerleading, no fluff."
  },
  SNARKY_COMMISH: {
    name: "Snarky Commish",
    directive: "Write like a sharp, dryly sarcastic league insider who keeps the spreadsheet receipts open in another tab. Prioritize clever understatement, ironic observations, and surgical factual burns (bench blunders, FAAB overpays, All-Play luck frauds) over loud banter. Dry wit > yelling."
  },
  RUTHLESS_ROAST: {
    name: "Ruthless Roast",
    directive: "Write in sharp, unfiltered league group-chat roast mode. Go straight for the jugular on weekly low scores, lopsided trades, catastrophic benching decisions, and playoff collapses using biting, surgical comedy. Stay 100% anchored to real receipts—never resort to generic hype-bro clichés."
  }
};

var AI_ANTI_SLOP_SYSTEM_BLOCK = [
  "STRICT EDITORIAL & ANTI-SLOP CONTRACT (MANDATORY):",
  "1. FORBIDDEN CLICHES — NEVER USE ANY OF THESE PHRASES OR VARIANTS: \"What a wild week\", \"Strap in\", \"Buckle up\", \"Without further ado\", \"Let's dive in\", \"Rollercoaster\", \"Statement win\", \"Absolute cinema\", \"Name a more iconic duo\", \"Chef's kiss\", \"Fantasy football is a cruel mistress\", \"Only time will tell\", \"At the end of the day\".",
  "2. FORMATTING: Maximum ONE (1) exclamation mark across the entire JSON response. ZERO emoji walls (do not put emojis in prose strings). Short, punchy sentences. Zero chatty preambles, zero throat-clearing intros, zero cheesy sign-offs.",
  "3. STRICT FACT GROUNDING: Every score, player name, player stat, FAAB dollar, record, division, seed, and magic number MUST appear verbatim in the supplied JSON payload. If a stat is missing, omit it—NEVER hallucinate NFL injuries, weather, or game scripts not present in the JSON."
].join("\n");

var POS_KING_BUCKETS_GAS = ["QB", "RB", "WR", "TE", "DL", "LB", "DB", "D/ST", "K"];

function normalizePosBucketGas_(rawPos, defaultPosId) {
  var pid = Number(defaultPosId);
  if (pid === 1) return "QB";
  if (pid === 2) return "RB";
  if (pid === 3) return "WR";
  if (pid === 4) return "TE";
  if (pid === 5) return "K";
  if (pid === 16) return "D/ST";
  if (pid === 9 || pid === 11) return "DL";
  if (pid === 10) return "LB";
  if (pid === 12 || pid === 13 || pid === 14) return "DB";
  var up = String(rawPos || "").toUpperCase().trim();
  if (up === "QB" || up === "RB" || up === "WR" || up === "TE" || up === "K") return up;
  if (up === "D/ST" || up === "DST" || up === "DEF") return "D/ST";
  if (up === "DL" || up === "DE" || up === "DT" || up === "EDR") return "DL";
  if (up === "LB") return "LB";
  if (up === "DB" || up === "CB" || up === "S" || up === "FS" || up === "SS") return "DB";
  return "";
}

function createEmptyPosKingsStore_(seasonYear) {
  var obj = { updatedAt: Date.now(), starters: {}, bench: {} };
  if (seasonYear) obj.seasonYear = Number(seasonYear);
  POS_KING_BUCKETS_GAS.forEach(function(b) {
    obj.starters[b] = [];
    obj.bench[b] = [];
  });
  return obj;
}

function mergeWeeklyPosKingsList_(existingArr, incomingArr, maxKeep) {
  var cap = maxKeep || 10;
  var byKey = {};
  (Array.isArray(existingArr) ? existingArr : []).concat(Array.isArray(incomingArr) ? incomingArr : []).forEach(function(row) {
    if (!row || !row.playerName || !(Number(row.pts) > 0)) return;
    var key = String(row.year || 0) + "_" + String(row.week || 0) + "_" + String(row.playerId || row.playerName) + "_" + String(row.teamId || "") + "_" + (row.isStarter ? "S" : "B");
    var cleanRow = {
      playerId: String(row.playerId || "").slice(0, 32),
      playerName: String(row.playerName || "").slice(0, 60),
      pos: String(row.pos || "").slice(0, 8),
      nflTeam: String(row.nflTeam || "NFL").slice(0, 8),
      teamId: String(row.teamId || "").slice(0, 32),
      year: Number(row.year) || 0,
      week: Number(row.week) || 0,
      pts: Math.round(Number(row.pts) * 100) / 100,
      slotId: Number(row.slotId !== undefined ? row.slotId : (row.isStarter ? 0 : 20)),
      isStarter: Boolean(row.isStarter)
    };
    if (!byKey[key] || cleanRow.pts > byKey[key].pts) {
      byKey[key] = cleanRow;
    }
  });
  var list = Object.keys(byKey).map(function(k) { return byKey[k]; });
  list.sort(function(a, b) {
    if (b.pts !== a.pts) return b.pts - a.pts;
    if (b.year !== a.year) return b.year - a.year;
    return b.week - a.week;
  });
  return list.slice(0, cap);
}

function sanitizeWeeklyPosKingsStore_(rawStore, fallbackSeasonYr) {
  var out = createEmptyPosKingsStore_(rawStore && rawStore.seasonYear ? rawStore.seasonYear : fallbackSeasonYr);
  if (rawStore && rawStore.startYear) out.startYear = Number(rawStore.startYear) || 2018;
  if (!rawStore || typeof rawStore !== "object") return out;
  var srcSt = (rawStore.starters && typeof rawStore.starters === "object") ? rawStore.starters : {};
  var srcBn = (rawStore.bench && typeof rawStore.bench === "object") ? rawStore.bench : {};
  POS_KING_BUCKETS_GAS.forEach(function(b) {
    out.starters[b] = mergeWeeklyPosKingsList_([], srcSt[b], 10);
    out.bench[b] = mergeWeeklyPosKingsList_([], srcBn[b], 10);
  });
  return out;
}

function mergeWeeklyPosKingsStores_(baseStore, incomingStore, replaceSameSeason) {
  var base = sanitizeWeeklyPosKingsStore_(baseStore, incomingStore && incomingStore.seasonYear);
  if (!incomingStore || typeof incomingStore !== "object") return base;
  var inc = sanitizeWeeklyPosKingsStore_(incomingStore, base.seasonYear);
  if (replaceSameSeason && inc.seasonYear && base.seasonYear && Number(inc.seasonYear) !== Number(base.seasonYear)) {
    return inc;
  }
  POS_KING_BUCKETS_GAS.forEach(function(b) {
    base.starters[b] = mergeWeeklyPosKingsList_(base.starters[b], inc.starters[b], 10);
    base.bench[b] = mergeWeeklyPosKingsList_(base.bench[b], inc.bench[b], 10);
  });
  base.updatedAt = Date.now();
  return base;
}

function ensureAiStructures_(stateObj) {
  if (!stateObj || typeof stateObj !== "object") return;
  if (!stateObj.settings || typeof stateObj.settings !== "object") {
    stateObj.settings = {};
  }
  var av = (stateObj.settings.aiVoice && typeof stateObj.settings.aiVoice === "object")
    ? stateObj.settings.aiVoice
    : {};
  var presetKey = String(av.preset || "SNARKY_COMMISH").toUpperCase();
  if (!AI_TONE_PRESETS[presetKey]) presetKey = "SNARKY_COMMISH";
  stateObj.settings.aiVoice = {
    preset: presetKey,
    customLore: String(av.customLore || "").slice(0, 1500),
    hofAiEnabled: Boolean(av.hofAiEnabled),
    updatedAt: Number(av.updatedAt) || 0
  };
  if (!stateObj.aiCache || typeof stateObj.aiCache !== "object") {
    stateObj.aiCache = {};
  }
  if (!stateObj.aiCache.weeklyRecaps || typeof stateObj.aiCache.weeklyRecaps !== "object") stateObj.aiCache.weeklyRecaps = {};
  if (!stateObj.aiCache.tradeGrades || typeof stateObj.aiCache.tradeGrades !== "object") stateObj.aiCache.tradeGrades = {};
  if (!stateObj.aiCache.playoffAdvisors || typeof stateObj.aiCache.playoffAdvisors !== "object") stateObj.aiCache.playoffAdvisors = {};
  if (!stateObj.aiCache.hofBios || typeof stateObj.aiCache.hofBios !== "object") stateObj.aiCache.hofBios = {};
}

function trimAiCache_(stateObj) {
  ensureAiStructures_(stateObj);
  var cache = stateObj.aiCache;
  var curSeason = Number((stateObj.espnData && stateObj.espnData.currentSeason) || 2026);
  // 1. Keep only active-season weekly recaps and playoff advisors in live aiCache
  Object.keys(cache.weeklyRecaps).forEach(function(k) {
    var m = String(k).match(/^(\d{4})_/);
    if (m && Number(m[1]) < curSeason) {
      delete cache.weeklyRecaps[k];
    }
  });
  Object.keys(cache.playoffAdvisors).forEach(function(k) {
    var m = String(k).match(/^(\d{4})_/);
    if (m && Number(m[1]) < curSeason) {
      delete cache.playoffAdvisors[k];
    }
  });
  // 2. Enforce strict <= 35 KB total storage cap by pruning oldest entries from advisors, then tradeGrades, then recaps
  var buckets = ["playoffAdvisors", "tradeGrades", "weeklyRecaps", "hofBios"];
  var guard = 0;
  while (JSON.stringify(cache).length > AI_CACHE_MAX_BYTES && guard < 200) {
    guard++;
    var evicted = false;
    for (var bi = 0; bi < buckets.length; bi++) {
      var bName = buckets[bi];
      var bObj = cache[bName] || {};
      var keys = Object.keys(bObj);
      if (keys.length === 0) continue;
      keys.sort(function(a, b) {
        return Number((bObj[a] && bObj[a].generatedAt) || 0) - Number((bObj[b] && bObj[b].generatedAt) || 0);
      });
      delete bObj[keys[0]];
      evicted = true;
      break;
    }
    if (!evicted) break;
  }
}

function enforceAntiSlopExclamationCap_(node, stateRef) {
  if (!stateRef) stateRef = { exclamUsed: 0 };
  if (typeof node === "string") {
    var cleaned = "";
    for (var i = 0; i < node.length; i++) {
      var ch = node.charAt(i);
      if (ch === "!") {
        if (stateRef.exclamUsed < 1) {
          stateRef.exclamUsed += 1;
          cleaned += ch;
        } else {
          cleaned += ".";
        }
      } else {
        cleaned += ch;
      }
    }
    return cleaned.replace(/\.\.+/g, ".").trim();
  }
  if (Array.isArray(node)) {
    return node.map(function(item) { return enforceAntiSlopExclamationCap_(item, stateRef); });
  }
  if (node && typeof node === "object") {
    var out = {};
    Object.keys(node).forEach(function(k) {
      out[k] = enforceAntiSlopExclamationCap_(node[k], stateRef);
    });
    return out;
  }
  return node;
}

function checkAndConsumeMemberAiQuota_() {
  try {
    var cache = CacheService.getScriptCache();
    var now = new Date();
    var hrKey = "gcm_ai_hr_" + now.getUTCFullYear() + "_" + (now.getUTCMonth() + 1) + "_" + now.getUTCDate() + "_" + now.getUTCHours();
    var curCount = Number(cache.get(hrKey) || 0);
    if (curCount >= AI_MEMBER_HOURLY_LIMIT) {
      return { allowed: false, used: curCount, limit: AI_MEMBER_HOURLY_LIMIT };
    }
    cache.put(hrKey, String(curCount + 1), 3700);
    return { allowed: true, used: curCount + 1, limit: AI_MEMBER_HOURLY_LIMIT };
  } catch (cacheErr) {
    return { allowed: true, used: 0, limit: AI_MEMBER_HOURLY_LIMIT };
  }
}

function callGeminiStructuredJson_(stateObj, taskDirective, telemetryPayload, responseSchema, opts) {
  opts = opts || {};
  ensureAiStructures_(stateObj);
  var props = PropertiesService.getScriptProperties();
  var apiKey = props ? String(props.getProperty("GEMINI_API_KEY") || "").trim() : "";
  if (!apiKey) {
    return {
      ok: false,
      code: "GEMINI_API_KEY_MISSING",
      error: "GEMINI_API_KEY is not configured in Apps Script PropertiesService. Add it in Project Settings -> Script Properties."
    };
  }

  var voiceCfg = stateObj.settings.aiVoice || { preset: "SNARKY_COMMISH", customLore: "" };
  var presetKey = AI_TONE_PRESETS[voiceCfg.preset] ? voiceCfg.preset : "SNARKY_COMMISH";
  var presetObj = AI_TONE_PRESETS[presetKey];

  var sysParts = [];
  if (opts.skipToneAndSlop) {
    sysParts.push("You are a strict JSON parser for a 10-team fantasy football keeper league (Gamecock Mafia). Return ONLY valid JSON matching the exact schema with zero preamble.");
  } else {
    sysParts.push("VOICE & TONE DIRECTIVE (" + presetObj.name + "):\n" + presetObj.directive);
    sysParts.push(AI_ANTI_SLOP_SYSTEM_BLOCK);
    if (voiceCfg.customLore && String(voiceCfg.customLore).trim()) {
      sysParts.push("COMMISSIONER LEAGUE LORE & STYLE NOTES:\n" + String(voiceCfg.customLore).trim());
    }
  }
  if (taskDirective) {
    sysParts.push("TASK INSTRUCTIONS:\n" + taskDirective);
  }

  var endpoint = "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODEL_NAME + ":generateContent?key=" + encodeURIComponent(apiKey);
  var reqPayload = {
    systemInstruction: {
      parts: [{ text: sysParts.join("\n\n") }]
    },
    contents: [
      {
        role: "user",
        parts: [{ text: "STRUCTURED LEAGUE TELEMETRY JSON:\n" + JSON.stringify(telemetryPayload) }]
      }
    ],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: responseSchema,
      temperature: opts.temperature !== undefined ? opts.temperature : 0.65
    }
  };

  var resp = UrlFetchApp.fetch(endpoint, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(reqPayload),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var text = resp.getContentText();
  if (code < 200 || code >= 300) {
    return {
      ok: false,
      code: "GEMINI_HTTP_" + code,
      error: "Gemini API returned HTTP " + code + ": " + String(text || "").slice(0, 260)
    };
  }

  var outer = JSON.parse(text);
  var cand = outer && outer.candidates && outer.candidates[0];
  var rawPart = cand && cand.content && cand.content.parts && cand.content.parts[0] && cand.content.parts[0].text;
  if (!rawPart) {
    return { ok: false, code: "GEMINI_EMPTY_RESPONSE", error: "Gemini returned an empty response." };
  }

  var parsed = null;
  try {
    parsed = JSON.parse(rawPart);
  } catch (parseErr) {
    var sIdx = rawPart.indexOf("{");
    var eIdx = rawPart.lastIndexOf("}");
    if (sIdx !== -1 && eIdx > sIdx) {
      parsed = JSON.parse(rawPart.slice(sIdx, eIdx + 1));
    } else {
      return { ok: false, code: "GEMINI_BAD_JSON", error: "Gemini response was not valid JSON." };
    }
  }

  if (!opts.skipToneAndSlop) {
    parsed = enforceAntiSlopExclamationCap_(parsed, { exclamUsed: 0 });
  }

  return {
    ok: true,
    data: parsed,
    model: GEMINI_MODEL_NAME,
    tonePreset: presetKey,
    toneName: presetObj.name
  };
}

/**
 * GET endpoint: Returns canonical live JSON state to authorized league members,
 * or snapshot metadata/preview for Commissioner Time Machine Rollback.
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

  var mode = (e && e.parameter && e.parameter.mode) ? String(e.parameter.mode).toUpperCase() : "";
  if (mode === "SNAPSHOT_LIST") {
    return jsonResponse({ ok: true, snapshots: listSnapshotsFromSheet(sheets.stateSheet) });
  }
  if (mode === "SNAPSHOT_PREVIEW") {
    var rowNum = Number(e.parameter.rowNumber) || 5;
    var snapState = readSnapshotStateAtRow(sheets.stateSheet, rowNum);
    return jsonResponse({ ok: Boolean(snapState), rowNumber: rowNum, snapshotState: sanitizeStateForClient(snapState) });
  }

  if (!stored) {
    return jsonResponse({ ok: true, empty: true, state: null });
  }
  return jsonResponse({ ok: true, empty: false, state: sanitizeStateForClient(stored), snapshots: listSnapshotsFromSheet(sheets.stateSheet) });
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

      var cleanFairnessVotes = { yes: [], no: [] };
      if (tr.fairnessVotes && typeof tr.fairnessVotes === "object") {
        if (Array.isArray(tr.fairnessVotes.yes)) {
          tr.fairnessVotes.yes.slice(0, 20).forEach(function(v) {
            var sv = String(v || "").trim().slice(0, 30);
            if (sv && cleanFairnessVotes.yes.indexOf(sv) === -1) cleanFairnessVotes.yes.push(sv);
          });
        }
        if (Array.isArray(tr.fairnessVotes.no)) {
          tr.fairnessVotes.no.slice(0, 20).forEach(function(v) {
            var sv = String(v || "").trim().slice(0, 30);
            if (sv && cleanFairnessVotes.no.indexOf(sv) === -1 && cleanFairnessVotes.yes.indexOf(sv) === -1) cleanFairnessVotes.no.push(sv);
          });
        }
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
        reactions: (tr.reactions && typeof tr.reactions === "object") ? tr.reactions : {},
        fairnessVotes: cleanFairnessVotes
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
        var partnerTargets = [cleanProp.teamA, cleanProp.teamB, cleanProp.teamC].filter(function(tid) {
          return Boolean(tid && tid !== cleanProp.proposedByTeam);
        });
        if (partnerTargets.length > 0) {
          dispatchOneSignalPush(baseState, {
            category: "DIRECT_TRADE",
            urgent: true,
            targetTeamIds: partnerTargets,
            title: "🤝 Incoming Trade Offer (" + propId + ")",
            message: cleanProp.headline
          });
        }
        return jsonResponse({ ok: true, assignedId: propId, state: sanitizeStateForClient(baseState) });
      }

      var assignedId = assignOfficialTradeId(baseState, primaryYear);
      var cleanTrade = sanitizeIncomingTradeObject(tr, primaryYear, assignedId);
      baseState.trades.push(cleanTrade);
      applyConfirmedPlayerTransfersToRosters(baseState, cleanTrade);
      pruneConflictingPendingProposals(baseState);
      baseState.updatedAt = Date.now();

      saveAndMirrorToSheets(sheets, baseState, "MEMBER_TRADE:" + assignedId);
      dispatchOneSignalPush(baseState, {
        category: "BROADCAST",
        title: "🚨 Official Trade Logged (" + assignedId + ")",
        message: cleanTrade.headline
      });
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
        if (subAction === "DECLINE" && propObj.proposedByTeam) {
          dispatchOneSignalPush(baseStateResp, {
            category: "DIRECT_TRADE",
            urgent: true,
            targetTeamIds: [propObj.proposedByTeam],
            title: "Trade Proposal Declined (" + propIdTarget + ")",
            message: propObj.headline + (replyNote ? (" — Note: " + replyNote) : "")
          });
        }
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

      if (isAuthorizedCommishOrCoCommish(stored, body).authorized) {
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
        dispatchOneSignalPush(baseStateResp, {
          category: "BROADCAST",
          title: "✅ Trade Confirmed (" + finalId + ")",
          message: propObj.headline
        });
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
    // ACTION 1F: LEAGUE MEMBER VOTES ON COMPLETED TRADE FAIRNESS (YES / NO)
    // ------------------------------------------------------------------------
    if (action === "VOTE_TRADE_FAIRNESS") {
      var baseStateVote = stored || body.fallbackBaseState;
      if (!baseStateVote || !Array.isArray(baseStateVote.trades) || !Array.isArray(baseStateVote.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      var voteTradeId = String(body.tradeId || "");
      var voterTeamId = String(body.teamId || "");
      var voteChoice = String(body.vote || "").toUpperCase(); // YES | NO | CLEAR
      var validVoter = baseStateVote.teams.some(function(t) { return t.id === voterTeamId; });
      if (!validVoter || ["YES", "NO", "CLEAR"].indexOf(voteChoice) === -1) {
        return jsonResponse({ ok: false, error: "Invalid trade fairness vote." });
      }
      for (var vti = 0; vti < baseStateVote.trades.length; vti++) {
        if (String(baseStateVote.trades[vti].id) === voteTradeId) {
          var vTr = baseStateVote.trades[vti];
          if (!vTr.fairnessVotes || typeof vTr.fairnessVotes !== "object") {
            vTr.fairnessVotes = { yes: [], no: [] };
          }
          var yesList = Array.isArray(vTr.fairnessVotes.yes) ? vTr.fairnessVotes.yes.filter(function(x) { return String(x) !== voterTeamId; }) : [];
          var noList = Array.isArray(vTr.fairnessVotes.no) ? vTr.fairnessVotes.no.filter(function(x) { return String(x) !== voterTeamId; }) : [];
          if (voteChoice === "YES") yesList.push(voterTeamId);
          else if (voteChoice === "NO") noList.push(voterTeamId);
          vTr.fairnessVotes = { yes: yesList, no: noList };
          baseStateVote.updatedAt = Date.now();
          saveAndMirrorToSheets(sheets, baseStateVote, "TRADE_FAIRNESS_VOTE:" + voteTradeId + ":" + voteChoice);
          return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateVote) });
        }
      }
      return jsonResponse({ ok: false, error: "Trade not found." });
    }

    // ------------------------------------------------------------------------
    // ACTION 1G: DESIGNATED DECISION-OWNER OR COMMISSIONER RESOLVES A POOL OR CONDITIONAL SWAP (A4, A5)
    // ------------------------------------------------------------------------
    if (action === "RESOLVE_TRADE_CONDITION") {
      var baseStateRes = stored || body.fallbackBaseState;
      if (!baseStateRes || !Array.isArray(baseStateRes.trades) || !Array.isArray(baseStateRes.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      var resKind = String(body.conditionKind || body.conditionType || "").toUpperCase(); // POOL | CONDITIONAL
      var resId = String(body.conditionId || "");
      var actorTeamId = String(body.actorTeamId || "");
      var resNote = String(body.resolutionNote || "").trim().slice(0, 220);
      var isCommishRes = isAuthorizedCommishOrCoCommish(stored, body).authorized;

      if (!resNote) {
        return jsonResponse({ ok: false, error: "A Resolution Audit Note is required to finalize a conditional or pool." });
      }

      var teamNameLookup = {};
      baseStateRes.teams.forEach(function(t) { teamNameLookup[t.id] = t.name; });
      var fmtTeam = function(tid) { return teamNameLookup[tid] || tid || "Commissioner"; };

      for (var rti = 0; rti < baseStateRes.trades.length; rti++) {
        var rTr = baseStateRes.trades[rti];
        if (resKind === "POOL" && Array.isArray(rTr.pools)) {
          for (var pi = 0; pi < rTr.pools.length; pi++) {
            var pool = rTr.pools[pi];
            if (String(pool.id) === resId) {
              var poolOwner = String(pool.poolHolder || "");
              if (!isCommishRes && (!actorTeamId || actorTeamId !== poolOwner)) {
                return jsonResponse({ ok: false, error: "Only the Pool Decision-Owner (" + fmtTeam(poolOwner) + ") or Commissioner can resolve this pool." });
              }
              var chosenOrig = String(body.resolvedOriginalOwner || body.chosenOriginalOwner || "");
              if (!Array.isArray(pool.candidateOriginalOwners) || pool.candidateOriginalOwners.indexOf(chosenOrig) === -1) {
                return jsonResponse({ ok: false, error: "Selected pick is not a valid candidate in this pool." });
              }
              pool.resolvedOriginalOwner = chosenOrig;
              pool.resolutionNote = resNote;
              pool.resolvedBy = isCommishRes && !actorTeamId ? "commish" : (actorTeamId || poolOwner);
              pool.resolvedAt = Date.now();
              appendServerActivityEvent(
                baseStateRes,
                "CONDITION_RESOLVED",
                "Condition Resolved: " + pool.year + " Rd " + pool.round + " Pool (" + rTr.id + ")",
                fmtTeam(pool.resolvedBy) + " finalized " + fmtTeam(pool.poolHolder) + " -> " + fmtTeam(pool.owedTo) + " using " + fmtTeam(chosenOrig) + "'s pick. Note: \"" + resNote + "\"",
                { year: Number(pool.year), tradeId: rTr.id, teamId: pool.owedTo, actorTeamId: pool.resolvedBy, resolutionNote: resNote }
              );
              baseStateRes.updatedAt = Date.now();
              saveAndMirrorToSheets(sheets, baseStateRes, "RESOLVE_POOL:" + resId);
              return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateRes) });
            }
          }
        } else if (resKind === "CONDITIONAL" && Array.isArray(rTr.conditionals)) {
          for (var ci = 0; ci < rTr.conditionals.length; ci++) {
            var cond = rTr.conditionals[ci];
            if (String(cond.id) === resId) {
              var decisionOwner = String(cond.favoredTeam || (cond.analysis && cond.analysis.favored) || cond.teamA || "");
              if (!isCommishRes && (!actorTeamId || actorTeamId !== decisionOwner)) {
                return jsonResponse({ ok: false, error: "Only the Designated Decision-Owner (" + fmtTeam(decisionOwner) + ") or Commissioner can resolve this condition." });
              }
              var chosenRd = Number(body.resolvedRound || 0);
              cond.isResolved = true;
              cond.resolvedRound = chosenRd > 0 ? chosenRd : null;
              cond.resolutionNote = resNote;
              cond.resolvedBy = isCommishRes && !actorTeamId ? "commish" : (actorTeamId || decisionOwner);
              cond.resolvedAt = Date.now();

              // Mirror direct transfer if a round swap was executed and not already added
              if (chosenRd > 0) {
                if (!Array.isArray(rTr.directTransfers)) rTr.directTransfers = [];
                var alreadyMirrored = rTr.directTransfers.some(function(dt) { return dt.fromConditionalId === cond.id; });
                if (!alreadyMirrored) {
                  rTr.directTransfers.push(
                    { year: cond.year, round: chosenRd, originalOwner: cond.teamA, fromOwner: cond.teamA, toOwner: cond.teamB, note: "Resolved Swap (" + cond.id + "): " + resNote, fromConditionalId: cond.id },
                    { year: cond.year, round: chosenRd, originalOwner: cond.teamB, fromOwner: cond.teamB, toOwner: cond.teamA, note: "Resolved Swap (" + cond.id + "): " + resNote, fromConditionalId: cond.id }
                  );
                }
              }

              var outcomeStr = chosenRd > 0 ? ("Swapped Round " + chosenRd) : "No Swap Triggered (Kept Original)";
              appendServerActivityEvent(
                baseStateRes,
                "CONDITION_RESOLVED",
                "Condition Resolved: " + cond.year + " Conditional Swap (" + rTr.id + ")",
                fmtTeam(cond.resolvedBy) + " resolved " + fmtTeam(cond.teamA) + " <-> " + fmtTeam(cond.teamB) + " (" + outcomeStr + "). Note: \"" + resNote + "\"",
                { year: Number(cond.year), tradeId: rTr.id, teamId: decisionOwner, actorTeamId: cond.resolvedBy, resolutionNote: resNote }
              );
              baseStateRes.updatedAt = Date.now();
              saveAndMirrorToSheets(sheets, baseStateRes, "RESOLVE_CONDITIONAL:" + resId);
              return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateRes) });
            }
          }
        }
      }
      return jsonResponse({ ok: false, error: "Condition or Pool not found." });
    }

    // ------------------------------------------------------------------------
    // ACTION 1H: INDIVIDUAL NFL PLAYER HALL OF FAME 7/10 MANAGER BALLOT VOTE (A7-A11)
    // ------------------------------------------------------------------------
    if (action === "VOTE_HOF_BALLOT") {
      var baseStateHof = stored || body.fallbackBaseState;
      if (!baseStateHof || !Array.isArray(baseStateHof.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      var hofPlayerId = String(body.playerId || "").trim().slice(0, 64);
      var hofPlayerName = String(body.playerName || hofPlayerId).trim().slice(0, 64);
      var hofPos = String(body.pos || "FLEX").trim().toUpperCase().slice(0, 12);
      var hofVoterId = String(body.teamId || "");
      var hofChoice = String(body.vote || "").toUpperCase(); // YES | NO | CLEAR | NOMINATE
      if (!hofPlayerId) {
        return jsonResponse({ ok: false, error: "Invalid player ID for Hall of Fame ballot." });
      }
      if (!baseStateHof.espnData || typeof baseStateHof.espnData !== "object") {
        baseStateHof.espnData = {};
      }
      if (!baseStateHof.espnData.playerHofBallots || typeof baseStateHof.espnData.playerHofBallots !== "object") {
        baseStateHof.espnData.playerHofBallots = {};
      }
      var existingBallot = baseStateHof.espnData.playerHofBallots[hofPlayerId] || {
        playerId: hofPlayerId,
        playerName: hofPlayerName,
        name: hofPlayerName,
        pos: hofPos,
        nominated: false,
        yes: [],
        no: [],
        inductedAt: 0
      };
      var isCommishHof = isAuthorizedCommishOrCoCommish(stored, body).authorized;
      if (hofChoice === "NOMINATE") {
        if (!isCommishHof) {
          return jsonResponse({ ok: false, error: "Only the Commissioner can toggle manual retirement nominations." });
        }
        existingBallot.nominated = Boolean(body.nominated);
      } else {
        var validHofVoter = baseStateHof.teams.some(function(t) { return t.id === hofVoterId; });
        if (!validHofVoter || ["YES", "NO", "CLEAR"].indexOf(hofChoice) === -1) {
          return jsonResponse({ ok: false, error: "Invalid Hall of Fame ballot vote." });
        }
        var hYes = Array.isArray(existingBallot.yes) ? existingBallot.yes.filter(function(x) { return String(x) !== hofVoterId; }) : [];
        var hNo = Array.isArray(existingBallot.no) ? existingBallot.no.filter(function(x) { return String(x) !== hofVoterId; }) : [];
        if (hofChoice === "YES") hYes.push(hofVoterId);
        else if (hofChoice === "NO") hNo.push(hofVoterId);
        existingBallot.yes = hYes;
        existingBallot.no = hNo;
        if (hYes.length >= 7 && !existingBallot.inductedAt) {
          existingBallot.inductedAt = Date.now();
          appendServerActivityEvent(
            baseStateHof,
            "HOF_INDUCTION",
            "Player Hall of Fame Induction: " + hofPlayerName + " (" + hofPos + ")",
            hofPlayerName + " achieved the 7/10 League Supermajority (" + hYes.length + "/10 YES) and has been enshrined in the Gamecock Mafia Player Hall of Fame!",
            { playerId: hofPlayerId, actorTeamId: hofVoterId }
          );
        } else if (hYes.length < 7) {
          existingBallot.inductedAt = 0;
        }
      }
      existingBallot.playerName = hofPlayerName || existingBallot.playerName || existingBallot.name;
      existingBallot.name = existingBallot.playerName;
      existingBallot.pos = hofPos || existingBallot.pos;
      baseStateHof.espnData.playerHofBallots[hofPlayerId] = existingBallot;
      baseStateHof.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateHof, "VOTE_HOF_BALLOT:" + hofPlayerId + ":" + hofChoice);
      return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateHof) });
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
      var isCommishCaller = isAuthorizedCommishOrCoCommish(stored, body).authorized;
      if (!isCommishCaller && existingProf.teamPinHash && /^[a-f0-9]{64}$/i.test(existingProf.teamPinHash)) {
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
        var shameMsg = shameTeamName + " selected the 'Upstate Kitties' colorway and activated the Sparkly Pink 8==D Curse!";
        baseStateProf.hallOfShame.unshift({
          id: "SHAME-" + Date.now(),
          teamId: teamId,
          teamName: shameTeamName,
          message: shameMsg,
          timestamp: Date.now(),
          dateStr: new Date().toLocaleDateString()
        });
        baseStateProf.hallOfShame = baseStateProf.hallOfShame.slice(0, 30);
        appendServerActivityEvent(baseStateProf, "HALL_OF_SHAME", "Hall of Shame Alert", shameMsg, { teamId: teamId });
      }

      // Enforce Champion Rename Decree lock on Phallic Feaster's customTeamName
      var feasterLocked = Boolean(existingProf.feasterRenamedByChamp);
      var resolvedCustomName = (feasterLocked && !isCommishCaller)
        ? String(existingProf.customTeamName || "").trim().slice(0, 24)
        : String(incoming.customTeamName !== undefined ? incoming.customTeamName : (existingProf.customTeamName || "")).trim().slice(0, 24);

      var resolvedStrategyTag = incoming.strategyTag !== undefined
        ? normalizeStrategyModeTag(incoming.strategyTag)
        : normalizeStrategyModeTag(existingProf.strategyTag || "HOLDING_SERVE");

      var validColors = [
        "sandstorm_noir", "chapel_sky", "honolulu_gridiron", "foundry_orange", "emerald_sound",
        "jets_of_the_west", "upstate_kitties", "pitch_black", "silly_philly", "iced_out", "horny_head", "merica", ""
      ];
      var resolvedColorTheme = (incoming.colorTheme !== undefined && validColors.indexOf(String(incoming.colorTheme || "")) !== -1)
        ? String(incoming.colorTheme || "")
        : String(existingProf.colorTheme || "");

      var resolvedEncryptedDraftPrep = incoming.clearDraftPrep
        ? ""
        : (incoming.encryptedDraftPrep !== undefined && typeof incoming.encryptedDraftPrep === "string"
            ? incoming.encryptedDraftPrep.slice(0, 120000)
            : String(existingProf.encryptedDraftPrep || ""));

      var existingPush = (existingProf.pushPrefs && typeof existingProf.pushPrefs === "object") ? existingProf.pushPrefs : {};
      var incomingPush = (incoming.pushPrefs && typeof incoming.pushPrefs === "object") ? incoming.pushPrefs : null;
      var sanitizedPushPrefs = {
        enabled: incomingPush ? Boolean(incomingPush.enabled) : Boolean(existingPush.enabled),
        subscriptionId: String((incomingPush && incomingPush.subscriptionId !== undefined ? incomingPush.subscriptionId : existingPush.subscriptionId) || "").trim().slice(0, 120),
        directTradeAlerts: incomingPush && incomingPush.directTradeAlerts !== undefined ? Boolean(incomingPush.directTradeAlerts) : (existingPush.directTradeAlerts !== false),
        conditionalAlerts: incomingPush && incomingPush.conditionalAlerts !== undefined ? Boolean(incomingPush.conditionalAlerts) : (existingPush.conditionalAlerts !== false),
        tradeBroadcasts: incomingPush && incomingPush.tradeBroadcasts !== undefined ? Boolean(incomingPush.tradeBroadcasts) : (existingPush.tradeBroadcasts !== false),
        hofAlerts: incomingPush && incomingPush.hofAlerts !== undefined ? Boolean(incomingPush.hofAlerts) : (existingPush.hofAlerts !== false),
        keeperAlerts: incomingPush && incomingPush.keeperAlerts !== undefined ? Boolean(incomingPush.keeperAlerts) : (existingPush.keeperAlerts !== false),
        dismissedPrompt: incomingPush && incomingPush.dismissedPrompt !== undefined ? Boolean(incomingPush.dismissedPrompt) : Boolean(existingPush.dismissedPrompt),
        updatedAt: Date.now()
      };

      baseStateProf.teamProfiles[teamId] = {
        tradeBlockNote: String(incoming.tradeBlockNote !== undefined ? incoming.tradeBlockNote : (existingProf.tradeBlockNote || "")).slice(0, 280),
        lookingFor: String(incoming.lookingFor !== undefined ? incoming.lookingFor : (existingProf.lookingFor || "")).slice(0, 180),
        needsChips: incoming.needsChips !== undefined ? sanitizedChips : (Array.isArray(existingProf.needsChips) ? existingProf.needsChips : []),
        pickTags: incoming.pickTags && typeof incoming.pickTags === "object" ? sanitizedTags : (existingProf.pickTags || {}),
        playerTags: incoming.playerTags && typeof incoming.playerTags === "object" ? sanitizedPlayerTags : (existingProf.playerTags || {}),
        strategyTag: resolvedStrategyTag,
        colorTheme: resolvedColorTheme,
        encryptedDraftPrep: resolvedEncryptedDraftPrep,
        pushPrefs: sanitizedPushPrefs,
        keepers: sanitizedKeepers,
        keepersConfirmedAt: Number(incoming.keepersConfirmedAt !== undefined ? incoming.keepersConfirmedAt : (existingProf.keepersConfirmedAt || 0)),
        customTeamName: resolvedCustomName,
        feasterRenamedByChamp: feasterLocked,
        feasterRenamedBy: String(existingProf.feasterRenamedBy || "").slice(0, 30),
        avatarType: ["MONOGRAM", "EMOJI", "IMAGE"].indexOf(incoming.avatarType) !== -1 ? incoming.avatarType : (existingProf.avatarType || "MONOGRAM"),
        avatarText: String(incoming.avatarText !== undefined ? incoming.avatarText : (existingProf.avatarText || "")).trim().slice(0, 4),
        avatarEmoji: String(incoming.avatarEmoji !== undefined ? incoming.avatarEmoji : (existingProf.avatarEmoji || "")).trim().slice(0, 12),
        avatarBg: String(incoming.avatarBg !== undefined ? incoming.avatarBg : (existingProf.avatarBg || "")).trim().slice(0, 24),
        avatarImageDataUrl: String(incoming.avatarImageDataUrl !== undefined ? incoming.avatarImageDataUrl : (existingProf.avatarImageDataUrl || "")).trim().slice(0, 35000),
        upstateKittyCursed: nowCursed,
        upstateKittyCursedAt: nowCursed ? (Number(existingProf.upstateKittyCursedAt) || Date.now()) : 0,
        flairBadge: String(existingProf.flairBadge || ""),
        teamPinVerifier: (incoming.teamPinVerifier && /^[a-f0-9]{64}$/i.test(incoming.teamPinVerifier))
          ? incoming.teamPinVerifier.toLowerCase()
          : (incoming.clearTeamPin ? "" : String(existingProf.teamPinVerifier || "")),
        teamPinHash: (incoming.teamPinHash && /^[a-f0-9]{64}$/i.test(incoming.teamPinHash))
          ? incoming.teamPinHash.toLowerCase()
          : (incoming.clearTeamPin ? "" : String(existingProf.teamPinHash || "")),
        updatedAt: Date.now()
      };
      if (Array.isArray(incoming.roster)) {
        if (!baseStateProf.espnData || typeof baseStateProf.espnData !== "object") {
          baseStateProf.espnData = { rosters: {} };
        }
        if (!baseStateProf.espnData.rosters || typeof baseStateProf.espnData.rosters !== "object") {
          baseStateProf.espnData.rosters = {};
        }
        baseStateProf.espnData.rosters[teamId] = incoming.roster.slice(0, 45).map(function(rp, rIdx) {
          return Object.assign({}, rp || {}, {
            id: String(rp && rp.id ? rp.id : ("p-" + teamId + "-" + (rIdx + 1))).slice(0, 48),
            espnId: Number(rp && rp.espnId) || 0,
            name: String(rp && rp.name ? rp.name : "").trim().slice(0, 60),
            pos: String(rp && rp.pos ? rp.pos : "FLEX").trim().toUpperCase().slice(0, 12),
            nflTeam: String(rp && rp.nflTeam ? rp.nflTeam : "NFL").trim().toUpperCase().slice(0, 8),
            slot: String(rp && rp.slot ? rp.slot : "BE").trim().toUpperCase().slice(0, 10),
            projPts: Number(rp && rp.projPts) || 0,
            totalPts: Number(rp && rp.totalPts) || 0
          });
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

      var isCommishDecree = isAuthorizedCommishOrCoCommish(stored, body).authorized;
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
        var feasterMsg = (cObj ? cObj.name : "The Overall Champion") + " exercised the Champion's Decree and renamed Phallic Feaster " + (fObj ? fObj.name : targetTid) + " to \"" + newFeasterName + "\"!";
        baseStateChamp.hallOfShame.unshift({
          id: "FEASTER-" + Date.now(),
          teamId: targetTid,
          teamName: fObj ? fObj.name : targetTid,
          message: feasterMsg,
          timestamp: Date.now(),
          dateStr: new Date().toLocaleDateString()
        });
        baseStateChamp.hallOfShame = baseStateChamp.hallOfShame.slice(0, 30);
        appendServerActivityEvent(baseStateChamp, "FEASTER_DECREE", "Phallic Feaster Team Name Decree", feasterMsg, { teamId: targetTid, actorTeamId: actorTid });
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
      var prevRecords = (existingEspn.records && typeof existingEspn.records === "object") ? existingEspn.records : {};
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
        espnTeamMap: (incomingEspn.espnTeamMap && typeof incomingEspn.espnTeamMap === "object")
          ? incomingEspn.espnTeamMap
          : (incomingEspn.teamSlotMap || existingEspn.teamSlotMap || {}),
        seasonOwnerOverrides: (incomingEspn.seasonOwnerOverrides && typeof incomingEspn.seasonOwnerOverrides === "object")
          ? incomingEspn.seasonOwnerOverrides
          : (existingEspn.seasonOwnerOverrides || {}),
        alumniManagers: Array.isArray(incomingEspn.alumniManagers)
          ? incomingEspn.alumniManagers
          : (existingEspn.alumniManagers || []),
        rosters: (incomingEspn.rosters && typeof incomingEspn.rosters === "object")
          ? incomingEspn.rosters
          : (existingEspn.rosters || {}),
        keepers: (incomingEspn.keepers && typeof incomingEspn.keepers === "object")
          ? incomingEspn.keepers
          : (existingEspn.keepers || {}),
        masterKeepersByTeam: (incomingEspn.masterKeepersByTeam && typeof incomingEspn.masterKeepersByTeam === "object")
          ? incomingEspn.masterKeepersByTeam
          : (existingEspn.masterKeepersByTeam || {}),
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
        draftResults: (incomingEspn.draftResults && typeof incomingEspn.draftResults === "object")
          ? incomingEspn.draftResults
          : (existingEspn.draftResults || {}),
        syncedMatchupKeys: (incomingEspn.syncedMatchupKeys && typeof incomingEspn.syncedMatchupKeys === "object")
          ? incomingEspn.syncedMatchupKeys
          : (existingEspn.syncedMatchupKeys || {}),
        rawTeams: Array.isArray(incomingEspn.rawTeams)
          ? incomingEspn.rawTeams.slice(0, 20)
          : (existingEspn.rawTeams || []),
        freeAgents: Array.isArray(incomingEspn.freeAgents)
          ? incomingEspn.freeAgents.slice(0, 250)
          : (existingEspn.freeAgents || []),
        positionalLeaders: (incomingEspn.positionalLeaders && typeof incomingEspn.positionalLeaders === "object")
          ? incomingEspn.positionalLeaders
          : (existingEspn.positionalLeaders || {}),
        positionalRecords: (incomingEspn.positionalRecords && typeof incomingEspn.positionalRecords === "object")
          ? incomingEspn.positionalRecords
          : (existingEspn.positionalRecords || []),
        playerHofBallots: (incomingEspn.playerHofBallots && typeof incomingEspn.playerHofBallots === "object")
          ? Object.assign({}, existingEspn.playerHofBallots || {}, incomingEspn.playerHofBallots)
          : (existingEspn.playerHofBallots || {}),
        hofState: (incomingEspn.hofState && typeof incomingEspn.hofState === "object")
          ? Object.assign({}, existingEspn.hofState || {}, incomingEspn.hofState)
          : (existingEspn.hofState || null),
        faabTransactions: (incomingEspn.faabTransactions && typeof incomingEspn.faabTransactions === "object")
          ? Object.assign({}, existingEspn.faabTransactions || {}, incomingEspn.faabTransactions)
          : (existingEspn.faabTransactions || {}),
        seasonWeeklyPosKings: (incomingEspn.seasonWeeklyPosKings && typeof incomingEspn.seasonWeeklyPosKings === "object")
          ? mergeWeeklyPosKingsStores_(existingEspn.seasonWeeklyPosKings, incomingEspn.seasonWeeklyPosKings, true)
          : sanitizeWeeklyPosKingsStore_(existingEspn.seasonWeeklyPosKings, Number(incomingEspn.currentSeason || existingEspn.currentSeason || 2026)),
        allTimeWeeklyPosKings: (incomingEspn.allTimeWeeklyPosKings && typeof incomingEspn.allTimeWeeklyPosKings === "object")
          ? mergeWeeklyPosKingsStores_(existingEspn.allTimeWeeklyPosKings, incomingEspn.allTimeWeeklyPosKings, false)
          : sanitizeWeeklyPosKingsStore_(existingEspn.allTimeWeeklyPosKings, null),
        weeklyBenchBlunders: (incomingEspn.weeklyBenchBlunders && typeof incomingEspn.weeklyBenchBlunders === "object")
          ? Object.assign({}, existingEspn.weeklyBenchBlunders || {}, incomingEspn.weeklyBenchBlunders)
          : (existingEspn.weeklyBenchBlunders || {})
      };
      // Wave 4 Q3: Import Keepers from ESPN as Master Source of Truth when present
      var masterKeepers = body.masterKeepersByTeam || incomingEspn.masterKeepersByTeam || incomingEspn.keepers || null;
      if (masterKeepers && typeof masterKeepers === "object" && body.syncMasterKeepers !== false) {
        if (!baseStateEspn.teamProfiles || typeof baseStateEspn.teamProfiles !== "object") {
          baseStateEspn.teamProfiles = {};
        }
        var maxKCap = Number((baseStateEspn.espnData.keeperRules && baseStateEspn.espnData.keeperRules.maxKeepers) || 12) || 12;
        Object.keys(masterKeepers).forEach(function(tid) {
          var kArr = Array.isArray(masterKeepers[tid]) ? masterKeepers[tid] : [];
          if (kArr.length > 0) {
            if (!baseStateEspn.teamProfiles[tid]) baseStateEspn.teamProfiles[tid] = {};
            var cleanMaster = [];
            kArr.slice(0, maxKCap).forEach(function(kid) {
              var ks = String(kid || "").trim().slice(0, 48);
              if (ks && cleanMaster.indexOf(ks) === -1) cleanMaster.push(ks);
            });
            baseStateEspn.teamProfiles[tid].keepers = cleanMaster;
            baseStateEspn.teamProfiles[tid].keepersConfirmedAt = Date.now();
            baseStateEspn.teamProfiles[tid].updatedAt = Date.now();
          }
        });
      }
      if (body.faabData && typeof body.faabData === "object") {
        var prevFaab = (baseStateEspn.faabData && typeof baseStateEspn.faabData === "object") ? baseStateEspn.faabData : {};
        baseStateEspn.faabData = {
          defaultBudget: Number(body.faabData.defaultBudget || prevFaab.defaultBudget) || 600,
          annualBudget: Number(body.faabData.annualBudget || prevFaab.annualBudget) || 600,
          seasons: Object.assign({}, prevFaab.seasons || {}, body.faabData.seasons || {}),
          records: (body.faabData.records && typeof body.faabData.records === "object")
            ? body.faabData.records
            : (prevFaab.records || { highestBids: [], mostSpentSeason: [], mostSavedSeason: [] })
        };
      }
      var nextRecs = baseStateEspn.espnData.records || {};
      if (nextRecs.highestWeekScore && Number(nextRecs.highestWeekScore.score || 0) > Number((prevRecords.highestWeekScore && prevRecords.highestWeekScore.score) || 0) && Number((prevRecords.highestWeekScore && prevRecords.highestWeekScore.score) || 0) > 0) {
        appendServerActivityEvent(baseStateEspn, "RECORD_SET", "New All-Time Single-Week Scoring Record!", (nextRecs.highestWeekScore.managerId || "Manager") + " posted " + nextRecs.highestWeekScore.score + " pts (Wk " + (nextRecs.highestWeekScore.week || "?") + ", " + (nextRecs.highestWeekScore.year || "") + ").", { year: nextRecs.highestWeekScore.year, teamId: nextRecs.highestWeekScore.managerId });
      }
      var syncedRosterCount = Object.keys(baseStateEspn.espnData.rosters || {}).length;
      var syncedFaCount = Array.isArray(baseStateEspn.espnData.freeAgents) ? baseStateEspn.espnData.freeAgents.length : 0;
      appendServerActivityEvent(
        baseStateEspn,
        "ESPN_SYNC",
        "ESPN League Sync Completed (" + baseStateEspn.espnData.currentSeason + ")",
        "Synced " + syncedRosterCount + " rosters, " + syncedFaCount + " Free Agents/Rookies, master keepers, optimal lineups, All-Time FAAB, standings, and record book.",
        { year: baseStateEspn.espnData.currentSeason }
      );
      // Auto-void open pending proposals if a staged player was dropped/moved on ESPN
      pruneConflictingPendingProposals(baseStateEspn);
      baseStateEspn.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateEspn, "ESPN_SYNC_UPDATE:" + (baseStateEspn.espnData.leagueId || "LIVE"));
      return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateEspn) });
    }

    // ------------------------------------------------------------------------
    // ACTION 1I: 4-STAGE PLAYER HALL OF FAME OFFSEASON FUNNEL (WAVE 2 & WAVE 3)
    // Stages: IDLE -> NOMINATIONS_OPEN -> VOTING_OPEN -> RESULTS_PUBLISHED
    // ------------------------------------------------------------------------
    if (action === "PUBLISH_HOF_CANDIDATES" || action === "PUBLISH_HOF_BALLOT" || action === "SUBMIT_HOF_VOTE" || action === "CLOSE_HOF_BALLOT") {
      var baseStateFunnel = stored || body.fallbackBaseState;
      if (!baseStateFunnel || !Array.isArray(baseStateFunnel.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      if (!baseStateFunnel.espnData || typeof baseStateFunnel.espnData !== "object") {
        baseStateFunnel.espnData = {};
      }
      var hofState = (baseStateFunnel.espnData.hofState && typeof baseStateFunnel.espnData.hofState === "object")
        ? baseStateFunnel.espnData.hofState
        : {
            stage: "IDLE",
            classYear: Number(baseStateFunnel.baseUpcomingYear) || 2027,
            maxClassSize: 3,
            candidates: [],
            ballot: [],
            votesByManager: {},
            inductedPlayers: [],
            carryoverPool: [],
            droppedPlayerIds: [],
            excludedActivePlayers: [],
            lastResults: null
          };
      if (!Array.isArray(hofState.candidates)) hofState.candidates = [];
      if (!Array.isArray(hofState.ballot)) hofState.ballot = [];
      if (!hofState.votesByManager || typeof hofState.votesByManager !== "object") hofState.votesByManager = {};
      if (!Array.isArray(hofState.inductedPlayers)) hofState.inductedPlayers = [];
      if (!Array.isArray(hofState.carryoverPool)) hofState.carryoverPool = [];
      if (!Array.isArray(hofState.droppedPlayerIds)) hofState.droppedPlayerIds = [];
      if (!Array.isArray(hofState.excludedActivePlayers)) hofState.excludedActivePlayers = [];
      var maxClassCap = Math.max(1, Math.min(5, Number((baseStateFunnel.hofState && baseStateFunnel.hofState.maxClassSize) || hofState.maxClassSize || body.maxClassSize) || 3));
      hofState.maxClassSize = maxClassCap;
      if (Array.isArray(body.excludedActivePlayers)) {
        hofState.excludedActivePlayers = body.excludedActivePlayers.slice(0, 250).map(function(x) { return String(x || "").trim(); }).filter(Boolean);
      }

      var isCommishFunnel = isAuthorizedCommishOrCoCommish(stored, body).authorized;

      if (action === "PUBLISH_HOF_CANDIDATES") {
        if (!isCommishFunnel) {
          return jsonResponse({ ok: false, error: "Only the Commissioner can publish the Official HOF Candidate Watchlist." });
        }
        var classYr = Number(body.classYear) || Number(hofState.classYear) || Number(baseStateFunnel.baseUpcomingYear) || 2027;
        var incomingCands = Array.isArray(body.candidates) ? body.candidates : [];
        var mergedCands = [];
        var seenCandIds = {};
        // Automatically include prior-year carryover finalists (up to 5 years on ballot) unless permanently dropped
        hofState.carryoverPool.forEach(function(co) {
          if (!co || !co.playerId) return;
          var pid = String(co.playerId);
          if (hofState.droppedPlayerIds.indexOf(pid) !== -1 || hofState.excludedActivePlayers.indexOf(pid) !== -1) return;
          if (Number(co.yearsOnBallot || 1) >= 5) return;
          seenCandIds[pid] = true;
          mergedCands.push(Object.assign({}, co, { isCarryover: true }));
        });
        incomingCands.forEach(function(c) {
          if (!c || !(c.playerId || c.id || c.name || c.playerName)) return;
          var pid = String(c.playerId || c.id || c.name || c.playerName).trim();
          if (seenCandIds[pid] || hofState.droppedPlayerIds.indexOf(pid) !== -1 || hofState.excludedActivePlayers.indexOf(pid) !== -1) return;
          seenCandIds[pid] = true;
          mergedCands.push({
            playerId: pid,
            playerName: String(c.playerName || c.name || pid).trim().slice(0, 64),
            name: String(c.playerName || c.name || pid).trim().slice(0, 64),
            pos: String(c.pos || "FLEX").trim().toUpperCase().slice(0, 12),
            legacyScore: Number(c.legacyScore) || 0,
            totalLeaguePts: Number(c.totalLeaguePts || c.totalPts) || 0,
            peakSummary: String(c.peakSummary || "").slice(0, 160),
            ringsCount: Number(c.ringsCount || c.rings) || 0,
            starterRingsCount: Number(c.starterRingsCount !== undefined ? c.starterRingsCount : (c.ringsCount || c.rings || 0)) || 0,
            pedigree: String(c.pedigree || "").slice(0, 120),
            faabCost: c.faabCost !== undefined ? Number(c.faabCost) : null,
            yearsOnBallot: Number(c.yearsOnBallot) || 1,
            isCarryover: Boolean(c.isCarryover)
          });
        });
        if (mergedCands.length === 0) {
          return jsonResponse({ ok: false, error: "Select at least 1 candidate to publish the Offseason Watchlist." });
        }
        hofState.classYear = classYr;
        hofState.stage = "NOMINATIONS_OPEN";
        hofState.candidates = mergedCands;
        hofState.ballot = [];
        hofState.votesByManager = {};
        hofState.updatedAt = Date.now();
        baseStateFunnel.espnData.hofState = hofState;
        appendServerActivityEvent(
          baseStateFunnel,
          "HOF_WATCHLIST",
          "Class of " + classYr + " HOF Candidate Watchlist Published",
          "Commissioner published " + mergedCands.length + " Official HOF Candidates for league debate.",
          { year: classYr }
        );
        dispatchOneSignalPush(baseStateFunnel, {
          category: "HOF",
          title: "🏛️ HOF Class of " + classYr + " Watchlist Open!",
          message: "The Commissioner published " + mergedCands.length + " Official Player HOF Candidates in The Vault."
        });
        baseStateFunnel.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseStateFunnel, "PUBLISH_HOF_CANDIDATES:" + classYr);
        return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateFunnel) });
      }

      if (action === "PUBLISH_HOF_BALLOT") {
        if (!isCommishFunnel) {
          return jsonResponse({ ok: false, error: "Only the Commissioner can lock and publish the Final 5-Man HOF Ballot." });
        }
        var incomingBallot = Array.isArray(body.ballot) ? body.ballot : [];
        if (incomingBallot.length < 1 || incomingBallot.length > 8) {
          return jsonResponse({ ok: false, error: "The Final HOF Ballot must contain between 1 and 8 finalists." });
        }
        var cleanBallot = incomingBallot.map(function(b) {
          var pid = String(b.playerId || b.id || b.playerName || b.name || "").trim();
          return {
            playerId: pid,
            playerName: String(b.playerName || b.name || pid).trim().slice(0, 64),
            name: String(b.playerName || b.name || pid).trim().slice(0, 64),
            pos: String(b.pos || "FLEX").trim().toUpperCase().slice(0, 12),
            legacyScore: Number(b.legacyScore) || 0,
            totalLeaguePts: Number(b.totalLeaguePts || b.totalPts) || 0,
            peakSummary: String(b.peakSummary || "").slice(0, 160),
            ringsCount: Number(b.ringsCount || b.rings) || 0,
            starterRingsCount: Number(b.starterRingsCount !== undefined ? b.starterRingsCount : (b.ringsCount || b.rings || 0)) || 0,
            pedigree: String(b.pedigree || "").slice(0, 120),
            faabCost: b.faabCost !== undefined ? Number(b.faabCost) : null,
            yearsOnBallot: Number(b.yearsOnBallot) || 1
          };
        });
        hofState.stage = "VOTING_OPEN";
        hofState.ballot = cleanBallot;
        hofState.votesByManager = {};
        hofState.updatedAt = Date.now();
        baseStateFunnel.espnData.hofState = hofState;
        appendServerActivityEvent(
          baseStateFunnel,
          "HOF_BALLOT_OPEN",
          "Class of " + hofState.classYear + " Official HOF Ballot Open!",
          "Manager voting is now open (select up to " + maxClassCap + " finalists; 7/10 supermajority required, max " + maxClassCap + " inductees).",
          { year: hofState.classYear }
        );
        dispatchOneSignalPush(baseStateFunnel, {
          category: "HOF",
          title: "🗳️ HOF Ballot is LIVE!",
          message: "Cast your vote (up to " + maxClassCap + " players) for the Class of " + hofState.classYear + " Hall of Fame."
        });
        baseStateFunnel.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseStateFunnel, "PUBLISH_HOF_BALLOT:" + hofState.classYear);
        return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateFunnel) });
      }

      if (action === "SUBMIT_HOF_VOTE") {
        if (hofState.stage !== "VOTING_OPEN") {
          return jsonResponse({ ok: false, error: "HOF ballot voting is not currently open." });
        }
        var voterId = String(body.teamId || body.actorTeamId || "").trim();
        var validMgr = baseStateFunnel.teams.some(function(t) { return t.id === voterId; });
        if (!validMgr) {
          return jsonResponse({ ok: false, error: "Invalid manager ID for HOF ballot." });
        }
        var picks = Array.isArray(body.selectedPlayerIds) ? body.selectedPlayerIds.map(String) : [];
        var uniquePicks = [];
        var validBallotIds = {};
        hofState.ballot.forEach(function(b) { validBallotIds[String(b.playerId)] = true; });
        picks.forEach(function(pid) {
          if (validBallotIds[pid] && uniquePicks.indexOf(pid) === -1) uniquePicks.push(pid);
        });
        if (uniquePicks.length > maxClassCap) {
          return jsonResponse({ ok: false, error: "Each manager may vote for at most " + maxClassCap + " finalists on the ballot." });
        }
        hofState.votesByManager[voterId] = {
          selectedPlayerIds: uniquePicks,
          submittedAt: Date.now()
        };
        hofState.updatedAt = Date.now();
        baseStateFunnel.espnData.hofState = hofState;
        baseStateFunnel.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseStateFunnel, "SUBMIT_HOF_VOTE:" + voterId);
        return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateFunnel) });
      }

      if (action === "CLOSE_HOF_BALLOT") {
        if (!isCommishFunnel) {
          return jsonResponse({ ok: false, error: "Only the Commissioner can close HOF voting and publish official results." });
        }
        var voteMap = hofState.votesByManager || {};
        var totalVoters = 10;
        var tally = hofState.ballot.map(function(b) {
          var pid = String(b.playerId);
          var voters = [];
          Object.keys(voteMap).forEach(function(mgrId) {
            var vEntry = voteMap[mgrId];
            var sel = (vEntry && Array.isArray(vEntry.selectedPlayerIds)) ? vEntry.selectedPlayerIds : (Array.isArray(vEntry) ? vEntry : []);
            if (sel.indexOf(pid) !== -1) voters.push(mgrId);
          });
          var cnt = voters.length;
          var pct = Math.round((cnt / totalVoters) * 100);
          return Object.assign({}, b, {
            voteCount: cnt,
            votePct: pct,
            voters: voters,
            metSupermajority: cnt >= 7
          });
        });

        // Sort by voteCount DESC, tie-break by legacyScore DESC
        tally.sort(function(a, b) {
          if (b.voteCount !== a.voteCount) return b.voteCount - a.voteCount;
          return (Number(b.legacyScore) || 0) - (Number(a.legacyScore) || 0);
        });

        var inductedThisYear = [];
        var nextCarryover = [];
        tally.forEach(function(item) {
          if (item.metSupermajority && inductedThisYear.length < maxClassCap) {
            item.outcome = "INDUCTED";
            var plaque = Object.assign({}, item, {
              classYear: hofState.classYear,
              inductedAt: Date.now()
            });
            inductedThisYear.push(plaque);
            var alreadyInducted = hofState.inductedPlayers.some(function(ip) { return String(ip.playerId) === String(item.playerId); });
            if (!alreadyInducted) hofState.inductedPlayers.push(plaque);
          } else if (item.voteCount < 2) {
            // < 2 votes (<20%) permanently drops finalist from future ballots
            item.outcome = "DROPPED_UNDER_20_PCT";
            if (hofState.droppedPlayerIds.indexOf(String(item.playerId)) === -1) {
              hofState.droppedPlayerIds.push(String(item.playerId));
            }
          } else {
            var nextYrs = (Number(item.yearsOnBallot) || 1) + 1;
            if (nextYrs > 5) {
              item.outcome = "EXPIRED_5_YEAR_CAP";
              if (hofState.droppedPlayerIds.indexOf(String(item.playerId)) === -1) {
                hofState.droppedPlayerIds.push(String(item.playerId));
              }
            } else {
              item.outcome = "CARRYOVER_NEXT_YEAR";
              nextCarryover.push(Object.assign({}, item, { yearsOnBallot: nextYrs, isCarryover: true }));
            }
          }
        });

        hofState.stage = body.setIdle ? "IDLE" : "RESULTS_PUBLISHED";
        hofState.carryoverPool = nextCarryover;
        hofState.lastResults = {
          classYear: hofState.classYear,
          closedAt: Date.now(),
          totalManagersVoted: Object.keys(voteMap).length,
          ballotResults: tally,
          inductedCount: inductedThisYear.length
        };
        hofState.updatedAt = Date.now();
        baseStateFunnel.espnData.hofState = hofState;

        var inductNames = inductedThisYear.map(function(p) { return p.playerName + " (" + p.votePct + "%)"; }).join(", ");
        appendServerActivityEvent(
          baseStateFunnel,
          "HOF_RESULTS",
          "Class of " + hofState.classYear + " HOF Ballot Results Official",
          inductedThisYear.length > 0
            ? ("Inducted into the Gamecock Mafia Player HOF: " + inductNames + "!")
            : ("No candidate reached the 7/10 (70%) supermajority for Class of " + hofState.classYear + "."),
          { year: hofState.classYear }
        );
        dispatchOneSignalPush(baseStateFunnel, {
          category: "HOF",
          title: "🏆 Class of " + hofState.classYear + " HOF Results Revealed!",
          message: inductedThisYear.length > 0
            ? ("Welcome to The Pantheon: " + inductNames)
            : "Vote percentages revealed in The Vault -> Player HOF."
        });
        baseStateFunnel.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseStateFunnel, "CLOSE_HOF_BALLOT:" + hofState.classYear);
        return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateFunnel) });
      }
    }

    // ------------------------------------------------------------------------
    // ACTION 1J: LEO & SYETH SYNCED CENTRAL BANKER MODE (UPDATE_DUES_STATUS)
    // Authenticated via Leo's or Syeth's personal Team PIN OR Commissioner/Co-Commish PIN
    // ------------------------------------------------------------------------
    if (action === "UPDATE_DUES_STATUS") {
      var baseStateDues = stored || body.fallbackBaseState;
      if (!baseStateDues || !Array.isArray(baseStateDues.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      var isCommishDues = isAuthorizedCommishOrCoCommish(stored, body).authorized;
      var actorForDues = String(body.actorTeamId || "").toLowerCase();
      var providedPinHash = body.teamPinHash ? String(body.teamPinHash).toLowerCase() : "";
      var isBankerAuthenticated = false;
      if (actorForDues === "leo" || actorForDues === "syeth") {
        var bProf = (baseStateDues.teamProfiles && baseStateDues.teamProfiles[actorForDues]) || {};
        if (bProf.teamPinHash && /^[a-f0-9]{64}$/i.test(bProf.teamPinHash)) {
          isBankerAuthenticated = (providedPinHash === bProf.teamPinHash.toLowerCase());
        } else {
          isBankerAuthenticated = true;
        }
      }
      if (!isCommishDues && !isBankerAuthenticated) {
        return jsonResponse({ ok: false, error: "Unauthorized: Only Central Banker Leo, Commissioner Syeth, or Commissioner Mode can update the Dues Ledger." });
      }

      if (!baseStateDues.duesLedger || typeof baseStateDues.duesLedger !== "object") {
        baseStateDues.duesLedger = {
          seasonYear: Number(baseStateDues.baseUpcomingYear) || 2027,
          buyInPerManager: 150,
          totalPot: 1500,
          isLockedFinal: false,
          lockedAt: 0,
          deadlineIso: "",
          managers: {}
        };
      }
      var ledger = baseStateDues.duesLedger;
      if (!ledger.managers || typeof ledger.managers !== "object") ledger.managers = {};

      if (body.seasonYear) ledger.seasonYear = Number(body.seasonYear) || ledger.seasonYear;
      if (body.deadlineIso !== undefined) ledger.deadlineIso = String(body.deadlineIso || "").slice(0, 40);
      if (body.isLockedFinal !== undefined) {
        ledger.isLockedFinal = Boolean(body.isLockedFinal);
        ledger.lockedAt = ledger.isLockedFinal ? (Number(ledger.lockedAt) || Date.now()) : 0;
      }

      var updaterLabel = isCommishDues ? "commish" : (actorForDues || "leo");
      if (body.managers && typeof body.managers === "object") {
        Object.keys(body.managers).forEach(function(tid) {
          var mIn = body.managers[tid] || {};
          var prevM = ledger.managers[tid] || {};
          ledger.managers[tid] = {
            paid: mIn.paid !== undefined ? Boolean(mIn.paid) : Boolean(prevM.paid),
            amountPaid: mIn.amountPaid !== undefined ? Number(mIn.amountPaid) : (mIn.paid ? 150 : 0),
            note: String(mIn.note !== undefined ? mIn.note : (prevM.note || "")).trim().slice(0, 80),
            dockedWins: Math.max(0, Number(mIn.dockedWins !== undefined ? mIn.dockedWins : (prevM.dockedWins || 0))),
            updatedAt: Date.now(),
            updatedBy: updaterLabel
          };
        });
      } else if (body.targetTeamId) {
        var tTid = String(body.targetTeamId);
        var prevTarget = ledger.managers[tTid] || {};
        ledger.managers[tTid] = {
          paid: body.paid !== undefined ? Boolean(body.paid) : Boolean(prevTarget.paid),
          amountPaid: body.amountPaid !== undefined ? Number(body.amountPaid) : (body.paid ? 150 : 0),
          note: String(body.note !== undefined ? body.note : (prevTarget.note || "")).trim().slice(0, 80),
          dockedWins: Math.max(0, Number(body.dockedWins !== undefined ? body.dockedWins : (prevTarget.dockedWins || 0))),
          updatedAt: Date.now(),
          updatedBy: updaterLabel
        };
      }

      ledger.teams = Object.assign({}, ledger.teams || {}, ledger.managers || {});
      var sYrKey = String(ledger.seasonYear || 2027);
      ledger[sYrKey] = { teams: Object.assign({}, ledger.teams) };
      ledger.updatedAt = Date.now();
      baseStateDues.duesLedger = ledger;
      baseStateDues.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateDues, "UPDATE_DUES_STATUS:" + updaterLabel.toUpperCase());
      return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateDues) });
    }

    // ------------------------------------------------------------------------
    // ACTION 1K: COMMISSIONER CUSTOM PUSH BLAST & ONESIGNAL CONFIG (SEC 5)
    // ------------------------------------------------------------------------
    if (action === "SEND_PUSH_BLAST" || action === "SAVE_ONESIGNAL_CONFIG") {
      var baseStatePush = stored || body.fallbackBaseState;
      if (!baseStatePush) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      var authPush = isAuthorizedCommishOrCoCommish(stored, body);
      if (!authPush.authorized) {
        return jsonResponse({ ok: false, error: "Only the Commissioner can configure OneSignal keys or send league push blasts." });
      }

      if (action === "SAVE_ONESIGNAL_CONFIG") {
        if (!authPush.isMasterCommish) {
          return jsonResponse({ ok: false, error: "Only Master Commissioner Syeth can modify OneSignal API keys." });
        }
        if (!baseStatePush.pushConfig || typeof baseStatePush.pushConfig !== "object") {
          baseStatePush.pushConfig = {};
        }
        if (body.oneSignalAppId !== undefined) {
          baseStatePush.pushConfig.oneSignalAppId = String(body.oneSignalAppId || "").trim().slice(0, 80);
          baseStatePush.oneSignalAppId = baseStatePush.pushConfig.oneSignalAppId;
        }
        if (body.oneSignalRestKey !== undefined && String(body.oneSignalRestKey).trim() !== "") {
          baseStatePush.pushConfig.oneSignalRestKey = String(body.oneSignalRestKey).trim().slice(0, 160);
          baseStatePush.oneSignalRestKey = baseStatePush.pushConfig.oneSignalRestKey;
        }
        try {
          var sp = PropertiesService.getScriptProperties();
          if (baseStatePush.pushConfig.oneSignalAppId) sp.setProperty("ONESIGNAL_APP_ID", baseStatePush.pushConfig.oneSignalAppId);
          if (baseStatePush.pushConfig.oneSignalRestKey) sp.setProperty("ONESIGNAL_REST_KEY", baseStatePush.pushConfig.oneSignalRestKey);
        } catch (ignoreSpErr) {}
        baseStatePush.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseStatePush, "SAVE_ONESIGNAL_CONFIG");
        return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStatePush) });
      }

      var audience = String(body.audience || "ALL").toUpperCase(); // ALL | UNLOCKED_KEEPERS | UNPAID_DUES
      var targetIds = null;
      if (audience === "UNLOCKED_KEEPERS") {
        targetIds = [];
        (baseStatePush.teams || []).forEach(function(t) {
          var pf = (baseStatePush.teamProfiles && baseStatePush.teamProfiles[t.id]) || {};
          var kCnt = Array.isArray(pf.keepers) ? pf.keepers.length : 0;
          if (kCnt < 12 || !pf.keepersConfirmedAt) targetIds.push(t.id);
        });
      } else if (audience === "UNPAID_DUES") {
        targetIds = [];
        var dMgrs = (baseStatePush.duesLedger && baseStatePush.duesLedger.managers) || {};
        (baseStatePush.teams || []).forEach(function(t) {
          if (!dMgrs[t.id] || !dMgrs[t.id].paid) targetIds.push(t.id);
        });
      }

      var blastTitle = String(body.title || "Gamecock Mafia Commissioner Alert").trim().slice(0, 90);
      var blastMsg = String(body.message || "").trim().slice(0, 240);
      if (!blastMsg) {
        return jsonResponse({ ok: false, error: "Enter a message to send a League Alert." });
      }
      var pushResult = dispatchOneSignalPush(baseStatePush, {
        category: "CUSTOM",
        urgent: Boolean(body.bypassQuietHours),
        targetTeamIds: targetIds,
        title: blastTitle,
        message: blastMsg
      });
      appendServerActivityEvent(
        baseStatePush,
        "COMMISH_ALERT",
        blastTitle + " (" + audience + ")",
        blastMsg,
        { actorTeamId: authPush.actorTeamId || "commish" }
      );
      baseStatePush.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStatePush, "SEND_PUSH_BLAST:" + audience);
      return jsonResponse({ ok: true, pushResult: pushResult, state: sanitizeStateForClient(baseStatePush) });
    }

    // ------------------------------------------------------------------------
    // ACTION 1L: 1-CLICK TIME MACHINE ROLLBACK (ADMIN_ROLLBACK — SEC 10)
    // Silent restore across 30 saved snapshots (no push spam) + ADMIN_ROLLBACK log
    // ------------------------------------------------------------------------
    if (action === "ADMIN_ROLLBACK") {
      var isCommishRb = isAuthorizedCommishOrCoCommish(stored, body).authorized;
      if (!isCommishRb) {
        return jsonResponse({ ok: false, error: "Invalid Commissioner PIN authorization for Time Machine Rollback." });
      }
      var targetRow = Number(body.rowNumber) || 5;
      if (targetRow < 5 || targetRow > 35) {
        return jsonResponse({ ok: false, error: "Snapshot row must be between 5 and 35." });
      }
      var restoredState = readSnapshotStateAtRow(sheets.stateSheet, targetRow);
      if (!restoredState || !Array.isArray(restoredState.teams)) {
        return jsonResponse({ ok: false, error: "Could not read valid state JSON from snapshot row " + targetRow + "." });
      }
      // Preserve current auth hashes, Co-Commissioners & OneSignal keys
      if (stored && stored.commishPinHash) restoredState.commishPinHash = stored.commishPinHash;
      if (stored && stored.coCommissioners) restoredState.coCommissioners = stored.coCommissioners;
      if (stored && stored.coCommishPinHashes) restoredState.coCommishPinHashes = stored.coCommishPinHashes;
      if (stored && stored.coCommishPinVerifiers) restoredState.coCommishPinVerifiers = stored.coCommishPinVerifiers;
      if (stored && stored.leaguePassphraseHash) restoredState.leaguePassphraseHash = stored.leaguePassphraseHash;
      if (stored && stored.oneSignalRestKey) restoredState.oneSignalRestKey = stored.oneSignalRestKey;
      if (stored && stored.pushConfig) restoredState.pushConfig = stored.pushConfig;

      var snapMeta = sheets.stateSheet.getRange(targetRow, 1, 1, 3).getValues()[0] || [];
      var snapLabel = String(snapMeta[2] || ("Row " + targetRow));
      var snapIso = String(snapMeta[1] || "");
      appendServerActivityEvent(
        restoredState,
        "ADMIN_ROLLBACK",
        "Time Machine Rollback Executed",
        "Commissioner restored league state to snapshot \"" + snapLabel + "\" (" + snapIso + ").",
        { actorTeamId: "commish" }
      );
      restoredState.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, restoredState, "ADMIN_ROLLBACK:ROW_" + targetRow);
      return jsonResponse({ ok: true, state: sanitizeStateForClient(restoredState), snapshots: listSnapshotsFromSheet(sheets.stateSheet) });
    }

    // ------------------------------------------------------------------------
    // ACTION 1M: 1-CLICK SEASON ROLLOVER WIZARD (SEASON_ROLLOVER — SEC 10)
    // Archives completed season, advances 2-year pick window, resets Dues Ledger, opens HOF Stage 1
    // ------------------------------------------------------------------------
    if (action === "SEASON_ROLLOVER") {
      var baseStateRoll = stored || body.fallbackBaseState;
      if (!baseStateRoll || !Array.isArray(baseStateRoll.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      var isCommishRoll = isAuthorizedCommishOrCoCommish(stored, body).authorized;
      if (!isCommishRoll) {
        return jsonResponse({ ok: false, error: "Only the Commissioner can execute the Season Rollover Wizard." });
      }
      var prevUpcoming = Number(baseStateRoll.baseUpcomingYear) || 2027;
      var nextUpcoming = Number(body.nextUpcomingYear) || (prevUpcoming + 1);
      baseStateRoll.baseUpcomingYear = nextUpcoming;
      baseStateRoll.activeYear = nextUpcoming;

      // Reset Central Banker Dues Ledger for the new season
      var resetMgrs = {};
      (baseStateRoll.teams || []).forEach(function(t) {
        resetMgrs[t.id] = { paid: false, amountPaid: 0, note: "", dockedWins: 0, updatedAt: Date.now() };
      });
      baseStateRoll.duesLedger = {
        seasonYear: nextUpcoming,
        buyInPerManager: 150,
        totalPot: 1500,
        isLockedFinal: false,
        lockedAt: 0,
        deadlineIso: "",
        managers: resetMgrs,
        updatedAt: Date.now()
      };

      // Open Stage 1 of Offseason Player HOF Funnel for the new class
      if (!baseStateRoll.espnData || typeof baseStateRoll.espnData !== "object") baseStateRoll.espnData = {};
      var prevHof = (baseStateRoll.espnData.hofState && typeof baseStateRoll.espnData.hofState === "object")
        ? baseStateRoll.espnData.hofState
        : { inductedPlayers: [], carryoverPool: [], droppedPlayerIds: [] };
      baseStateRoll.espnData.hofState = {
        stage: "IDLE",
        scoutingOpen: true,
        classYear: nextUpcoming,
        candidates: Array.isArray(prevHof.carryoverPool) ? prevHof.carryoverPool.slice() : [],
        ballot: [],
        votesByManager: {},
        inductedPlayers: Array.isArray(prevHof.inductedPlayers) ? prevHof.inductedPlayers : [],
        carryoverPool: Array.isArray(prevHof.carryoverPool) ? prevHof.carryoverPool : [],
        droppedPlayerIds: Array.isArray(prevHof.droppedPlayerIds) ? prevHof.droppedPlayerIds : [],
        lastResults: prevHof.lastResults || null,
        updatedAt: Date.now()
      };

      // Roll prior-season AI Weekly Recaps into ESPN season archive and clear active-season aiCache
      ensureAiStructures_(baseStateRoll);
      var prevSeasonKey = String(prevUpcoming - 1);
      if (baseStateRoll.espnData.seasons && baseStateRoll.espnData.seasons[prevSeasonKey]) {
        var archivedRecaps = {};
        Object.keys(baseStateRoll.aiCache.weeklyRecaps || {}).forEach(function(rk) {
          if (String(rk).indexOf(prevSeasonKey + "_") === 0) {
            archivedRecaps[rk] = baseStateRoll.aiCache.weeklyRecaps[rk];
          }
        });
        if (Object.keys(archivedRecaps).length > 0) {
          baseStateRoll.espnData.seasons[prevSeasonKey].archivedRecaps = Object.assign(
            {},
            baseStateRoll.espnData.seasons[prevSeasonKey].archivedRecaps || {},
            archivedRecaps
          );
        }
      }
      baseStateRoll.aiCache.weeklyRecaps = {};
      baseStateRoll.aiCache.playoffAdvisors = {};
      trimAiCache_(baseStateRoll);

      appendServerActivityEvent(
        baseStateRoll,
        "SEASON_ROLLOVER",
        "Season Rollover Complete: Advanced to " + nextUpcoming + "–" + (nextUpcoming + 1) + " Window",
        "Archived " + (prevUpcoming - 1) + " standings/awards, rolled draft pick window to " + nextUpcoming + "–" + (nextUpcoming + 1) + ", reset Central Banker Dues Ledger, and opened HOF Class of " + nextUpcoming + " Scouting.",
        { year: nextUpcoming, actorTeamId: "commish" }
      );
      baseStateRoll.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateRoll, "SEASON_ROLLOVER:" + nextUpcoming);
      return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateRoll) });
    }

    // ------------------------------------------------------------------------
    // ACTION 1N: CO-COMMISSIONER GOVERNANCE & INDIVIDUAL CO-COMMISH PINS (WAVE 4)
    // Master Commissioner Syeth assigns Co-Commissioners; Co-Commissioners set/use separate PINs
    // ------------------------------------------------------------------------
    if (action === "SET_CO_COMMISSIONERS" || action === "SET_MY_CO_COMMISH_PIN") {
      var baseStateCo = stored || body.fallbackBaseState;
      if (!baseStateCo || !Array.isArray(baseStateCo.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      if (!Array.isArray(baseStateCo.coCommissioners)) baseStateCo.coCommissioners = [];
      if (!baseStateCo.coCommishPinHashes || typeof baseStateCo.coCommishPinHashes !== "object") baseStateCo.coCommishPinHashes = {};
      if (!baseStateCo.coCommishPinVerifiers || typeof baseStateCo.coCommishPinVerifiers !== "object") baseStateCo.coCommishPinVerifiers = {};

      var authCo = isAuthorizedCommishOrCoCommish(stored, body);

      if (action === "SET_CO_COMMISSIONERS") {
        if (!authCo.isMasterCommish) {
          return jsonResponse({ ok: false, error: "Only Master Commissioner Syeth can assign or revoke Co-Commissioner powers." });
        }
        var validMgrSet = {};
        baseStateCo.teams.forEach(function(t) { if (t && t.id && t.id !== "syeth") validMgrSet[t.id] = true; });
        var incomingCoList = Array.isArray(body.coCommissioners) ? body.coCommissioners : [];
        var cleanCoList = [];
        incomingCoList.forEach(function(cid) {
          var sCid = String(cid || "").toLowerCase().trim();
          if (validMgrSet[sCid] && cleanCoList.indexOf(sCid) === -1) cleanCoList.push(sCid);
        });
        baseStateCo.coCommissioners = cleanCoList;

        // Clean up hashes/verifiers for revoked Co-Commissioners
        Object.keys(baseStateCo.coCommishPinHashes).forEach(function(k) {
          if (cleanCoList.indexOf(k) === -1) delete baseStateCo.coCommishPinHashes[k];
        });
        Object.keys(baseStateCo.coCommishPinVerifiers).forEach(function(k) {
          if (cleanCoList.indexOf(k) === -1) delete baseStateCo.coCommishPinVerifiers[k];
        });

        // Optional PIN reset or assignment by Syeth for a specific Co-Commissioner
        if (body.resetCoCommishPinTeamId) {
          var rTid = String(body.resetCoCommishPinTeamId).toLowerCase().trim();
          delete baseStateCo.coCommishPinHashes[rTid];
          delete baseStateCo.coCommishPinVerifiers[rTid];
        }
        if (body.setPinForTeamId && body.coCommishPinHash && body.coCommishPinVerifier) {
          var sTid = String(body.setPinForTeamId).toLowerCase().trim();
          if (cleanCoList.indexOf(sTid) !== -1 && /^[a-f0-9]{64}$/i.test(body.coCommishPinHash) && /^[a-f0-9]{64}$/i.test(body.coCommishPinVerifier)) {
            baseStateCo.coCommishPinHashes[sTid] = String(body.coCommishPinHash).toLowerCase();
            baseStateCo.coCommishPinVerifiers[sTid] = String(body.coCommishPinVerifier).toLowerCase();
          }
        }

        baseStateCo.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseStateCo, "SET_CO_COMMISSIONERS:" + (cleanCoList.join(",") || "NONE"));
        return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateCo) });
      }

      if (action === "SET_MY_CO_COMMISH_PIN") {
        var targetCoTid = String(body.actorTeamId || body.teamId || "").toLowerCase().trim();
        if (!targetCoTid || baseStateCo.coCommissioners.indexOf(targetCoTid) === -1) {
          return jsonResponse({ ok: false, error: "This franchise is not assigned Co-Commissioner powers." });
        }
        var existingCoHash = baseStateCo.coCommishPinHashes[targetCoTid] || "";
        var prevCoHash = body.prevCoCommishPinHash ? String(body.prevCoCommishPinHash).toLowerCase() : "";
        if (existingCoHash && !authCo.isMasterCommish && prevCoHash !== existingCoHash.toLowerCase()) {
          return jsonResponse({ ok: false, error: "Current Co-Commissioner PIN required to change PIN (or ask Syeth to reset it)." });
        }
        var newCoHash = String(body.coCommishPinHash || "").toLowerCase();
        var newCoVerifier = String(body.coCommishPinVerifier || "").toLowerCase();
        if (!/^[a-f0-9]{64}$/i.test(newCoHash) || !/^[a-f0-9]{64}$/i.test(newCoVerifier)) {
          return jsonResponse({ ok: false, error: "Invalid Co-Commissioner PIN hash payload." });
        }
        baseStateCo.coCommishPinHashes[targetCoTid] = newCoHash;
        baseStateCo.coCommishPinVerifiers[targetCoTid] = newCoVerifier;
        baseStateCo.updatedAt = Date.now();
        saveAndMirrorToSheets(sheets, baseStateCo, "SET_CO_COMMISH_PIN:" + targetCoTid);
        return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateCo) });
      }
    }

    // ------------------------------------------------------------------------
    // ACTION 1O: NATIVE GEMINI AI VOICE SETTINGS (COMMISSIONER / CO-COMMISH)
    // ------------------------------------------------------------------------
    if (action === "SAVE_AI_VOICE_SETTINGS") {
      var authAiCfg = isAuthorizedCommishOrCoCommish(stored, body);
      if (!authAiCfg.authorized) {
        return jsonResponse({ ok: false, error: "Commissioner or Co-Commissioner PIN required to update AI Voice Settings." });
      }
      var baseStateAiCfg = stored || body.fallbackBaseState;
      if (!baseStateAiCfg || !Array.isArray(baseStateAiCfg.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      ensureAiStructures_(baseStateAiCfg);
      var incVoice = (body.aiVoice && typeof body.aiVoice === "object") ? body.aiVoice : {};
      var nextPreset = String(incVoice.preset || baseStateAiCfg.settings.aiVoice.preset || "SNARKY_COMMISH").toUpperCase();
      if (!AI_TONE_PRESETS[nextPreset]) nextPreset = "SNARKY_COMMISH";
      baseStateAiCfg.settings.aiVoice = {
        preset: nextPreset,
        customLore: String(incVoice.customLore !== undefined ? incVoice.customLore : baseStateAiCfg.settings.aiVoice.customLore || "").slice(0, 1500),
        hofAiEnabled: incVoice.hofAiEnabled !== undefined ? Boolean(incVoice.hofAiEnabled) : Boolean(baseStateAiCfg.settings.aiVoice.hofAiEnabled),
        updatedAt: Date.now()
      };
      baseStateAiCfg.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateAiCfg, "SAVE_AI_VOICE:" + nextPreset);
      return jsonResponse({ ok: true, state: sanitizeStateForClient(baseStateAiCfg) });
    }

    // ------------------------------------------------------------------------
    // ACTION 1P: AI_GENERATE_WEEKLY_RECAP (COMMISSIONER / CO-COMMISH PIN)
    // ------------------------------------------------------------------------
    if (action === "AI_GENERATE_WEEKLY_RECAP") {
      var authRecap = isAuthorizedCommishOrCoCommish(stored, body);
      if (!authRecap.authorized) {
        return jsonResponse({ ok: false, error: "Commissioner or Co-Commissioner PIN required to generate Weekly Recaps." });
      }
      var baseStateRecap = stored || body.fallbackBaseState;
      if (!baseStateRecap || !Array.isArray(baseStateRecap.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      ensureAiStructures_(baseStateRecap);
      var recapSeason = Number(body.season || (baseStateRecap.espnData && baseStateRecap.espnData.currentSeason) || 2026);
      var recapWeek = Number(body.week) || 1;
      var recapKey = recapSeason + "_" + recapWeek;
      var forceRegenRecap = Boolean(body.forceRegenerate);

      if (!forceRegenRecap && baseStateRecap.aiCache.weeklyRecaps[recapKey]) {
        return jsonResponse({ ok: true, cached: true, recapKey: recapKey, recap: baseStateRecap.aiCache.weeklyRecaps[recapKey], aiCache: baseStateRecap.aiCache });
      }

      var recapSchema = {
        type: "OBJECT",
        properties: {
          headline: { type: "STRING" },
          subhead: { type: "STRING" },
          matchupDigest: {
            type: "ARRAY",
            items: {
              type: "OBJECT",
              properties: {
                matchup: { type: "STRING" },
                summary: { type: "STRING" }
              },
              required: ["matchup", "summary"]
            }
          },
          fraudAndBenchWatch: { type: "STRING" },
          playoffPulse: { type: "STRING" }
        },
        required: ["headline", "subhead", "matchupDigest", "fraudAndBenchWatch", "playoffPulse"]
      };

      var recapDirective = [
        "Generate a concise, punchy Week " + recapWeek + " (" + recapSeason + " season) Gamecock Mafia fantasy football recap strictly from the supplied JSON telemetry.",
        "Total copy across ALL fields combined MUST be <= 300 words.",
        "matchupDigest must have 1 crisp sentence per played matchup in the payload (exact teams and scores from JSON).",
        "fraudAndBenchWatch highlights All-Play luck frauds, worst bench blunders, King of the Week positional explosion, or notable FAAB bids/busts ONLY if present in the JSON.",
        "playoffPulse summarizes the biggest playoff seed/division race swing based ONLY on the JSON standings."
      ].join(" ");

      var gemRecap = callGeminiStructuredJson_(baseStateRecap, recapDirective, body.telemetry || {}, recapSchema, { temperature: 0.72 });
      if (!gemRecap.ok) {
        return jsonResponse({ ok: false, code: gemRecap.code, error: gemRecap.error });
      }

      var savedRecap = {
        season: recapSeason,
        week: recapWeek,
        headline: String(gemRecap.data.headline || "").slice(0, 180),
        subhead: String(gemRecap.data.subhead || "").slice(0, 260),
        matchupDigest: Array.isArray(gemRecap.data.matchupDigest) ? gemRecap.data.matchupDigest.slice(0, 6).map(function(m) {
          return {
            matchup: String((m && m.matchup) || "").slice(0, 120),
            summary: String((m && m.summary) || "").slice(0, 280)
          };
        }) : [],
        fraudAndBenchWatch: String(gemRecap.data.fraudAndBenchWatch || "").slice(0, 520),
        playoffPulse: String(gemRecap.data.playoffPulse || "").slice(0, 420),
        tonePreset: gemRecap.tonePreset,
        toneName: gemRecap.toneName,
        model: gemRecap.model,
        generatedBy: authRecap.actorTeamId || "commish",
        generatedAt: Date.now()
      };

      baseStateRecap.aiCache.weeklyRecaps[recapKey] = savedRecap;
      trimAiCache_(baseStateRecap);
      baseStateRecap.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateRecap, "AI_WEEKLY_RECAP:" + recapKey);
      return jsonResponse({ ok: true, cached: false, recapKey: recapKey, recap: savedRecap, aiCache: baseStateRecap.aiCache, state: sanitizeStateForClient(baseStateRecap) });
    }

    // ------------------------------------------------------------------------
    // ACTION 1Q: AI_GRADE_TRADE (CACHE-FIRST + 25/HR MEMBER THROTTLE)
    // ------------------------------------------------------------------------
    if (action === "AI_GRADE_TRADE") {
      var baseStateTradeAi = stored || body.fallbackBaseState;
      if (!baseStateTradeAi || !Array.isArray(baseStateTradeAi.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      ensureAiStructures_(baseStateTradeAi);
      var tKey = String(body.tradeId || body.cacheKey || "").trim().slice(0, 80);
      if (!tKey) {
        return jsonResponse({ ok: false, error: "Missing tradeId/cacheKey for AI trade grade." });
      }
      var authTradeAi = isAuthorizedCommishOrCoCommish(stored, body);
      var forceTradeRegen = Boolean(body.forceRegenerate && authTradeAi.authorized);

      if (!forceTradeRegen && baseStateTradeAi.aiCache.tradeGrades[tKey]) {
        return jsonResponse({ ok: true, cached: true, tradeKey: tKey, grade: baseStateTradeAi.aiCache.tradeGrades[tKey], aiCache: baseStateTradeAi.aiCache });
      }

      var quotaTrade = checkAndConsumeMemberAiQuota_();
      if (!quotaTrade.allowed) {
        return jsonResponse({ ok: false, code: "AI_RATE_LIMIT", error: "Hourly league AI limit (" + quotaTrade.limit + " uncached calls/hr) reached. Cached grades still load instantly—try again next hour." });
      }

      var tradeSchema = {
        type: "OBJECT",
        properties: {
          teamAGrade: { type: "STRING" },
          teamBGrade: { type: "STRING" },
          winnerTag: { type: "STRING" },
          verdict: { type: "STRING" },
          keeperCapNote: { type: "STRING" }
        },
        required: ["teamAGrade", "teamBGrade", "winnerTag", "verdict", "keeperCapNote"]
      };

      var tradeDirective = [
        "Grade and roast this Gamecock Mafia keeper-league trade strictly using the provided JSON telemetry (players, positional starter ranks, 12-keeper cap impact, draft picks, conditional/Salmon pool rules, 3-Pillar Power Ranks, and Contender/Rebuilder status).",
        "teamAGrade and teamBGrade should be crisp letter grades (e.g., A-, B+, C, D-).",
        "winnerTag is a punchy 2-6 word tag naming the winner or calling it even.",
        "Total copy across ALL fields combined MUST be <= 120 words."
      ].join(" ");

      var gemTrade = callGeminiStructuredJson_(baseStateTradeAi, tradeDirective, body.telemetry || {}, tradeSchema, { temperature: 0.7 });
      if (!gemTrade.ok) {
        return jsonResponse({ ok: false, code: gemTrade.code, error: gemTrade.error });
      }

      var savedGrade = {
        tradeKey: tKey,
        teamAGrade: String(gemTrade.data.teamAGrade || "B").slice(0, 12),
        teamBGrade: String(gemTrade.data.teamBGrade || "B").slice(0, 12),
        winnerTag: String(gemTrade.data.winnerTag || "").slice(0, 80),
        verdict: String(gemTrade.data.verdict || "").slice(0, 480),
        keeperCapNote: String(gemTrade.data.keeperCapNote || "").slice(0, 260),
        tonePreset: gemTrade.tonePreset,
        toneName: gemTrade.toneName,
        model: gemTrade.model,
        generatedAt: Date.now()
      };

      baseStateTradeAi.aiCache.tradeGrades[tKey] = savedGrade;
      trimAiCache_(baseStateTradeAi);
      baseStateTradeAi.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateTradeAi, "AI_GRADE_TRADE:" + tKey);
      return jsonResponse({ ok: true, cached: false, tradeKey: tKey, grade: savedGrade, aiCache: baseStateTradeAi.aiCache });
    }

    // ------------------------------------------------------------------------
    // ACTION 1R: AI_PLAYOFF_ADVISOR (CACHE-FIRST + 25/HR MEMBER THROTTLE)
    // ------------------------------------------------------------------------
    if (action === "AI_PLAYOFF_ADVISOR") {
      var baseStatePoAi = stored || body.fallbackBaseState;
      if (!baseStatePoAi || !Array.isArray(baseStatePoAi.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      ensureAiStructures_(baseStatePoAi);
      var poSeason = Number(body.season || (baseStatePoAi.espnData && baseStatePoAi.espnData.currentSeason) || 2026);
      var poMgr = String(body.managerId || "").toLowerCase().trim().slice(0, 24);
      var poWeek = Number(body.week) || 1;
      if (!poMgr) {
        return jsonResponse({ ok: false, error: "Missing managerId for Playoff Advisor." });
      }
      var poKey = poSeason + "_" + poMgr + "_W" + poWeek;
      var authPoAi = isAuthorizedCommishOrCoCommish(stored, body);
      var forcePoRegen = Boolean(body.forceRegenerate && authPoAi.authorized);

      if (!forcePoRegen && baseStatePoAi.aiCache.playoffAdvisors[poKey]) {
        return jsonResponse({ ok: true, cached: true, advisorKey: poKey, brief: baseStatePoAi.aiCache.playoffAdvisors[poKey], aiCache: baseStatePoAi.aiCache });
      }

      var quotaPo = checkAndConsumeMemberAiQuota_();
      if (!quotaPo.allowed) {
        return jsonResponse({ ok: false, code: "AI_RATE_LIMIT", error: "Hourly league AI limit (" + quotaPo.limit + " uncached calls/hr) reached. Try again next hour." });
      }

      var poSchema = {
        type: "OBJECT",
        properties: {
          leverageBullet: { type: "STRING" },
          rootingBullet: { type: "STRING" },
          rosterAngleBullet: { type: "STRING" }
        },
        required: ["leverageBullet", "rootingBullet", "rosterAngleBullet"]
      };

      var poDirective = [
        "Produce a sharp 3-bullet Playoff Path advisor brief for this Gamecock Mafia manager strictly from the supplied JSON telemetry (exact seed, overall & divisional record, magic/tragic numbers, top rooting guide games, remaining opponents, and starter-tier surplus/deficits).",
        "Remember League 30215 locks Division Winners to Seeds #1-#2 and breaks ALL ties by Divisional Record BEFORE Points For.",
        "Total copy across ALL 3 bullets combined MUST be <= 130 words."
      ].join(" ");

      var gemPo = callGeminiStructuredJson_(baseStatePoAi, poDirective, body.telemetry || {}, poSchema, { temperature: 0.65 });
      if (!gemPo.ok) {
        return jsonResponse({ ok: false, code: gemPo.code, error: gemPo.error });
      }

      var savedBrief = {
        advisorKey: poKey,
        season: poSeason,
        managerId: poMgr,
        week: poWeek,
        leverageBullet: String(gemPo.data.leverageBullet || "").slice(0, 320),
        rootingBullet: String(gemPo.data.rootingBullet || "").slice(0, 320),
        rosterAngleBullet: String(gemPo.data.rosterAngleBullet || "").slice(0, 320),
        tonePreset: gemPo.tonePreset,
        toneName: gemPo.toneName,
        model: gemPo.model,
        generatedAt: Date.now()
      };

      baseStatePoAi.aiCache.playoffAdvisors[poKey] = savedBrief;
      trimAiCache_(baseStatePoAi);
      baseStatePoAi.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStatePoAi, "AI_PLAYOFF_ADVISOR:" + poKey);
      return jsonResponse({ ok: true, cached: false, advisorKey: poKey, brief: savedBrief, aiCache: baseStatePoAi.aiCache });
    }

    // ------------------------------------------------------------------------
    // ACTION 1S: AI_GENERATE_HOF_BIO (STANDBY FEATURE FLAG + COMMISH PIN)
    // ------------------------------------------------------------------------
    if (action === "AI_GENERATE_HOF_BIO") {
      var authHofAi = isAuthorizedCommishOrCoCommish(stored, body);
      if (!authHofAi.authorized) {
        return jsonResponse({ ok: false, error: "Commissioner or Co-Commissioner PIN required to generate Hall of Fame Plaque Bios." });
      }
      var baseStateHofAi = stored || body.fallbackBaseState;
      if (!baseStateHofAi || !Array.isArray(baseStateHofAi.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      ensureAiStructures_(baseStateHofAi);
      if (!baseStateHofAi.settings.aiVoice.hofAiEnabled) {
        return jsonResponse({ ok: false, code: "HOF_AI_DISABLED", error: "HOF Bio AI is currently in Standby mode. Enable HOF_AI_ENABLED in Commissioner AI Voice Settings when inductions begin." });
      }
      var hofPid = String(body.playerId || "").trim().slice(0, 64);
      if (!hofPid) {
        return jsonResponse({ ok: false, error: "Missing playerId for HOF Bio." });
      }
      var forceHofRegen = Boolean(body.forceRegenerate);
      if (!forceHofRegen && baseStateHofAi.aiCache.hofBios[hofPid]) {
        return jsonResponse({ ok: true, cached: true, playerId: hofPid, bio: baseStateHofAi.aiCache.hofBios[hofPid], aiCache: baseStateHofAi.aiCache });
      }

      var hofSchema = {
        type: "OBJECT",
        properties: {
          plaqueHeadline: { type: "STRING" },
          inductionBio: { type: "STRING" },
          signatureStatLine: { type: "STRING" }
        },
        required: ["plaqueHeadline", "inductionBio", "signatureStatLine"]
      };

      var hofDirective = "Write a prestigious yet grounded Gamecock Mafia Hall of Fame bronze plaque inscription strictly from the supplied player career telemetry JSON. Total copy <= 120 words.";
      var gemHof = callGeminiStructuredJson_(baseStateHofAi, hofDirective, body.telemetry || {}, hofSchema, { temperature: 0.6 });
      if (!gemHof.ok) {
        return jsonResponse({ ok: false, code: gemHof.code, error: gemHof.error });
      }

      var savedBio = {
        playerId: hofPid,
        plaqueHeadline: String(gemHof.data.plaqueHeadline || "").slice(0, 120),
        inductionBio: String(gemHof.data.inductionBio || "").slice(0, 560),
        signatureStatLine: String(gemHof.data.signatureStatLine || "").slice(0, 180),
        tonePreset: gemHof.tonePreset,
        toneName: gemHof.toneName,
        model: gemHof.model,
        generatedAt: Date.now()
      };

      baseStateHofAi.aiCache.hofBios[hofPid] = savedBio;
      trimAiCache_(baseStateHofAi);
      baseStateHofAi.updatedAt = Date.now();
      saveAndMirrorToSheets(sheets, baseStateHofAi, "AI_HOF_BIO:" + hofPid);
      return jsonResponse({ ok: true, cached: false, playerId: hofPid, bio: savedBio, aiCache: baseStateHofAi.aiCache });
    }

    // ------------------------------------------------------------------------
    // ACTION 1T: AI_PARSE_TRADE_TEXT (STRICT PRE-FILL ONLY — NEVER AUTO-COMMITS)
    // ------------------------------------------------------------------------
    if (action === "AI_PARSE_TRADE_TEXT") {
      var baseStateParse = stored || body.fallbackBaseState;
      if (!baseStateParse || !Array.isArray(baseStateParse.teams)) {
        return jsonResponse({ ok: false, error: "Initialize database first." });
      }
      var rawTradeText = String(body.rawText || "").trim().slice(0, 600);
      if (!rawTradeText) {
        return jsonResponse({ ok: false, error: "Paste group-chat trade text first." });
      }

      var quotaParse = checkAndConsumeMemberAiQuota_();
      if (!quotaParse.allowed) {
        return jsonResponse({ ok: false, code: "AI_RATE_LIMIT", error: "Hourly league AI limit (" + quotaParse.limit + " calls/hr) reached. Please enter the trade manually." });
      }

      var baseYr = Number(baseStateParse.baseUpcomingYear) || 2027;
      var validYears = [baseYr, baseYr + 1];
      var validTeamIds = (baseStateParse.teams || []).map(function(t) { return String(t.id); });
      var teamDirectory = (baseStateParse.teams || []).map(function(t) {
        var prof = (baseStateParse.teamProfiles && baseStateParse.teamProfiles[t.id]) || {};
        return {
          id: t.id,
          name: t.name,
          customTeamName: prof.customTeamName || ""
        };
      });

      var parseSchema = {
        type: "OBJECT",
        properties: {
          teamA: { type: "STRING" },
          teamB: { type: "STRING" },
          summary: { type: "STRING" },
          playersMoved: {
            type: "ARRAY",
            items: {
              type: "OBJECT",
              properties: {
                name: { type: "STRING" },
                fromOwner: { type: "STRING" },
                toOwner: { type: "STRING" }
              },
              required: ["name", "fromOwner", "toOwner"]
            }
          },
          picksMoved: {
            type: "ARRAY",
            items: {
              type: "OBJECT",
              properties: {
                year: { type: "INTEGER" },
                round: { type: "INTEGER" },
                originalOwner: { type: "STRING" },
                fromOwner: { type: "STRING" },
                toOwner: { type: "STRING" }
              },
              required: ["year", "round", "originalOwner", "fromOwner", "toOwner"]
            }
          },
          isConditional: { type: "BOOLEAN" },
          conditionalRule: { type: "STRING" }
        },
        required: ["teamA", "teamB", "summary", "playersMoved", "picksMoved", "isConditional", "conditionalRule"]
      };

      var parseDirective = [
        "Parse the raw group-chat trade text into structured fields for pre-filling a trade form.",
        "Constrain teamA, teamB, originalOwner, fromOwner, and toOwner STRICTLY to one of these exact valid manager IDs: " + validTeamIds.join(", ") + ".",
        "Known manager aliases: Alex = geltz, Shane = boss, David/Maryland = mary.",
        "Constrain pick years STRICTLY to the active 2-year rolling draft window: " + validYears.join(" or ") + " (default to " + baseYr + " if year omitted). Rounds must be integers 1..20.",
        "If the text mentions a pick swap, best-of/worst-of pool, or conditional protection (e.g. '9/10 swap favoring Corbyn'), set isConditional=true and write the exact rule in conditionalRule.",
        "Expand obvious NFL player shorthand (e.g. AJB -> A.J. Brown, D. Monty -> David Montgomery) when matching roster names in the payload."
      ].join(" ");

      var parsePayload = {
        rawTradeText: rawTradeText,
        validManagers: teamDirectory,
        activeDraftYears: validYears,
        rosterHints: body.rosterHints || {}
      };

      var gemParse = callGeminiStructuredJson_(baseStateParse, parseDirective, parsePayload, parseSchema, { temperature: 0.1, skipToneAndSlop: true });
      if (!gemParse.ok) {
        return jsonResponse({ ok: false, code: gemParse.code, error: gemParse.error });
      }

      var pd = gemParse.data || {};
      var validSet = {};
      validTeamIds.forEach(function(id) { validSet[id] = true; });
      var cleanId = function(v) {
        var s = String(v || "").toLowerCase().trim();
        return validSet[s] ? s : "";
      };

      var cleanTeamA = cleanId(pd.teamA);
      var cleanTeamB = cleanId(pd.teamB);
      var cleanPicks = (Array.isArray(pd.picksMoved) ? pd.picksMoved : []).map(function(pk) {
        var yrNum = Number(pk && pk.year);
        if (validYears.indexOf(yrNum) === -1) yrNum = baseYr;
        var rdNum = Math.min(20, Math.max(1, Math.round(Number(pk && pk.round) || 1)));
        var orig = cleanId(pk && pk.originalOwner) || cleanId(pk && pk.fromOwner) || cleanTeamA;
        var from = cleanId(pk && pk.fromOwner) || orig;
        var to = cleanId(pk && pk.toOwner) || (from === cleanTeamA ? cleanTeamB : cleanTeamA);
        if (!orig || !to || from === to) return null;
        return { year: yrNum, round: rdNum, originalOwner: orig, fromOwner: from, toOwner: to };
      }).filter(Boolean);

      var cleanPlayers = (Array.isArray(pd.playersMoved) ? pd.playersMoved : []).map(function(pl) {
        var nm = String((pl && pl.name) || "").trim().slice(0, 64);
        var from = cleanId(pl && pl.fromOwner);
        var to = cleanId(pl && pl.toOwner) || (from === cleanTeamA ? cleanTeamB : cleanTeamA);
        if (!nm || !from || !to || from === to) return null;
        return { name: nm, fromOwner: from, toOwner: to };
      }).filter(Boolean);

      return jsonResponse({
        ok: true,
        parsed: {
          teamA: cleanTeamA,
          teamB: cleanTeamB,
          summary: String(pd.summary || "").trim().slice(0, 220),
          playersMoved: cleanPlayers,
          picksMoved: cleanPicks,
          isConditional: Boolean(pd.isConditional),
          conditionalRule: String(pd.conditionalRule || "").trim().slice(0, 240)
        }
      });
    }

    // ------------------------------------------------------------------------
    // ACTION 2: COMMISSIONER FULL STATE SYNC (ORDER LOCKS, RESOLUTIONS, EDITS)
    // ------------------------------------------------------------------------
    if (action === "ADMIN_FULL_SYNC") {
      var authAdmin = isAuthorizedCommishOrCoCommish(stored, body);

      if (!authAdmin.authorized) {
        return jsonResponse({ ok: false, error: "Invalid Commissioner / Co-Commissioner PIN authorization." });
      }

      var newState = body.state;
      if (!newState || !Array.isArray(newState.teams) || !Array.isArray(newState.trades)) {
        return jsonResponse({ ok: false, error: "Invalid state payload." });
      }

      var syncReason = String(body.reason || "ADMIN_SYNC");
      if (!authAdmin.isMasterCommish && (syncReason === "UPDATE_PIN" || syncReason === "UPDATE_PASSPHRASE" || syncReason.indexOf("RESET") === 0)) {
        return jsonResponse({ ok: false, error: "Only Master Commissioner Syeth can perform master security/reset actions." });
      }
      var isFullReset = authAdmin.isMasterCommish && (syncReason.indexOf("RESET") === 0 || syncReason.indexOf("IMPORT") === 0);
      var baseSnapshotTs = Number(body.clientSnapshotUpdatedAt || newState.updatedAt || 0);

      // Preserve concurrent manager teamProfile saves (including encryptedDraftPrep, keepers, avatar, PIN, strategyTag, pushPrefs)
      if (!isFullReset && stored && stored.teamProfiles && typeof stored.teamProfiles === "object") {
        if (!newState.teamProfiles || typeof newState.teamProfiles !== "object") newState.teamProfiles = {};
        Object.keys(stored.teamProfiles).forEach(function(tid) {
          var sProf = stored.teamProfiles[tid];
          var nProf = newState.teamProfiles[tid];
          if (!sProf) return;
          if (!nProf || (Number(sProf.updatedAt) || 0) > (Number(nProf.updatedAt) || 0)) {
            newState.teamProfiles[tid] = sProf;
          } else {
            if (sProf.teamPinHash && !nProf.teamPinHash && nProf.teamPinVerifier === sProf.teamPinVerifier) {
              nProf.teamPinHash = sProf.teamPinHash;
            }
            if (sProf.encryptedDraftPrep && !nProf.encryptedDraftPrep) {
              nProf.encryptedDraftPrep = sProf.encryptedDraftPrep;
            }
            if (sProf.pushPrefs && !nProf.pushPrefs) nProf.pushPrefs = sProf.pushPrefs;
            if (sProf.strategyTag && !nProf.strategyTag) nProf.strategyTag = normalizeStrategyModeTag(sProf.strategyTag);
            if (sProf.playerTags && (!nProf.playerTags || Object.keys(nProf.playerTags).length === 0)) nProf.playerTags = sProf.playerTags;
          }
        });
      }

      // Preserve concurrent pending trade proposals submitted while Commissioner modal was open
      if (!isFullReset && stored && Array.isArray(stored.pendingTrades) && stored.pendingTrades.length > 0) {
        if (!Array.isArray(newState.pendingTrades)) newState.pendingTrades = [];
        var existingPendingIds = {};
        newState.pendingTrades.forEach(function(pt) { if (pt && pt.id) existingPendingIds[String(pt.id)] = true; });
        stored.pendingTrades.forEach(function(spt) {
          if (spt && spt.id && !existingPendingIds[String(spt.id)] && Number(spt.createdAt || 0) > baseSnapshotTs) {
            newState.pendingTrades.push(spt);
          }
        });
        pruneConflictingPendingProposals(newState);
      }

      // Preserve concurrent trade fairness votes / reactions if stored was updated more recently
      if (!isFullReset && stored && Array.isArray(stored.trades) && Number(stored.updatedAt || 0) > baseSnapshotTs) {
        var storedTradeById = {};
        stored.trades.forEach(function(st) { if (st && st.id) storedTradeById[String(st.id)] = st; });
        newState.trades.forEach(function(nt) {
          var stMatch = nt && nt.id ? storedTradeById[String(nt.id)] : null;
          if (stMatch) {
            if (stMatch.fairnessVotes && (!nt.fairnessVotes || ((nt.fairnessVotes.yes || []).length + (nt.fairnessVotes.no || []).length === 0))) {
              nt.fairnessVotes = stMatch.fairnessVotes;
            }
            if (stMatch.reactions && (!nt.reactions || Object.keys(nt.reactions).length === 0)) {
              nt.reactions = stMatch.reactions;
            }
          }
        });
      }

      // Preserve concurrent playerHofBallots & hofState on espnData
      if (!isFullReset && stored && stored.espnData && typeof stored.espnData === "object") {
        if (!newState.espnData || typeof newState.espnData !== "object") newState.espnData = {};
        if (stored.espnData.playerHofBallots) {
          newState.espnData.playerHofBallots = Object.assign({}, stored.espnData.playerHofBallots, newState.espnData.playerHofBallots || {});
        }
        if (stored.espnData.hofState && !newState.espnData.hofState) {
          newState.espnData.hofState = stored.espnData.hofState;
        }
      }

      // Preserve concurrent duesLedger updates by Leo/Syeth if newer
      if (!isFullReset && stored && stored.duesLedger && typeof stored.duesLedger === "object") {
        if (!newState.duesLedger || (Number(stored.duesLedger.updatedAt) || 0) > (Number(newState.duesLedger.updatedAt) || 0)) {
          newState.duesLedger = stored.duesLedger;
        }
      }

      // Preserve concurrent activityLog entries created after baseSnapshotTs
      if (!isFullReset && stored && Array.isArray(stored.activityLog) && stored.activityLog.length > 0) {
        if (!Array.isArray(newState.activityLog)) newState.activityLog = [];
        var seenActIds = {};
        newState.activityLog.forEach(function(ae) { if (ae && ae.id) seenActIds[String(ae.id)] = true; });
        stored.activityLog.forEach(function(sae) {
          if (sae && sae.id && !seenActIds[String(sae.id)] && Number(sae.timestamp || 0) > baseSnapshotTs) {
            newState.activityLog.push(sae);
          }
        });
        newState.activityLog.sort(function(a, b) { return Number(b.timestamp || 0) - Number(a.timestamp || 0); });
        newState.activityLog = newState.activityLog.slice(0, 100);
      }

      // Guard master security credentials & Co-Commissioners list
      if (!authAdmin.isMasterCommish && stored) {
        newState.commishPinHash = stored.commishPinHash || DEFAULT_COMMISH_PIN_HASH;
        newState.commishPinVerifier = stored.commishPinVerifier || newState.commishPinVerifier;
        newState.leaguePassphraseHash = stored.leaguePassphraseHash || DEFAULT_PASSPHRASE_HASH;
        newState.leaguePassphraseVerifier = stored.leaguePassphraseVerifier || newState.leaguePassphraseVerifier;
        newState.coCommissioners = Array.isArray(stored.coCommissioners) ? stored.coCommissioners : [];
        newState.coCommishPinHashes = stored.coCommishPinHashes || {};
        newState.coCommishPinVerifiers = stored.coCommishPinVerifiers || {};
      } else {
        if (!newState.commishPinHash && stored && stored.commishPinHash) {
          newState.commishPinHash = stored.commishPinHash;
        }
        if (!newState.leaguePassphraseHash && stored && stored.leaguePassphraseHash) {
          newState.leaguePassphraseHash = stored.leaguePassphraseHash;
        }
        if (!Array.isArray(newState.coCommissioners) && stored && Array.isArray(stored.coCommissioners)) {
          newState.coCommissioners = stored.coCommissioners;
        }
        var nextCoHashes = Object.assign({}, (stored && stored.coCommishPinHashes) || {}, newState.coCommishPinHashes || {});
        var activeCo = Array.isArray(newState.coCommissioners) ? newState.coCommissioners : [];
        Object.keys(nextCoHashes).forEach(function(k) {
          if (activeCo.indexOf(k) === -1) delete nextCoHashes[k];
        });
        newState.coCommishPinHashes = nextCoHashes;
      }

      if (!newState.oneSignalRestKey && stored && stored.oneSignalRestKey) {
        newState.oneSignalRestKey = stored.oneSignalRestKey;
      }
      if (stored && stored.pushConfig && stored.pushConfig.oneSignalRestKey) {
        if (!newState.pushConfig || typeof newState.pushConfig !== "object") newState.pushConfig = {};
        if (!newState.pushConfig.oneSignalRestKey) {
          newState.pushConfig.oneSignalRestKey = stored.pushConfig.oneSignalRestKey;
        }
      }
      // Preserve Wave 4 settings.aiVoice and server-generated aiCache safely
      ensureAiStructures_(newState);
      if (stored && stored.aiCache && typeof stored.aiCache === "object") {
        ensureAiStructures_(stored);
        ["weeklyRecaps", "tradeGrades", "playoffAdvisors", "hofBios"].forEach(function(bk) {
          newState.aiCache[bk] = Object.assign({}, stored.aiCache[bk] || {}, newState.aiCache[bk] || {});
        });
      }
      if (stored && stored.settings && stored.settings.aiVoice && (!newState.settings.aiVoice.updatedAt || newState.settings.aiVoice.updatedAt < Number(stored.settings.aiVoice.updatedAt || 0))) {
        newState.settings.aiVoice = Object.assign({}, stored.settings.aiVoice);
      }
      trimAiCache_(newState);
      delete newState.cloudSyncUrl;
      delete newState.geminiApiKey;
      delete newState.GEMINI_API_KEY;
      if (newState.settings) {
        delete newState.settings.geminiApiKey;
        delete newState.settings.GEMINI_API_KEY;
      }

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
 * Writes canonical JSON (chunk-safe across columns B2+ and D5+ up to 45 chunks ~2MB) + rolling 30-snapshot history + 5 human-readable Google Sheet tabs!
 */
function saveAndMirrorToSheets(sheets, state, actionLabel) {
  var nowIso = new Date().toISOString();
  var jsonStr = JSON.stringify(state);

  // Split jsonStr into 44,000-character chunks (safe below 50,000-char Google Sheets cell limit)
  var primaryChunk = jsonStr.slice(0, CELL_CHUNK_SIZE);
  var overflowChunks = [];
  for (var pos = CELL_CHUNK_SIZE; pos < jsonStr.length; pos += CELL_CHUNK_SIZE) {
    overflowChunks.push(jsonStr.slice(pos, pos + CELL_CHUNK_SIZE));
  }
  var curMaxCols = sheets.stateSheet.getMaxColumns();
  // Pad trailing empty strings across all existing overflow columns so shrinking payloads never leave stale chunks
  var padTarget = Math.max(15, curMaxCols - 4, overflowChunks.length + 3);
  while (overflowChunks.length < padTarget) overflowChunks.push("");

  // Ensure LIVE_STATE_JSON has enough columns for snapshot row (4 metadata/primary cols + overflowChunks.length)
  var neededCols = 4 + overflowChunks.length;
  if (curMaxCols < neededCols) {
    sheets.stateSheet.insertColumnsAfter(curMaxCols, neededCols - curMaxCols);
  }

  // 1A. Update LIVE_STATE_JSON CURRENT_STATE row (B2 primary, C2 ISO, D2+ overflow chunks)
  sheets.stateSheet.getRange("B2:C2").setValues([[primaryChunk, nowIso]]);
  sheets.stateSheet.getRange(2, 4, 1, overflowChunks.length).setValues([overflowChunks]);

  // 1B. Insert rolling 30-row Time Machine snapshot at Row 5 (A5:C5 metadata, D5 primary, E5+ overflow chunks)
  sheets.stateSheet.insertRowBefore(5);
  var snapRowValues = [state.updatedAt || Date.now(), nowIso, safeSheetCell(actionLabel), primaryChunk].concat(overflowChunks);
  sheets.stateSheet.getRange(5, 1, 1, snapRowValues.length).setValues([snapRowValues]);
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
    "Conditional Swaps", "Fairness Review (Yes/No)", "Notes"
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
      var notePart = p.resolutionNote ? (" [Note by " + tName(p.resolvedBy || p.poolHolder) + ": " + p.resolutionNote + "]") : "";
      var st = p.resolvedOriginalOwner ? ("Finalized -> " + tName(p.resolvedOriginalOwner) + notePart) : "PENDING FINAL";
      return p.year + " Rd " + p.round + " Pool (" + tName(p.poolHolder) + " owes 1 to " + tName(p.owedTo) + " | " + st + ")";
    }).join("; ");
    var conds = (tr.conditionals || []).map(function(c) {
      var notePart = c.resolutionNote ? (" [Note by " + tName(c.resolvedBy || c.favoredTeam) + ": " + c.resolutionNote + "]") : "";
      var st = c.isResolved ? ((c.resolvedRound ? "Swapped Rd " + c.resolvedRound : "Settled") + notePart) : "PENDING FINAL";
      return c.year + " Rd " + (c.roundsInvolved || []).join("/") + " (" + st + ": " + (c.summary || "") + ")";
    }).join("; ");
    var fv = (tr.fairnessVotes && typeof tr.fairnessVotes === "object") ? tr.fairnessVotes : { yes: [], no: [] };
    var yesCnt = Array.isArray(fv.yes) ? fv.yes.length : 0;
    var noCnt = Array.isArray(fv.no) ? fv.no.length : 0;
    var fairnessSummary = (yesCnt + noCnt) > 0 ? ("Yes: " + yesCnt + " / No: " + noCnt) : "";

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
      safeSheetCell(fairnessSummary),
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
      var ruleAndNote = (p.ruleLabel || "") + (p.resolutionNote ? (" | Resolved Note: " + p.resolutionNote) : "");
      pickRows.push([
        p.year, p.round, p.resolvedOriginalOwner ? "FINALIZED POOL" : "PENDING FINAL POOL",
        safeSheetCell(tName(p.poolHolder)), safeSheetCell(tName(p.owedTo)),
        safeSheetCell(ruleAndNote), safeSheetCell(tr.id || "")
      ]);
    });
    (tr.conditionals || []).forEach(function(c) {
      var condRuleAndNote = (c.summary || "") + (c.resolutionNote ? (" | Resolved Note: " + c.resolutionNote) : "");
      pickRows.push([
        c.year, safeSheetCell((c.roundsInvolved || []).join("/")), c.isResolved ? "FINALIZED SWAP" : "PENDING FINAL SWAP",
        safeSheetCell(tName(c.teamA) + " <-> " + tName(c.teamB)), safeSheetCell("Favors " + tName(c.favoredTeam)),
        safeSheetCell(condRuleAndNote), safeSheetCell(tr.id || "")
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
          Number(pl.pts !== undefined ? pl.pts : pl.totalPts) || 0,
          safeSheetCell(isKeeper),
          safeSheetCell(tagStatus)
        ]);
      });
    });
    sheets.espnRostersSheet.getRange(1, 1, rosterRows.length, rosterRows[0].length).setValues(rosterRows);
    sheets.espnRostersSheet.getRange(1, 1, 1, rosterRows[0].length).setFontWeight("bold");
  }

  // 5. Rebuild ESPN_HISTORY_BACKUP tab with exact Wave 2 Sec 2C schema:
  // Season Year, Franchise Slot, Manager Name, Regular Season W-L-T, Divisional W-L, PF, PA, Combined Expected Wins, Net Luck +/-, Max PF, Coaching Efficiency %, Final Finish, Champion?
  if (sheets.espnHistorySheet && state.espnData && (state.espnData.seasonStandings || state.espnData.seasons)) {
    sheets.espnHistorySheet.clearContents();
    var histRows = [[
      "Season Year", "Franchise Slot", "Manager Name", "Regular Season W-L-T", "Divisional W-L",
      "PF", "PA", "Combined Expected Wins", "Net Luck +/-", "Max PF",
      "Coaching Efficiency %", "Final Finish", "Champion?"
    ]];
    var standingsByYr = state.espnData.seasonStandings || {};
    var seasonsByYr = state.espnData.seasons || {};
    var allYrs = Object.keys(Object.assign({}, standingsByYr, seasonsByYr)).sort().reverse();
    allYrs.forEach(function(yr) {
      var sObj = seasonsByYr[yr] || {};
      var rows = Array.isArray(standingsByYr[yr]) ? standingsByYr[yr] : (Array.isArray(sObj.standings) ? sObj.standings : []);
      rows.forEach(function(r) {
        var w = Number(r.wins) || 0;
        var l = Number(r.losses) || 0;
        var t = Number(r.ties) || 0;
        var wltStr = w + "-" + l + (t > 0 ? ("-" + t) : "");
        var divStr = (r.divWins !== undefined && r.divLosses !== undefined) ? (r.divWins + "-" + r.divLosses + (r.divTies ? "-" + r.divTies : "")) : (r.divisionRecord || "—");
        var apW = Number(r.allPlayWins) || 0;
        var apL = Number(r.allPlayLosses) || 0;
        var regGames = w + l + t;
        var apTotalGames = apW + apL;
        // Scale All-Play to regular season games if ESPN included postseason weeks in allPlayWins
        var apScale = (apTotalGames > 0 && regGames > 0 && apTotalGames > regGames * 9) ? ((regGames * 9) / apTotalGames) : 1;
        var apExpWins = (apW * apScale) / 9;
        var medWins = r.medianWins !== undefined ? Number(r.medianWins) : null;
        var combExp = r.expectedWins !== undefined
          ? Number(r.expectedWins)
          : (medWins !== null ? (0.5 * apExpWins + 0.5 * medWins) : apExpWins);
        var netLuck = r.netLuck !== undefined ? Number(r.netLuck) : (w + 0.5 * t - combExp);
        var luckFormatted = (netLuck >= 0 ? "+" : "") + netLuck.toFixed(2);
        var pfVal = Number(r.pointsFor !== undefined ? r.pointsFor : r.pf) || 0;
        var paVal = Number(r.pointsAgainst !== undefined ? r.pointsAgainst : r.pa) || 0;
        var maxPfVal = r.maxPf ? Number(r.maxPf).toFixed(1) : "—";
        var effVal = r.coachingEfficiency ? (Number(r.coachingEfficiency).toFixed(1) + "%") : "—";
        var mgrLabel = tName(r.managerId || r.teamId) || r.ownerName || r.teamName || "";
        var isChampBool = Boolean(r.isChampion || r.isChamp || (sObj.champion && String(sObj.champion) === String(r.managerId || r.teamId)));

        histRows.push([
          Number(yr) || yr,
          Number(r.espnTeamId) || "",
          safeSheetCell(mgrLabel),
          safeSheetCell(wltStr),
          safeSheetCell(divStr),
          pfVal,
          paVal,
          Number(combExp.toFixed(2)),
          safeSheetCell(luckFormatted),
          safeSheetCell(maxPfVal),
          safeSheetCell(effVal),
          Number(r.finalStanding || r.finalRank || r.rank) || 0,
          isChampBool ? "CHAMPION 🏆" : ""
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
      "?view=mTeam&view=mRoster&view=mStandings&view=mMatchupScore&view=mSettings&view=mTransactions2&view=mBoxscore";

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

    // 2. Sync Live Rosters, Master Keepers & Positional Ranks
    ed.rosters = ed.rosters || {};
    var rosteredEspnIds = {};
    var allLeaguePlayersByPos = {};
    var prevSeasonRosters = (ed.seasons && ed.seasons[String(seasonYr - 1)] && ed.seasons[String(seasonYr - 1)].rosters) || {};
    var syncedMasterKeepers = {};
    var syncedKeptPlayerList = [];
    data.teams.forEach(function(t) {
      var slotKey = String(t.id);
      var mgrId = slotMap[slotKey];
      if (!mgrId) return;
      var prevTeamRoster = Array.isArray(prevSeasonRosters[mgrId]) ? prevSeasonRosters[mgrId] : [];
      var prevPidSet = {};
      prevTeamRoster.forEach(function(pr) {
        if (pr && pr.espnId) prevPidSet[String(pr.espnId)] = true;
      });
      var entries = (t.roster && Array.isArray(t.roster.entries)) ? t.roster.entries : [];
      if (entries.length > 0) {
        var teamKeepIds = [];
        var mappedRoster = entries.map(function(e) {
          var p = (e.playerPoolEntry && e.playerPoolEntry.player) || {};
          var pid = p.id || e.playerId;
          if (pid) rosteredEspnIds[String(pid)] = true;
          var pos = posMap[p.defaultPositionId] || "FLEX";
          var pts = Math.round(Number((e.playerPoolEntry && e.playerPoolEntry.appliedStatTotal) || 0) * 10) / 10;
          var pStats = Array.isArray(p.stats) ? p.stats : [];
          var actSt = null;
          var prjSt = null;
          var priorSt = null;
          for (var psi = 0; psi < pStats.length; psi++) {
            var stItm = pStats[psi];
            if (!stItm || Number(stItm.statSplitTypeId) !== 0) continue;
            if (Number(stItm.seasonId) === seasonYr && Number(stItm.statSourceId) === 0) actSt = stItm;
            if (Number(stItm.seasonId) === seasonYr && Number(stItm.statSourceId) === 1) prjSt = stItm;
            if (Number(stItm.seasonId) === (seasonYr - 1) && Number(stItm.statSourceId) === 0) priorSt = stItm;
          }
          if (pts === 0 && actSt && Number(actSt.appliedTotal) > 0) {
            pts = Math.round(Number(actSt.appliedTotal) * 10) / 10;
          }
          var priorPts = Math.round(Number((priorSt && priorSt.appliedTotal) || 0) * 10) / 10;
          var projPts = Math.round(Number((prjSt && prjSt.appliedTotal) || 0) * 10) / 10;
          var avgPts = Math.round((Number(actSt && actSt.appliedAverage) || (pts > 0 ? pts / 17 : Number((priorSt && priorSt.appliedAverage) || 0))) * 10) / 10;
          var acqType = String(e.acquisitionType || "ROSTER").toUpperCase();
          var isKeepFlag = (acqType === "KEEP") || (Boolean(prevPidSet[String(pid)]) && acqType !== "ADD" && acqType !== "TRADE");
          var keeperSrc = acqType === "KEEP" ? "ESPN_KEEP" : (isKeepFlag ? "YOY_CONTINUITY" : "");
          var isRookie = Boolean((p.draftYear && Number(p.draftYear) >= seasonYr) || p.experience === 0 || (p.seasonOutlook && String(p.seasonOutlook).toLowerCase().indexOf("rookie") !== -1));
          var ownObj = p.ownership || {};
          var plId = "espn_" + pid;
          if (isKeepFlag && teamKeepIds.length < 12) {
            teamKeepIds.push(plId);
            syncedKeptPlayerList.push({
              managerId: mgrId,
              teamId: mgrId,
              playerId: plId,
              espnId: Number(pid) || 0,
              name: p.fullName || "Unknown Player",
              pos: pos,
              source: keeperSrc || "ESPN_KEEP"
            });
          }
          var plObj = {
            id: plId,
            espnId: pid,
            name: p.fullName || "Unknown Player",
            pos: pos,
            nflTeam: teamNflMap[p.proTeamId] || "NFL",
            slot: slotNameMap[e.lineupSlotId] || "BE",
            pts: pts,
            totalPts: pts,
            priorPts: priorPts,
            projPts: projPts,
            avgPts: avgPts,
            posRank: 0,
            injuryStatus: String(p.injuryStatus || (p.injured ? "INJURED" : "ACTIVE")).slice(0, 24),
            pctOwned: Math.round((Number(ownObj.percentOwned) || 0) * 10) / 10,
            pctStarted: Math.round((Number(ownObj.percentStarted) || 0) * 10) / 10,
            adp: Math.round((Number(ownObj.averageDraftPosition) || 0) * 10) / 10,
            isRookie: isRookie,
            isKeeper: isKeepFlag,
            keeperSource: keeperSrc,
            acquiredVia: acqType
          };
          if (!allLeaguePlayersByPos[pos]) allLeaguePlayersByPos[pos] = [];
          allLeaguePlayersByPos[pos].push(plObj);
          return plObj;
        });
        ed.rosters[mgrId] = mappedRoster;
        if (teamKeepIds.length > 0) {
          syncedMasterKeepers[mgrId] = teamKeepIds;
        }
      }
    });
    Object.keys(allLeaguePlayersByPos).forEach(function(posKey) {
      allLeaguePlayersByPos[posKey].sort(function(a, b) {
        var bScore = (b.totalPts || 0) > 0 ? b.totalPts : ((b.priorPts || 0) || (b.projPts || 0));
        var aScore = (a.totalPts || 0) > 0 ? a.totalPts : ((a.priorPts || 0) || (a.projPts || 0));
        return bScore - aScore;
      });
      allLeaguePlayersByPos[posKey].forEach(function(pl, idx) { pl.posRank = idx + 1; });
    });

    // Master Keeper Source of Truth write-through when ESPN keeper tags / YoY continuity are detected
    if (Object.keys(syncedMasterKeepers).length > 0) {
      ed.masterKeepersByTeam = syncedMasterKeepers;
      if (!currentState.teamProfiles || typeof currentState.teamProfiles !== "object") currentState.teamProfiles = {};
      Object.keys(syncedMasterKeepers).forEach(function(mKey) {
        if (!currentState.teamProfiles[mKey] || typeof currentState.teamProfiles[mKey] !== "object") {
          currentState.teamProfiles[mKey] = {};
        }
        currentState.teamProfiles[mKey].keepers = syncedMasterKeepers[mKey].slice(0, 12);
      });
    }

    // 2B. Sync Top 250 Available Unrostered Free Agents & Incoming Rookies via kona_player_info
    try {
      var faUrl = "https://lm-api-reads.fantasy.espn.com/apis/v3/games/ffl/seasons/" +
        seasonYr + "/segments/0/leagues/" + leagueId + "?view=kona_player_info";
      var faHeaders = Object.assign({}, headers, {
        "x-fantasy-filter": JSON.stringify({
          players: {
            filterStatus: { value: ["FREEAGENT", "WAIVERS"] },
            limit: 320,
            sortPercOwned: { sortAsc: false, sortPriority: 1 },
            sortAppliedStatTotal: { sortAsc: false, sortPriority: 2, value: "10" + seasonYr }
          }
        })
      });
      var faResp = UrlFetchApp.fetch(faUrl, { method: "get", headers: faHeaders, muteHttpExceptions: true });
      if (faResp.getResponseCode() === 200) {
        var faData = JSON.parse(faResp.getContentText());
        var faRaw = Array.isArray(faData && faData.players) ? faData.players : [];
        var mappedFas = [];
        for (var fi = 0; fi < faRaw.length; fi++) {
          var pl = faRaw[fi];
          var pObj = (pl && pl.player) || pl || {};
          var pid = Number(pObj.id || pl.id) || 0;
          if (!pid || rosteredEspnIds[String(pid)]) continue;
          var pStats = Array.isArray(pObj.stats) ? pObj.stats : [];
          var actSt = null;
          var prjSt = null;
          var priorSt = null;
          for (var fsi = 0; fsi < pStats.length; fsi++) {
            var fst = pStats[fsi];
            if (!fst || Number(fst.statSplitTypeId) !== 0) continue;
            if (Number(fst.seasonId) === seasonYr && Number(fst.statSourceId) === 0) actSt = fst;
            if (Number(fst.seasonId) === seasonYr && Number(fst.statSourceId) === 1) prjSt = fst;
            if (Number(fst.seasonId) === (seasonYr - 1) && Number(fst.statSourceId) === 0) priorSt = fst;
          }
          var tot = Math.round((Number((actSt && actSt.appliedTotal) || 0)) * 10) / 10;
          var priorTot = Math.round((Number((priorSt && priorSt.appliedTotal) || 0)) * 10) / 10;
          var avg = Math.round((Number(actSt && actSt.appliedAverage) || (tot > 0 ? tot / 17 : Number((priorSt && priorSt.appliedAverage) || 0))) * 10) / 10;
          var prj = Math.round((Number(prjSt && prjSt.appliedTotal) || priorTot || tot) * 10) / 10;
          var ownObj = pObj.ownership || {};
          var pctOwn = Math.round((Number(ownObj.percentOwned) || 0) * 10) / 10;
          var isRookie = Boolean((pObj.draftYear && Number(pObj.draftYear) >= seasonYr) || pObj.experience === 0);
          if (!pObj.fullName) continue;
          if (tot <= 0 && prj <= 0 && priorTot <= 0 && pctOwn <= 0 && !isRookie) continue;
          mappedFas.push({
            id: "espn_fa_" + pid,
            espnId: pid,
            name: pObj.fullName,
            pos: posMap[pObj.defaultPositionId] || "FLEX",
            nflTeam: teamNflMap[pObj.proTeamId] || "FA",
            slot: "FA",
            status: "FREE_AGENT",
            totalPts: tot,
            pts: tot,
            priorPts: priorTot,
            projPts: prj,
            avgPts: avg,
            injuryStatus: String(pObj.injuryStatus || "ACTIVE").slice(0, 24),
            percentOwned: pctOwn,
            pctOwned: pctOwn,
            pctStarted: Math.round((Number(ownObj.percentStarted) || 0) * 10) / 10,
            adp: Math.round((Number(ownObj.averageDraftPosition) || 0) * 10) / 10,
            isRookie: isRookie
          });
          if (mappedFas.length >= 250) break;
        }
        if (mappedFas.length > 0) {
          ed.freeAgents = mappedFas.slice(0, 250);
        }
      }
    } catch (faSyncErr) {
      Logger.log("Free agent sync skipped: " + faSyncErr);
    }

    // 3. Sync Weekly Matchups (Preserving Future Unplayed Schedule for Playoff Path + True Max PF when boxscores exist)
    var regWeeks = Number(data.settings && data.settings.scheduleSettings && data.settings.scheduleSettings.matchupPeriodCount) || 14;
    var playoffTeamCnt = Number(data.settings && data.settings.scheduleSettings && data.settings.scheduleSettings.playoffTeamCount) || 6;
    var annualFaabBudget = Number(data.settings && data.settings.acquisitionSettings && data.settings.acquisitionSettings.acquisitionBudget) || 600;
    var currentSp = Number(data.scoringPeriodId || (data.status && data.status.currentMatchupPeriod) || 0);
    var finalSp = Number((data.status && data.status.finalScoringPeriod) || 17);
    var slotCountsForMaxPf = (data.settings && data.settings.rosterSettings && data.settings.rosterSettings.lineupSlotCounts) || {};

    function computeBoxscoreOptimalPts(entries) {
      if (!Array.isArray(entries) || entries.length === 0) return null;
      var activeSlotOrder = [0, 2, 4, 6, 23, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];
      var candidates = entries.map(function(en) {
        var p = (en && en.playerPoolEntry && en.playerPoolEntry.player) || {};
        var pts = Number((en && en.playerPoolEntry && en.playerPoolEntry.appliedStatTotal) || 0);
        var elig = Array.isArray(p.eligibleSlots) ? p.eligibleSlots : [];
        return { pts: pts, elig: elig, used: false, lineupSlotId: Number(en.lineupSlotId) };
      });
      var benchSum = 0;
      candidates.forEach(function(c) {
        if (c.lineupSlotId === 20 || c.lineupSlotId === 21) benchSum += Math.max(0, c.pts);
      });
      candidates.sort(function(a, b) { return b.pts - a.pts; });
      var optSum = 0;
      // Assign stricter single-position slots before multi-position FLEX/OP/DP slots
      var orderedSlots = [0, 2, 4, 6, 8, 9, 12, 13, 10, 11, 14, 16, 17, 23, 15, 7];
      for (var osi = 0; osi < orderedSlots.length; osi++) {
        var sId = orderedSlots[osi];
        var needed = Number(slotCountsForMaxPf[String(sId)] || 0);
        if (needed <= 0) continue;
        for (var cnt = 0; cnt < needed; cnt++) {
          for (var ci = 0; ci < candidates.length; ci++) {
            var cand = candidates[ci];
            if (!cand.used && cand.pts > 0 && cand.elig.indexOf(sId) !== -1) {
              cand.used = true;
              optSum += cand.pts;
              break;
            }
          }
        }
      }
      return { optPts: optSum, benchPts: benchSum };
    }

    var allPlayBySlot = {};
    var medianBySlot = {};
    var toiletStatsBySlot = {};
    var maxPfBySlot = {};
    data.teams.forEach(function(t) {
      var sk = String(t.id);
      allPlayBySlot[sk] = { w: 0, l: 0 };
      medianBySlot[sk] = { w: 0, l: 0 };
      toiletStatsBySlot[sk] = { wins: 0, pts: 0 };
      maxPfBySlot[sk] = { maxPf: 0, actualPf: 0, benchPts: 0, weeksCounted: 0 };
    });

    if (!ed.headToHead || typeof ed.headToHead !== "object") ed.headToHead = {};
    if (!ed.syncedMatchupKeys || typeof ed.syncedMatchupKeys !== "object") ed.syncedMatchupKeys = {};
    if (!ed.records || typeof ed.records !== "object") ed.records = {};

    var schedule = Array.isArray(data.schedule) ? data.schedule : [];
    var scoresByWeek = {};
    var syncedSeasonMatchups = [];
    function extractSidePointsForWeek(side, wkNum) {
      if (!side) return 0;
      var tot = Number(side.totalPoints || 0);
      var bySp = (side.pointsByScoringPeriod && side.pointsByScoringPeriod[String(wkNum)] !== undefined)
        ? Number(side.pointsByScoringPeriod[String(wkNum)] || 0)
        : 0;
      var live = Number(side.totalPointsLive || 0);
      var boxTot = Number((side.rosterForCurrentScoringPeriod && side.rosterForCurrentScoringPeriod.appliedStatTotal) ||
                          (side.rosterForMatchupPeriod && side.rosterForMatchupPeriod.appliedStatTotal) || 0);
      return tot > 0 ? tot : Math.max(bySp, live, boxTot);
    }

    schedule.forEach(function(m) {
      if (!m || !m.home || !m.away) return;
      var wk = Number(m.matchupPeriodId) || 0;
      if (wk <= 0) return;
      var hScore = extractSidePointsForWeek(m.home, wk);
      var aScore = extractSidePointsForWeek(m.away, wk);
      var hSlot = String(m.home.teamId);
      var aSlot = String(m.away.teamId);
      var hMgr = slotMap[hSlot] || "";
      var aMgr = slotMap[aSlot] || "";
      var mTier = String(m.playoffTierType || "NONE").toUpperCase();
      var isPlayedGame = Boolean(hScore > 0 || aScore > 0);

      // Preserve future unplayed regular-season games (played: false) so Playoff Path has remaining schedule!
      if (!isPlayedGame && wk > regWeeks) return;
      if (!hMgr || !aMgr) return;

      syncedSeasonMatchups.push({
        week: wk,
        homeSlot: Number(m.home.teamId) || 0,
        homeTeamId: hMgr,
        homeManagerId: hMgr,
        homeScore: Math.round(hScore * 10) / 10,
        awaySlot: Number(m.away.teamId) || 0,
        awayTeamId: aMgr,
        awayManagerId: aMgr,
        awayScore: Math.round(aScore * 10) / 10,
        played: isPlayedGame,
        tier: mTier,
        matchupType: mTier
      });

      if (!isPlayedGame) return;

      if (wk <= regWeeks) {
        if (!scoresByWeek[wk]) scoresByWeek[wk] = [];
        scoresByWeek[wk].push({ slot: hSlot, score: hScore });
        scoresByWeek[wk].push({ slot: aSlot, score: aScore });

        // Compute boxscore optimal points when entries are present for regular-season weeks
        var hBoxEntries = (m.home.rosterForCurrentScoringPeriod && m.home.rosterForCurrentScoringPeriod.entries) ||
                          (m.home.rosterForMatchupPeriod && m.home.rosterForMatchupPeriod.entries) || null;
        var aBoxEntries = (m.away.rosterForCurrentScoringPeriod && m.away.rosterForCurrentScoringPeriod.entries) ||
                          (m.away.rosterForMatchupPeriod && m.away.rosterForMatchupPeriod.entries) || null;
        var hOpt = computeBoxscoreOptimalPts(hBoxEntries);
        if (hOpt && hOpt.optPts > 0 && maxPfBySlot[hSlot]) {
          maxPfBySlot[hSlot].maxPf += Math.max(hScore, hOpt.optPts);
          maxPfBySlot[hSlot].actualPf += hScore;
          maxPfBySlot[hSlot].benchPts += hOpt.benchPts;
          maxPfBySlot[hSlot].weeksCounted += 1;
        }
        var aOpt = computeBoxscoreOptimalPts(aBoxEntries);
        if (aOpt && aOpt.optPts > 0 && maxPfBySlot[aSlot]) {
          maxPfBySlot[aSlot].maxPf += Math.max(aScore, aOpt.optPts);
          maxPfBySlot[aSlot].actualPf += aScore;
          maxPfBySlot[aSlot].benchPts += aOpt.benchPts;
          maxPfBySlot[aSlot].weeksCounted += 1;
        }
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
          var prevHigh = ed.records.highestWeekScore ? Number(ed.records.highestWeekScore.score || 0) : 0;
          if (!ed.records.highestWeekScore || sp.sc > prevHigh) {
            ed.records.highestWeekScore = { managerId: sp.mgr, score: Math.round(sp.sc * 100) / 100, year: seasonYr, week: wk, opponentId: sp.opp };
            if (prevHigh > 0) {
              appendServerActivityEvent(currentState, "RECORD_SET", "New All-Time Weekly High Score!", sp.mgr + " dropped " + ed.records.highestWeekScore.score + " pts in Week " + wk + " (" + seasonYr + ").", { year: seasonYr, teamId: sp.mgr });
            }
          }
          var prevLow = (ed.records.lowestWeekScore && ed.records.lowestWeekScore.score) ? Number(ed.records.lowestWeekScore.score) : 0;
          if (sp.sc > 0 && (!ed.records.lowestWeekScore || !prevLow || sp.sc < prevLow)) {
            ed.records.lowestWeekScore = { managerId: sp.mgr, score: Math.round(sp.sc * 100) / 100, year: seasonYr, week: wk, opponentId: sp.opp };
            if (prevLow > 0) {
              appendServerActivityEvent(currentState, "RECORD_SET", "New All-Time Weekly Low Score Alert", sp.mgr + " scored just " + ed.records.lowestWeekScore.score + " pts in Week " + wk + " (" + seasonYr + ").", { year: seasonYr, teamId: sp.mgr });
            }
          }
        });
        var margin = Math.round(Math.abs(hScore - aScore) * 100) / 100;
        if (margin > 0) {
          var winMgr = hScore > aScore ? hMgr : aMgr;
          var loseMgr = hScore > aScore ? aMgr : hMgr;
          var prevBlowout = ed.records.biggestBlowout ? Number(ed.records.biggestBlowout.margin || 0) : 0;
          if (!ed.records.biggestBlowout || margin > prevBlowout) {
            ed.records.biggestBlowout = { winnerId: winMgr, loserId: loseMgr, margin: margin, year: seasonYr, week: wk };
            if (prevBlowout > 0) {
              appendServerActivityEvent(currentState, "RECORD_SET", "New Biggest Blowout Record!", winMgr + " defeated " + loseMgr + " by " + margin + " pts (Week " + wk + ", " + seasonYr + ").", { year: seasonYr, teamId: winMgr });
            }
          }
          var prevClosest = (ed.records.closestWin && ed.records.closestWin.margin) ? Number(ed.records.closestWin.margin) : 0;
          if (!ed.records.closestWin || !prevClosest || margin < prevClosest) {
            ed.records.closestWin = { winnerId: winMgr, loserId: loseMgr, margin: margin, year: seasonYr, week: wk };
            if (prevClosest > 0) {
              appendServerActivityEvent(currentState, "RECORD_SET", "New Closest Nailbiter Record!", winMgr + " edged " + loseMgr + " by " + margin + " pts (Week " + wk + ", " + seasonYr + ").", { year: seasonYr, teamId: winMgr });
            }
          }
        }
      }
    });

    Object.keys(scoresByWeek).forEach(function(wkKey) {
      var list = scoresByWeek[wkKey].slice();
      for (var i = 0; i < list.length; i++) {
        for (var j = 0; j < list.length; j++) {
          if (i === j) continue;
          if (list[i].score > list[j].score) allPlayBySlot[list[i].slot].w += 1;
          else if (list[i].score < list[j].score) allPlayBySlot[list[i].slot].l += 1;
        }
      }
      // Weekly League Median: Top 5 of 10 earn 1.0 median win, Bottom 5 earn 0.0
      list.sort(function(a, b) { return b.score - a.score; });
      for (var mIdx = 0; mIdx < list.length; mIdx++) {
        var sl = list[mIdx].slot;
        if (!medianBySlot[sl]) medianBySlot[sl] = { w: 0, l: 0 };
        if (mIdx < 5) medianBySlot[sl].w += 1;
        else medianBySlot[sl].l += 1;
      }
    });

    // 4. Build Live Season Standings & Official Completed-Season Gate
    var maxGamesPlayed = 0;
    var hasWeeklyScores = Object.keys(scoresByWeek).length > 0;
    var prevStandingsByMgr = {};
    var prevStList = (ed.seasons && ed.seasons[String(seasonYr)] && Array.isArray(ed.seasons[String(seasonYr)].standings))
      ? ed.seasons[String(seasonYr)].standings
      : [];
    prevStList.forEach(function(ps) {
      if (ps && ps.managerId) prevStandingsByMgr[ps.managerId] = ps;
    });
    var rawTeamRows = data.teams.map(function(t) {
      var slotKey = String(t.id);
      var mgrId = slotMap[slotKey] || "";
      var prevRow = prevStandingsByMgr[mgrId] || {};
      var recOverall = (t.record && t.record.overall) || {};
      var recDiv = (t.record && t.record.division) || {};
      var txCounter = t.transactionCounter || {};
      var acqCnt = Number(txCounter.acquisitions || 0);
      var dropCnt = Number(txCounter.drops || 0);
      var trdCnt = Number(txCounter.trades || 0);
      var irCnt = Number(txCounter.moveToIR || 0);
      var faabSpent = Number(txCounter.acquisitionBudgetSpent || 0);
      var w = Number(recOverall.wins || 0);
      var l = Number(recOverall.losses || 0);
      var ties = Number(recOverall.ties || 0);
      var divW = Number(recDiv.wins || 0);
      var divL = Number(recDiv.losses || 0);
      var divT = Number(recDiv.ties || 0);
      if (w + l + ties > maxGamesPlayed) maxGamesPlayed = w + l + ties;
      var pf = Math.round(Number(recOverall.pointsFor || 0) * 10) / 10;
      var pa = Math.round(Number(recOverall.pointsAgainst || 0) * 10) / 10;
      var seed = Number(t.playoffSeed || 0);
      var finalRank = Number(t.rankCalculatedFinal || 0);
      var ap = allPlayBySlot[slotKey] || { w: 0, l: 0 };
      var med = medianBySlot[slotKey] || { w: 0, l: 0 };
      var apExpWins = ap.w / 9;
      var combExpWins = hasWeeklyScores ? (0.5 * apExpWins + 0.5 * med.w) : apExpWins;
      var netLuck = (w + 0.5 * ties) - combExpWins;
      var tb = toiletStatsBySlot[slotKey] || { wins: 0, pts: 0 };
      var mpRec = maxPfBySlot[slotKey] || { maxPf: 0, actualPf: 0, benchPts: 0, weeksCounted: 0 };
      var hasTrueMaxPf = mpRec.weeksCounted > 0 && mpRec.maxPf > 0;
      var maxPfVal = hasTrueMaxPf
        ? Math.round(mpRec.maxPf * 10) / 10
        : (Number(prevRow.maxPointsFor || prevRow.maxPf) || 0);
      var effVal = hasTrueMaxPf
        ? Math.min(100, Math.round((mpRec.actualPf / mpRec.maxPf) * 1000) / 10)
        : (Number(prevRow.coachingEfficiency) || 0);
      var benchVal = hasTrueMaxPf
        ? Math.round(mpRec.benchPts * 10) / 10
        : (Number(prevRow.benchPoints) || 0);
      var tNameFull = ((t.location || "") + " " + (t.nickname || "")).trim() || (t.name || ("Slot " + slotKey));
      return {
        espnTeamId: Number(t.id),
        divisionId: Number(t.divisionId || 0),
        teamId: mgrId,
        managerId: mgrId,
        currentManagerId: mgrId,
        teamName: tNameFull,
        wins: w,
        losses: l,
        ties: ties,
        divWins: divW,
        divLosses: divL,
        divTies: divT,
        divisionRecord: divW + "-" + divL + (divT > 0 ? "-" + divT : ""),
        pf: pf,
        pa: pa,
        pointsFor: pf,
        pointsAgainst: pa,
        maxPf: maxPfVal,
        maxPointsFor: maxPfVal,
        coachingEfficiency: effVal,
        benchPoints: benchVal,
        hasTrueMaxPf: Boolean(hasTrueMaxPf || prevRow.hasTrueMaxPf),
        playoffSeed: seed,
        rank: finalRank || seed || 0,
        finalRank: finalRank || seed || 0,
        finalStanding: finalRank || seed || 0,
        allPlayWins: ap.w,
        allPlayLosses: ap.l,
        medianWins: med.w,
        medianLosses: med.l,
        expectedWins: Math.round(combExpWins * 100) / 100,
        netLuck: Math.round(netLuck * 100) / 100,
        toiletWins: tb.wins,
        toiletPoints: Math.round(tb.pts * 10) / 10,
        faabSpent: faabSpent,
        faabBudget: annualFaabBudget,
        acquisitions: acqCnt,
        drops: dropCnt,
        espnTrades: trdCnt,
        moveToIR: irCnt,
        hasTransactions: Boolean(acqCnt > 0 || dropCnt > 0 || trdCnt > 0 || faabSpent > 0),
        isChampion: false
      };
    });

    rawTeamRows.sort(function(a, b) {
      var ra = a.rank || 99;
      var rb = b.rank || 99;
      if (ra !== rb) return ra - rb;
      if (b.wins !== a.wins) return b.wins - a.wins;
      if (b.divWins !== a.divWins) return b.divWins - a.divWins;
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
      if (!feasterId && rawTeamRows.length >= 10) {
        feasterId = rawTeamRows[rawTeamRows.length - 1].managerId || "";
      }
    }

    // 4B. Wave 4 Repair: Full Scoring-Period Sweep for Weekly Boxscores (Positional Kings + True Max PF) & mTransactions2 FAAB Waivers
    var weeklyPtsByPid = {};
    var playerMetaByPid = {};
    Object.keys(ed.rosters || {}).forEach(function(mKey) {
      (ed.rosters[mKey] || []).forEach(function(rp) {
        if (rp && rp.espnId) {
          playerMetaByPid[String(rp.espnId)] = {
            name: rp.name || ("Player #" + rp.espnId),
            pos: normalizePosBucketGas_(rp.pos, 0) || String(rp.pos || ""),
            nflTeam: rp.nflTeam || "NFL"
          };
        }
      });
    });

    var maxSweepSp = Math.min(18, Math.max(currentSp || 1, maxGamesPlayed || 1));
    var seasonKingsBuilder = createEmptyPosKingsStore_(seasonYr);
    var weeklyBenchBlundersMap = (ed.weeklyBenchBlunders && typeof ed.weeklyBenchBlunders === "object") ? Object.assign({}, ed.weeklyBenchBlunders) : {};
    var rebuiltMaxPfBySlot = {};
    data.teams.forEach(function(t) {
      rebuiltMaxPfBySlot[String(t.id)] = { maxPf: 0, actualPf: 0, benchPts: 0, weeksCounted: 0 };
    });

    function recordBoxscoreEntriesGas_(wkNum, teamSlotStr, sideScore, entries) {
      if (!Array.isArray(entries) || entries.length === 0) return;
      var mgrIdForSlot = slotMap[teamSlotStr] || "";
      if (wkNum <= regWeeks) {
        var wkOpt = computeBoxscoreOptimalPts(entries);
        if (wkOpt && wkOpt.optPts > 0 && rebuiltMaxPfBySlot[teamSlotStr]) {
          rebuiltMaxPfBySlot[teamSlotStr].maxPf += Math.max(sideScore, wkOpt.optPts);
          rebuiltMaxPfBySlot[teamSlotStr].actualPf += sideScore;
          rebuiltMaxPfBySlot[teamSlotStr].benchPts += wkOpt.benchPts;
          rebuiltMaxPfBySlot[teamSlotStr].weeksCounted += 1;
        }
      }
      entries.forEach(function(en) {
        var pObj = (en && en.playerPoolEntry && en.playerPoolEntry.player) || {};
        var pid = Number((en && en.playerId) || pObj.id) || 0;
        if (pid <= 0) return;
        var pidStr = String(pid);
        var slotId = Number(en && en.lineupSlotId !== undefined ? en.lineupSlotId : 20);
        var pts = Math.round(Number((en && en.playerPoolEntry && en.playerPoolEntry.appliedStatTotal) || 0) * 100) / 100;
        var pName = pObj.fullName || (playerMetaByPid[pidStr] && playerMetaByPid[pidStr].name) || ("Player #" + pidStr);
        var posBucket = normalizePosBucketGas_(playerMetaByPid[pidStr] && playerMetaByPid[pidStr].pos, pObj.defaultPositionId);
        var nflTm = (playerMetaByPid[pidStr] && playerMetaByPid[pidStr].nflTeam) || "NFL";
        playerMetaByPid[pidStr] = { name: pName, pos: posBucket || "", nflTeam: nflTm };
        if (!weeklyPtsByPid[pidStr]) weeklyPtsByPid[pidStr] = {};
        weeklyPtsByPid[pidStr][wkNum] = pts;

        if (mgrIdForSlot && posBucket && pts > 0 && POS_KING_BUCKETS_GAS.indexOf(posBucket) !== -1) {
          var isStarter = (slotId !== 20 && slotId !== 21);
          var kingRow = {
            playerId: "espn_p_" + pidStr,
            playerName: pName,
            pos: posBucket,
            nflTeam: nflTm,
            teamId: mgrIdForSlot,
            year: seasonYr,
            week: wkNum,
            pts: pts,
            slotId: slotId,
            isStarter: isStarter
          };
          if (isStarter) {
            seasonKingsBuilder.starters[posBucket].push(kingRow);
          } else if (slotId === 20) {
            seasonKingsBuilder.bench[posBucket].push(kingRow);
            var bkKey = seasonYr + "_" + wkNum;
            if (!weeklyBenchBlundersMap[bkKey]) weeklyBenchBlundersMap[bkKey] = [];
            weeklyBenchBlundersMap[bkKey].push({
              playerName: pName, pos: posBucket, teamId: mgrIdForSlot, pts: pts, week: wkNum, year: seasonYr
            });
          }
        }
      });
    }

    // Sweep completed weeks 1..maxSweepSp for real weekly boxscores & Weekly Positional Kings
    for (var bxWk = 1; bxWk <= maxSweepSp; bxWk++) {
      try {
        var bxUrl = "https://lm-api-reads.fantasy.espn.com/apis/v3/games/ffl/seasons/" +
          seasonYr + "/segments/0/leagues/" + leagueId +
          "?scoringPeriodId=" + bxWk + "&view=mBoxscore&view=mMatchup";
        var bxResp = UrlFetchApp.fetch(bxUrl, { method: "get", headers: headers, muteHttpExceptions: true });
        if (bxResp.getResponseCode() === 200) {
          var bxJson = JSON.parse(bxResp.getContentText());
          var bxSched = (bxJson && Array.isArray(bxJson.schedule)) ? bxJson.schedule : [];
          bxSched.forEach(function(bm) {
            if (!bm || Number(bm.matchupPeriodId) !== bxWk) return;
            [bm.home, bm.away].forEach(function(side) {
              if (!side || !side.teamId) return;
              var entries = (side.rosterForCurrentScoringPeriod && side.rosterForCurrentScoringPeriod.entries) ||
                            (side.rosterForMatchupPeriod && side.rosterForMatchupPeriod.entries) || [];
              var sidePts = extractSidePointsForWeek(side, bxWk);
              if (entries.length > 0 && sidePts > 0) {
                recordBoxscoreEntriesGas_(bxWk, String(side.teamId), sidePts, entries);
              }
            });
          });
        }
      } catch (bxErr) {}
    }

    // Apply rebuilt true coaching efficiency onto rawTeamRows when weekly boxscores were swept
    rawTeamRows.forEach(function(rRow) {
      var rb = rebuiltMaxPfBySlot[String(rRow.espnTeamId)];
      if (rb && rb.weeksCounted >= 1 && rb.maxPf > 0) {
        rRow.hasTrueMaxPf = true;
        rRow.maxPointsFor = Math.round(rb.maxPf * 10) / 10;
        rRow.maxPf = rRow.maxPointsFor;
        rRow.benchPoints = Math.round(rb.benchPts * 10) / 10;
        rRow.coachingEfficiency = Math.min(100, Math.round((rb.actualPf / rb.maxPf) * 1000) / 10);
      }
    });

    // Trim compact Weekly Positional Kings (Top 10 Starter + Top 10 Bench per position) & Bench Blunders (Top 3/week)
    ed.seasonWeeklyPosKings = sanitizeWeeklyPosKingsStore_(seasonKingsBuilder, seasonYr);
    ed.allTimeWeeklyPosKings = mergeWeeklyPosKingsStores_(ed.allTimeWeeklyPosKings, ed.seasonWeeklyPosKings, false);
    Object.keys(weeklyBenchBlundersMap).forEach(function(bkKey) {
      var arr = Array.isArray(weeklyBenchBlundersMap[bkKey]) ? weeklyBenchBlundersMap[bkKey] : [];
      arr.sort(function(a, b) { return Number(b.pts || 0) - Number(a.pts || 0); });
      weeklyBenchBlundersMap[bkKey] = arr.slice(0, 3);
    });
    ed.weeklyBenchBlunders = weeklyBenchBlundersMap;

    // Checkpoint 1 — Robust 3-Tier Season acquisitionBudget Denominator (never 0/NaN)
    var settingsAcqBudget = Number(data.settings && data.settings.acquisitionSettings && data.settings.acquisitionSettings.acquisitionBudget) || 0;
    var teamSumMaxBudget = 0;
    (data.teams || []).forEach(function(tm) {
      var tc = (tm && tm.transactionCounter) || {};
      var spVal = Number(tc.acquisitionBudgetSpent) || 0;
      var remVal = Number(tc.acquisitionBudgetRemaining) || 0;
      if (spVal + remVal > teamSumMaxBudget) teamSumMaxBudget = spVal + remVal;
    });
    annualFaabBudget = settingsAcqBudget > 0 ? settingsAcqBudget : (teamSumMaxBudget > 0 ? teamSumMaxBudget : 600);

    // Checkpoint 2 — Full-Season mTransactions2 Sweep across scoringPeriodId 1..maxSweepSp
    var txById = {};
    var rawTxs = [];
    function ingestRawTxList_(arr) {
      if (!Array.isArray(arr)) return;
      arr.forEach(function(tx) {
        if (!tx) return;
        var tid = String(tx.id || (String(tx.scoringPeriodId || 0) + "_" + String(tx.teamId || 0) + "_" + String(tx.processDate || tx.proposedDate || 0) + "_" + String(tx.bidAmount || 0)));
        if (!txById[tid]) {
          txById[tid] = true;
          rawTxs.push(tx);
        }
      });
    }
    ingestRawTxList_(data.transactions);

    for (var txSp = 1; txSp <= maxSweepSp; txSp++) {
      try {
        var txUrlSp = "https://lm-api-reads.fantasy.espn.com/apis/v3/games/ffl/seasons/" +
          seasonYr + "/segments/0/leagues/" + leagueId + "?scoringPeriodId=" + txSp + "&view=mTransactions2";
        var txFilterHeaders = Object.assign({}, headers, {
          "x-fantasy-filter": JSON.stringify({ transactions: { filterType: { value: ["WAIVER", "FREEAGENT", "TRADE_ACCEPTED"] } } })
        });
        var txRespSp = UrlFetchApp.fetch(txUrlSp, { method: "get", headers: txFilterHeaders, muteHttpExceptions: true });
        if (txRespSp.getResponseCode() !== 200) {
          txRespSp = UrlFetchApp.fetch(txUrlSp, { method: "get", headers: headers, muteHttpExceptions: true });
        }
        if (txRespSp.getResponseCode() === 200) {
          var txJsonSp = JSON.parse(txRespSp.getContentText());
          if (txJsonSp && Array.isArray(txJsonSp.transactions)) {
            ingestRawTxList_(txJsonSp.transactions);
          }
        }
      } catch (txSpErr) {}
    }

    var faabTeamsMap = {};
    rawTeamRows.forEach(function(stRow) {
      if (!stRow.managerId) return;
      faabTeamsMap[stRow.managerId] = {
        budget: annualFaabBudget,
        spent: Number(stRow.faabSpent) || 0,
        remaining: Math.max(0, annualFaabBudget - (Number(stRow.faabSpent) || 0)),
        moves: Number(stRow.acquisitions) || 0,
        zeroDollarAdds: 0,
        maxSingleBid: 0,
        maxSingleBidPct: 0,
        maxBidPlayer: "",
        totalWaiverPts: 0,
        bestBargainPts: 0,
        bestBargainBid: 0,
        bestBargainPlayer: "",
        worstBusterBid: 0,
        worstBusterBidPct: 0,
        worstBusterPts: 0,
        worstBusterPtsPerDollar: null,
        worstBusterPlayer: ""
      };
    });

    // Checkpoint 3 — Read parent tx.bidAmount and child ADD item playerId/toTeamId
    var addEvents = [];
    var dropEvents = [];
    var paidWaiverCount = 0;
    var busterCandidateCount = 0;
    var busterFloor = Math.ceil(annualFaabBudget * 0.25);

    rawTxs.forEach(function(tx) {
      if (!tx || String(tx.status || "").toUpperCase() !== "EXECUTED") return;
      var txType = String(tx.type || "").toUpperCase();
      var wk = Math.max(1, Number(tx.scoringPeriodId || tx.matchupPeriodId) || 1);
      var bid = Math.max(0, Number(tx.bidAmount || 0));
      var ts = Number(tx.processDate || tx.proposedDate) || 0;
      (Array.isArray(tx.items) ? tx.items : []).forEach(function(it) {
        var itType = String((it && it.type) || "").toUpperCase();
        var pidStr = String((it && it.playerId) || "").trim();
        if (!pidStr || pidStr === "0") return;
        if (itType === "ADD" && (txType === "WAIVER" || txType === "FREEAGENT" || txType.indexOf("WAIVER") !== -1 || txType.indexOf("FREEAGENT") !== -1)) {
          var toMgr = slotMap[String(it.toTeamId || tx.teamId)] || "";
          if (toMgr) {
            addEvents.push({ mgrId: toMgr, pid: pidStr, bid: bid, wk: wk, ts: ts, isWaiver: txType.indexOf("WAIVER") !== -1 });
          }
        } else if (itType === "DROP" || (itType === "TRADE" && it.fromTeamId)) {
          var fromMgr = slotMap[String(it.fromTeamId || tx.teamId)] || "";
          if (fromMgr) dropEvents.push({ mgrId: fromMgr, pid: pidStr, wk: wk, ts: ts });
        }
      });
    });
    addEvents.sort(function(a, b) { return (a.wk - b.wk) || (a.ts - b.ts); });
    dropEvents.sort(function(a, b) { return (a.wk - b.wk) || (a.ts - b.ts); });

    var bargainCap = Math.max(0, Math.round(annualFaabBudget * 0.02));
    var hasWkBox = Object.keys(weeklyPtsByPid).length > 0;
    var executedClaims = [];
    addEvents.forEach(function(ad) {
      var tRec = faabTeamsMap[ad.mgrId];
      if (!tRec) return;
      var nextDrop = null;
      for (var di = 0; di < dropEvents.length; di++) {
        var dr = dropEvents[di];
        if (dr.mgrId === ad.mgrId && dr.pid === ad.pid && (dr.wk > ad.wk || (dr.wk === ad.wk && dr.ts > ad.ts))) {
          nextDrop = dr;
          break;
        }
      }
      var endWk = nextDrop ? nextDrop.wk : 18;
      // Checkpoint 4 — Normalized String(playerId) join & explicit 0.0-point validity
      var rPts = 0.0;
      if (hasWkBox && weeklyPtsByPid[ad.pid]) {
        for (var w = ad.wk; w < endWk; w++) {
          var wVal = weeklyPtsByPid[ad.pid][w];
          if (wVal !== undefined && wVal !== null && !isNaN(Number(wVal))) rPts += Number(wVal);
        }
        if (ad.wk === endWk && weeklyPtsByPid[ad.pid][ad.wk] !== undefined) {
          rPts += Number(weeklyPtsByPid[ad.pid][ad.wk]) || 0;
        }
      }
      rPts = Math.round(rPts * 10) / 10;
      var pName = (playerMetaByPid[ad.pid] && playerMetaByPid[ad.pid].name) || ("Player #" + ad.pid);
      var bidPct = annualFaabBudget > 0 ? Math.round((ad.bid / annualFaabBudget) * 1000) / 10 : 0;

      if (ad.bid === 0) tRec.zeroDollarAdds += 1;
      if (ad.bid > 0) paidWaiverCount += 1;
      if (ad.bid > 0 && ad.bid > tRec.maxSingleBid) {
        tRec.maxSingleBid = ad.bid;
        tRec.maxSingleBidPct = bidPct;
        tRec.maxBidPlayer = pName;
      }
      tRec.totalWaiverPts = Math.round((tRec.totalWaiverPts + rPts) * 10) / 10;
      if (ad.bid <= bargainCap && rPts > tRec.bestBargainPts) {
        tRec.bestBargainPts = rPts;
        tRec.bestBargainBid = ad.bid;
        tRec.bestBargainPlayer = pName;
      }
      if (ad.bid >= busterFloor) {
        busterCandidateCount += 1;
        var ppd = Math.round((rPts / Math.max(1, ad.bid)) * 1000) / 1000;
        if (tRec.worstBusterPtsPerDollar === null || ppd < tRec.worstBusterPtsPerDollar || (ppd === tRec.worstBusterPtsPerDollar && ad.bid > tRec.worstBusterBid)) {
          tRec.worstBusterPtsPerDollar = ppd;
          tRec.worstBusterBid = ad.bid;
          tRec.worstBusterBidPct = bidPct;
          tRec.worstBusterPts = rPts;
          tRec.worstBusterPlayer = pName;
        }
      }
      executedClaims.push({
        teamId: ad.mgrId,
        playerId: "espn_p_" + ad.pid,
        playerName: pName,
        bid: ad.bid,
        bidPct: bidPct,
        addWeek: ad.wk,
        dropWeek: nextDrop ? nextDrop.wk : null,
        rosteredPts: rPts
      });
    });

    var seasonFaabObj = {
      isFaabEra: true,
      synced: true,
      budget: annualFaabBudget,
      hasExecutedClaims: executedClaims.length > 0,
      usedRosFallback: !hasWkBox,
      health: {
        budget: annualFaabBudget,
        paidWaiverCount: paidWaiverCount,
        busterCandidateCount: busterCandidateCount,
        sweptWeeks: maxSweepSp
      },
      teams: faabTeamsMap,
      claims: executedClaims.slice(0, 250)
    };
    if (!currentState.faabData || typeof currentState.faabData !== "object") {
      currentState.faabData = { defaultBudget: 600, annualBudget: annualFaabBudget, seasons: {}, records: { highestBids: [], mostSpentSeason: [], mostSavedSeason: [] } };
    }
    if (!currentState.faabData.seasons || typeof currentState.faabData.seasons !== "object") {
      currentState.faabData.seasons = {};
    }
    currentState.faabData.seasons[String(seasonYr)] = seasonFaabObj;

    if (!ed.seasons || typeof ed.seasons !== "object") ed.seasons = {};
    if (!ed.seasonStandings || typeof ed.seasonStandings !== "object") ed.seasonStandings = {};
    var prevSeasonObj = ed.seasons[String(seasonYr)] || {};
    ed.seasons[String(seasonYr)] = {
      year: seasonYr,
      seasonCompleted: isOfficiallyCompleted,
      playoffTeamCount: playoffTeamCnt,
      faabBudget: annualFaabBudget,
      champion: champId || prevSeasonObj.champion || "",
      runnerUp: runnerUpId || prevSeasonObj.runnerUp || "",
      toiletBowlWinner: toiletWinnerId || prevSeasonObj.toiletBowlWinner || "",
      phallicFeaster: feasterId || prevSeasonObj.phallicFeaster || "",
      standings: rawTeamRows,
      rosters: ed.rosters,
      faab: seasonFaabObj,
      matchups: syncedSeasonMatchups.length > 0 ? syncedSeasonMatchups : (prevSeasonObj.matchups || []),
      keptPlayers: syncedKeptPlayerList.length > 0 ? syncedKeptPlayerList : (Array.isArray(prevSeasonObj.keptPlayers) ? prevSeasonObj.keptPlayers : []),
      draftPicks: Array.isArray(prevSeasonObj.draftPicks) ? prevSeasonObj.draftPicks : undefined,
      archivedRecaps: prevSeasonObj.archivedRecaps || undefined
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

    var topRow = rawTeamRows[0];
    var topSummary = topRow ? (" • #1 Seed: " + (topRow.teamName || topRow.managerId || "TBD") + " (" + topRow.wins + "-" + topRow.losses + ")") : "";
    appendServerActivityEvent(
      currentState,
      "WEEKLY_SYNC",
      "Weekly ESPN Auto-Sync (" + seasonYr + (currentSp ? " • Wk " + currentSp : "") + ")",
      "Updated all 10 rosters, standings, All-Play & H2H records" + topSummary + ".",
      { year: seasonYr }
    );

    ed.lastSyncedAt = Date.now();
    currentState.updatedAt = Date.now();
    saveAndMirrorToSheets(sheets, currentState, "WEEKLY_ESPN_AUTO_SYNC");
  } catch (err) {
    Logger.log("autoWeeklyEspnSync error: " + err);
  } finally {
    lock.releaseLock();
  }
}

