/**
 * Sid's Gym — Google Apps Script Web App  (v3.3 — TWO-WAY SYNC + SETS MIRROR)
 * ---------------------------------------------------------------------------
 * This script is the cloud source of truth for the Sid's Gym app.
 * The website both READS (recall on load) and WRITES (save after changes)
 * the full app state through this single Web App URL.
 *
 * HOW IT STORES DATA
 *   - A hidden sheet named "_state" holds the entire app state as one JSON
 *     blob in cell A1. This is lossless and round-trips every data store
 *     (sessions, health, library, customMetrics, bodyData).
 *   - A human-readable "Sessions" sheet, "Health" sheet, and "Sets" sheet
 *     are also written so you can browse your data as a normal spreadsheet.
 *     These are for VIEWING only — the app reads from "_state".
 *   - The Sets sheet is fully rebuilt on every save, so deleted workouts
 *     are automatically removed from it.
 *
 * SETS SHEET COLUMNS
 *   Session ID | Date | Exercise Name | Set # | Reps | Weight | Unit | Notes
 *
 * ENDPOINTS
 *   GET   ?action=ping        -> { status:'ok', version:'3.3' }      (connection test)
 *   GET   ?action=getAll      -> { status:'ok', state:{...} }        (recall everything)
 *   POST  { action:'saveAll', state:{...} }                          (save everything)
 *   POST  { type:'fullSync', sessions, health, bodyData }            (back-compat)
 *   POST  { type:'session'|'health'|'measurement', ... }             (back-compat: merge one item)
 *
 * DEPLOY (one time):
 *   1. Open your Google Sheet → Extensions → Apps Script.
 *   2. Delete any old code, paste THIS file in full, Save.
 *   3. Deploy → New deployment → type "Web app".
 *        Execute as: Me
 *        Who has access: Anyone
 *   4. Authorize when prompted, copy the /exec Web App URL.
 *   5. Paste that URL into the app (Health tab → Google Sheets Sync).
 *
 *   IMPORTANT: any time you EDIT this script you must redeploy
 *   (Deploy → Manage deployments → Edit → New version) for changes to apply.
 * ---------------------------------------------------------------------------
 */

var VERSION    = '3.3';
var STATE_SHEET = '_state';   // hidden JSON blob — the real source of truth
var SESS_SHEET  = 'Sessions'; // human-readable mirror
var HEALTH_SHEET= 'Health';   // human-readable mirror
var SETS_SHEET  = 'Sets';     // human-readable sets/exercises mirror

// ── ENTRY POINTS ────────────────────────────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ping';
  try {
    if (action === 'getAll') {
      return json({ status: 'ok', version: VERSION, state: readState() });
    }
    // default: ping / connection test
    return json({ status: 'ok', version: VERSION });
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    // New full two-way payload
    if (body.action === 'saveAll' || body.type === 'fullSync') {
      var incoming = body.state || {
        sessions:      body.sessions || {},
        health:        body.health   || {},
        bodyData:      body.bodyData || null,
        library:       body.library  || null,
        customMetrics: body.customMetrics || null
      };
      var merged = mergeState(readState(), incoming);
      writeState(merged);
      return json({ status: 'ok', version: VERSION, state: merged });
    }

    // Back-compat: single-item merges (still supported, app no longer relies on them)
    var state = readState();
    if (body.type === 'session' && body.sessionId) {
      state.sessions = state.sessions || {};
      state.sessions[body.sessionId] = body.session;
      writeState(state);
      return json({ status: 'ok', version: VERSION });
    }
    if (body.type === 'health' && body.metrics) {
      state.health = state.health || {};
      var dateKey = body.metrics.date || new Date().toISOString().slice(0, 10);
      state.health[dateKey] = body.metrics;
      writeState(state);
      return json({ status: 'ok', version: VERSION });
    }
    if (body.type === 'measurement' && body.measurement) {
      state.bodyData = state.bodyData || { height: null, measurements: {} };
      state.bodyData.measurements = state.bodyData.measurements || {};
      var m = body.measurement;
      var arr = state.bodyData.measurements[m.id] || [];
      // avoid duplicates by timestamp
      if (!arr.some(function (x) { return x.ts === m.ts; })) {
        arr.push({ val: m.val, date: m.date, ts: m.ts });
      }
      state.bodyData.measurements[m.id] = arr;
      writeState(state);
      return json({ status: 'ok', version: VERSION });
    }

    return json({ status: 'error', message: 'Unknown payload', received: body.type || body.action || null });
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  }
}

// ── STATE READ / WRITE (JSON blob in _state!A1) ──────────────────────────────
function readState() {
  var sh = getSheet(STATE_SHEET, true);
  var raw = sh.getRange(1, 1).getValue();
  if (!raw) return emptyState();
  try {
    var parsed = JSON.parse(raw);
    // ensure all keys exist
    return mergeState(emptyState(), parsed);
  } catch (err) {
    return emptyState();
  }
}

function writeState(state) {
  var sh = getSheet(STATE_SHEET, true);
  state.updatedAt = new Date().toISOString();
  sh.getRange(1, 1).setValue(JSON.stringify(state));
  // Also refresh the human-readable mirrors (best-effort)
  try { writeReadableMirrors(state); } catch (err) {}
}

function emptyState() {
  return {
    sessions: {},
    health: {},
    library: null,
    customMetrics: null,
    bodyData: { height: null, measurements: {} },
    updatedAt: null
  };
}

/**
 * Merge two states. Cloud + incoming. For keyed collections (sessions, health)
 * we keep the union; on conflicts the incoming (most recent client) wins.
 * library / customMetrics / bodyData take incoming when provided.
 */
function mergeState(base, incoming) {
  base = base || emptyState();
  incoming = incoming || {};
  var out = {
    sessions: {},
    health: {},
    library: base.library || null,
    customMetrics: base.customMetrics || null,
    bodyData: base.bodyData || { height: null, measurements: {} }
  };

  // sessions — union, incoming wins on conflict
  Object.keys(base.sessions || {}).forEach(function (k) { out.sessions[k] = base.sessions[k]; });
  Object.keys(incoming.sessions || {}).forEach(function (k) { out.sessions[k] = incoming.sessions[k]; });

  // health — union, incoming wins
  Object.keys(base.health || {}).forEach(function (k) { out.health[k] = base.health[k]; });
  Object.keys(incoming.health || {}).forEach(function (k) { out.health[k] = incoming.health[k]; });

  if (incoming.library) out.library = incoming.library;
  if (incoming.customMetrics) out.customMetrics = incoming.customMetrics;
  if (incoming.bodyData) out.bodyData = mergeBody(out.bodyData, incoming.bodyData);

  return out;
}

function mergeBody(base, incoming) {
  base = base || { height: null, measurements: {} };
  incoming = incoming || {};
  var out = { height: incoming.height != null ? incoming.height : base.height, measurements: {} };
  var meas = {};
  (function copy(src) {
    Object.keys(src || {}).forEach(function (id) {
      meas[id] = meas[id] || [];
      (src[id] || []).forEach(function (entry) {
        if (!meas[id].some(function (x) { return x.ts === entry.ts; })) meas[id].push(entry);
      });
    });
  })(base.measurements);
  (function copy(src) {
    Object.keys(src || {}).forEach(function (id) {
      meas[id] = meas[id] || [];
      (src[id] || []).forEach(function (entry) {
        if (!meas[id].some(function (x) { return x.ts === entry.ts; })) meas[id].push(entry);
      });
    });
  })(incoming.measurements);
  out.measurements = meas;
  return out;
}

// ── HUMAN-READABLE MIRRORS (optional, for browsing in the sheet) ─────────────
function writeReadableMirrors(state) {
  // Sessions sheet
  var ss = getSheet(SESS_SHEET, false);
  ss.clearContents();
  ss.getRange(1, 1, 1, 7).setValues([['Session ID', 'Date', 'Type', 'Duration (s)', 'Started', 'Finished', 'Summary']]);
  var rows = [];
  Object.keys(state.sessions || {}).sort().forEach(function (id) {
    var s = state.sessions[id] || {};
    var summary = '';
    if (s.cardio) {
      summary = 'cardio ' + (s.cardio.subType || '') + (s.cardio.dist ? ' ' + s.cardio.dist + 'mi' : '');
    } else if (s.exercises) {
      summary = s.exercises.length + ' exercises';
    }
    rows.push([id, s.date || '', s.type || '', s.duration || '', fmt(s.startedAt), fmt(s.finishedAt), summary]);
  });
  if (rows.length) ss.getRange(2, 1, rows.length, 7).setValues(rows);

  // Health sheet
  var hs = getSheet(HEALTH_SHEET, false);
  hs.clearContents();
  hs.getRange(1, 1, 1, 7).setValues([['Date', 'RHR', 'HRV', 'Recovery', 'VO2', 'Sleep', 'Bodyweight']]);
  var hrows = [];
  Object.keys(state.health || {}).sort().forEach(function (d) {
    var m = state.health[d] || {};
    hrows.push([d, m.rhr || '', m.hrv || '', m.recovery || '', m.vo2 || '', m.sleep || '', m.bw || '']);
  });
  if (hrows.length) hs.getRange(2, 1, hrows.length, 7).setValues(hrows);

  // Sets sheet — fully rebuilt every save so deletions are reflected automatically
  writeSetsSheet(state);
}

/**
 * Rebuild the Sets sheet from scratch using the current state.
 * Each strength exercise set gets its own row.
 * Cardio sessions are written as a single summary row (no individual sets).
 * Because the sheet is fully cleared and rewritten, any workout deleted
 * from the app will automatically disappear from this sheet on the next save.
 *
 * Columns:
 *   Session ID | Date | Exercise Name | Set # | Reps | Weight | Unit | Notes
 */
function writeSetsSheet(state) {
  var ws = getSheet(SETS_SHEET, false);
  ws.clearContents();
  ws.getRange(1, 1, 1, 8).setValues([
    ['Session ID', 'Date', 'Exercise Name', 'Set #', 'Reps', 'Weight', 'Unit', 'Notes']
  ]);

  var setsRows = [];

  Object.keys(state.sessions || {}).sort().forEach(function (id) {
    var s = state.sessions[id] || {};
    var date = s.date || '';

    // Strength / resistance session — iterate exercises and their sets
    if (s.exercises && s.exercises.length) {
      s.exercises.forEach(function (ex) {
        var exName = ex.name || ex.id || 'Unknown';
        var sets = ex.sets || [];
        if (sets.length === 0) {
          // Exercise logged but no set data
          setsRows.push([id, date, exName, '', '', '', '', '']);
        } else {
          sets.forEach(function (set, idx) {
            var setNum  = (set.setNum  != null) ? set.setNum  : (idx + 1);
            var reps    = (set.reps    != null) ? set.reps    : '';
            var weight  = (set.weight  != null) ? set.weight  : '';
            var unit    = set.unit    || 'lbs';
            var notes   = set.notes   || '';
            setsRows.push([id, date, exName, setNum, reps, weight, unit, notes]);
          });
        }
      });

    // Cardio session — single summary row
    } else if (s.cardio) {
      var c = s.cardio;
      var cardioName = 'Cardio' + (c.subType ? ' — ' + c.subType : '');
      var cardioDetail = [
        c.dist  ? c.dist + ' mi'            : '',
        c.time  ? c.time + ' min'           : '',
        c.pace  ? 'pace ' + c.pace          : '',
        c.cals  ? c.cals + ' kcal'          : ''
      ].filter(Boolean).join(', ');
      setsRows.push([id, date, cardioName, '', '', '', '', cardioDetail]);
    }
  });

  if (setsRows.length) {
    ws.getRange(2, 1, setsRows.length, 8).setValues(setsRows);
  }
}

function fmt(ts) {
  if (!ts) return '';
  try { return new Date(Number(ts)).toISOString(); } catch (e) { return ''; }
}

// ── HELPERS ──────────────────────────────────────────────────────────────────
function getSheet(name, hidden) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (hidden) sh.hideSheet();
  }
  return sh;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
