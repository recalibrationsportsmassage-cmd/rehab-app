/**
 * Recalibration rehab app — backend v2
 * Lives inside the Google Sheet (Extensions → Apps Script).
 * Deploy as Web App: Execute as Me, access: Anyone.
 */

const COACH_PASSWORD = "Mushr00ms";

/**
 * Tabs that are never clients. Matching is case-insensitive and ignores
 * extra spaces. Any tab whose name starts with "_" is also skipped, so you
 * can hide a working tab by renaming it "_scratch" without editing this list.
 */
const SKIP_TABS = [
  "Library", "Template", "Log", "Pain Scale", "Painscale",
  "Exercise Matrix", "Matrix", "Notes", "Settings"
];

const LOG_TAB = "Log";
const TEMPLATE_TAB = "Template";
const LIBRARY_TAB = "Library";

// Client tab headers (found by name, order irrelevant)
const H = {
  START: "START", END: "END", STATUS: "STATUS", NUM: "#",
  EXERCISE: "EXERCISE", TYPE: "TYPE", CUES: "CUES", VIDEO: "VIDEO",
  SETS: "SETS", TARGET: "TARGET", REPS: "REP RANGE", UNIT: "REP UNIT",
  SIDE: "PER SIDE", FREQ: "FREQUENCY",
  SESSIONS: "SESSIONS LOGGED", LAST: "LAST LOGGED", BEN: "BEN'S NOTES"
};

// Log tab column order
const LOG_COLS = ["TIMESTAMP","CLIENT","WEEK START","ROW","EXERCISE","SESSION","SET","LOAD","REPS","INTENSITY","NOTE"];

/* ---------------- helpers ---------------- */

/**
 * Tidy a value typed by a human: strips non-breaking and zero-width spaces,
 * collapses runs of whitespace, trims the ends. Bad data entry in the Exercise
 * Matrix (a trailing space, a double space) then can't break matching.
 */
function clean(s){
  return String(s == null ? "" : s)
    .replace(/[   ]/g, " ")   // non-breaking spaces
    .replace(/[​-‍﻿]/g, "")   // zero-width junk
    .replace(/\s+/g, " ")
    .trim();
}

function slugify(s){
  return clean(s).toLowerCase().replace(/[^a-z0-9\s-]/g,"").replace(/\s+/g,"-");
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }

/* ---------------- caching ----------------
   Reading a client tab plus the whole Log is the slow part. Cache the built
   program per client and throw the cache away whenever that client's data
   changes, so reads are fast but never stale.                             */

/* Short enough that a stale read can't linger, long enough to absorb the
   repeated calls a page makes while you click around. Writes clear it anyway. */
var CACHE_SECONDS = 120;

function cacheKey(clientName){ return "prog:" + clientName; }

function bumpVersion(clientName){
  try{ CacheService.getScriptCache().remove(cacheKey(clientName)); }catch(e){}
}

function cacheGet(clientName){
  try{
    var v = CacheService.getScriptCache().get(cacheKey(clientName));
    return v ? JSON.parse(v) : null;
  }catch(e){ return null; }
}

function cachePut(clientName, obj){
  try{
    var s = JSON.stringify(obj);
    // script cache tops out around 100KB per entry
    if(s.length < 95000) CacheService.getScriptCache().put(cacheKey(clientName), s, CACHE_SECONDS);
  }catch(e){}
}

function norm(s){ return clean(s).toLowerCase(); }

function isSkipped(name){
  if(norm(name).indexOf("_") === 0) return true;
  var n = norm(name);
  for(var i=0;i<SKIP_TABS.length;i++){
    if(norm(SKIP_TABS[i]) === n) return true;
  }
  // a client tab must look like a name: at least two words, no digits
  if(n.split(" ").length < 2) return true;
  if(/\d/.test(n)) return true;
  return false;
}

function clientSheets(){
  return ss().getSheets().filter(function(sh){ return !isSkipped(sh.getName()); });
}

function findSheet(clientSlug){
  var m = clientSheets().filter(function(sh){ return slugify(sh.getName()) === clientSlug; });
  return m.length ? m[0] : null;
}

function headerMap(sheet){
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var map = {};
  headers.forEach(function(h,i){ map[String(h).trim().toUpperCase()] = i; });
  return map;
}

function toIso(v){
  if(v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v || "");
}

function splitCues(text){
  if(!text) return [];
  return String(text)
    .split(/\n|(?:^|\s)[-–•]\s+/)
    .map(function(s){ return s.trim().replace(/^[-–•]\s*/,""); })
    .filter(function(s){ return s.length > 1; });
}

function freqNum(v){
  var s = clean(v);
  if(!s) return 1;
  if(/daily/i.test(s)) return 7;
  if(/as\s*need/i.test(s)) return 1;
  var n = parseInt(s.replace(/[^0-9]/g,""),10);
  return n > 0 ? n : 1;
}

/* A single save can cover at most this many weeks, to stop a mistyped due date
   writing hundreds of rows. */
var MAX_WEEKS = 12;

/**
 * Split a start/due range into one block per week.
 *
 * Rounds up, so a 10-day range becomes two weeks with a short second one.
 * With repeat off (or no due date) it returns a single block spanning the
 * whole range, which is how saving worked before.
 */
function weekSchedule(startIso, endIso, repeat){
  var s = startIso ? new Date(startIso) : null;
  var e = endIso   ? new Date(endIso)   : null;
  if(!s) return [{ start:"", end: e || "" }];
  if(!e || !repeat) return [{ start:s, end: e || "" }];

  var days = Math.round((e.getTime() - s.getTime()) / 86400000) + 1;
  if(days < 1) days = 1;
  var n = Math.min(MAX_WEEKS, Math.max(1, Math.ceil(days / 7)));

  var out = [];
  for(var i = 0; i < n; i++){
    var ws = new Date(s.getTime() + i * 7 * 86400000);
    var we = new Date(ws.getTime() + 6 * 86400000);
    if(we.getTime() > e.getTime()) we = e;
    out.push({ start: ws, end: we });
  }
  return out;
}

/**
 * If a cell has a dropdown, snap the incoming value to whichever option it
 * matches loosely — so "3×" lands as "3x", "daily" as "Daily", and so on.
 * Returns the original value if there's no dropdown or no sensible match.
 */
function matchValidation(cell, val){
  var rule;
  try{ rule = cell.getDataValidation(); }catch(e){ return val; }
  if(!rule) return val;

  // Tick box columns store TRUE/FALSE — unless the box was given custom
  // "checked"/"unchecked" values, in which case use those instead.
  var type;
  try{ type = rule.getCriteriaType(); }catch(e){ type = null; }
  if(type && String(type) === "CHECKBOX"){
    var on = String(val).trim().toLowerCase();
    var yes = (["yes","y","true","1"].indexOf(on) !== -1);
    var cv = [];
    try{ cv = rule.getCriteriaValues() || []; }catch(e){}
    if(cv.length >= 2) return yes ? cv[0] : cv[1];   // custom tick box values
    if(cv.length === 1) return yes ? cv[0] : false;
    return yes;
  }

  var opts;
  try{
    var crit = rule.getCriteriaValues();
    opts = (crit && crit[0] && crit[0].length) ? crit[0] : null;
  }catch(e){ return val; }
  if(!opts) return val;

  /* Normalise before stripping punctuation, otherwise "×" is thrown away and
     "3×" can never match "3x". Map the multiplication sign to a letter x first. */
  var loose = function(s){
    return String(s).toLowerCase()
      .replace(/[×✕✖]/g, "x")
      .replace(/[^a-z0-9]/g, "");
  };
  var want = loose(val);

  for(var i=0;i<opts.length;i++){
    if(String(opts[i]) === String(val)) return opts[i];
  }
  for(var j=0;j<opts.length;j++){
    if(loose(opts[j]) === want) return opts[j];
  }
  return val;
}

/**
 * Decide how an exercise behaves in the client app.
 *
 * The Matrix says Mobility, Strength or Both. "Both" (and a blank Type) can't
 * be resolved from the Matrix alone, so fall back to the intensity target you
 * set — Low/Mod is mobility, anything with RIR is strength. That's the same
 * decision you made when you put it in one section or the other.
 */
function typeFor(typeRaw, target){
  var t = clean(typeRaw).toLowerCase();
  if(t.indexOf("mob") === 0) return "mobility";
  if(t.indexOf("str") === 0) return "strength";

  var g = clean(target).toLowerCase();
  if(/rir/.test(g)) return "strength";
  if(/^(low|mod)/.test(g)) return "mobility";
  return "strength";   // safe default: asks for load, nothing is hidden
}

/** PER SIDE is a tick box — always write a real boolean, never "Yes"/"No". */
function sideBool(v){
  var s = String(v == null ? "" : v).trim().toLowerCase();
  return (["yes","y","true","1"].indexOf(s) !== -1);
}

/**
 * A cell that should hold a number but might contain a Date, because the column
 * was formatted as a date at some point. Sheets stores 2 as 1900-01-01, so we
 * can convert it back rather than losing the value.
 */
function toNum(v){
  if(v instanceof Date){
    var epoch = new Date(1899, 11, 30).getTime();
    var days = Math.round((v.getTime() - epoch) / 86400000);
    return (days >= 0 && days < 1000) ? days : 0;   // sane session counts only
  }
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function linkFrom(richCell, rawValue){
  var video = "";
  if(richCell){
    video = richCell.getLinkUrl() || "";
    if(!video){
      var runs = richCell.getRuns();
      for(var r=0; r<runs.length && !video; r++) video = runs[r].getLinkUrl() || "";
    }
  }
  if(!video){
    var raw = clean(rawValue);
    if(/^https?:\/\//i.test(raw)) video = raw;
  }
  return video;
}

/* ---------------- log tab ---------------- */

function logSheet(){
  var sh = ss().getSheetByName(LOG_TAB);
  if(!sh){
    sh = ss().insertSheet(LOG_TAB);
    sh.getRange(1,1,1,LOG_COLS.length).setValues([LOG_COLS]);
  }
  return sh;
}

/** All log rows for one client, keyed by "row|session" -> {sets:{setNo:[l,r,i]}, note:""} */
function readLog(clientName){
  var sh = logSheet();
  var last = sh.getLastRow();
  if(last < 2) return {};
  var vals = sh.getRange(2,1,last-1,LOG_COLS.length).getValues();
  var out = {};
  for(var i=0;i<vals.length;i++){
    var v = vals[i];
    if(String(v[1]).trim() !== clientName) continue;
    var key = String(v[3]) + "|" + String(v[5]);
    if(!out[key]) out[key] = { sets:{}, note:"", date:"" };
    out[key].sets[String(v[6])] = [String(v[7]), String(v[8]), String(v[9])];
    if(v[10]) out[key].note = String(v[10]);
    out[key].date = toIso(v[0]);
  }
  return out;
}

/* ---------------- client app: read program ---------------- */

function getProgram(clientSlug, skipCache){
  var sheet = findSheet(clientSlug);
  if(!sheet) return { error:"not_found" };

  if(!skipCache){
    var hit = cacheGet(sheet.getName());
    if(hit) return hit;
  }

  var m = headerMap(sheet);
  var lastRow = sheet.getLastRow();
  if(lastRow < 2) return { name: sheet.getName(), weeks: [] };

  var nCols = sheet.getLastColumn();
  var values = sheet.getRange(2,1,lastRow-1,nCols).getValues();
  var rich = sheet.getRange(2,1,lastRow-1,nCols).getRichTextValues();
  var log = readLog(sheet.getName());

  var byStart = {};
  values.forEach(function(row,i){
    var name = clean(row[m[H.EXERCISE]]);
    if(!name) return;
    var start = toIso(row[m[H.START]]);
    if(!start) return;
    var end = toIso(row[m[H.END]]);
    var sheetRow = i + 2;

    var typeRaw = (H.TYPE in m) ? clean(row[m[H.TYPE]]) : "";
    var sideVal = (H.SIDE in m) ? row[m[H.SIDE]] : false;
    var freqRaw = (H.FREQ in m) ? clean(row[m[H.FREQ]]) : "";

    var sessions = {};
    var totalSessions = freqNum(freqRaw);
    for(var s=1; s<=totalSessions; s++){
      var k = sheetRow + "|" + s;
      if(log[k]) sessions[s] = log[k];
    }

    var ex = {
      row: sheetRow,
      name: name,
      type: typeFor(typeRaw, row[m[H.TARGET]]),
      cues: splitCues(row[m[H.CUES]]),
      video: linkFrom(rich[i][m[H.VIDEO]], row[m[H.VIDEO]]),
      sets: Math.max(1, Math.min(parseInt(row[m[H.SETS]],10) || 1, 5)),
      target: clean(row[m[H.TARGET]]),
      reps: clean(row[m[H.REPS]]),
      unit: (H.UNIT in m) ? clean(row[m[H.UNIT]]) : "",
      perSide: sideVal === true || ["TRUE","YES","Y"].indexOf(String(sideVal).trim().toUpperCase()) !== -1,
      freq: freqRaw,
      freqNum: totalSessions,
      sessions: sessions
    };

    if(!byStart[start]) byStart[start] = { start:start, end:end, exercises:[] };
    if(end && !byStart[start].end) byStart[start].end = end;
    byStart[start].exercises.push(ex);
  });

  var weeks = Object.keys(byStart).sort().map(function(k){ return byStart[k]; });
  var out = { name: sheet.getName(), weeks: weeks };
  cachePut(sheet.getName(), out);
  return out;
}

/* ---------------- client app: save a session ---------------- */

function saveSession(body){
  var sheet = findSheet(String(body.client||"").toLowerCase().trim());
  if(!sheet) return { error:"not_found" };

  var row = parseInt(body.row,10);
  if(!row || row < 2 || row > sheet.getLastRow()) return { error:"bad_row" };

  var session = parseInt(body.session,10) || 1;
  var sets = body.sets || [];
  var note = String(body.note||"").trim();
  var clientName = sheet.getName();

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var m = headerMap(sheet);
    var weekStart = toIso(sheet.getRange(row, m[H.START]+1).getValue());
    var exName = clean(sheet.getRange(row, m[H.EXERCISE]+1).getValue());

    var lg = logSheet();
    var last = lg.getLastRow();

    // remove any previous rows for this client/row/session so re-saving overwrites
    if(last >= 2){
      var vals = lg.getRange(2,1,last-1,LOG_COLS.length).getValues();
      for(var i=vals.length-1; i>=0; i--){
        if(String(vals[i][1]).trim() === clientName &&
           String(vals[i][3]) == String(row) &&
           String(vals[i][5]) == String(session)){
          lg.deleteRow(i+2);
        }
      }
    }

    var now = new Date();
    var num = function(v){ return (v !== "" && v !== null && !isNaN(v)) ? Number(v) : (v || ""); };
    var out = sets.map(function(s,idx){
      return [now, clientName, weekStart, row, exName, session, idx+1,
              num(s[0]), num(s[1]), num(s[2]), idx === 0 ? note : ""];
    });
    if(out.length) lg.getRange(lg.getLastRow()+1, 1, out.length, LOG_COLS.length).setValues(out);

    // roll-up columns on the client tab
    var log = readLog(clientName);
    var count = 0, latest = "";
    Object.keys(log).forEach(function(k){
      if(k.split("|")[0] == String(row)){
        count++;
        if(log[k].date > latest) latest = log[k].date;
      }
    });
    if(H.SESSIONS in m){
      var sc = sheet.getRange(row, m[H.SESSIONS]+1);
      try{ sc.setNumberFormat("0"); }catch(e){}   // stop a date-formatted column mangling the count
      sc.setValue(count);
    }
    if(H.LAST in m && latest){
      var lc = sheet.getRange(row, m[H.LAST]+1);
      try{ lc.setNumberFormat("dd/mm/yyyy"); }catch(e){}
      lc.setValue(new Date(latest));
    }
  } finally {
    lock.releaseLock();
  }
  bumpVersion(clientName);
  return { ok:true };
}

/* ---------------- archiving ----------------
   Kept in script properties rather than the sheet, so archiving never touches
   a tab name — a client's link keeps working either way.                    */

function archivedList(){
  try{
    var raw = PropertiesService.getScriptProperties().getProperty("archived");
    return raw ? JSON.parse(raw) : [];
  }catch(e){ return []; }
}

function setArchived(name, on){
  name = String(name||"").trim();
  if(!name) return { error:"bad_name" };
  var list = archivedList();
  var i = list.indexOf(name);
  if(on && i === -1) list.push(name);
  if(!on && i !== -1) list.splice(i,1);
  try{
    PropertiesService.getScriptProperties().setProperty("archived", JSON.stringify(list));
  }catch(e){ return { error:"save_failed" }; }
  return { ok:true, archived: on };
}

/* ---------------- coach: clients ---------------- */

function listClients(){
  var out = [];
  var arch = archivedList();
  clientSheets().forEach(function(sh){
    var m = headerMap(sh);
    var last = sh.getLastRow();
    var wkLabel = "No weeks yet", due = "", done = 0, total = 0, starts = [];
    if(last >= 2 && (H.START in m)){
      var vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
      vals.forEach(function(r){
        var st = toIso(r[m[H.START]]);
        if(!st || !String(r[m[H.EXERCISE]]||"").trim()) return;
        if(starts.indexOf(st) === -1) starts.push(st);
      });
      starts.sort();
      if(starts.length){
        var latest = starts[starts.length-1];
        wkLabel = "Week " + starts.length;
        vals.forEach(function(r){
          if(toIso(r[m[H.START]]) !== latest) return;
          if(!String(r[m[H.EXERCISE]]||"").trim()) return;
          var f = freqNum(r[m[H.FREQ]]);
          total += f;
          done += Math.min(toNum(r[m[H.SESSIONS]]), f);   // never count more than prescribed
          var e = toIso(r[m[H.END]]);
          if(e && (!due || e < due)) due = e;
        });
      }
    }
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var isArchived = arch.indexOf(sh.getName()) !== -1;
    var status = "open";
    if(isArchived) status = "archived";
    else if(total > 0 && done >= total) status = "done";
    else if(due && due < today && done < total) status = "late";
    out.push({ name: sh.getName(), slug: slugify(sh.getName()), week: wkLabel,
               due: due, done: done, total: total, status: status,
               archived: isArchived, weeks: starts });
  });
  return { clients: out };
}

/* ---------------- coach: library ---------------- */

function getLibrary(){
  var sh = ss().getSheetByName(LIBRARY_TAB);
  if(!sh) return { library: [] };
  var last = sh.getLastRow();
  if(last < 2) return { library: [] };
  var m = headerMap(sh);
  var nameCol = ("EXERCISE" in m) ? m["EXERCISE"] :
                ("MOBILITY EXERCISE" in m) ? m["MOBILITY EXERCISE"] : 0;
  var typeCol = ("TYPE" in m) ? m["TYPE"] : 6;
  var vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  var out = [], seen = {};
  vals.forEach(function(r){
    var n = clean(r[nameCol]);
    if(!n || seen[n]) return;
    seen[n] = 1;
    var t = clean(r[typeCol]);
    // blank stays blank — don't guess, so a missing Type is visible in the UI
    // Both means it can be prescribed as either — offered in both sections
    var type = /both/i.test(t) ? "Both"
             : /mob/i.test(t) ? "Mobility"
             : /str/i.test(t) ? "Strength" : "";
    out.push({ name:n, type:type });
  });
  out.sort(function(a,b){ return a.name.toLowerCase().localeCompare(b.name.toLowerCase()); });
  return { library: out };
}

/* ---------------- coach: create client ---------------- */

function createClient(name){
  name = clean(name);
  // must look like a person: at least two words, letters only
  if(name.split(" ").length < 2) return { error:"needs_two_names" };
  if(name.length < 3) return { error:"bad_name" };
  if(/\d/.test(name)) return { error:"bad_name" };
  if(ss().getSheetByName(name)) return { error:"exists" };

  var tpl = ss().getSheetByName(TEMPLATE_TAB);
  if(!tpl) return { error:"no_template" };

  var sh = tpl.copyTo(ss());
  sh.setName(name);

  // The Template may carry sample values in its first rows. Clear everything
  // a coach would type, but leave the formula columns and dropdowns intact.
  var m = headerMap(sh);
  var last = sh.getMaxRows();
  [H.START,H.END,H.EXERCISE,H.SETS,H.TARGET,H.REPS,H.UNIT,H.SIDE,H.FREQ,H.BEN,
   H.SESSIONS,H.LAST].forEach(function(key){
    if(key in m) sh.getRange(2, m[key]+1, last-1, 1).clearContent();
  });

  ss().setActiveSheet(sh);
  ss().moveActiveSheet(ss().getSheets().length);
  return { ok:true, name:name, slug:slugify(name) };
}

/* ---------------- coach: save a week ---------------- */

/** Reject bad rows before touching the sheet, so nothing half-writes. */
function validateRows(rows){
  var problems = [];
  rows.forEach(function(r,i){
    var label = clean(r.ex) || ("row " + (i+1));
    if(!clean(r.ex)) problems.push(label + ": no exercise selected");
    var sets = clean(r.sets);
    if(sets !== "" && !/^[1-5]$/.test(sets)) problems.push(label + ": sets must be a whole number 1–5, got \"" + sets + "\"");
    // free text is fine when the unit is Text — "As many as possible" etc.
    var unit = clean(r.unit).toLowerCase();
    var reps = clean(r.reps);
    if(reps && unit !== "text" && !/^[0-9]+(\s*[-–]\s*[0-9]+)?$/.test(reps))
      problems.push(label + ": reps must be a number or range like 8-10 — pick unit \"Text\" if you want words. Got \"" + reps + "\"");
    if(!clean(r.freq)) problems.push(label + ": no frequency set");
  });
  return problems;
}

function saveWeek(body){
  var sheet = findSheet(String(body.client||"").toLowerCase().trim());
  if(!sheet) return { error:"not_found" };
  var rows = body.rows || [];
  if(!rows.length) return { error:"no_rows" };

  var problems = validateRows(rows);
  if(problems.length) return { error:"validation", problems: problems };

  /* One block per week when the range spans more than seven days. Each block
     is a separate week in the sheet, so the client logs them independently and
     week-on-week comparisons work across the whole period. */
  var schedule = weekSchedule(body.start, body.end, body.repeat !== false);
  var expanded = [];
  schedule.forEach(function(wk){
    rows.forEach(function(r){
      var copy = {};
      for(var k in r) copy[k] = r[k];
      copy._start = wk.start;
      copy._end   = wk.end;
      expanded.push(copy);
    });
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var m = headerMap(sheet);
    var failures = [];

    /* Don't use getLastRow(): the Template has formulas filled down, so it
       reports the last row containing anything (often several hundred).
       Find the last row that actually has an exercise name instead. */
    var startRow = 2;
    if(H.EXERCISE in m){
      var exCol = m[H.EXERCISE] + 1;
      var deep = sheet.getMaxRows();
      var exVals = sheet.getRange(1, exCol, deep, 1).getValues();
      for(var r = exVals.length - 1; r >= 1; r--){
        if(String(exVals[r][0]).trim() !== ""){ startRow = r + 2; break; }
      }
    }
    // make room if we're writing past the current sheet size
    var needed = startRow + expanded.length - 1;
    if(needed > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows());

    expanded.forEach(function(r,i){
      var row = startRow + i;
      // Skip blanks: a column with "Reject the input" validation throws if you
      // write "" into it, which would kill the whole save.
      /* PER SIDE has collected conflicting rules over time (a Yes/No list, then
         a tick box). Write the boolean; if the cell rejects it, replace the rule
         with a clean tick box so the column heals itself as weeks are saved.  */
      var putSide = function(row, mm, v){
        if(!(H.SIDE in mm)) return;
        var cell = sheet.getRange(row, mm[H.SIDE]+1);
        var b = sideBool(v);
        try{ cell.setNumberFormat("@"); }catch(e){}
        try{
          cell.setValue(matchValidation(cell, b));
        }catch(err){
          try{
            cell.setDataValidation(
              SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build()
            );
            cell.setValue(b);
          }catch(err2){
            try{ cell.setDataValidation(null); cell.setValue(b); }catch(err3){}
          }
        }
      };

      /* Write one cell, defensively.
         - forces a sane number format first: a column accidentally formatted as
           a date turns 1 into "31-12-1899", which is where "1-1" came from
         - if a validation rule rejects the value, drop the rule for that cell
           rather than failing the whole save                                */
      var put = function(key,val,fmt){
        if(!(key in m)) return true;
        if(val === undefined || val === null || val === "") return true;
        var cell = sheet.getRange(row, m[key]+1);
        try{ cell.setNumberFormat(fmt || "@"); }catch(e){}
        val = matchValidation(cell, val);
        try{
          cell.setValue(val);
          return true;
        }catch(err){
          try{
            cell.setDataValidation(null);
            cell.setValue(val);
            return true;
          }catch(err2){
            failures.push(String(r.ex||key) + ": " + (err2.message || err2));
            return false;
          }
        }
      };
      put(H.START, r._start, "dd/mm/yyyy");
      put(H.END,   r._end,   "dd/mm/yyyy");
      put(H.EXERCISE, r.ex, "@");
      /* SETS: the dropdown holds text ("1","2","3"), so writing the number 1
         trips validation and a decimal format shows it as 1.00. Let
         matchValidation snap it to the dropdown's own value. */
      put(H.SETS, clean(r.sets), "0");
      put(H.TARGET, r.target, "@");
      put(H.REPS, r.reps, "@");
      put(H.UNIT, r.unit, "@");
      putSide(row, m, r.side);
      put(H.FREQ, r.freq, "@");
      put(H.BEN, r.note || "", "@");
    });

    // copy the formula columns (#, STATUS, TYPE, CUES, VIDEO) down from row 2
    if(startRow > 2){
      [H.NUM, H.STATUS, H.TYPE, H.CUES, H.VIDEO].forEach(function(key){
        if(!(key in m)) return;
        var col = m[key] + 1;
        var src = sheet.getRange(2, col);
        if(src.getFormula()){
          src.copyTo(sheet.getRange(startRow, col, expanded.length, 1));
        }
      });
    }
    SpreadsheetApp.flush();

    /* Read the exercise names back. A mismatch is usually the cell's dropdown
       rejecting the value silently, so retry once without the rule before
       giving up — and if it still fails, say what's actually in the cell.   */
    var exCol2 = m[H.EXERCISE] + 1;
    var loose = function(s){ return String(s).toLowerCase().replace(/[^a-z0-9]/g,""); };
    var written = sheet.getRange(startRow, exCol2, expanded.length, 1).getValues();

    var retried = false;
    expanded.forEach(function(r,i){
      if(loose(written[i][0]) === loose(r.ex)) return;
      var cell = sheet.getRange(startRow + i, exCol2);
      try{
        cell.setDataValidation(null);
        cell.setNumberFormat("@");
        cell.setValue(r.ex);
        retried = true;
      }catch(e){}
    });
    if(retried){
      SpreadsheetApp.flush();
      written = sheet.getRange(startRow, exCol2, expanded.length, 1).getValues();
    }
    var seen = {};
    expanded.forEach(function(r,i){
      if(loose(written[i][0]) !== loose(r.ex) && !seen[r.ex]){
        seen[r.ex] = 1;   // report each exercise once, not once per week
        failures.push(r.ex + ": didn't save — the cell ended up as \"" +
                      String(written[i][0]) + "\"");
      }
    });
    if(failures.length){
      bumpVersion(sheet.getName());
      return { error:"partial", problems: failures,
               added: expanded.length - failures.length, firstRow: startRow };
    }
  } finally {
    lock.releaseLock();
  }
  bumpVersion(sheet.getName());
  return { ok:true, added: expanded.length, firstRow: startRow,
           weeks: schedule.length, perWeek: rows.length };
}

/* ---------------- maintenance ----------------
   Run these by hand from the Apps Script editor (pick the function from the
   dropdown and press Run). They are not exposed to the web app.          */

/**
 * Wipe one week for one client, so it can be rebuilt from the coach app.
 *
 * Clears the row contents rather than deleting the rows, because the Log
 * references row numbers — deleting would shift every row below and attach
 * old log entries to the wrong exercise. Matching Log rows are removed too.
 *
 * Edit the two values at the top, then Run. It logs what it did and changes
 * nothing if the week can't be found.
 */
function clearWeekForClient(){
  var CLIENT_NAME = "Billy Quan";   // exact tab name
  var WEEK_NUMBER = 2;              // 1 = earliest start date

  var sheet = ss().getSheetByName(CLIENT_NAME);
  if(!sheet){ Logger.log('No tab named "' + CLIENT_NAME + '"'); return; }

  var m = headerMap(sheet);
  var last = sheet.getLastRow();
  if(last < 2){ Logger.log("Nothing in that tab"); return; }

  var vals = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();

  // distinct start dates, oldest first — the same grouping the apps use
  var starts = [];
  vals.forEach(function(r){
    var st = toIso(r[m[H.START]]);
    if(st && clean(r[m[H.EXERCISE]]) && starts.indexOf(st) === -1) starts.push(st);
  });
  starts.sort();

  if(WEEK_NUMBER < 1 || WEEK_NUMBER > starts.length){
    Logger.log("That client has " + starts.length + " weeks: " + starts.join(", "));
    return;
  }
  var target = starts[WEEK_NUMBER - 1];

  var rowsCleared = [];
  var cols = [H.START,H.END,H.EXERCISE,H.SETS,H.TARGET,H.REPS,H.UNIT,H.SIDE,
              H.FREQ,H.BEN,H.SESSIONS,H.LAST];
  vals.forEach(function(r,i){
    if(toIso(r[m[H.START]]) !== target) return;
    var row = i + 2;
    cols.forEach(function(key){
      if(key in m) sheet.getRange(row, m[key] + 1).clearContent();
    });
    rowsCleared.push(row);
  });

  // drop the client's log entries for those rows
  var lg = logSheet();
  var lastLog = lg.getLastRow();
  var logsRemoved = 0;
  if(lastLog >= 2){
    var lv = lg.getRange(2, 1, lastLog - 1, LOG_COLS.length).getValues();
    for(var i = lv.length - 1; i >= 0; i--){
      if(clean(lv[i][1]) !== CLIENT_NAME) continue;
      if(rowsCleared.indexOf(Number(lv[i][3])) === -1) continue;
      lg.deleteRow(i + 2);
      logsRemoved++;
    }
  }

  bumpVersion(CLIENT_NAME);
  Logger.log("Week " + WEEK_NUMBER + " (" + target + ") cleared for " + CLIENT_NAME +
             "\nSheet rows cleared: " + rowsCleared.join(", ") +
             "\nLog entries removed: " + logsRemoved +
             "\nWeeks now: " + (starts.length - 1));
}

/* ---------------- coach: progress ---------------- */

function getProgress(clientSlug, fresh){
  var sheet = findSheet(clientSlug);
  if(!sheet) return { error:"not_found" };
  var prog = getProgram(clientSlug, fresh);
  var log = readLog(sheet.getName());

  prog.weeks.forEach(function(w){
    w.exercises.forEach(function(ex){
      var done = 0, loads = [], reps = [], ints = [], notes = [], work = [];
      for(var s=1; s<=ex.freqNum; s++){
        var k = ex.row + "|" + s;
        if(!log[k]) continue;
        done++;
        Object.keys(log[k].sets).forEach(function(sn){
          var v = log[k].sets[sn];
          var l = (v[0] !== "" && !isNaN(v[0])) ? Number(v[0]) : null;
          var r = (v[1] !== "" && !isNaN(v[1])) ? Number(v[1]) : null;
          if(l !== null) loads.push(l);
          if(r !== null) reps.push(r);
          if(v[2] !== "") ints.push(v[2]);
          // work done in this one set — averaged later, so set counts don't skew it
          if(r !== null) work.push((l !== null && l > 0) ? l * r : r);
        });
        if(log[k].note) notes.push(log[k].note);
      }
      var avg = function(a){ return a.length ? a.reduce(function(x,y){return x+y;},0)/a.length : null; };
      ex.done = done;
      ex.avgLoad = avg(loads);
      ex.avgReps = avg(reps);
      // the single figure both apps compare on — average work per set
      ex.avgPerSet = avg(work);
      ex.loaded = loads.length > 0;
      ex.avgInt = ints.length ? ints[ints.length-1] : "";
      ex.notes = notes;
      delete ex.sessions;
      delete ex.cues;
    });
  });
  return prog;
}

/* ---------------- routing ---------------- */

/* Every response goes out as JSON, including crashes — otherwise Apps Script
   returns an HTML error page and the app has nothing useful to show. */

function doGet(e){
  try{
    var p = e.parameter || {};
    var action = p.action || "program";

    var fresh = String(p.fresh||"") === "1";

    if(action === "program"){
      return json(getProgram(String(p.client||"").toLowerCase().trim(), fresh));
    }

    // everything below is coach-only
    if(p.pw !== COACH_PASSWORD) return json({ error:"auth" });

    if(action === "clients")  return json(listClients());
    if(action === "library")  return json(getLibrary());
    if(action === "progress") return json(getProgress(String(p.client||"").toLowerCase().trim(), fresh));
    return json({ error:"unknown_action" });
  }catch(err){
    return json({ error:"server", detail: String(err && err.message || err) });
  }
}

function doPost(e){
  try{
    var body;
    try{ body = JSON.parse(e.postData.contents); }
    catch(err){ return json({ error:"bad_request" }); }

    var action = body.action || "session";

    if(action === "session") return json(saveSession(body));

    if(body.pw !== COACH_PASSWORD) return json({ error:"auth" });

    if(action === "createClient") return json(createClient(body.name));
    if(action === "saveWeek")     return json(saveWeek(body));
    if(action === "archive")      return json(setArchived(body.name, !!body.archived));
    return json({ error:"unknown_action" });
  }catch(err){
    return json({ error:"server", detail: String(err && err.message || err) });
  }
}
