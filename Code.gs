/**
 * Recalibration rehab app — backend.
 * Lives inside the Google Sheet (Extensions → Apps Script).
 * Deploy as Web App: Execute as Me, access: Anyone.
 */

// Tabs that are never client programs
const SKIP_TABS = ["Library", "Template", "Pain Scale"];

// Expected headers (row 1 of every client tab). Order doesn't matter —
// columns are found by name.
const H = {
  START: "START", END: "END", EXERCISE: "EXERCISE", CUES: "CUES",
  VIDEO: "VIDEO", SETS: "SETS", TARGET: "TARGET", REPS: "REP RANGE",
  PROGRESS: "PROGRESS", NOTE: "CLIENT NOTE",
  S: [["S1 LOAD","S1 REPS","S1 RIR"],["S2 LOAD","S2 REPS","S2 RIR"],["S3 LOAD","S3 REPS","S3 RIR"]]
};

function slugify(s) {
  return String(s).toLowerCase().trim().replace(/\s+/g, "-");
}

function findSheet(clientSlug) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().find(sh =>
    SKIP_TABS.indexOf(sh.getName()) === -1 && slugify(sh.getName()) === clientSlug
  ) || null;
}

function headerMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[String(h).trim().toUpperCase()] = i; });
  return map;
}

function toIso(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v || "");
}

function splitCues(text) {
  if (!text) return [];
  return String(text)
    .split(/\n|(?:^|\s)[-–•]\s+/)
    .map(s => s.trim().replace(/^[-–•]\s*/, ""))
    .filter(s => s.length > 1);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const client = (e.parameter.client || "").toLowerCase().trim();
  const sheet = findSheet(client);
  if (!sheet) return json({ error: "not_found" });

  const m = headerMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return json({ name: sheet.getName(), weeks: [] });

  const nCols = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, nCols).getValues();
  const rich = sheet.getRange(2, 1, lastRow - 1, nCols).getRichTextValues();

  const weeksByStart = {};
  values.forEach((row, i) => {
    const name = String(row[m[H.EXERCISE]] || "").trim();
    if (!name) return;
    const start = toIso(row[m[H.START]]);
    if (!start) return; // rows without a start date are skipped
    const end = toIso(row[m[H.END]]);

    // video: prefer the hyperlink behind the cell text, else plain URL text
    let video = "";
    const cell = rich[i][m[H.VIDEO]];
    if (cell) {
      video = cell.getLinkUrl() || "";
      if (!video) {
        const runs = cell.getRuns();
        for (let r = 0; r < runs.length && !video; r++) video = runs[r].getLinkUrl() || "";
      }
    }
    if (!video) {
      const raw = String(row[m[H.VIDEO]] || "").trim();
      if (/^https?:\/\//i.test(raw)) video = raw;
    }

    const logged = H.S.map(cols => {
      if (!(cols[0] in m)) return null;
      return [
        String(row[m[cols[0]]] ?? ""),
        String(row[m[cols[1]]] ?? ""),
        String(row[m[cols[2]]] ?? "")
      ];
    }).filter(x => x !== null);

    const ex = {
      row: i + 2, // 1-based sheet row
      name: name,
      cues: splitCues(row[m[H.CUES]]),
      video: video,
      sets: Math.min(parseInt(row[m[H.SETS]], 10) || 1, logged.length || 3),
      target: String(row[m[H.TARGET]] || "").trim(),
      reps: String(row[m[H.REPS]] || "").trim(),
      logged: logged,
      note: String(row[m[H.NOTE]] || "")
    };

    if (!weeksByStart[start]) weeksByStart[start] = { start: start, end: end, exercises: [] };
    if (end && !weeksByStart[start].end) weeksByStart[start].end = end;
    weeksByStart[start].exercises.push(ex);
  });

  const weeks = Object.keys(weeksByStart).sort().map(k => weeksByStart[k]);
  return json({ name: sheet.getName(), weeks: weeks });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json({ error: "bad_request" }); }

  const sheet = findSheet((body.client || "").toLowerCase().trim());
  if (!sheet) return json({ error: "not_found" });

  const row = parseInt(body.row, 10);
  if (!row || row < 2 || row > sheet.getLastRow()) return json({ error: "bad_row" });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const m = headerMap(sheet);
    const sets = body.sets || [];
    // store numeric strings as numbers so the sheet stays clean
    const num = v => (v !== "" && v !== null && !isNaN(v)) ? Number(v) : (v || "");

    sets.forEach((s, i) => {
      const cols = H.S[i];
      if (!cols || !(cols[0] in m)) return;
      sheet.getRange(row, m[cols[0]] + 1).setValue(num(s[0]));
      sheet.getRange(row, m[cols[1]] + 1).setValue(num(s[1]));
      sheet.getRange(row, m[cols[2]] + 1).setValue(num(s[2]));
    });

    if (H.NOTE in m) sheet.getRange(row, m[H.NOTE] + 1).setValue(body.note || "");
    if (H.PROGRESS in m && body.progress) sheet.getRange(row, m[H.PROGRESS] + 1).setValue(body.progress);
  } finally {
    lock.releaseLock();
  }

  return json({ ok: true });
}
