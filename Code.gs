/**
 * UP Idea Lab — Google Apps Script backend (optional)
 * Saves anonymous student ideas to the Google Sheet this script is attached to.
 * NOT run by GitHub Pages — copy this into Apps Script.
 *
 * Setup (5 min):
 * 1. Make a new Google Sheet ("UP Idea Lab 2027") > Extensions > Apps Script.
 * 2. Paste this file in, Save.
 * 3. Deploy > New deployment > Web app.
 *      Execute as: Me    Who has access: Anyone
 * 4. Copy the web app URL into SCRIPT_URL near the top of the <script> in index.html.
 * Export: File > Download > .csv in the Sheet, or open the page with ?teacher for .csv / .md.
 */
const HEADERS = ["Timestamp", "Project title", "Who it's for", "What's needed",
                 "Approximate timeline", "What you need to learn"];
const KEYS = ["timestamp", "title", "audience", "needed", "timeline", "learn"];
const MAX_LEN = 1000;

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean_(v) {
  let s = String(v == null ? "" : v).slice(0, MAX_LEN).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s; // block formula injection
  return s;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = JSON.parse(e.postData.contents);
    if (!data.title) return json_({ ok: false, error: "Missing project title" });
    const row = KEYS.map(k => k === "timestamp" ? new Date() : clean_(data[k]));
    sheet_().appendRow(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  const values = sheet_().getDataRange().getValues().slice(1);
  const ideas = values.map(r => {
    const o = {};
    KEYS.forEach((k, i) => o[k] = r[i] instanceof Date ? r[i].toISOString() : String(r[i]).replace(/^'/, ""));
    return o;
  });
  return json_({ ok: true, ideas });
}
