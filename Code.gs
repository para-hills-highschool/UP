/**
 * UP Idea Lab — Google Apps Script backend
 * Saves anonymous student ideas to a Google Sheet.
 * Page: https://para-hills-highschool.github.io/UP/
 *
 * Setup:
 * 1. Create a Google Sheet > Extensions > Apps Script, paste this file in, Save.
 *    (Or use a standalone script and paste the Sheet's ID into SHEET_ID.)
 * 2. Deploy > New deployment > Web app.  Execute as: Me.  Who has access: Anyone.
 * 3. Copy the web app URL (ends in /exec) into SCRIPT_URL in index.html.
 *
 * The page calls this with fetch (GET ?action=list / ?action=add). JSONP (?callback=) also works.
 */
const HEADERS = ["Timestamp", "Project title", "Who it's for", "What's needed",
                 "Approximate timeline", "What you need to learn"];
const KEYS = ["timestamp", "title", "audience", "needed", "timeline", "learn"];
const MAX_LEN = 1000;
const SHEET_ID = "1BvxcvWXNb6XxXTISpQ9z9MR6VClFn7uHDRVaHt1hwRE";

function sheet_() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

function clean_(v) {
  let s = String(v == null ? "" : v).slice(0, MAX_LEN).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s; // block formula injection
  return s;
}

function add_(data) {
  if (!data.title) return { ok: false, error: "Missing project title" };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    sheet_().appendRow(KEYS.map(k => k === "timestamp" ? new Date() : clean_(data[k])));
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

function list_() {
  return sheet_().getDataRange().getValues().slice(1).map(r => {
    const o = {};
    KEYS.forEach((k, i) => o[k] = r[i] instanceof Date ? r[i].toISOString() : String(r[i]).replace(/^'/, ""));
    return o;
  });
}

function out_(obj, callback) {
  const body = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + body + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const result = p.action === "add" ? add_(p) : { ok: true, ideas: list_() };
  return out_(result, p.callback);
}

function doPost(e) {
  try { return out_(add_(JSON.parse(e.postData.contents))); }
  catch (err) { return out_({ ok: false, error: String(err) }); }
}

// Run once from the editor to authorise and add the header row.
function setup() {
  sheet_();
}
