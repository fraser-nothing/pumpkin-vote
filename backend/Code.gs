/**
 * Pumpkin Reckoning — vote store (Google Apps Script web app).
 *
 * Setup (about 2 minutes):
 *   1. Go to https://script.google.com → New project, paste this whole file over Code.gs, save.
 *   2. Deploy → New deployment → type "Web app".
 *      Execute as: Me.   Who has access: Anyone.
 *   3. Authorise when asked, then copy the Web app URL (ends in /exec) into config.js on the site.
 *
 * On the first vote it creates a Google Sheet called "Pumpkin Votes" in your Drive.
 * Every ballot is one row; voting again with the same name replaces that person's row.
 * Tallies are only returned to the site once the deadline has passed.
 */

const DEADLINE = new Date('2026-11-05T23:59:59Z'); // Thursday 5 November, 23:59 UK time (GMT)
const PUMPKIN_COUNT = 21;
const POINTS = { first: 3, second: 2, third: 1 };

function doGet(e) {
  return json_(status_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (new Date() > DEADLINE) return json_({ ok: false, error: 'Voting has closed.' });

    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const name = String(body.name || '').trim().replace(/\s+/g, ' ').slice(0, 60);
    const picks = [body.first, body.second, body.third].map(Number);

    if (name.length < 2) return json_({ ok: false, error: 'Please tell us your name.' });
    if (!picks.every(n => Number.isInteger(n) && n >= 1 && n <= PUMPKIN_COUNT)) {
      return json_({ ok: false, error: 'Pick a pumpkin for 1st, 2nd and 3rd.' });
    }
    if (new Set(picks).size !== 3) return json_({ ok: false, error: 'Your three picks must be different pumpkins.' });

    const sheet = sheet_();
    const key = name.toLowerCase();
    const rows = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).toLowerCase() === key) { rowIndex = i + 1; break; }
    }
    const row = [new Date(), name, picks[0], picks[1], picks[2]];
    if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);

    return json_({ ok: true, updated: rowIndex > 0, count: Math.max(sheet.getLastRow() - 1, 0) });
  } catch (err) {
    return json_({ ok: false, error: 'Something went bump in the night. Please try again.' });
  } finally {
    lock.releaseLock();
  }
}

function status_() {
  const sheet = sheet_();
  const rows = sheet.getDataRange().getValues().slice(1);
  const closed = new Date() > DEADLINE;
  const out = { ok: true, count: rows.length, closed: closed, deadline: DEADLINE.toISOString() };
  if (closed) out.results = tally_(rows);
  return out;
}

function tally_(rows) {
  const t = {};
  for (let n = 1; n <= PUMPKIN_COUNT; n++) t[n] = { n: n, points: 0, first: 0, second: 0, third: 0 };
  rows.forEach(r => {
    const f = t[r[2]], s = t[r[3]], h = t[r[4]];
    if (f) { f.points += POINTS.first; f.first++; }
    if (s) { s.points += POINTS.second; s.second++; }
    if (h) { h.points += POINTS.third; h.third++; }
  });
  return Object.values(t).sort((a, b) =>
    b.points - a.points || b.first - a.first || b.second - a.second || a.n - b.n);
}

function sheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  let ss;
  if (id) ss = SpreadsheetApp.openById(id);
  else {
    ss = SpreadsheetApp.create('Pumpkin Votes');
    props.setProperty('SHEET_ID', ss.getId());
    ss.getSheets()[0].appendRow(['Time', 'Name', '1st (3 pts)', '2nd (2 pts)', '3rd (1 pt)']);
  }
  return ss.getSheets()[0];
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
