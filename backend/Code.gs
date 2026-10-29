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
 *
 * Results email: a few minutes after the deadline, the winners and full table are emailed to
 * RESULTS_EMAIL. The timer is set up automatically on the first visit to the site. To check it
 * works, pick sendResultsEmail in the editor's function menu and press Run (it sends the current tally).
 */

const DEADLINE = new Date('2026-11-05T23:59:59Z'); // Thursday 5 November, 23:59 UK time (GMT)
const RESULTS_EMAIL = 'fraser.stone@nothing.tech';
const PUMPKIN_COUNT = 21;
const POINTS = { first: 3, second: 2, third: 1 };
const NAMES = [
  'The Squinting Sentinel', 'Baby Frankenstein', 'Knife-Mouth Mortimer', 'The Cat-o’-Lantern',
  'Hollow-Eyed Hector', 'The AI Apparition', 'The KPI Horror', 'The Spewing Creature',
  'Itsy Bitsy', 'The Fanged Fiend', 'Weeping Wanda', 'Tongue-Tied Terror',
  'Spiral-Eyed Specter', 'Classic Jack', 'The Lark Facepalm', 'Cyclops Lashes',
  'Sully Wazowski', 'Gap-Toothed Grinner', 'TARS from Interstellar', 'The Gnarled Goblin', 'The Bauhaus Grin'
];

function doGet(e) {
  ensureResultsTrigger_();
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

// Runs once, a few minutes after the deadline (or by hand from the editor to preview).
function sendResultsEmail() {
  const sheet = sheet_();
  const rows = sheet.getDataRange().getValues().slice(1);
  const res = tally_(rows);
  const label = r => 'No. ' + r.n + ' · ' + NAMES[r.n - 1];
  const final = new Date() > DEADLINE;

  const podium = ['1st', '2nd', '3rd'].map((place, i) =>
    '<p style="font-size:18px;margin:6px 0"><b>' + place + ':</b> ' + label(res[i]) +
    ' <span style="color:#888">(' + res[i].points + ' pts, ' + res[i].first + ' first-place)</span></p>').join('');
  const table = res.map((r, i) =>
    '<tr><td>' + (i + 1) + '</td><td>' + label(r) + '</td><td align="right">' + r.points + '</td><td align="right">' +
    r.first + '</td><td align="right">' + r.second + '</td><td align="right">' + r.third + '</td></tr>').join('');
  const html =
    '<h2 style="color:#ff6a00">Nothing Pumpkin Competition: ' + (final ? 'final results' : 'results so far') + '</h2>' +
    '<p>' + rows.length + ' votes counted.</p>' + podium +
    '<table cellpadding="6" style="border-collapse:collapse;margin-top:16px">' +
    '<tr style="color:#888"><td>#</td><td>Pumpkin</td><td>Pts</td><td>1st</td><td>2nd</td><td>3rd</td></tr>' + table + '</table>' +
    '<p style="margin-top:16px"><a href="' + sheet.getParent().getUrl() + '">Every ballot is in the Pumpkin Votes sheet</a></p>';

  MailApp.sendEmail({
    to: RESULTS_EMAIL,
    subject: '🎃 Pumpkin Competition ' + (final ? 'winners' : 'results so far') + ': ' + NAMES[res[0].n - 1] + ' leads with ' + res[0].points + ' pts',
    htmlBody: html
  });
}

function ensureResultsTrigger_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('RESULTS_TRIGGER') || new Date() > DEADLINE) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    if (props.getProperty('RESULTS_TRIGGER')) return;
    ScriptApp.newTrigger('sendResultsEmail').timeBased().at(new Date(DEADLINE.getTime() + 5 * 60 * 1000)).create();
    props.setProperty('RESULTS_TRIGGER', '1');
  } finally {
    lock.releaseLock();
  }
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
