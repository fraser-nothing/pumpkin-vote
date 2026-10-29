/**
 * Nothing Pumpkin Competition — voting page and vote store (Google Apps Script web app).
 *
 * The project has two files: this one (Code.gs) and Index.html (the voting page).
 * Setup:
 *   1. Pick "setup" in the function menu and press Run; allow the permissions it asks for.
 *   2. Deploy → New deployment → type "Web app".
 *      Execute as: Me.   Who has access: Anyone within Nothing Technology Limited.
 *   3. Share the Web app URL (ends in /exec), or point the GitHub Pages site at it.
 *
 * Voters are identified by their Nothing Google account, so it's one vote each automatically;
 * voting again replaces their earlier vote. Votes go into a Google Sheet called "Pumpkin Votes".
 * Tallies are only sent to the page once the deadline has passed.
 *
 * Results email: a few minutes after the deadline, the winners and full table are emailed to
 * RESULTS_EMAIL. To preview it, pick sendResultsEmail in the function menu and press Run.
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
  'Sully Wazowski', 'Gap-Toothed Grinner', 'TARS from Interstellar', 'The Gnarled Goblin', 'The Classic Grinner'
];

function doGet() {
  ensureResultsTrigger_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Nothing Pumpkin Competition')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// Called from the page: who is voting, how many votes so far, their current vote, and results once closed.
function getState() {
  const email = voter_();
  const rows = sheet_().getDataRange().getValues().slice(1);
  const closed = new Date() > DEADLINE;
  const mineRow = rows.find(r => String(r[1]).toLowerCase() === email);
  const out = {
    email: email,
    count: rows.length,
    closed: closed,
    mine: mineRow ? { first: Number(mineRow[2]), second: Number(mineRow[3]), third: Number(mineRow[4]) } : null
  };
  if (closed) out.results = tally_(rows);
  return out;
}

// Called from the page when someone hits Vote.
function submitVote(v) {
  const email = voter_();
  if (!email) throw new Error('Please sign in with your Nothing Google account to vote.');
  if (new Date() > DEADLINE) throw new Error('Voting has closed.');
  const picks = [v && v.first, v && v.second, v && v.third].map(Number);
  if (!picks.every(n => Number.isInteger(n) && n >= 1 && n <= PUMPKIN_COUNT)) {
    throw new Error('Pick a pumpkin for 1st, 2nd and 3rd.');
  }
  if (new Set(picks).size !== 3) throw new Error('Your three picks must be different pumpkins.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = sheet_();
    const rows = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).toLowerCase() === email) { rowIndex = i + 1; break; }
    }
    const row = [new Date(), email, picks[0], picks[1], picks[2]];
    if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);
    return { updated: rowIndex > 0, count: Math.max(sheet.getLastRow() - 1, 0) };
  } finally {
    lock.releaseLock();
  }
}

function voter_() {
  return String(Session.getActiveUser().getEmail() || '').toLowerCase();
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

// Run once from the editor: asks for permissions, creates the sheet and sets the results-email timer.
function setup() {
  sheet_();
  ensureResultsTrigger_();
}

function sheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  let ss;
  if (id) ss = SpreadsheetApp.openById(id);
  else {
    ss = SpreadsheetApp.create('Pumpkin Votes');
    props.setProperty('SHEET_ID', ss.getId());
    ss.getSheets()[0].appendRow(['Time', 'Voter', '1st (3 pts)', '2nd (2 pts)', '3rd (1 pt)']);
  }
  return ss.getSheets()[0];
}
