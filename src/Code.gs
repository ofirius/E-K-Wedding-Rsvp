/**
 * Smart RSVP – Google Apps Script (v4: one row per person, grouped into households)
 * Single deployment: Execute as: Me, Access: Anyone.
 *
 * Sheet "Guests" columns (ONE ROW PER PERSON):
 * ID | Household | Name | Status | Remarks | Responded At | Updated By
 *
 * - Household: any label shared by everyone in the same family/party, e.g. "Santos Family".
 *   Rows with the same Household are shown together, and anyone in it can update everyone's answer.
 *   Leave it blank for a person invited on their own.
 * - Name: the person. Leave Name BLANK (but fill Household) to reserve an open seat that the
 *   household can name themselves.
 * - Seats reserved for a household = number of its rows (named people + open seats).
 * - Status: Pending / Attending / Declined (per person, maintained by the app).
 * - Remarks: one household message, stored on the household's FIRST row.
 * - Responded At / Updated By: who last changed that person's answer, filled by the app.
 * - ID: filled automatically (leave blank).
 */

const SHEET_NAME = 'Guests';
const OLD_SHEET_NAME = 'Guests (old layout)';
const HEADERS = ['ID', 'Household', 'Name', 'Status', 'Remarks', 'Responded At', 'Updated By'];
const STATUS = { PENDING: 'Pending', ATTENDING: 'Attending', DECLINED: 'Declined' };
const MAX_SEARCH_RESULTS = 5;
const MIN_QUERY_LENGTH = 3;
const MAX_SEARCHES_PER_MINUTE = 60;

/* ---------- Routing ---------- */

function doGet(e) {
  const page = ((e && e.parameter && e.parameter.page) || '').toLowerCase();
  if (page === 'dashboard') {
    const tpl = HtmlService.createTemplateFromFile('Dashboard');
    tpl.key = (e.parameter.key || '');
    return tpl.evaluate()
      .setTitle('RSVP Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'ics') return icsOutput_();

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Wedding RSVP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- One-time setup (run manually from the editor) ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet && isOldLayout_(sheet)) {
    throw new Error('Your "Guests" sheet still uses the old one-row-per-household layout. Run migrateToPersonRows() instead of setup().');
  }
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  initSheet_(sheet);
  ensureDashboardKey_();
}

/** Optional: fill in IDs + default status for rows you added by hand. (The app also does this lazily.) */
function generateIds() {
  const sheet = getSheet_();
  readRows_().forEach(r => {
    if (!r.ID) sheet.getRange(r.row, 1).setValue(newId_());
    if (r.Name && !r.Status) sheet.getRange(r.row, col_('Status')).setValue(STATUS.PENDING);
  });
}

/**
 * One-time conversion from the old layout
 *   ID | Name | Seats Reserved | Party Names | Status | Seats Confirmed | Attending | Not Attending | Remarks | Responded At
 * to one row per person. The old sheet is kept, renamed "Guests (old layout)".
 */
function migrateToPersonRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(SHEET_NAME);
  if (!old || !isOldLayout_(old)) throw new Error('No old-layout "Guests" sheet found, nothing to migrate.');

  const OLD = ['ID', 'Name', 'Seats Reserved', 'Party Names', 'Status', 'Seats Confirmed',
               'Attending', 'Not Attending', 'Remarks', 'Responded At'];
  const last = old.getLastRow();
  const values = last >= 2 ? old.getRange(2, 1, last - 1, OLD.length).getValues() : [];
  const out = [];
  const usedLabels = new Set();

  values.forEach(v => {
    const o = Object.fromEntries(OLD.map((h, i) => [h, v[i]]));
    const primary = String(o['Name'] || '').trim();
    if (!primary) return;

    let label = primary, n = 2;
    while (usedLabels.has(norm_(label))) label = primary + ' (' + (n++) + ')';
    usedLabels.add(norm_(label));

    const reserved = Math.max(1, Number(o['Seats Reserved']) || 1);
    const party = splitNames_(o['Party Names']);
    if (!party.some(x => norm_(x) === norm_(primary))) party.unshift(primary);
    const names = party.slice(0, reserved);
    const att = splitNames_(o['Attending']).map(norm_);
    const notAtt = splitNames_(o['Not Attending']).map(norm_);

    names.forEach((nm, i) => {
      const k = norm_(nm);
      const status = att.includes(k) ? STATUS.ATTENDING : notAtt.includes(k) ? STATUS.DECLINED : STATUS.PENDING;
      const answered = status !== STATUS.PENDING;
      out.push([newId_(), label, nm, status, i === 0 ? String(o['Remarks'] || '') : '',
                answered ? o['Responded At'] : '', answered ? primary : '']);
    });
    for (let i = names.length; i < reserved; i++) out.push([newId_(), label, '', '', '', '', '']); // open seats
  });

  let backupName = OLD_SHEET_NAME;
  if (ss.getSheetByName(backupName)) backupName += ' ' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd-HHmmss');
  old.setName(backupName);

  const sheet = ss.insertSheet(SHEET_NAME);
  initSheet_(sheet);
  if (out.length) sheet.getRange(2, 1, out.length, HEADERS.length).setValues(out);
  ensureDashboardKey_();
  Logger.log('Migrated ' + out.length + ' rows. Old sheet kept as "' + backupName + '".');
}

function initSheet_(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(STATUS), true).build();
  sheet.getRange(2, col_('Status'), 1000, 1).setDataValidation(rule);
}

function ensureDashboardKey_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('DASHBOARD_KEY')) {
    props.setProperty('DASHBOARD_KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 12));
  }
  Logger.log('Dashboard key: ' + props.getProperty('DASHBOARD_KEY'));
}

function isOldLayout_(sheet) {
  if (sheet.getLastColumn() < 1) return false;
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String).includes('Party Names');
}

/* ---------- Calendar file (.ics) ---------- */
// Keep in sync with CONFIG in Index.html. 2:00 PM Philippine time (UTC+8) = 06:00Z; the 4-hour end time is an assumption.
const EVENT = {
  title: "Emman & Kath's Wedding",
  startUtc: '20261128T060000Z',
  endUtc: '20261128T120000Z',
  ceremony: 'Our Lady of Lourdes Parish, Tagaytay, Cavite',   // 'Church name, address'  (used as the calendar location)
  reception: 'Joselito\'s Resort and Events Place, Amadeo, Cavite'    // 'Venue name, address'   (added to the calendar description)
};

function icsOutput_() {
  const esc = s => String(s || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\n/g, '\\n');
  const lines = [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Emman and Kath//RSVP//EN', 'CALSCALE:GREGORIAN',
    'BEGIN:VEVENT',
    'UID:emman-kath-wedding-2026@rsvp',
    'DTSTAMP:' + Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'"),
    'DTSTART:' + EVENT.startUtc,
    'DTEND:' + EVENT.endUtc,
    'SUMMARY:' + esc(EVENT.title)
  ];
  if (EVENT.ceremony) lines.push('LOCATION:' + esc(EVENT.ceremony));
  const desc = EVENT.title + (EVENT.reception ? '\nReception: ' + EVENT.reception : '');
  lines.push('DESCRIPTION:' + esc(desc));
  lines.push('END:VEVENT', 'END:VCALENDAR');
  return ContentService.createTextOutput(lines.join('\r\n')).setMimeType(ContentService.MimeType.ICAL);
}

/* ---------- Guest API ---------- */

/** Name search over individual people. Returns at most 5 {id, name, hint}; nothing else about the guest is exposed. */
function searchGuests(query) {
  const q = norm_(query);
  if (q.replace(/ /g, '').length < MIN_QUERY_LENGTH) return [];
  rateLimit_();

  const tokens = q.split(' ');
  const hits = readRows_()
    .filter(r => r.Name && tokens.every(t => norm_(r.Name).includes(t)))
    .slice(0, MAX_SEARCH_RESULTS);
  if (!hits.length) return [];
  ensureIds_(hits);

  // if two people share a name, show their household label so they can tell them apart
  const counts = {};
  hits.forEach(r => { const k = norm_(r.Name); counts[k] = (counts[k] || 0) + 1; });
  return hits.map(r => ({ id: r.ID, name: r.Name, hint: counts[norm_(r.Name)] > 1 ? r.Household : '' }));
}

/** Everyone in the searcher's household with their current answers. The searcher is listed first. */
function getHousehold(id) {
  const rows = readRows_();
  const me = rows.find(r => r.ID && r.ID === String(id || '').trim());
  if (!me || !me.Name) return null;

  const mates = rows.filter(r => groupKey_(r) === groupKey_(me)); // sheet order
  ensureIds_(mates);
  const ordered = [me].concat(mates.filter(r => r !== me));

  let latest = null;
  mates.forEach(r => {
    if (!r['Responded At']) return;
    const d = new Date(r['Responded At']);
    if (!isNaN(d) && (!latest || d > latest.d)) latest = { d: d, by: String(r['Updated By'] || '') };
  });

  return {
    name: me.Name,
    household: me.Household,
    seatsReserved: mates.length,                      // named people + open seats, includes the searcher
    remarks: String(mates[0].Remarks || ''),
    lastUpdate: latest
      ? { by: latest.by, at: Utilities.formatDate(latest.d, Session.getScriptTimeZone(), 'MMM d, h:mm a') }
      : null,
    members: ordered.map(r => ({
      id: r.ID,
      name: r.Name,                                   // '' = open seat
      isMe: r === me,
      attending: statusOf_(r) === STATUS.ATTENDING ? true : statusOf_(r) === STATUS.DECLINED ? false : null
    }))
  };
}

/**
 * payload = { members: [{id, name, attending: true|false}], remarks }
 * Anyone in the household can save answers for everyone in it. Only rows whose answer actually changed
 * get a new Responded At / Updated By, so "last updated by" stays truthful.
 * Open seats (blank Name in the sheet) are named from payload.members[].name.
 */
function submitRsvp(id, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rows = readRows_();
    const me = rows.find(r => r.ID && r.ID === String(id || '').trim());
    if (!me || !me.Name) throw new Error('Guest not found.');
    const mates = rows.filter(r => groupKey_(r) === groupKey_(me));
    const byId = new Map(mates.map(r => [r.ID, r]));

    const entries = [];
    (payload.members || []).forEach(m => {
      const r = byId.get(String((m && m.id) || '').trim());
      if (!r) return;                                                   // ignore anything outside this household
      const name = r.Name || cleanText_(m.name, 80).replace(/;/g, ' '); // open seats can be named by the guest
      if (!name) return;
      if (m.attending !== true && m.attending !== false) {
        throw new Error('Please choose Attending or Not Attending for every guest.');
      }
      entries.push({ r: r, name: name, status: m.attending ? STATUS.ATTENDING : STATUS.DECLINED });
    });
    if (!entries.length) throw new Error('Please enter at least one name.');
    if (mates.some(r => r.Name && !entries.some(e => e.r === r))) {
      throw new Error('Please answer for everyone in the invitation.');
    }

    const sheet = getSheet_();
    const remarks = cleanText_(payload.remarks, 500);
    const first = mates[0];
    const now = new Date();
    let firstTouched = false;

    entries.forEach(e => {
      const changed = statusOf_(e.r) !== e.status || e.r.Name !== e.name;
      if (!changed) return;
      if (e.r === first) firstTouched = true;
      // Name, Status, Remarks, Responded At, Updated By (contiguous)
      sheet.getRange(e.r.row, col_('Name'), 1, 5).setValues([[
        e.name, e.status, e.r === first ? remarks : e.r.Remarks, now, me.Name
      ]]);
    });
    if (!firstTouched && remarks !== String(first.Remarks || '')) {
      sheet.getRange(first.row, col_('Remarks')).setValue(remarks);
    }
    return { ok: true, attending: entries.filter(e => e.status === STATUS.ATTENDING).length };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Dashboard API ---------- */

function getDashboardData(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('DASHBOARD_KEY');
  if (!key || key !== expected) throw new Error('Unauthorized.');

  const groups = new Map();
  readRows_().forEach(r => {
    const k = groupKey_(r);
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(r);
  });

  const totals = { households: 0, invitees: 0, attending: 0, declined: 0, pending: 0, seatsReserved: 0, openSeats: 0 };
  const households = [];
  const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, h:mm a');

  groups.forEach(list => {
    const named = list.filter(r => r.Name);
    if (!named.length) return;                          // a household with only open seats and nobody named yet
    const yes = named.filter(r => statusOf_(r) === STATUS.ATTENDING).map(r => r.Name);
    const no = named.filter(r => statusOf_(r) === STATUS.DECLINED).map(r => r.Name);
    const wait = named.filter(r => statusOf_(r) === STATUS.PENDING).map(r => r.Name);
    const open = list.length - named.length;

    totals.households++;
    totals.invitees += named.length;
    totals.attending += yes.length;
    totals.declined += no.length;
    totals.pending += wait.length;
    totals.seatsReserved += list.length;
    totals.openSeats += open;

    let latest = null;
    list.forEach(r => {
      if (!r['Responded At']) return;
      const d = new Date(r['Responded At']);
      if (!isNaN(d) && (!latest || d > latest.d)) latest = { d: d, by: String(r['Updated By'] || '') };
    });

    households.push({
      label: list[0].Household || named[0].Name,
      status: !wait.length ? (yes.length ? 'Attending' : 'Declined') : (yes.length || no.length ? 'Partial' : 'Pending'),
      seats: list.length,
      openSeats: open,
      attending: yes,
      notAttending: no,
      pending: wait,
      remarks: String(list[0].Remarks || ''),
      updatedBy: latest ? latest.by : '',
      respondedAt: latest ? fmt(latest.d) : ''
    });
  });

  return { totals: totals, households: households };
}

/* ---------- Helpers ---------- */

function col_(header) { return HEADERS.indexOf(header) + 1; }

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Run setup() first.');
  if (isOldLayout_(sheet)) throw new Error('The Guests sheet uses the old layout. Run migrateToPersonRows() once.');
  return sheet;
}

/** Every non-empty row as an object, plus .row (sheet row number). Rows with neither Name nor Household are skipped. */
function readRows_() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const rows = [];
  values.forEach((v, i) => {
    const o = Object.fromEntries(HEADERS.map((h, j) => [h, v[j]]));
    o.row = i + 2;
    o.ID = String(o.ID || '').trim();
    o.Household = String(o.Household || '').trim();
    o.Name = String(o.Name || '').trim();
    if (!o.Name && !o.Household) return;
    rows.push(o);
  });
  return rows;
}

/** Rows sharing a Household label form one household; a row with no Household is its own household. */
function groupKey_(r) { return r.Household ? 'h:' + norm_(r.Household) : 'r:' + r.row; }

function statusOf_(r) {
  const s = String(r.Status || '').trim();
  return s === STATUS.ATTENDING || s === STATUS.DECLINED ? s : STATUS.PENDING;
}

/** Give any row without an ID one. Re-checks the cell under a lock so two guests can't assign different IDs. */
function ensureIds_(rows) {
  const missing = rows.filter(r => !r.ID);
  if (!missing.length) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    missing.forEach(r => {
      const cell = sheet.getRange(r.row, 1);
      const cur = String(cell.getValue() || '').trim();
      if (cur) { r.ID = cur; return; }
      r.ID = newId_();
      cell.setValue(r.ID);
    });
  } finally {
    lock.releaseLock();
  }
}

function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function splitNames_(s) { return String(s || '').split(';').map(x => x.trim()).filter(Boolean); }

/** trim, collapse whitespace, cap length, and drop leading =+-@ so a cell can never become a formula */
function cleanText_(s, max) {
  return String(s || '').replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ')
    .replace(/^[=+\-@\s]+/, '').trim().slice(0, max);
}

/** lowercase, strip accents/punctuation, collapse spaces ("Peña" matches "pena") */
function norm_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}

function rateLimit_() {
  const cache = CacheService.getScriptCache();
  const key = 'searches_' + Math.floor(Date.now() / 60000);
  const n = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(n), 90);
  if (n > MAX_SEARCHES_PER_MINUTE) throw new Error('Lots of people are searching right now. Please try again in a minute.');
}
