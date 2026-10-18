/**
 * Smart RSVP – Google Apps Script (v4: one row per person, grouped into households)
 * Single deployment: Execute as: Me, Access: Anyone.
 *
 * Sheet "Guests" columns (ONE ROW PER PERSON):
 * ID | Household | Name | Status | Remarks | Responded At | Updated By | Groups
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
 * - Groups (optional): one or more labels separated by ";", e.g. "HS Friends; Work". A group is a VIEW-ONLY
 *   circle: anyone who lists a label can see the name and status of everyone else who lists it, but can only
 *   change answers inside their own Household. Add the "Groups" header in column H if your sheet doesn't have it.
 */

const GUEST_SHEET = 'Guests';
const LEGACY_SHEET_NAME = 'Guests (old layout)';
const GUEST_COLUMNS = ['ID', 'Household', 'Name', 'Status', 'Remarks', 'Responded At', 'Updated By', 'Groups'];
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
  if (page === 'ics') return buildIcsResponse_();

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Wedding RSVP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- One-time setup (run manually from the editor) ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GUEST_SHEET);
  if (sheet && hasLegacyLayout_(sheet)) {
    throw new Error('Your "Guests" sheet still uses the old one-row-per-household layout. Run migrateToPersonRows() instead of setup().');
  }
  if (!sheet) sheet = ss.insertSheet(GUEST_SHEET);
  formatGuestSheet_(sheet);
  ensureDashboardKey_();
}

/** Optional: fill in IDs + default status for rows you added by hand. (The app also does this lazily.) */
function fillMissingIds() {
  const sheet = guestSheet_();
  readGuestRows_().forEach(r => {
    if (!r.ID) sheet.getRange(r.row, 1).setValue(randomId_());
    if (r.Name && !r.Status) sheet.getRange(r.row, columnOf_('Status')).setValue(STATUS.PENDING);
  });
}

/**
 * One-time conversion from the old layout
 *   ID | Name | Seats Reserved | Party Names | Status | Seats Confirmed | Attending | Not Attending | Remarks | Responded At
 * to one row per person. The old sheet is kept, renamed "Guests (old layout)".
 */
function migrateToPersonRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(GUEST_SHEET);
  if (!old || !hasLegacyLayout_(old)) throw new Error('No old-layout "Guests" sheet found, nothing to migrate.');

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
    while (usedLabels.has(normalize_(label))) label = primary + ' (' + (n++) + ')';
    usedLabels.add(normalize_(label));

    const reserved = Math.max(1, Number(o['Seats Reserved']) || 1);
    const party = splitNames_(o['Party Names']);
    if (!party.some(x => normalize_(x) === normalize_(primary))) party.unshift(primary);
    const names = party.slice(0, reserved);
    const att = splitNames_(o['Attending']).map(normalize_);
    const notAtt = splitNames_(o['Not Attending']).map(normalize_);

    names.forEach((nm, i) => {
      const k = normalize_(nm);
      const status = att.includes(k) ? STATUS.ATTENDING : notAtt.includes(k) ? STATUS.DECLINED : STATUS.PENDING;
      const answered = status !== STATUS.PENDING;
      out.push([randomId_(), label, nm, status, i === 0 ? String(o['Remarks'] || '') : '',
                answered ? o['Responded At'] : '', answered ? primary : '', '']);
    });
    for (let i = names.length; i < reserved; i++) out.push([randomId_(), label, '', '', '', '', '', '']); // open seats
  });

  let backupName = LEGACY_SHEET_NAME;
  if (ss.getSheetByName(backupName)) backupName += ' ' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd-HHmmss');
  old.setName(backupName);

  const sheet = ss.insertSheet(GUEST_SHEET);
  formatGuestSheet_(sheet);
  if (out.length) sheet.getRange(2, 1, out.length, GUEST_COLUMNS.length).setValues(out);
  ensureDashboardKey_();
  Logger.log('Migrated ' + out.length + ' rows. Old sheet kept as "' + backupName + '".');
}

function formatGuestSheet_(sheet) {
  sheet.getRange(1, 1, 1, GUEST_COLUMNS.length).setValues([GUEST_COLUMNS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(STATUS), true).build();
  sheet.getRange(2, columnOf_('Status'), 1000, 1).setDataValidation(rule);
}

function ensureDashboardKey_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('DASHBOARD_KEY')) {
    props.setProperty('DASHBOARD_KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 12));
  }
  Logger.log('Dashboard key: ' + props.getProperty('DASHBOARD_KEY'));
}

function hasLegacyLayout_(sheet) {
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

function buildIcsResponse_() {
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
  const q = normalize_(query);
  if (q.replace(/ /g, '').length < MIN_QUERY_LENGTH) return [];
  enforceSearchLimit_();

  const tokens = q.split(' ');
  const hits = readGuestRows_()
    .filter(r => r.Name && tokens.every(t => normalize_(r.Name).includes(t)))
    .slice(0, MAX_SEARCH_RESULTS);
  if (!hits.length) return [];
  assignMissingIds_(hits);

  // if two people share a name, show their household label so they can tell them apart
  const counts = {};
  hits.forEach(r => { const k = normalize_(r.Name); counts[k] = (counts[k] || 0) + 1; });
  return hits.map(r => ({ id: r.ID, name: r.Name, hint: counts[normalize_(r.Name)] > 1 ? r.Household : '' }));
}

/** Everyone in the searcher's household with their current answers. The searcher is listed first. */
function getHousehold(id) {
  const rows = readGuestRows_();
  const searcher = rows.find(r => r.ID && r.ID === String(id || '').trim());
  if (!searcher || !searcher.Name) return null;

  const householdRows = rows.filter(r => householdKey_(r) === householdKey_(searcher)); // sheet order
  assignMissingIds_(householdRows);
  const searcherFirst = [searcher].concat(householdRows.filter(r => r !== searcher));

  let newest = null;
  householdRows.forEach(r => {
    if (!r['Responded At']) return;
    const d = new Date(r['Responded At']);
    if (!isNaN(d) && (!newest || d > newest.d)) newest = { d: d, by: String(r['Updated By'] || '') };
  });

  return {
    name: searcher.Name,
    household: searcher.Household,
    seatsReserved: householdRows.length,                      // named people + open seats, includes the searcher
    groups: searcher.groupLabels.map(label => ({ label: label })), // view-only circles this person belongs to
    remarks: String(householdRows[0].Remarks || ''),
    lastUpdate: newest
      ? { by: newest.by, at: Utilities.formatDate(newest.d, Session.getScriptTimeZone(), 'MMM d, h:mm a') }
      : null,
    members: searcherFirst.map(r => ({
      id: r.ID,
      name: r.Name,                                   // '' = open seat
      isSelf: r === searcher,
      attending: statusOf_(r) === STATUS.ATTENDING ? true : statusOf_(r) === STATUS.DECLINED ? false : null
    }))
  };
}

/**
 * View-only list of everyone in one of the searcher's groups, with their current status.
 * A group can only be opened by someone who is in it, so labels can't be guessed to peek at other circles.
 */
function getGroupMembers(id, label) {
  const rows = readGuestRows_();
  const searcher = rows.find(r => r.ID && r.ID === String(id || '').trim());
  if (!searcher || !searcher.Name) throw new Error('Guest not found.');

  const wanted = normalize_(label);
  const groupLabel = searcher.groupLabels.find(l => normalize_(l) === wanted);
  if (!groupLabel) throw new Error('You are not part of that group.');

  const members = rows
    .filter(r => r.Name && r.groupLabels.some(l => normalize_(l) === wanted))
    .map(r => ({ name: r.Name, isSelf: r === searcher, status: statusOf_(r) }));
  const count = status => members.filter(m => m.status === status).length;

  return {
    label: groupLabel,
    members: members,
    counts: { attending: count(STATUS.ATTENDING), declined: count(STATUS.DECLINED), pending: count(STATUS.PENDING) }
  };
}

/**
 * payload = { members: [{id, name, attending: true|false|null}], remarks }
 * attending null = no answer yet: that guest stays Pending. The searcher's own answer is required (new or existing).
 * Anyone in the household can save answers for everyone in it. Only rows whose answer actually changed
 * get a new Responded At / Updated By, so "last updated by" stays truthful.
 * Open seats (blank Name in the sheet) are named from payload.members[].name.
 */
function submitRsvp(id, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rows = readGuestRows_();
    const searcher = rows.find(r => r.ID && r.ID === String(id || '').trim());
    if (!searcher || !searcher.Name) throw new Error('Guest not found.');
    const householdRows = rows.filter(r => householdKey_(r) === householdKey_(searcher));
    const rowById = new Map(householdRows.map(r => [r.ID, r]));

    const answers = [];
    (payload.members || []).forEach(m => {
      const r = rowById.get(String((m && m.id) || '').trim());
      if (!r) return;                                                   // ignore anything outside this household
      const name = r.Name || sanitizeText_(m.name, 80).replace(/;/g, ' '); // open seats can be named by the guest
      if (!name) return;
      const answered = m.attending === true || m.attending === false;
      if (!answered && m.attending != null) throw new Error('Invalid answer.');
      // a guest left unanswered keeps the status they already have (Pending for a new one)
      answers.push({
        r: r, name: name, answered: answered,
        status: answered ? (m.attending ? STATUS.ATTENDING : STATUS.DECLINED) : statusOf_(r)
      });
    });
    if (!answers.length) throw new Error('Please enter at least one name.');
    // the person filling this in must have an answer: given now, or already recorded earlier
    const selfAnswered = answers.some(e => e.r === searcher && e.answered) || statusOf_(searcher) !== STATUS.PENDING;
    if (!selfAnswered) throw new Error('Please answer for yourself first.');

    const sheet = guestSheet_();
    const remarks = sanitizeText_(payload.remarks, 500);
    const firstRow = householdRows[0];
    const now = new Date();
    let firstRowChanged = false;

    answers.forEach(e => {
      const statusChanged = e.answered && statusOf_(e.r) !== e.status;
      const nameChanged = e.r.Name !== e.name;
      if (!statusChanged && !nameChanged) return;
      if (e.r === firstRow) firstRowChanged = true;
      // Name, Status, Remarks, Responded At, Updated By (contiguous).
      // Naming a seat without answering for it doesn't count as a response, so those two columns are kept.
      sheet.getRange(e.r.row, columnOf_('Name'), 1, 5).setValues([[
        e.name, e.status, e.r === firstRow ? remarks : e.r.Remarks,
        statusChanged ? now : e.r['Responded At'], statusChanged ? searcher.Name : e.r['Updated By']
      ]]);
    });
    if (!firstRowChanged && remarks !== String(firstRow.Remarks || '')) {
      sheet.getRange(firstRow.row, columnOf_('Remarks')).setValue(remarks);
    }
    return {
      ok: true,
      attending: answers.filter(e => e.answered && e.status === STATUS.ATTENDING).length,
      waiting: answers.filter(e => !e.answered).length
    };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Dashboard API ---------- */

function getDashboardData(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('DASHBOARD_KEY');
  if (!key || key !== expected) throw new Error('Unauthorized.');

  const rowsByHousehold = new Map();
  readGuestRows_().forEach(r => {
    const k = householdKey_(r);
    if (!rowsByHousehold.has(k)) rowsByHousehold.set(k, []);
    rowsByHousehold.get(k).push(r);
  });

  const totals = { households: 0, invitees: 0, attending: 0, declined: 0, pending: 0, seatsReserved: 0, openSeats: 0 };
  const households = [];
  const formatTime = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, h:mm a');

  rowsByHousehold.forEach(householdRows => {
    const named = householdRows.filter(r => r.Name);
    if (!named.length) return;                          // a household with only open seats and nobody named yet
    const attendingNames = named.filter(r => statusOf_(r) === STATUS.ATTENDING).map(r => r.Name);
    const declinedNames = named.filter(r => statusOf_(r) === STATUS.DECLINED).map(r => r.Name);
    const pendingNames = named.filter(r => statusOf_(r) === STATUS.PENDING).map(r => r.Name);
    const openSeatCount = householdRows.length - named.length;

    totals.households++;
    totals.invitees += named.length;
    totals.attending += attendingNames.length;
    totals.declined += declinedNames.length;
    totals.pending += pendingNames.length;
    totals.seatsReserved += householdRows.length;
    totals.openSeats += openSeatCount;

    let newest = null;
    householdRows.forEach(r => {
      if (!r['Responded At']) return;
      const d = new Date(r['Responded At']);
      if (!isNaN(d) && (!newest || d > newest.d)) newest = { d: d, by: String(r['Updated By'] || '') };
    });

    households.push({
      label: householdRows[0].Household || named[0].Name,
      status: !pendingNames.length ? (attendingNames.length ? 'Attending' : 'Declined') : (attendingNames.length || declinedNames.length ? 'Partial' : 'Pending'),
      seats: householdRows.length,
      openSeats: openSeatCount,
      attending: attendingNames,
      notAttending: declinedNames,
      pending: pendingNames,
      remarks: String(householdRows[0].Remarks || ''),
      updatedBy: newest ? newest.by : '',
      respondedAt: newest ? formatTime(newest.d) : ''
    });
  });

  return { totals: totals, households: households };
}

/* ---------- Helpers ---------- */

function columnOf_(header) { return GUEST_COLUMNS.indexOf(header) + 1; }

function guestSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GUEST_SHEET);
  if (!sheet) throw new Error('Run setup() first.');
  if (hasLegacyLayout_(sheet)) throw new Error('The Guests sheet uses the old layout. Run migrateToPersonRows() once.');
  return sheet;
}

/** Every non-empty row as an object, plus .row (sheet row number). Rows with neither Name nor Household are skipped. */
function readGuestRows_() {
  const sheet = guestSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, GUEST_COLUMNS.length).getValues();
  const rows = [];
  values.forEach((v, i) => {
    const o = Object.fromEntries(GUEST_COLUMNS.map((h, j) => [h, v[j]]));
    o.row = i + 2;
    o.ID = String(o.ID || '').trim();
    o.Household = String(o.Household || '').trim();
    o.Name = String(o.Name || '').trim();
    o.groupLabels = splitGroupLabels_(o.Groups);   // [] when the person is in no group
    if (!o.Name && !o.Household) return;
    rows.push(o);
  });
  return rows;
}

/** Rows sharing a Household label form one household; a row with no Household is its own household. */
function householdKey_(r) { return r.Household ? 'h:' + normalize_(r.Household) : 'r:' + r.row; }

function statusOf_(r) {
  const s = String(r.Status || '').trim();
  return s === STATUS.ATTENDING || s === STATUS.DECLINED ? s : STATUS.PENDING;
}

/** Give any row without an ID one. Re-checks the cell under a lock so two guests can't assign different IDs. */
function assignMissingIds_(rows) {
  const missing = rows.filter(r => !r.ID);
  if (!missing.length) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = guestSheet_();
    missing.forEach(r => {
      const cell = sheet.getRange(r.row, 1);
      const cur = String(cell.getValue() || '').trim();
      if (cur) { r.ID = cur; return; }
      r.ID = randomId_();
      cell.setValue(r.ID);
    });
  } finally {
    lock.releaseLock();
  }
}

function randomId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function splitNames_(s) { return String(s || '').split(';').map(x => x.trim()).filter(Boolean); }

function splitGroupLabels_(s) { return String(s || '').split(/[;,]/).map(x => x.trim()).filter(Boolean); }

/** trim, collapse whitespace, cap length, and drop leading =+-@ so a cell can never become a formula */
function sanitizeText_(s, max) {
  return String(s || '').replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ')
    .replace(/^[=+\-@\s]+/, '').trim().slice(0, max);
}

/** lowercase, strip accents/punctuation, collapse spaces ("Peña" matches "pena") */
function normalize_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}

function enforceSearchLimit_() {
  const cache = CacheService.getScriptCache();
  const key = 'searches_' + Math.floor(Date.now() / 60000);
  const n = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(n), 90);
  if (n > MAX_SEARCHES_PER_MINUTE) throw new Error('Lots of people are searching right now. Please try again in a minute.');
}
