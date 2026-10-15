/**
 * Smart RSVP – Google Apps Script skeleton (v3)
 * Single deployment: Execute as: Me, Access: Anyone.
 *
 * Sheet "Guests" columns (one row per invitation / household):
 * ID | Name | Seats Reserved | Party Names | Status | Seats Confirmed | Attending | Not Attending | Remarks | Responded At
 *
 * - Party Names: optional, you pre-fill the names for the reserved seats, separated by ";"
 *   e.g. "Maria Santos; Jose Santos; Baby Ana". Guests can fix typos and fill empty seats.
 * - Attending / Not Attending: filled by the app, names separated by ";"
 */

const SHEET_NAME = 'Guests';
const HEADERS = ['ID', 'Name', 'Seats Reserved', 'Party Names', 'Status', 'Seats Confirmed',
                 'Attending', 'Not Attending', 'Remarks', 'Responded At'];
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
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(STATUS), true).build();
  sheet.getRange(2, col_('Status'), 1000, 1).setDataValidation(rule);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('DASHBOARD_KEY')) {
    props.setProperty('DASHBOARD_KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 12));
  }
  Logger.log('Dashboard key: ' + props.getProperty('DASHBOARD_KEY'));
}

/** Optional: fill in IDs + default status for rows you added. (Search also assigns IDs lazily.) */
function generateIds() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return;
  const range = sheet.getRange(2, 1, last - 1, HEADERS.length);
  const rows = range.getValues();
  const iID = HEADERS.indexOf('ID'), iName = HEADERS.indexOf('Name'), iStatus = HEADERS.indexOf('Status');
  rows.forEach(r => {
    if (r[iName] && !r[iID]) r[iID] = newId_();
    if (r[iName] && !r[iStatus]) r[iStatus] = STATUS.PENDING;
  });
  range.setValues(rows);
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

/** Name search. Returns at most 5 {id, name}; nothing else about the guest is exposed. */
function searchGuests(query) {
  const q = norm_(query);
  if (q.replace(/ /g, '').length < MIN_QUERY_LENGTH) return [];
  rateLimit_();

  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const iID = HEADERS.indexOf('ID'), iName = HEADERS.indexOf('Name'), iParty = HEADERS.indexOf('Party Names');
  const tokens = q.split(' ');
  const out = [];

  for (let i = 0; i < values.length && out.length < MAX_SEARCH_RESULTS; i++) {
    const name = values[i][iName];
    if (!name) continue;
    // match on the invitee name OR any pre-filled party member, so "Jose" finds the Santos household
    const haystack = norm_(name + ' ' + String(values[i][iParty] || '').replace(/;/g, ' '));
    if (tokens.every(t => haystack.includes(t))) {
      let id = values[i][iID];
      if (!id) {
        id = newId_();
        sheet.getRange(i + 2, 1).setValue(id);
      }
      out.push({ id: String(id), name: String(name) });
    }
  }
  return out;
}

function getGuest(id) {
  const found = findRow_(id);
  if (!found) return null;
  const g = found.obj;
  const reserved = Math.max(1, Number(g['Seats Reserved']) || 1); // includes the primary invitee
  const primary = String(g['Name']).trim();

  // Seat 1 is always the primary invitee, whether or not Party Names lists them
  const party = splitNames_(g['Party Names']);
  if (!party.some(n => norm_(n) === norm_(primary))) party.unshift(primary);
  const att = splitNames_(g['Attending']);
  const notAtt = splitNames_(g['Not Attending']);

  const members = party.slice(0, reserved).map(n => ({
    name: n,
    attending: att.includes(n) ? true : notAtt.includes(n) ? false : null // null = not answered yet
  }));
  while (members.length < reserved) members.push({ name: '', attending: null });

  return {
    name: primary,
    seatsReserved: reserved,
    status: g['Status'] || STATUS.PENDING,
    remarks: g['Remarks'] || '',
    members: members
  };
}

/**
 * payload = { members: [{name, attending: true|false}], remarks }
 * Every named seat needs an explicit answer; blank-name seats are ignored.
 * Status is derived: at least one attending -> Attending, otherwise Declined.
 */
function submitRsvp(id, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const found = findRow_(id);
    if (!found) throw new Error('Guest not found.');
    const reserved = Math.max(1, Number(found.obj['Seats Reserved']) || 1);

    const raw = (payload.members || []).slice(0, reserved)
      .map(m => ({ name: cleanText_(m.name, 80).replace(/;/g, ' '), attending: m.attending }))
      .filter(m => m.name);
    if (!raw.length) throw new Error('Please enter at least one name.');
    if (raw.some(m => m.attending !== true && m.attending !== false)) {
      throw new Error('Please choose Attending or Not Attending for every guest.');
    }

    const attending = raw.filter(m => m.attending).map(m => m.name);
    const notAttending = raw.filter(m => !m.attending).map(m => m.name);
    const status = attending.length ? STATUS.ATTENDING : STATUS.DECLINED;
    const remarks = cleanText_(payload.remarks, 500);

    // Party Names, Status, Seats Confirmed, Attending, Not Attending, Remarks, Responded At (contiguous columns)
    getSheet_().getRange(found.row, col_('Party Names'), 1, 7).setValues([[
      raw.map(m => m.name).join('; '), status, attending.length,
      attending.join('; '), notAttending.join('; '), remarks, new Date()
    ]]);
    return { ok: true, status: status, attending: attending.length };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Dashboard API ---------- */

function getDashboardData(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('DASHBOARD_KEY');
  if (!key || key !== expected) throw new Error('Unauthorized.');

  const guests = readAll_().filter(g => g['Name']);
  const sum = (arr, f) => arr.reduce((t, g) => t + (Number(g[f]) || 0), 0);

  const attending = guests.filter(g => g['Status'] === STATUS.ATTENDING);
  const declined = guests.filter(g => g['Status'] === STATUS.DECLINED);
  const pending = guests.filter(g => !g['Status'] || g['Status'] === STATUS.PENDING);
  const responded = attending.concat(declined);

  return {
    totals: {
      invitations: guests.length,
      attending: attending.length,
      declined: declined.length,
      pending: pending.length,
      seatsReserved: sum(guests, 'Seats Reserved'),
      seatsConfirmed: sum(attending, 'Seats Confirmed'),
      seatsPending: sum(pending, 'Seats Reserved'),
      seatsReleased: sum(responded, 'Seats Reserved') - sum(responded, 'Seats Confirmed')
    },
    guests: guests.map(g => ({
      name: g['Name'],
      status: g['Status'] || STATUS.PENDING,
      seatsReserved: g['Seats Reserved'],
      seatsConfirmed: g['Seats Confirmed'],
      attending: g['Attending'],
      notAttending: g['Not Attending'],
      remarks: g['Remarks'],
      respondedAt: g['Responded At']
        ? Utilities.formatDate(new Date(g['Responded At']), Session.getScriptTimeZone(), 'MMM d, h:mm a')
        : ''
    }))
  };
}

/* ---------- Helpers ---------- */

function col_(header) { return HEADERS.indexOf(header) + 1; }

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Run setup() first.');
  return sheet;
}

function readAll_() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  return values.map(r => Object.fromEntries(HEADERS.map((h, i) => [h, r[i]])));
}

function findRow_(id) {
  if (!id) return null;
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const idx = values.findIndex(r => String(r[0]).trim() === String(id).trim());
  if (idx === -1) return null;
  return { row: idx + 2, obj: Object.fromEntries(HEADERS.map((h, i) => [h, values[idx][i]])) };
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
