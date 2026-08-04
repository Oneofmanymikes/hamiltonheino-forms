/**********************************************************************
 * Heino Doessing — Canvass shift signup + week tracker
 *
 * Adapted from the votehafiz.ca system (signup-webapp + phonebank-webapp),
 * trimmed to canvassing only.
 *
 * Two surfaces, both served by the same web app as the four forms:
 *   ?form=shift            public — pick a canvass shift and sign up
 *   ?view=week&key=<KEY>   staff  — 7-day grid of who is signed up
 *
 * Tabs (created lazily, like the form tabs):
 *   Shifts        one row per shift slot, with capacity + running signed-up count
 *   ShiftSignups  one row per person per shift
 **********************************************************************/

/*** CONFIG ***/
const SHIFTS_TAB        = 'Shifts';
const SHIFT_SIGNUPS_TAB = 'ShiftSignups';
const SHIFT_ACTIVITY    = 'Canvass';
const SHIFT_CAPACITY    = 20;          // per shift; set 0 for unlimited
const SHIFT_LOCATION    = '';          // optional: shown to volunteers, e.g. 'Campaign office, 123 Main St E'

// The schedule. Weekdays run two blocks; weekends add a 10 AM.
// Times are 24h 'HH:MM'. Changing these changes future generated shifts only —
// already-generated rows keep their times (regenerating never edits existing rows).
const SHIFT_BLOCKS_WEEKDAY = [
  { start: '13:00', end: '15:00', label: 'Afternoon' },
  { start: '17:00', end: '19:00', label: 'Evening'   }
];
const SHIFT_BLOCKS_WEEKEND = [
  { start: '10:00', end: '12:00', label: 'Morning'   },
  { start: '13:00', end: '15:00', label: 'Afternoon' },
  { start: '17:00', end: '19:00', label: 'Evening'   }
];

/*** COLUMN MAPS ***/
const SH_ID = 1, SH_DATE = 2, SH_DAY = 3, SH_START = 4, SH_END = 5,
      SH_ACTIVITY = 6, SH_DESC = 7, SH_CAPACITY = 8, SH_SIGNED_UP = 9;
const SH_COLS = 9;

const SU_ID = 1, SU_SHIFT_ID = 2, SU_DATE = 3, SU_TIME = 4, SU_ACTIVITY = 5,
      SU_FIRST = 6, SU_LAST = 7, SU_EMAIL = 8, SU_PHONE = 9, SU_POSTAL = 10,
      SU_CONSENT = 11, SU_SIGNED_AT = 12, SU_SOURCE = 13, SU_STATUS = 14, SU_NOTES = 15;
const SU_COLS = 15;

const SHIFT_HEADERS = ['Shift ID', 'Date', 'Day', 'Start', 'End', 'Activity',
                       'Description', 'Capacity', 'Signed Up'];
const SIGNUP_HEADERS = ['Signup ID', 'Shift ID', 'Date', 'Time', 'Activity',
                        'First name', 'Last name', 'Email', 'Phone', 'Postal code',
                        'Consent', 'Signed up at', 'Source', 'Status', 'Notes'];

/*** SCHEDULE ***/

/** Blocks for a day of week. 0=Sun … 6=Sat. Canvass runs every day. */
function shiftsForDayOfWeek_(dow) {
  const weekend = (dow === 0 || dow === 6);
  return (weekend ? SHIFT_BLOCKS_WEEKEND : SHIFT_BLOCKS_WEEKDAY).map(function (b) {
    return {
      start: b.start,
      end: b.end,
      activity: SHIFT_ACTIVITY,
      desc: b.label + ' canvass',
      capacity: SHIFT_CAPACITY
    };
  });
}

/**
 * Create shift rows for every day in [startDate, endDate].
 * Idempotent: a shift whose ID already exists is skipped, so re-running after
 * extending the range only adds the new days. Never edits or deletes existing rows.
 *
 * Run from the Apps Script editor:  generateShifts('2026-08-05', '2026-10-04')
 * Both arguments are 'YYYY-MM-DD'. Omit endDate to generate 60 days from start.
 */
function generateShifts(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ensureShiftsSheet_(ss);
  const tz = Session.getScriptTimeZone();

  // parseYmdLocal_ avoids the UTC-midnight off-by-one that new Date('YYYY-MM-DD')
  // causes in a timezone behind UTC (it would land the range one day early).
  const start = startDateStr ? parseYmdLocal_(startDateStr) : new Date();
  start.setHours(0, 0, 0, 0);
  let end;
  if (endDateStr) {
    end = parseYmdLocal_(endDateStr);
  } else {
    end = new Date(start);
    end.setDate(end.getDate() + 60);
  }
  end.setHours(0, 0, 0, 0);

  const existing = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, SH_ID, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      if (r[0]) existing[r[0]] = true;
    });
  }

  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const newRows = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    const dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    const dayLabel = dayNames[d.getDay()];
    shiftsForDayOfWeek_(d.getDay()).forEach(function (s) {
      const id = dateStr + '_' + s.start.replace(':', '') + '_' + s.activity.toLowerCase();
      if (existing[id]) return;
      newRows.push([id, new Date(d), dayLabel, s.start, s.end,
                    s.activity, s.desc, s.capacity, 0]);
    });
  }

  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, SH_COLS).setValues(newRows);
    sh.getRange(2, 1, sh.getLastRow() - 1, SH_COLS)
      .sort([{ column: SH_DATE, ascending: true }, { column: SH_START, ascending: true }]);
  }
  return { added: newRows.length,
           from: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
           to: Utilities.formatDate(end, tz, 'yyyy-MM-dd') };
}

/*** PUBLIC API (called by Shift.html via google.script.run) ***/

/**
 * Upcoming shifts, grouped by date, with remaining capacity.
 * Only returns shifts that are today or later and not already full.
 */
function listOpenShifts(daysAhead) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHIFTS_TAB);
  if (!sh || sh.getLastRow() < 2) return [];

  const tz = Session.getScriptTimeZone();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const horizon = new Date(today);
  horizon.setDate(horizon.getDate() + (daysAhead || 21));

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, SH_COLS).getValues();
  const byDate = {};
  const order = [];

  rows.forEach(function (r) {
    if (!r[SH_ID - 1]) return;
    const dv = r[SH_DATE - 1];
    const d = (dv instanceof Date) ? dv : parseYmdLocal_(String(dv));
    if (!d || isNaN(d.getTime())) return;
    const dayStart = new Date(d); dayStart.setHours(0, 0, 0, 0);
    if (dayStart < today || dayStart > horizon) return;

    const capacity = Number(r[SH_CAPACITY - 1]) || 0;
    const signed = Number(r[SH_SIGNED_UP - 1]) || 0;
    const remaining = capacity > 0 ? (capacity - signed) : 999;
    if (remaining <= 0) return;                       // full — hide it

    const dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    if (!byDate[dateStr]) {
      byDate[dateStr] = { date: dateStr,
                          label: Utilities.formatDate(d, tz, 'EEEE, MMMM d'),
                          shifts: [] };
      order.push(dateStr);
    }
    byDate[dateStr].shifts.push({
      shiftId: r[SH_ID - 1],
      start: String(r[SH_START - 1] || ''),
      end: String(r[SH_END - 1] || ''),
      label: prettyRange_(r[SH_START - 1], r[SH_END - 1]),
      desc: String(r[SH_DESC - 1] || ''),
      remaining: remaining,
      tight: (capacity > 0 && remaining <= 3)         // show "only N left"
    });
  });

  return order.sort().map(function (k) { return byDate[k]; });
}

/**
 * Record a shift signup. Called by the public form.
 * Returns {ok:true, shift:'…'} or {ok:false, error:'…'} — the form shows `error` verbatim.
 */
function submitShiftSignup(data) {
  data = data || {};
  if (data._hp) return { ok: true };                  // honeypot filled = bot; silently drop

  const first = String(data.first || '').trim();
  const last  = String(data.last  || '').trim();
  const email = String(data.email || '').trim();
  const phone = String(data.phone || '').trim();
  const shiftIds = (data.shiftIds || []).filter(String);

  if (!first || !last) return { ok: false, error: 'Please enter your first and last name.' };
  if (!email)          return { ok: false, error: 'Please enter your email address.' };
  if (!phone)          return { ok: false, error: 'Please enter a phone number.' };
  if (!shiftIds.length) return { ok: false, error: 'Please choose at least one shift.' };
  if (!data.consent)   return { ok: false, error: 'Please agree to be contacted.' };

  // Capacity is checked and incremented under a lock so two people submitting at
  // the same moment cannot both take the last spot.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return { ok: false, error: 'The system is busy. Please try again in a moment.' };
  }

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ensureShiftsSheet_(ss);
    const su = ensureSignupsSheet_(ss);
    const tz = Session.getScriptTimeZone();

    const shiftRows = sh.getLastRow() >= 2
      ? sh.getRange(2, 1, sh.getLastRow() - 1, SH_COLS).getValues() : [];

    // Existing signups for this email, so a refresh-and-resubmit doesn't double-book.
    const already = {};
    if (su.getLastRow() >= 2) {
      su.getRange(2, 1, su.getLastRow() - 1, SU_COLS).getValues().forEach(function (r) {
        if (String(r[SU_EMAIL - 1]).toLowerCase() === email.toLowerCase()) {
          already[r[SU_SHIFT_ID - 1]] = true;
        }
      });
    }

    const booked = [], skipped = [], newRows = [], countUpdates = [];
    const now = new Date();

    shiftIds.forEach(function (id) {
      let idx = -1;
      for (let i = 0; i < shiftRows.length; i++) {
        if (shiftRows[i][SH_ID - 1] === id) { idx = i; break; }
      }
      if (idx === -1) { skipped.push('that shift is no longer listed'); return; }
      if (already[id]) { skipped.push('you were already signed up for one of those'); return; }

      const r = shiftRows[idx];
      const capacity = Number(r[SH_CAPACITY - 1]) || 0;
      const signed = Number(r[SH_SIGNED_UP - 1]) || 0;
      if (capacity > 0 && signed >= capacity) {
        skipped.push('one of those shifts filled up');
        return;
      }

      const dv = r[SH_DATE - 1];
      const d = (dv instanceof Date) ? dv : parseYmdLocal_(String(dv));
      const human = Utilities.formatDate(d, tz, 'EEEE, MMMM d') + ' · ' +
                    prettyRange_(r[SH_START - 1], r[SH_END - 1]);

      newRows.push([
        Utilities.getUuid().slice(0, 8), id, new Date(d), String(r[SH_START - 1] || ''),
        r[SH_ACTIVITY - 1] || SHIFT_ACTIVITY, first, last, email, phone,
        String(data.postal || '').trim(), 'Yes', now,
        String(data.source || WEBSITE_URL), 'Scheduled', ''
      ]);
      countUpdates.push({ row: idx + 2, value: signed + 1 });
      shiftRows[idx][SH_SIGNED_UP - 1] = signed + 1;   // keep local copy in step
      booked.push(human);
    });

    if (!booked.length) {
      return { ok: false,
               error: skipped.length
                 ? 'Sorry — ' + skipped[0] + '. Please pick another shift.'
                 : 'Sorry, we could not record that. Please try again.' };
    }

    su.getRange(su.getLastRow() + 1, 1, newRows.length, SU_COLS).setValues(newRows);
    countUpdates.forEach(function (u) {
      sh.getRange(u.row, SH_SIGNED_UP).setValue(u.value);
    });

    SpreadsheetApp.flush();
    notifyShiftSignup_(first, last, email, phone, booked);
    return { ok: true, booked: booked };

  } finally {
    lock.releaseLock();
  }
}

/*** STAFF WEEK TRACKER ***/

/**
 * 7 days from startDateStr, each with its shifts and who signed up.
 * Staff-only: this returns volunteer phone numbers and emails, and the web app
 * is public, so the guard must run before anything is read.
 */
function getWeekGrid(startDateStr, token) {
  requireStaff_(token);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHIFTS_TAB);
  const su = ss.getSheetByName(SHIFT_SIGNUPS_TAB);
  const tz = Session.getScriptTimeZone();

  const start = startDateStr ? parseYmdLocal_(startDateStr) : startOfWeek_(new Date());
  start.setHours(0, 0, 0, 0);

  const days = [], byDate = {};
  for (let i = 0; i < 7; i++) {
    const d = new Date(start); d.setDate(d.getDate() + i);
    const day = { date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
                  label: Utilities.formatDate(d, tz, 'EEE MMM d'),
                  slots: [], total: 0 };
    days.push(day);
    byDate[day.date] = day;
  }

  // Shift slots first, so empty shifts still show as columns in the grid.
  const slotByKey = {};
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, SH_COLS).getValues().forEach(function (r) {
      if (!r[SH_ID - 1]) return;
      const dv = r[SH_DATE - 1];
      const d = (dv instanceof Date) ? dv : parseYmdLocal_(String(dv));
      const ds = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      if (!byDate[ds]) return;
      const slot = { shiftId: r[SH_ID - 1],
                     start: String(r[SH_START - 1] || ''),
                     label: prettyRange_(r[SH_START - 1], r[SH_END - 1]),
                     capacity: Number(r[SH_CAPACITY - 1]) || 0,
                     people: [] };
      byDate[ds].slots.push(slot);
      slotByKey[r[SH_ID - 1]] = slot;
    });
  }

  if (su && su.getLastRow() >= 2) {
    su.getRange(2, 1, su.getLastRow() - 1, SU_COLS).getValues().forEach(function (r) {
      const slot = slotByKey[r[SU_SHIFT_ID - 1]];
      if (!slot) return;
      if (String(r[SU_STATUS - 1] || '') === 'Cancelled') return;
      slot.people.push({ first: r[SU_FIRST - 1] || '', last: r[SU_LAST - 1] || '',
                         email: r[SU_EMAIL - 1] || '', phone: r[SU_PHONE - 1] || '' });
    });
  }

  days.forEach(function (day) {
    day.slots.sort(function (a, b) { return String(a.start).localeCompare(String(b.start)); });
    day.slots.forEach(function (s) {
      s.people.sort(function (a, b) { return String(a.first).localeCompare(String(b.first)); });
      day.total += s.people.length;
    });
  });

  return { start: days[0].date, days: days };
}

/*** HELPERS ***/

function ensureShiftsSheet_(ss) {
  let sh = ss.getSheetByName(SHIFTS_TAB);
  if (!sh) {
    sh = ss.insertSheet(SHIFTS_TAB);
    sh.getRange(1, 1, 1, SH_COLS).setValues([SHIFT_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureSignupsSheet_(ss) {
  let su = ss.getSheetByName(SHIFT_SIGNUPS_TAB);
  if (!su) {
    su = ss.insertSheet(SHIFT_SIGNUPS_TAB);
    su.getRange(1, 1, 1, SU_COLS).setValues([SIGNUP_HEADERS]).setFontWeight('bold');
    su.setFrozenRows(1);
  }
  return su;
}

/** 'YYYY-MM-DD' → local midnight (new Date(str) would parse as UTC and shift the day). */
function parseYmdLocal_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) { const d = new Date(s); return isNaN(d.getTime()) ? null : d; }
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** Monday of the week containing d. */
function startOfWeek_(d) {
  const out = new Date(d);
  const dow = out.getDay();                 // 0=Sun
  out.setDate(out.getDate() - ((dow + 6) % 7));
  out.setHours(0, 0, 0, 0);
  return out;
}

/** '13:00','15:00' → '1:00 – 3:00 PM' */
function prettyRange_(start, end) {
  return prettyTime_(start) + ' – ' + prettyTime_(end);
}

function prettyTime_(hhmm) {
  const m = String(hhmm || '').match(/^(\d{1,2}):(\d{2})/);
  if (!m) return String(hhmm || '');
  let h = Number(m[1]);
  const ampm = h >= 12 ? 'PM' : 'AM';
  if (h === 0) h = 12; else if (h > 12) h -= 12;
  return h + ':' + m[2] + ' ' + ampm;
}

function notifyShiftSignup_(first, last, email, phone, booked) {
  const body =
    'New canvass shift signup for ' + CAMPAIGN + ':\n\n' +
    'Name: ' + first + ' ' + last + '\n' +
    'Email: ' + email + '\n' +
    'Phone: ' + phone + '\n\n' +
    'Shifts:\n' + booked.map(function (b) { return '  • ' + b; }).join('\n') + '\n\n' +
    'Sheet: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit\n';

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: '[Canvass shift] ' + first + ' ' + last,
    body: body,
    replyTo: email || NOTIFY_EMAIL,
    name: CAMPAIGN + ' Forms'
  });

  // Confirmation to the volunteer. Best-effort: a bounce here must not fail the signup.
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'You are signed up to canvass with ' + CANDIDATE,
      body: 'Hi ' + first + ',\n\n' +
            'Thanks for signing up to canvass with the ' + CANDIDATE + ' campaign. You are booked for:\n\n' +
            booked.map(function (b) { return '  • ' + b; }).join('\n') + '\n\n' +
            (SHIFT_LOCATION ? 'Where: ' + SHIFT_LOCATION + '\n\n' : '') +
            'Someone from the campaign will be in touch with the details before your shift. ' +
            'If you need to change or cancel, just reply to this email.\n\n' +
            'Thank you,\n' + CAMPAIGN + '\n' + WEBSITE_URL + '\n',
      replyTo: NOTIFY_EMAIL,
      name: CAMPAIGN
    });
  } catch (err) {
    // ignore — the signup is already recorded and the campaign has been notified
  }
}
