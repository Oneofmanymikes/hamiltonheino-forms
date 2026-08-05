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
      SU_CONSENT = 11, SU_SIGNED_AT = 12, SU_SOURCE = 13, SU_STATUS = 14, SU_NOTES = 15,
      // --- reminder / confirm / attendance / thank-you loop ---
      SU_REMINDED_AT = 16, SU_CONFIRMED_AT = 17, SU_ATTENDED_AT = 18, SU_THANKED_AT = 19;
const SU_COLS = 19;

/* Signup lifecycle. A volunteer moves Scheduled → Confirmed → Arrived →
 * Completed. No-show and Cancelled are the two ways out. Attendance is what
 * makes the "Active volunteer" segment mean something: without it, someone who
 * booked once and never turned up looks identical to your best canvasser. */
const SS_SCHEDULED = 'Scheduled', SS_CONFIRMED = 'Confirmed',
      SS_ARRIVED   = 'Arrived',   SS_COMPLETED = 'Completed',
      SS_NOSHOW    = 'No-show',   SS_CANCELLED = 'Cancelled';
const SS_OPEN = [SS_SCHEDULED, SS_CONFIRMED, SS_ARRIVED];

const REMINDER_LEAD_HOURS = 24;    // email this far ahead of the shift
const THANKYOU_WINDOW_DAYS = 7;    // completed within this window → thank-you call

const SHIFT_HEADERS = ['Shift ID', 'Date', 'Day', 'Start', 'End', 'Activity',
                       'Description', 'Capacity', 'Signed Up'];
const SIGNUP_HEADERS = ['Signup ID', 'Shift ID', 'Date', 'Time', 'Activity',
                        'First name', 'Last name', 'Email', 'Phone', 'Postal code',
                        'Consent', 'Signed up at', 'Source', 'Status', 'Notes',
                        'Reminded at', 'Confirmed at', 'Attended at', 'Thanked at'];

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

/** Admin-triggered generation, so nobody has to open the script editor. */
function adminGenerateShifts(token, startDateStr, endDateStr) {
  requireAdmin_(token);
  try {
    const res = generateShifts(startDateStr, endDateStr);
    return { ok: true, message: 'Added ' + res.added + ' shift' +
                                (res.added === 1 ? '' : 's') +
                                ' from ' + res.from + ' to ' + res.to + '.' };
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

/*** PUBLIC API (called by Shift.html via google.script.run) ***/

/**
 * Upcoming shifts, grouped by date, with remaining capacity.
 * Only returns shifts that are today or later and not already full.
 */
function listOpenShifts(daysAhead) {
  // Cached briefly: the phone bank asks for this on EVERY contact, and it reads
  // the whole Shifts tab each time. 60s is short enough that a shift filling up
  // is reflected almost immediately, and capacity is enforced again at submit
  // time under a lock, so a stale read can never oversubscribe a shift.
  const cacheKey = 'openShifts_' + (daysAhead || 21);
  const cache = CacheService.getScriptCache();
  const hit = cache.get(cacheKey);
  if (hit) {
    try { return JSON.parse(hit); } catch (err) { /* fall through and rebuild */ }
  }

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
      start: hhmm_(r[SH_START - 1]),
      end: hhmm_(r[SH_END - 1]),
      label: prettyRange_(r[SH_START - 1], r[SH_END - 1]),
      desc: String(r[SH_DESC - 1] || ''),
      remaining: remaining,
      tight: (capacity > 0 && remaining <= 3)         // show "only N left"
    });
  });

  const result = order.sort().map(function (k) { return byDate[k]; });
  try { cache.put(cacheKey, JSON.stringify(result), 60); } catch (err) { /* >100KB — skip cache */ }
  return result;
}

/**
 * "Have we met?" — a returning volunteer types their phone number and skips
 * re-entering everything.
 *
 * PRIVACY: this is a PUBLIC endpoint, so it deliberately returns only a first
 * name — never the surname, email, postal code, or anything else. Otherwise
 * anyone could dial through phone numbers and harvest the volunteer list. The
 * rest of the record is filled in server-side at submit time (see backfill in
 * submitShiftSignup), so the browser never sees it.
 */
function lookupByPhone(phone) {
  const rec = findPersonByPhone_(phone);
  if (!rec) return { found: false };
  return { found: true, first: rec.first };
}

/**
 * Most recent record for a phone number, searched across shift signups, the
 * call list, and the form tabs. Returns null if we have not seen them.
 * Internal only — never expose the result of this to the browser directly.
 */
function findPersonByPhone_(phone) {
  const key = normPhone_(phone);
  if (!key || key.length < 10) return null;

  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. Someone who has signed up for a shift before — the richest record.
  const su = ss.getSheetByName(SHIFT_SIGNUPS_TAB);
  if (su && su.getLastRow() >= 2) {
    const rows = su.getRange(2, 1, su.getLastRow() - 1, SU_COLS).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {          // newest first
      if (normPhone_(rows[i][SU_PHONE - 1]) === key) {
        return { first: rows[i][SU_FIRST - 1], last: rows[i][SU_LAST - 1],
                 email: rows[i][SU_EMAIL - 1], postal: rows[i][SU_POSTAL - 1] };
      }
    }
  }

  // 2. The recruitment call list.
  const cl = ss.getSheetByName(CALL_LIST_TAB);
  if (cl && cl.getLastRow() >= 2) {
    const rows = cl.getRange(2, 1, cl.getLastRow() - 1, CL_COLS).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (normPhone_(rows[i][CL_PHONE - 1]) === key) {
        return { first: rows[i][CL_FIRST - 1], last: rows[i][CL_LAST - 1],
                 email: rows[i][CL_EMAIL - 1], postal: rows[i][CL_POSTAL - 1] };
      }
    }
  }

  // 3. Anyone who filled in one of the four public forms.
  const tabs = ['Volunteers', 'CommitToVote', 'SignRequests', 'Donations'];
  for (let t = 0; t < tabs.length; t++) {
    const sh = ss.getSheetByName(tabs[t]);
    if (!sh || sh.getLastRow() < 2) continue;
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                      .map(function (h) { return String(h).trim().toLowerCase(); });
    const iPhone = headers.indexOf('phone');
    if (iPhone === -1) continue;
    const iFirst = headers.indexOf('first name'), iLast = headers.indexOf('last name'),
          iEmail = headers.indexOf('email'), iPostal = headers.indexOf('postal code');
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (normPhone_(rows[i][iPhone]) === key) {
        return { first:  iFirst  !== -1 ? rows[i][iFirst]  : '',
                 last:   iLast   !== -1 ? rows[i][iLast]   : '',
                 email:  iEmail  !== -1 ? rows[i][iEmail]  : '',
                 postal: iPostal !== -1 ? rows[i][iPostal] : '' };
      }
    }
  }
  return null;
}

/**
 * Record a shift signup. Called by the public form.
 * Returns {ok:true, shift:'…'} or {ok:false, error:'…'} — the form shows `error` verbatim.
 */
function submitShiftSignup(data) {
  data = data || {};
  if (data._hp) return { ok: true };                  // honeypot filled = bot; silently drop

  let first = String(data.first || '').trim();
  let last  = String(data.last  || '').trim();
  let email = String(data.email || '').trim();
  const phone = String(data.phone || '').trim();
  const shiftIds = (data.shiftIds || []).filter(String);

  // Returning volunteer: they gave us a recognised phone number and the form
  // therefore never asked for the rest. Fill it in from what we already hold —
  // server-side, so these details are never sent to the browser.
  if (phone && (!first || !last || !email)) {
    const known = findPersonByPhone_(phone);
    if (known) {
      first = first || String(known.first || '').trim();
      last  = last  || String(known.last  || '').trim();
      email = email || String(known.email || '').trim();
      if (!data.postal) data.postal = known.postal || '';
    }
  }

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
        Utilities.getUuid().slice(0, 8), id, new Date(d), hhmm_(r[SH_START - 1]),
        r[SH_ACTIVITY - 1] || SHIFT_ACTIVITY,
        sanitizeCell_(first), sanitizeCell_(last), sanitizeCell_(email), sanitizeCell_(phone),
        sanitizeCell_(String(data.postal || '').trim()), 'Yes', now,
        sanitizeCell_(String(data.source || WEBSITE_URL)), SS_SCHEDULED, '',
        '', '', '', ''   // reminded / confirmed / attended / thanked
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
                     start: hhmm_(r[SH_START - 1]),
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
      const st = String(r[SU_STATUS - 1] || SS_SCHEDULED);
      slot.people.push({ signupId: r[SU_ID - 1],
                         first: r[SU_FIRST - 1] || '', last: r[SU_LAST - 1] || '',
                         email: r[SU_EMAIL - 1] || '', phone: r[SU_PHONE - 1] || '',
                         status: st,
                         cancelled: st === SS_CANCELLED,
                         confirmed: !!r[SU_CONFIRMED_AT - 1],
                         attended: st === SS_ARRIVED || st === SS_COMPLETED,
                         noShow: st === SS_NOSHOW });
    });
  }

  days.forEach(function (day) {
    day.slots.sort(function (a, b) { return String(a.start).localeCompare(String(b.start)); });
    day.slots.forEach(function (s) {
      s.people.sort(function (a, b) { return String(a.first).localeCompare(String(b.first)); });
      // Cancelled people are still listed (struck through) but must not inflate
      // the headcount a organiser plans around.
      s.active = s.people.filter(function (p) { return !p.cancelled; }).length;
      day.total += s.active;
    });
  });

  return { start: days[0].date, days: days };
}

/*** REMINDERS ***/

/**
 * Email everyone whose shift starts within the next REMINDER_LEAD_HOURS and who
 * has not already been reminded. Idempotent — the "Reminded at" stamp means
 * running twice never double-emails, so a trigger firing late or twice is safe.
 *
 * Install on an hourly trigger with installReminderTrigger().
 */
function sendShiftReminders() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const su = ensureSignupsSheet_(ss);
  if (su.getLastRow() < 2) return { sent: 0 };

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const horizon = new Date(now.getTime() + REMINDER_LEAD_HOURS * 3600 * 1000);
  const n = su.getLastRow() - 1;
  const rows = su.getRange(2, 1, n, SU_COLS).getValues();
  let sent = 0;

  for (let i = 0; i < n; i++) {
    const r = rows[i];
    if (!r[SU_ID - 1]) continue;
    if (SS_OPEN.indexOf(String(r[SU_STATUS - 1] || SS_SCHEDULED)) === -1) continue;
    if (r[SU_REMINDED_AT - 1]) continue;                 // already reminded
    const email = String(r[SU_EMAIL - 1] || '').trim();
    if (!email) continue;

    const when = shiftStart_(r[SU_DATE - 1], r[SU_TIME - 1]);
    if (!when || when < now || when > horizon) continue;  // not in the window

    const whenTxt = Utilities.formatDate(when, tz, 'EEEE, MMMM d') + ' at ' +
                    prettyTime_(r[SU_TIME - 1]);
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'Reminder: your canvass shift ' + Utilities.formatDate(when, tz, 'EEEE') +
                 ' at ' + prettyTime_(r[SU_TIME - 1]),
        body: 'Hi ' + String(r[SU_FIRST - 1] || '') + ',\n\n' +
              'A quick reminder that you are booked to canvass with the ' + CANDIDATE +
              ' campaign:\n\n' +
              '  ' + whenTxt + '\n' +
              (SHIFT_LOCATION ? '  ' + SHIFT_LOCATION + '\n' : '') + '\n' +
              'Wear comfortable shoes and dress for the weather. We will pair you up ' +
              'and show you everything on the day — no experience needed.\n\n' +
              'If you can no longer make it, just reply to this email and we will ' +
              'free up your place.\n\n' +
              'Thank you,\n' + CAMPAIGN + '\n' + WEBSITE_URL + '\n',
        replyTo: NOTIFY_EMAIL,
        name: CAMPAIGN
      });
      su.getRange(i + 2, SU_REMINDED_AT).setValue(new Date());
      sent++;
    } catch (err) {
      // A single bad address must not stop the rest of the run.
      logActivity_('system', 'reminder_failed', 'signup', r[SU_ID - 1], String(err.message || err));
    }
  }

  if (sent) logActivity_('system', 'reminders_sent', 'shift', '', sent + ' reminder(s)');
  return { sent: sent };
}

/** Hourly trigger for the reminders. Safe to re-run; replaces any existing one. */
function installReminderTrigger() {
  uninstallReminderTrigger();
  ScriptApp.newTrigger('sendShiftReminders').timeBased().everyHours(1).create();
  return 'Shift reminders will now run hourly.';
}

function uninstallReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendShiftReminders') ScriptApp.deleteTrigger(t);
  });
  return 'Reminder trigger removed.';
}

/** Daily trigger that lapses stale contacts. */
function installLapseTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'markLapsedContacts') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('markLapsedContacts').timeBased().everyDays(1).atHour(4).create();
  return 'Lapse sweep will run daily at 4 AM.';
}

/** Turn both scheduled jobs on in one go (admin button). */
function adminInstallTriggers(token) {
  requireAdmin_(token);
  const a = installReminderTrigger();
  const b = installLapseTrigger();
  logActivity_(requireStaff_(token).email, 'triggers_installed', 'system', '', a + ' ' + b);
  return { ok: true, message: a + ' ' + b };
}

/*** ATTENDANCE ***/

/**
 * Set a signup's status. This is how attendance gets recorded:
 *   Confirmed  they said yes on a confirmation call
 *   Arrived    they turned up
 *   Completed  they did the shift
 *   No-show    they did not turn up
 *   Cancelled  they pulled out (frees the place)
 *
 * Cancelling decrements the shift's signed-up count so the place is genuinely
 * released; un-cancelling puts it back. That symmetry is the whole reason
 * capacity stays honest over a long campaign.
 */
function setSignupStatus(token, signupId, status) {
  const user = requireStaff_(token);
  const valid = [SS_SCHEDULED, SS_CONFIRMED, SS_ARRIVED, SS_COMPLETED, SS_NOSHOW, SS_CANCELLED];
  if (valid.indexOf(status) === -1) return { ok: false, error: 'Unknown status.' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const su = ensureSignupsSheet_(ss);
  const sh = ensureShiftsSheet_(ss);
  const row = findSignupRow_(su, signupId);
  if (row === -1) return { ok: false, error: 'That signup could not be found.' };

  const rec = su.getRange(row, 1, 1, SU_COLS).getValues()[0];
  const wasCancelled = String(rec[SU_STATUS - 1] || '') === SS_CANCELLED;
  const nowCancelled = (status === SS_CANCELLED);
  const now = new Date();

  su.getRange(row, SU_STATUS).setValue(status);
  if (status === SS_CONFIRMED && !rec[SU_CONFIRMED_AT - 1]) {
    su.getRange(row, SU_CONFIRMED_AT).setValue(now);
  }
  if ((status === SS_ARRIVED || status === SS_COMPLETED) && !rec[SU_ATTENDED_AT - 1]) {
    su.getRange(row, SU_ATTENDED_AT).setValue(now);
  }

  // Keep the shift's capacity count in step with reality.
  if (wasCancelled !== nowCancelled) {
    const shRow = findShiftRow_(sh, rec[SU_SHIFT_ID - 1]);
    if (shRow !== -1) {
      const cur = Number(sh.getRange(shRow, SH_SIGNED_UP).getValue()) || 0;
      sh.getRange(shRow, SH_SIGNED_UP).setValue(Math.max(0, cur + (nowCancelled ? -1 : 1)));
    }
  }

  logActivity_(user.email, 'signup_status', 'signup', signupId,
               String(rec[SU_FIRST - 1] || '') + ' ' + String(rec[SU_LAST - 1] || '') +
               ' → ' + status);
  return { ok: true, status: status };
}

function findSignupRow_(su, signupId) {
  if (!signupId || su.getLastRow() < 2) return -1;
  const vals = su.getRange(2, SU_ID, su.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(signupId)) return i + 2;
  }
  return -1;
}

function findShiftRow_(sh, shiftId) {
  if (!shiftId || sh.getLastRow() < 2) return -1;
  const vals = sh.getRange(2, SH_ID, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(shiftId)) return i + 2;
  }
  return -1;
}

/** Local Date for a shift's start, from its stored date + 'HH:MM'. */
function shiftStart_(dateVal, timeVal) {
  const d = (dateVal instanceof Date) ? new Date(dateVal) : parseYmdLocal_(String(dateVal));
  if (!d || isNaN(d.getTime())) return null;
  const m = String(hhmm_(timeVal)).match(/^(\d{1,2}):(\d{2})/);
  d.setHours(m ? Number(m[1]) : 0, m ? Number(m[2]) : 0, 0, 0);
  return d;
}

/** Top volunteers by completed shifts — for recognition and re-recruiting. */
function getAttendanceLeaderboard(token, limit) {
  requireStaff_(token);
  const su = ensureSignupsSheet_(SpreadsheetApp.openById(SHEET_ID));
  if (su.getLastRow() < 2) return [];

  const tally = {};
  su.getRange(2, 1, su.getLastRow() - 1, SU_COLS).getValues().forEach(function (r) {
    if (!r[SU_ID - 1]) return;
    const st = String(r[SU_STATUS - 1] || '');
    const key = normPhone_(r[SU_PHONE - 1]) || String(r[SU_EMAIL - 1] || '').toLowerCase();
    if (!key) return;
    if (!tally[key]) {
      tally[key] = { name: String(r[SU_FIRST - 1] || '') + ' ' + String(r[SU_LAST - 1] || ''),
                     completed: 0, noShows: 0, booked: 0 };
    }
    tally[key].booked++;
    if (st === SS_COMPLETED || st === SS_ARRIVED) tally[key].completed++;
    if (st === SS_NOSHOW) tally[key].noShows++;
  });

  return Object.keys(tally).map(function (k) { return tally[k]; })
    .filter(function (v) { return v.completed > 0 || v.noShows > 0; })
    .sort(function (a, b) { return b.completed - a.completed || a.noShows - b.noShows; })
    .slice(0, Number(limit) || 15);
}

/*** HELPERS ***/

function ensureShiftsSheet_(ss) {
  let sh = ss.getSheetByName(SHIFTS_TAB);
  if (!sh) {
    sh = ss.insertSheet(SHIFTS_TAB);
    sh.getRange(1, 1, 1, SH_COLS).setValues([SHIFT_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // Start/End must stay TEXT. Left alone, Sheets parses '13:00' into a datetime on
  // the 1899-12-30 epoch, and every downstream read gets a Date instead of 'HH:MM'.
  sh.getRange(2, SH_START, sh.getMaxRows() - 1, 2).setNumberFormat('@');
  return sh;
}

/**
 * Normalize a Start/End cell to 'HH:MM' whatever Sheets handed back —
 * a Date (coerced), a string, or something else.
 */
function hhmm_(v) {
  if (v instanceof Date) {
    return ('0' + v.getHours()).slice(-2) + ':' + ('0' + v.getMinutes()).slice(-2);
  }
  const m = String(v == null ? '' : v).match(/^(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : String(v == null ? '' : v);
}

/**
 * Rewrite Start/End on every existing shift row as plain text.
 * Start comes from the shift ID (…_1300_canvass), which is immune to the
 * coercion; End is looked up from the block definitions by matching start,
 * so no arithmetic assumption about shift length is baked in.
 * Safe to re-run.
 */
function repairShiftTimes() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ensureShiftsSheet_(ss);
  if (sh.getLastRow() < 2) return { fixed: 0 };

  const endByStart = {};
  SHIFT_BLOCKS_WEEKEND.concat(SHIFT_BLOCKS_WEEKDAY).forEach(function (b) {
    endByStart[b.start] = b.end;
  });

  const n = sh.getLastRow() - 1;
  const ids = sh.getRange(2, SH_ID, n, 1).getValues();
  const out = [];
  let fixed = 0;

  for (let i = 0; i < n; i++) {
    const m = String(ids[i][0] || '').match(/_(\d{2})(\d{2})_/);
    if (!m) { out.push(['', '']); continue; }
    const start = m[1] + ':' + m[2];
    out.push([start, endByStart[start] || '']);
    fixed++;
  }

  const rng = sh.getRange(2, SH_START, n, 2);
  rng.setNumberFormat('@');
  rng.setValues(out);
  return { fixed: fixed };
}

function ensureSignupsSheet_(ss) {
  let su = ss.getSheetByName(SHIFT_SIGNUPS_TAB);
  if (!su) {
    su = ss.insertSheet(SHIFT_SIGNUPS_TAB);
    su.getRange(1, 1, 1, SU_COLS).setValues([SIGNUP_HEADERS]).setFontWeight('bold');
    su.setFrozenRows(1);
    return su;
  }
  // Migration: the reminder/confirm/attendance columns were added later.
  if (su.getLastColumn() < SU_COLS) {
    su.getRange(1, 1, 1, SU_COLS).setValues([SIGNUP_HEADERS]).setFontWeight('bold');
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

function prettyTime_(v) {
  const hhmm = hhmm_(v);
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
