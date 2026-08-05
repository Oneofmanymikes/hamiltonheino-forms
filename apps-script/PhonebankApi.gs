/**********************************************************************
 * Heino Doessing — Volunteer recruitment phone bank
 *
 * Manual-dial: the app serves one contact at a time with a script and
 * disposition buttons; the caller dials on their own phone. No Twilio.
 *
 * The call list is built from the campaign's OWN form submissions — people who
 * already gave consent to be contacted — so it works with no external data.
 *
 *   ?view=phonebank   staff — sign in, get a contact, record the outcome
 *
 * The payoff of keeping this in the same project as the shift system: when
 * someone says yes on the phone, the caller books them into a canvass shift
 * in the same screen, and it lands in ShiftSignups like a web signup.
 **********************************************************************/

const CALL_LIST_TAB = 'CallList';

const CL_ID = 1, CL_VANID = 2, CL_FIRST = 3, CL_LAST = 4, CL_PHONE = 5,
      CL_EMAIL = 6, CL_POSTAL = 7, CL_ADDRESS = 8, CL_AGE = 9,
      CL_SOURCE = 10, CL_CONSENT = 11, CL_PRIORITY = 12,
      CL_STATUS = 13, CL_ATTEMPTS = 14, CL_LAST_CALLED = 15, CL_LAST_BY = 16,
      CL_CALLBACK_AT = 17, CL_NOTES = 18, CL_CLAIMED_BY = 19, CL_CLAIMED_AT = 20,
      CL_BOOKED = 21;
const CL_COLS = 21;
const CALL_LIST_HEADERS = ['Contact ID', 'VAN ID', 'First name', 'Last name', 'Phone',
  'Email', 'Postal code', 'Address', 'Age', 'Source', 'Email consent', 'Priority',
  'Status', 'Attempts', 'Last called', 'Last called by', 'Callback at', 'Notes',
  'Claimed by', 'Claimed at', 'Booked shift'];

// Email consent. Voter-file records never asked to hear from us, so they are
// callable but must not be swept into email sends.
const CONSENT_YES = 'Yes (form)', CONSENT_NONE = 'No (voter file)';

// Priority decides who gets called first. Someone who filled in a form this week
// is a far warmer ask than a name off the voter file.
const PRI_FORM = 1, PRI_VOTERFILE = 2;

// Statuses. 'New' and 'Callback' and 'No answer' are callable; the rest are done.
const CS_NEW = 'New', CS_CALLBACK = 'Callback', CS_NO_ANSWER = 'No answer',
      CS_BOOKED = 'Booked', CS_MAYBE = 'Maybe later', CS_DECLINED = 'Declined',
      CS_WRONG = 'Wrong number', CS_DNC = 'Do not call';

const MAX_ATTEMPTS = 4;              // stop calling after this many no-answers
const CLAIM_MINUTES = 8;             // a contact held this long by an idle caller is released

// Which form tabs to harvest contacts from. Every one of these forms has a
// required consent checkbox, so everybody here has agreed to be contacted.
const CALL_SOURCE_TABS = ['Volunteers', 'CommitToVote', 'SignRequests', 'Donations'];

// The static voter-file export. Different headers entirely — see buildCallList.
const MASTER_LIST_TAB = 'Master List';

/*** LIST BUILDING ***/

/**
 * Build or refresh the call list from the form submission tabs.
 * Idempotent: existing contacts keep their status, attempts, and notes —
 * only genuinely new people are appended. Safe to run any time; run it again
 * after a burst of new signups.
 *
 * Skips: anyone with no phone number, anyone whose consent was 'No', anyone
 * already signed up for a canvass shift, and anyone marked Do not call.
 */
function buildCallList() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const cl = ensureCallListSheet_(ss);

  // Existing contacts, keyed by normalized phone, so a refresh never duplicates.
  const seen = {};
  if (cl.getLastRow() >= 2) {
    cl.getRange(2, 1, cl.getLastRow() - 1, CL_COLS).getValues().forEach(function (r) {
      const k = normPhone_(r[CL_PHONE - 1]);
      if (k) seen[k] = true;
    });
  }

  // People who already have a shift don't need recruiting.
  const su = ss.getSheetByName(SHIFT_SIGNUPS_TAB);
  if (su && su.getLastRow() >= 2) {
    su.getRange(2, 1, su.getLastRow() - 1, SU_COLS).getValues().forEach(function (r) {
      const k = normPhone_(r[SU_PHONE - 1]);
      if (k) seen[k] = true;
    });
  }

  const rows = [];

  // ---- 1. Form tabs. Everyone here ticked a consent box. Warm, so called first.
  CALL_SOURCE_TABS.forEach(function (tabName) {
    const sh = ss.getSheetByName(tabName);
    if (!sh || sh.getLastRow() < 2) return;

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                      .map(function (h) { return String(h).trim().toLowerCase(); });
    const iFirst   = headers.indexOf('first name');
    const iLast    = headers.indexOf('last name');
    const iPhone   = headers.indexOf('phone');
    const iEmail   = headers.indexOf('email');
    const iPostal  = headers.indexOf('postal code');
    const iConsent = headers.indexOf('consent');
    if (iPhone === -1) return;                       // no phone column, nothing to call

    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(function (r) {
      const phone = normPhone_(r[iPhone]);
      if (!phone || phone.length < 10) return;
      if (iConsent !== -1 && String(r[iConsent]).trim().toLowerCase() === 'no') return;
      if (seen[phone]) return;
      seen[phone] = true;

      rows.push(newCallRow_({
        first:  iFirst  !== -1 ? r[iFirst]  : '',
        last:   iLast   !== -1 ? r[iLast]   : '',
        phone:  phone,
        email:  iEmail  !== -1 ? r[iEmail]  : '',
        postal: iPostal !== -1 ? r[iPostal] : '',
        source: tabName, consent: CONSENT_YES, priority: PRI_FORM
      }));
    });
  });

  // ---- 2. Master List (voter-file export). Different shape entirely:
  //         one 'Name' column, three phone columns, 'Preferred Email', no consent.
  const ml = ss.getSheetByName(MASTER_LIST_TAB);
  if (ml && ml.getLastRow() >= 2) {
    const H = ml.getRange(1, 1, 1, ml.getLastColumn()).getValues()[0]
                .map(function (h) { return String(h).trim().toLowerCase(); });
    const iVan     = H.indexOf('van id');
    const iName    = H.indexOf('name');
    const iAddress = H.indexOf('address');
    const iPostal  = H.indexOf('postal code');
    const iAge     = H.indexOf('age');
    const iEmail   = H.indexOf('preferred email');
    // Cell first — most likely to be answered and least likely to be a shared landline.
    const phoneCols = ['cell phone', 'phone', 'landline phone']
                        .map(function (h) { return H.indexOf(h); })
                        .filter(function (i) { return i !== -1; });

    ml.getRange(2, 1, ml.getLastRow() - 1, ml.getLastColumn()).getValues().forEach(function (r) {
      let phone = '';
      for (let i = 0; i < phoneCols.length; i++) {
        const p = normPhone_(r[phoneCols[i]]);
        if (p && p.length === 10) { phone = p; break; }
      }
      if (!phone) return;                            // no reachable number
      if (seen[phone]) return;                       // already have them from a form
      seen[phone] = true;

      const nm = splitName_(iName !== -1 ? r[iName] : '');
      rows.push(newCallRow_({
        vanId:   iVan     !== -1 ? r[iVan]     : '',
        first:   nm.first, last: nm.last,
        phone:   phone,
        email:   iEmail   !== -1 ? r[iEmail]   : '',
        postal:  iPostal  !== -1 ? r[iPostal]  : '',
        address: iAddress !== -1 ? r[iAddress] : '',
        age:     iAge     !== -1 ? r[iAge]     : '',
        source:  MASTER_LIST_TAB, consent: CONSENT_NONE, priority: PRI_VOTERFILE
      }));
    });
  }

  if (rows.length) {
    cl.getRange(cl.getLastRow() + 1, 1, rows.length, CL_COLS).setValues(rows);
    cl.getRange(2, 1, cl.getLastRow() - 1, CL_COLS)
      .sort([{ column: CL_PRIORITY, ascending: true }]);
  }
  return { added: rows.length, total: Math.max(0, cl.getLastRow() - 1) };
}

/** One CallList row from a normalized person object. */
function newCallRow_(p) {
  return [
    Utilities.getUuid().slice(0, 8),
    p.vanId || '', String(p.first || '').trim(), String(p.last || '').trim(),
    formatPhone_(p.phone), p.email || '', p.postal || '', p.address || '', p.age || '',
    p.source || '', p.consent || CONSENT_NONE, p.priority || PRI_VOTERFILE,
    CS_NEW, 0, '', '', '', '', '', '', ''
  ];
}

/**
 * Voter files write names either "Last, First" or "First Last" — handle both.
 * Everything after the first token is treated as the surname, so compound
 * surnames survive ("Maria Van Der Berg" → Maria / Van Der Berg).
 */
function splitName_(full) {
  const s = String(full || '').trim().replace(/\s+/g, ' ');
  if (!s) return { first: '', last: '' };
  if (s.indexOf(',') !== -1) {
    const parts = s.split(',');
    return { first: (parts[1] || '').trim(), last: (parts[0] || '').trim() };
  }
  const bits = s.split(' ');
  if (bits.length === 1) return { first: bits[0], last: '' };
  return { first: bits[0], last: bits.slice(1).join(' ') };
}

/** Admin-triggered refresh, so nobody has to open the script editor. */
function adminBuildCallList(token) {
  requireAdmin_(token);
  try {
    const res = buildCallList();
    return { ok: true, message: 'Added ' + res.added + ' new contact' +
                                (res.added === 1 ? '' : 's') +
                                '. The call list now holds ' + res.total + '.' };
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

/*** CALLER API (Phonebank.html) ***/

/**
 * Claim and return the next contact to call, plus the shifts that can be
 * offered. Returns {done:true} when the list is exhausted.
 *
 * Claiming prevents two callers being handed the same person: a claim older
 * than CLAIM_MINUTES is treated as abandoned and can be taken over.
 */
function getNextContact(token) {
  const user = requireStaff_(token);

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (err) { return { error: 'Busy, please try again.' }; }

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const cl = ensureCallListSheet_(ss);
    if (cl.getLastRow() < 2) {
      return { done: true, message: 'The call list is empty. Run buildCallList() to populate it.' };
    }

    const data = cl.getRange(2, 1, cl.getLastRow() - 1, CL_COLS).getValues();
    const now = new Date();
    const staleBefore = new Date(now.getTime() - CLAIM_MINUTES * 60000);
    const skipped = loadSkips_(user.email);

    let bestRow = -1, bestScore = -1;
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (!r[CL_ID - 1]) continue;
      if (skipped[String(r[CL_ID - 1])]) continue;      // this caller skipped them

      const status = String(r[CL_STATUS - 1] || CS_NEW);
      if ([CS_BOOKED, CS_DECLINED, CS_WRONG, CS_DNC].indexOf(status) !== -1) continue;
      if (Number(r[CL_ATTEMPTS - 1] || 0) >= MAX_ATTEMPTS) continue;

      // Held by somebody else and still fresh → skip.
      const claimedBy = String(r[CL_CLAIMED_BY - 1] || '');
      const claimedAt = r[CL_CLAIMED_AT - 1];
      if (claimedBy && claimedBy !== user.email &&
          claimedAt instanceof Date && claimedAt > staleBefore) continue;

      // Callbacks that are due outrank everything; then new; then retries.
      let score;
      if (status === CS_CALLBACK) {
        const due = r[CL_CALLBACK_AT - 1];
        if (due instanceof Date && due > now) continue;   // not due yet
        score = 3000;
      } else if (status === CS_NEW) {
        score = 2000;
      } else if (status === CS_MAYBE) {
        score = 500;
      } else {
        score = 1000 - Number(r[CL_ATTEMPTS - 1] || 0);   // fewer attempts first
      }
      // Within the same status, warm form signups outrank cold voter-file names.
      if (Number(r[CL_PRIORITY - 1] || PRI_VOTERFILE) === PRI_FORM) score += 100;

      if (score > bestScore) { bestScore = score; bestRow = i + 2; }
    }

    if (bestRow === -1) {
      const nSkipped = Object.keys(skipped).length;
      return { done: true,
               skippedCount: nSkipped,
               message: nSkipped
                 ? 'Nobody left to call. You skipped ' + nSkipped + ' contact' +
                   (nSkipped === 1 ? '' : 's') + ' — you can bring them back below.'
                 : 'Nobody left to call right now. Great work!' };
    }

    cl.getRange(bestRow, CL_CLAIMED_BY).setValue(user.email);
    cl.getRange(bestRow, CL_CLAIMED_AT).setValue(now);
    SpreadsheetApp.flush();

    const r = data[bestRow - 2];
    return {
      done: false,
      contactId: r[CL_ID - 1],
      first: r[CL_FIRST - 1] || '',
      last: r[CL_LAST - 1] || '',
      phone: r[CL_PHONE - 1] || '',
      email: r[CL_EMAIL - 1] || '',
      postal: r[CL_POSTAL - 1] || '',
      address: r[CL_ADDRESS - 1] || '',
      age: r[CL_AGE - 1] || '',
      source: r[CL_SOURCE - 1] || '',
      warm: Number(r[CL_PRIORITY - 1] || PRI_VOTERFILE) === PRI_FORM,
      emailConsent: r[CL_CONSENT - 1] || '',
      status: r[CL_STATUS - 1] || CS_NEW,
      attempts: Number(r[CL_ATTEMPTS - 1] || 0),
      notes: r[CL_NOTES - 1] || '',
      shifts: listOpenShifts(14)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Record what happened on a call.
 * outcome: 'booked' | 'maybe' | 'declined' | 'noanswer' | 'callback' | 'wrong' | 'dnc'
 * If outcome is 'booked', shiftIds are signed up exactly as a web signup would be,
 * so the person appears in the week tracker and gets the same confirmation email.
 */
function saveCallOutcome(token, payload) {
  const user = requireStaff_(token);
  payload = payload || {};

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const cl = ensureCallListSheet_(ss);
  const row = findCallRow_(cl, payload.contactId);
  if (row === -1) return { ok: false, error: 'That contact could not be found.' };

  const rec = cl.getRange(row, 1, 1, CL_COLS).getValues()[0];
  const now = new Date();
  const outcome = String(payload.outcome || '').toLowerCase();
  let status = CS_NO_ANSWER, booked = [];

  if (outcome === 'booked') {
    const ids = (payload.shiftIds || []).filter(String);
    if (!ids.length) return { ok: false, error: 'Pick at least one shift to book them into.' };
    const res = submitShiftSignup({
      first: rec[CL_FIRST - 1], last: rec[CL_LAST - 1],
      email: rec[CL_EMAIL - 1], phone: rec[CL_PHONE - 1],
      postal: rec[CL_POSTAL - 1], consent: true,
      shiftIds: ids, source: 'Phone bank (' + user.email + ')'
    });
    if (!res.ok) return { ok: false, error: res.error };
    booked = res.booked || [];
    status = CS_BOOKED;
    cl.getRange(row, CL_BOOKED).setValue(booked.join(' | '));
  } else if (outcome === 'maybe')    { status = CS_MAYBE;
  } else if (outcome === 'declined') { status = CS_DECLINED;
  } else if (outcome === 'wrong')    { status = CS_WRONG;
  } else if (outcome === 'dnc')      { status = CS_DNC;
  } else if (outcome === 'callback') {
    status = CS_CALLBACK;
    if (payload.callbackAt) {
      const d = new Date(payload.callbackAt);
      if (!isNaN(d.getTime())) cl.getRange(row, CL_CALLBACK_AT).setValue(d);
    }
  }

  // Only a genuine dial attempt increments the counter.
  if (['noanswer', 'callback'].indexOf(outcome) !== -1) {
    cl.getRange(row, CL_ATTEMPTS).setValue(Number(rec[CL_ATTEMPTS - 1] || 0) + 1);
  }

  cl.getRange(row, CL_STATUS).setValue(status);
  cl.getRange(row, CL_LAST_CALLED).setValue(now);
  cl.getRange(row, CL_LAST_BY).setValue(user.email);
  if (payload.notes != null && String(payload.notes).trim()) {
    const prev = String(rec[CL_NOTES - 1] || '');
    const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d') +
                  ' (' + user.name + '): ' + String(payload.notes).trim();
    cl.getRange(row, CL_NOTES).setValue(prev ? prev + '\n' + stamp : stamp);
  }
  // Release the claim so the next caller can pick it up if it comes back around.
  cl.getRange(row, CL_CLAIMED_BY).setValue('');
  cl.getRange(row, CL_CLAIMED_AT).setValue('');

  return { ok: true, status: status, booked: booked };
}

/** Put a contact back without recording an outcome (caller closed the tab, etc.). */
function releaseContact(token, contactId) {
  requireStaff_(token);
  const cl = ensureCallListSheet_(SpreadsheetApp.openById(SHEET_ID));
  const row = findCallRow_(cl, contactId);
  if (row !== -1) {
    cl.getRange(row, CL_CLAIMED_BY).setValue('');
    cl.getRange(row, CL_CLAIMED_AT).setValue('');
  }
  return { ok: true };
}

/**
 * Skip a contact for good (for this caller).
 *
 * Releasing alone is not enough: the contact keeps its status and priority, so
 * the very next getNextContact would hand back the same person. The skip is
 * therefore recorded per caller in Script Properties, which survives reloads
 * and new sessions. Other callers still see the contact — one person skipping
 * is not a decision for the whole team. To actually remove someone from the
 * list for everyone, use the "Not interested" or "Do not call" outcome instead.
 */
function skipContact(token, contactId) {
  const user = requireStaff_(token);
  if (contactId) {
    const key = skipKey_(user.email);
    const props = PropertiesService.getScriptProperties();
    let ids = [];
    try { ids = JSON.parse(props.getProperty(key) || '[]'); } catch (err) { ids = []; }
    if (ids.indexOf(String(contactId)) === -1) ids.push(String(contactId));
    // Keep it bounded — Script Properties values cap at ~9KB.
    if (ids.length > 400) ids = ids.slice(ids.length - 400);
    props.setProperty(key, JSON.stringify(ids));
  }
  releaseContact(token, contactId);
  return { ok: true };
}

/** Bring this caller's skipped contacts back into their queue. */
function clearMySkips(token) {
  const user = requireStaff_(token);
  PropertiesService.getScriptProperties().deleteProperty(skipKey_(user.email));
  return { ok: true, message: 'Skipped contacts are back in your queue.' };
}

function skipKey_(email) { return 'skip_' + String(email).toLowerCase(); }

function loadSkips_(email) {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(skipKey_(email));
    const arr = JSON.parse(raw || '[]');
    const set = {};
    arr.forEach(function (id) { set[String(id)] = true; });
    return set;
  } catch (err) { return {}; }
}

/** Counts for the header strip. */
function getPhonebankStats(token) {
  requireStaff_(token);
  const cl = ensureCallListSheet_(SpreadsheetApp.openById(SHEET_ID));
  const out = { total: 0, remaining: 0, booked: 0, calledToday: 0 };
  if (cl.getLastRow() < 2) return out;

  const today = new Date(); today.setHours(0, 0, 0, 0);
  cl.getRange(2, 1, cl.getLastRow() - 1, CL_COLS).getValues().forEach(function (r) {
    if (!r[CL_ID - 1]) return;
    out.total++;
    const st = String(r[CL_STATUS - 1] || CS_NEW);
    if (st === CS_BOOKED) out.booked++;
    if ([CS_BOOKED, CS_DECLINED, CS_WRONG, CS_DNC].indexOf(st) === -1 &&
        Number(r[CL_ATTEMPTS - 1] || 0) < MAX_ATTEMPTS) out.remaining++;
    const lc = r[CL_LAST_CALLED - 1];
    if (lc instanceof Date && lc >= today) out.calledToday++;
  });
  return out;
}

/**
 * Admin overview of the call list: how many sit at each status, and the most
 * recent calls with who made them. Lets the campaign see results without
 * opening the spreadsheet.
 */
function adminCallListSummary(token) {
  requireAdmin_(token);
  const cl = ensureCallListSheet_(SpreadsheetApp.openById(SHEET_ID));
  const out = { byStatus: {}, bySource: {}, recent: [], total: 0 };
  if (cl.getLastRow() < 2) return out;

  const tz = Session.getScriptTimeZone();
  const rows = cl.getRange(2, 1, cl.getLastRow() - 1, CL_COLS).getValues();
  const called = [];

  rows.forEach(function (r) {
    if (!r[CL_ID - 1]) return;
    out.total++;
    const st = String(r[CL_STATUS - 1] || CS_NEW);
    out.byStatus[st] = (out.byStatus[st] || 0) + 1;
    const src = String(r[CL_SOURCE - 1] || '?');
    out.bySource[src] = (out.bySource[src] || 0) + 1;
    if (r[CL_LAST_CALLED - 1] instanceof Date) {
      called.push({
        when: r[CL_LAST_CALLED - 1],
        contactId: r[CL_ID - 1],
        name: String(r[CL_FIRST - 1] || '') + ' ' + String(r[CL_LAST - 1] || ''),
        status: st,
        by: String(r[CL_LAST_BY - 1] || ''),
        attempts: Number(r[CL_ATTEMPTS - 1] || 0)
      });
    }
  });

  called.sort(function (a, b) { return b.when - a.when; });
  out.recent = called.slice(0, 25).map(function (c) {
    c.when = Utilities.formatDate(c.when, tz, 'EEE MMM d, h:mm a');
    return c;
  });
  return out;
}

/**
 * Put a contact back to untouched — clears status, attempts, call history and
 * any claim. For undoing a mis-click, or resetting after testing.
 */
function adminResetContact(token, contactId) {
  requireAdmin_(token);
  const cl = ensureCallListSheet_(SpreadsheetApp.openById(SHEET_ID));
  const row = findCallRow_(cl, contactId);
  if (row === -1) return { ok: false, error: 'No contact with that ID.' };
  cl.getRange(row, CL_STATUS).setValue(CS_NEW);
  cl.getRange(row, CL_ATTEMPTS).setValue(0);
  [CL_LAST_CALLED, CL_LAST_BY, CL_CALLBACK_AT, CL_CLAIMED_BY, CL_CLAIMED_AT, CL_BOOKED]
    .forEach(function (c) { cl.getRange(row, c).setValue(''); });
  return { ok: true, message: 'Reset ' + contactId + ' to New.' };
}

/*** HELPERS ***/

function ensureCallListSheet_(ss) {
  let cl = ss.getSheetByName(CALL_LIST_TAB);
  if (!cl) {
    cl = ss.insertSheet(CALL_LIST_TAB);
    cl.getRange(1, 1, 1, CL_COLS).setValues([CALL_LIST_HEADERS]).setFontWeight('bold');
    cl.setFrozenRows(1);
  }
  return cl;
}

function findCallRow_(cl, contactId) {
  if (!contactId || cl.getLastRow() < 2) return -1;
  const vals = cl.getRange(2, CL_ID, cl.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(contactId)) return i + 2;
  }
  return -1;
}

/** Digits only, dropping a leading country code, so 1-905-555-0142 == (905) 555-0142. */
function normPhone_(v) {
  let d = String(v == null ? '' : v).replace(/\D/g, '');
  if (d.length === 11 && d.charAt(0) === '1') d = d.slice(1);
  return d;
}

function formatPhone_(digits) {
  const d = normPhone_(digits);
  if (d.length !== 10) return String(digits || '');
  return '(' + d.slice(0, 3) + ') ' + d.slice(3, 6) + '-' + d.slice(6);
}
