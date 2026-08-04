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

const CL_ID = 1, CL_FIRST = 2, CL_LAST = 3, CL_PHONE = 4, CL_EMAIL = 5,
      CL_POSTAL = 6, CL_SOURCE = 7, CL_STATUS = 8, CL_ATTEMPTS = 9,
      CL_LAST_CALLED = 10, CL_LAST_BY = 11, CL_CALLBACK_AT = 12, CL_NOTES = 13,
      CL_CLAIMED_BY = 14, CL_CLAIMED_AT = 15, CL_BOOKED = 16;
const CL_COLS = 16;
const CALL_LIST_HEADERS = ['Contact ID', 'First name', 'Last name', 'Phone', 'Email',
  'Postal code', 'Source', 'Status', 'Attempts', 'Last called', 'Last called by',
  'Callback at', 'Notes', 'Claimed by', 'Claimed at', 'Booked shift'];

// Statuses. 'New' and 'Callback' and 'No answer' are callable; the rest are done.
const CS_NEW = 'New', CS_CALLBACK = 'Callback', CS_NO_ANSWER = 'No answer',
      CS_BOOKED = 'Booked', CS_MAYBE = 'Maybe later', CS_DECLINED = 'Declined',
      CS_WRONG = 'Wrong number', CS_DNC = 'Do not call';

const MAX_ATTEMPTS = 4;              // stop calling after this many no-answers
const CLAIM_MINUTES = 8;             // a contact held this long by an idle caller is released

// Which form tabs to harvest contacts from. Every one of these forms has a
// required consent checkbox, so everybody here has agreed to be contacted.
const CALL_SOURCE_TABS = ['CommitToVote', 'Volunteers', 'SignRequests', 'Donations'];

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

      rows.push([
        Utilities.getUuid().slice(0, 8),
        iFirst  !== -1 ? r[iFirst]  : '',
        iLast   !== -1 ? r[iLast]   : '',
        formatPhone_(phone),
        iEmail  !== -1 ? r[iEmail]  : '',
        iPostal !== -1 ? r[iPostal] : '',
        tabName, CS_NEW, 0, '', '', '', '', '', '', ''
      ]);
    });
  });

  if (rows.length) {
    cl.getRange(cl.getLastRow() + 1, 1, rows.length, CL_COLS).setValues(rows);
  }
  return { added: rows.length, total: Math.max(0, cl.getLastRow() - 1) };
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

    let bestRow = -1, bestScore = -1;
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (!r[CL_ID - 1]) continue;

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
        score = 300;
      } else if (status === CS_NEW) {
        score = 200;
      } else if (status === CS_MAYBE) {
        score = 50;
      } else {
        score = 100 - Number(r[CL_ATTEMPTS - 1] || 0);    // fewer attempts first
      }

      if (score > bestScore) { bestScore = score; bestRow = i + 2; }
    }

    if (bestRow === -1) {
      return { done: true, message: 'Nobody left to call right now. Great work!' };
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
      source: r[CL_SOURCE - 1] || '',
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
