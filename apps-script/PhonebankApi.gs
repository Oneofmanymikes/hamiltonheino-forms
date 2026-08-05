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
      CL_BOOKED = 21,
      // --- ported from the Hafiz phone bank: lapse + call-limit machinery ---
      CL_STAGE = 22,          // New / Active / Lapsed
      CL_NA_STREAK = 23,      // consecutive no-answers
      CL_TRIED_SLOTS = 24,    // time-of-day slots already tried with no answer
      CL_HOLD_UNTIL = 25,     // suppressed from every queue until this moment
      CL_FOLLOWUP_AT = 26,    // resurface on/after this date
      CL_SEGMENT = 27;        // which call list they belong to (see below)
const CL_COLS = 27;
const CALL_LIST_HEADERS = ['Contact ID', 'VAN ID', 'First name', 'Last name', 'Phone',
  'Email', 'Postal code', 'Address', 'Age', 'Source', 'Email consent', 'Priority',
  'Status', 'Attempts', 'Last called', 'Last called by', 'Callback at', 'Notes',
  'Claimed by', 'Claimed at', 'Booked shift',
  'Stage', 'No-answer streak', 'Tried slots', 'Hold until', 'Follow up at', 'Segment'];

/* ---------------------------------------------------------------------------
 * SEGMENTS — three genuinely different conversations, so the caller can pick
 * which one they are having rather than switching mental gears every call.
 *
 *   Recruit        never signed up with us. Cold-ish: voter file, or someone
 *                  who only committed to vote / asked for a sign / pledged.
 *                  Ask: would you come out and canvass?
 *   New volunteer  filled in the volunteer form but has never booked a shift.
 *                  Warm onboarding: welcome them, get a first shift in the diary.
 *   Active         has at least one canvass shift booked. Confirm, re-book,
 *                  thank. Never treat these as a cold ask.
 *
 * Segment is recomputed by refreshSegments(), and set to Active the moment a
 * booking is made (on the phone or through the website).
 * ------------------------------------------------------------------------ */
const SEG_RECRUIT = 'Recruit';
const SEG_ONBOARD = 'New volunteer';
const SEG_ACTIVE  = 'Active volunteer';

/* ---------------------------------------------------------------------------
 * LAPSE + CALL-LIMIT RULES (ported from votehafiz.ca's phone bank)
 *
 * 1. Four consecutive no-answers auto-lapses a contact: they leave the normal
 *    queue and only reappear in the dedicated "Lapsed" mission.
 * 2. Every no-answer puts them on hold until 6 AM the next day, so nobody gets
 *    rung twice in an evening.
 * 3. Time-of-day rotation. Without it, a contact could burn all four attempts
 *    in the same weekday-morning slot and lapse without ever being tried on an
 *    evening or a weekend. A slot that produced no answer is remembered, and
 *    the contact is hidden during that slot next time round.
 * 4. Thirty days with no contact also lapses them (run markLapsedContacts on a
 *    daily trigger).
 * 5. Reaching a person clears all of it: streak zeroed, slots cleared, stage
 *    back to Active.
 * ------------------------------------------------------------------------ */
const NO_ANSWER_LAPSE_THRESHOLD = 4;   // consecutive no-answers before lapsing
const NO_ANSWER_HOLD_HOUR = 6;         // released at 6 AM local next day
const NO_ANSWER_FOLLOWUP_DAYS = 1;     // resurface the following morning
const LAPSE_AFTER_DAYS = 30;           // untouched for this long → Lapsed
const OUTCOME_DEDUPE_MS = 15000;       // ignore an identical re-submit within 15s

const ES_NEW = 'New', ES_ACTIVE = 'Active', ES_LAPSED = 'Lapsed';

// Time-of-day slots. Weekday splits three ways; weekend splits at 1 PM.
const SLOT_WD_AM = 'WD-AM', SLOT_WD_PM = 'WD-PM', SLOT_WD_EVE = 'WD-EVE',
      SLOT_WE_AM = 'WE-AM', SLOT_WE_PM = 'WE-PM';
const SLOT_WD_PM_HOUR = 12, SLOT_WD_EVE_HOUR = 17, SLOT_WE_PM_HOUR = 13;
const SLOT_LABELS = {
  'WD-AM': 'weekday morning', 'WD-PM': 'weekday afternoon', 'WD-EVE': 'weekday evening',
  'WE-AM': 'weekend morning', 'WE-PM': 'weekend afternoon'
};

// Queue modes ("missions").
const MODE_ELIGIBLE = 'eligible';   // default — everything callable, Lapsed excluded
const MODE_LAPSED   = 'lapsed';     // re-engagement: Lapsed contacts only
const MODE_CALLBACK = 'callback';   // promised call-backs that are due
const MODE_RECRUIT  = 'recruit';    // segment: Recruit
const MODE_ONBOARD  = 'onboard';    // segment: New volunteer
const MODE_ACTIVE   = 'active';     // segment: Active volunteer

const MODE_CONFIRM  = 'confirm';    // has a shift today/tomorrow, not yet confirmed
const MODE_THANKYOU = 'thankyou';   // did a shift recently, not yet thanked

/**
 * Phone → shift context, for the confirm and thank-you missions.
 * Returns { confirm: {phone: {...}}, thank: {phone: {...}} }.
 *
 * Both missions are DATE-driven, so they deliberately ignore the time-of-day
 * rotation and the no-answer hold: a shift-confirmation call that waits until
 * tomorrow morning is worthless if the shift is this afternoon.
 */
function shiftCallContext_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const su = ss.getSheetByName(SHIFT_SIGNUPS_TAB);
  const out = { confirm: {}, thank: {} };
  if (!su || su.getLastRow() < 2) return out;

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayStart = new Date(now); todayStart.setHours(0, 0, 0, 0);
  const cutoff = new Date(todayStart); cutoff.setDate(cutoff.getDate() + 2);   // today + tomorrow
  const thankFrom = new Date(todayStart);
  thankFrom.setDate(thankFrom.getDate() - THANKYOU_WINDOW_DAYS);

  su.getRange(2, 1, su.getLastRow() - 1, SU_COLS).getValues().forEach(function (r) {
    if (!r[SU_ID - 1]) return;
    const phone = normPhone_(r[SU_PHONE - 1]);
    if (!phone) return;
    const st = String(r[SU_STATUS - 1] || SS_SCHEDULED);
    const when = shiftStart_(r[SU_DATE - 1], r[SU_TIME - 1]);
    if (!when) return;

    const label = Utilities.formatDate(when, tz, 'EEEE, MMMM d') + ' at ' +
                  prettyTime_(r[SU_TIME - 1]);

    // Confirm: an open booking that starts today or tomorrow and is unconfirmed.
    if (SS_OPEN.indexOf(st) !== -1 && !r[SU_CONFIRMED_AT - 1] &&
        when >= todayStart && when < cutoff) {
      out.confirm[phone] = { signupId: r[SU_ID - 1], label: label, when: when.getTime() };
    }

    // Thank you: turned up recently and has not been thanked.
    if ((st === SS_COMPLETED || st === SS_ARRIVED) && !r[SU_THANKED_AT - 1] &&
        when >= thankFrom && when <= now) {
      out.thank[phone] = { signupId: r[SU_ID - 1], label: label, when: when.getTime() };
    }
  });
  return out;
}

/** Which segment does a mode restrict to? null = no segment restriction. */
function segmentForMode_(mode) {
  if (mode === MODE_RECRUIT) return SEG_RECRUIT;
  if (mode === MODE_ONBOARD) return SEG_ONBOARD;
  if (mode === MODE_ACTIVE)  return SEG_ACTIVE;
  return null;
}

/** Which time-of-day slot are we in right now? */
function currentSlot_(now) {
  const d = now || new Date();
  const weekend = (d.getDay() === 0 || d.getDay() === 6);
  const h = d.getHours();
  if (weekend) return h < SLOT_WE_PM_HOUR ? SLOT_WE_AM : SLOT_WE_PM;
  if (h < SLOT_WD_PM_HOUR) return SLOT_WD_AM;
  if (h < SLOT_WD_EVE_HOUR) return SLOT_WD_PM;
  return SLOT_WD_EVE;
}

// Email consent. Voter-file records never asked to hear from us, so they are
// callable but must not be swept into email sends.
const CONSENT_YES = 'Yes (form)', CONSENT_NONE = 'No (voter file)';

// Priority decides who gets called first. Someone who filled in a form this week
// is a far warmer ask than a name off the voter file.
const PRI_FORM = 1, PRI_VOTERFILE = 2;

// Statuses. 'New' and 'Callback' and 'No answer' are callable; the rest are done.
const CS_NEW = 'New', CS_CALLBACK = 'Callback', CS_NO_ANSWER = 'No answer',
      CS_VOICEMAIL = 'Voicemail',   // counts like a no-answer toward lapsing
      CS_BOOKED = 'Booked', CS_MAYBE = 'Maybe later', CS_DECLINED = 'Declined',
      CS_WRONG = 'Wrong number', CS_DNC = 'Do not call';

const MAX_ATTEMPTS = 4;              // stop calling after this many no-answers
const CLAIM_MINUTES = 8;             // a contact held this long by an idle caller is released
const SKIP_TTL_DAYS = 7;             // a caller's skip lasts a week, then the contact returns

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
        source: tabName, consent: CONSENT_YES, priority: PRI_FORM,
        segment: (tabName === 'Volunteers') ? SEG_ONBOARD : SEG_RECRUIT
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
        source:  MASTER_LIST_TAB, consent: CONSENT_NONE, priority: PRI_VOTERFILE,
        segment: SEG_RECRUIT
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
    CS_NEW, 0, '', '', '', '', '', '', '',
    ES_NEW, 0, '', '', '', p.segment || SEG_RECRUIT
  ];
}

/**
 * Recompute every contact's segment.
 *   anyone with a shift booking            → Active volunteer
 *   anyone who came from the Volunteers form → New volunteer
 *   everyone else                          → Recruit
 * Safe to re-run; call it after a rebuild or a batch of website signups.
 */
function refreshSegments() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const cl = ensureCallListSheet_(ss);
  if (cl.getLastRow() < 2) return { updated: 0 };

  // Phones that have booked at least one canvass shift.
  const booked = {};
  const su = ss.getSheetByName(SHIFT_SIGNUPS_TAB);
  if (su && su.getLastRow() >= 2) {
    su.getRange(2, 1, su.getLastRow() - 1, SU_COLS).getValues().forEach(function (r) {
      if (String(r[SU_STATUS - 1] || '') === 'Cancelled') return;
      const k = normPhone_(r[SU_PHONE - 1]);
      if (k) booked[k] = true;
    });
  }

  // Phones that came in through the volunteer form.
  const volunteered = {};
  const vs = ss.getSheetByName('Volunteers');
  if (vs && vs.getLastRow() >= 2) {
    const H = vs.getRange(1, 1, 1, vs.getLastColumn()).getValues()[0]
                .map(function (h) { return String(h).trim().toLowerCase(); });
    const iPhone = H.indexOf('phone');
    if (iPhone !== -1) {
      vs.getRange(2, 1, vs.getLastRow() - 1, vs.getLastColumn()).getValues().forEach(function (r) {
        const k = normPhone_(r[iPhone]);
        if (k) volunteered[k] = true;
      });
    }
  }

  const n = cl.getLastRow() - 1;
  const rows = cl.getRange(2, 1, n, CL_COLS).getValues();
  const out = [];
  let updated = 0;

  for (let i = 0; i < n; i++) {
    const key = normPhone_(rows[i][CL_PHONE - 1]);
    const src = String(rows[i][CL_SOURCE - 1] || '');
    let seg = SEG_RECRUIT;
    if (key && booked[key]) seg = SEG_ACTIVE;
    else if ((key && volunteered[key]) || src === 'Volunteers') seg = SEG_ONBOARD;
    if (String(rows[i][CL_SEGMENT - 1] || '') !== seg) updated++;
    out.push([seg]);
  }
  cl.getRange(2, CL_SEGMENT, n, 1).setValues(out);
  return { updated: updated, total: n };
}

/**
 * Integrity repair: give any row a Contact ID if it has lost one.
 * A row with a blank ID is invisible to the queue and to reporting, but its
 * phone number still blocks the rebuild from re-adding the person — so they
 * silently vanish from the list. This puts them back.
 */
function repairCallListIds() {
  const cl = ensureCallListSheet_(SpreadsheetApp.openById(SHEET_ID));
  const n = cl.getLastRow() - 1;
  if (n < 1) return { repaired: 0 };

  const rows = cl.getRange(2, 1, n, CL_COLS).getValues();
  const out = [];
  let repaired = 0;
  for (let i = 0; i < n; i++) {
    let id = String(rows[i][CL_ID - 1] || '').trim();
    const hasData = String(rows[i][CL_PHONE - 1] || '').trim() !== '';
    if (!id && hasData) { id = Utilities.getUuid().slice(0, 8); repaired++; }
    out.push([id]);
  }
  cl.getRange(2, CL_ID, n, 1).setValues(out);
  return { repaired: repaired, rows: n };
}

function adminRepairCallList(token) {
  requireAdmin_(token);
  const ids = repairCallListIds();
  const segs = refreshSegments();
  return { ok: true,
           message: 'Repaired ' + ids.repaired + ' missing ID' +
                    (ids.repaired === 1 ? '' : 's') + ' across ' + ids.rows +
                    ' rows; refreshed ' + segs.updated + ' segment' +
                    (segs.updated === 1 ? '' : 's') + '.' };
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
function getNextContact(token, mode) {
  const user = requireStaff_(token);
  mode = mode || MODE_ELIGIBLE;

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
    const slot = currentSlot_(now);
    const wantSegment = segmentForMode_(mode);
    // Date-driven missions: the shift is imminent (or just happened), so these
    // bypass the slot rotation and the overnight hold entirely.
    const dateDriven = (mode === MODE_CONFIRM || mode === MODE_THANKYOU);
    const shiftCtx = dateDriven ? shiftCallContext_() : null;
    const wanted = dateDriven ? (mode === MODE_CONFIRM ? shiftCtx.confirm : shiftCtx.thank) : null;
    let heldBack = 0, lapsedCount = 0;

    let bestRow = -1, bestScore = -1;
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (!r[CL_ID - 1]) continue;
      if (skipped[String(r[CL_ID - 1])]) continue;      // this caller skipped them

      // --- date-driven missions: only people with an imminent/just-past shift ---
      if (dateDriven && !wanted[normPhone_(r[CL_PHONE - 1])]) continue;

      // --- segment filter: which conversation is this caller having? ---
      if (wantSegment && String(r[CL_SEGMENT - 1] || SEG_RECRUIT) !== wantSegment) continue;

      const status = String(r[CL_STATUS - 1] || CS_NEW);
      // Active volunteers are meant to be re-called (confirm / re-book), so a
      // previous booking does not close them out of the Active list.
      const terminal = (String(r[CL_SEGMENT - 1] || '') === SEG_ACTIVE)
        ? [CS_DECLINED, CS_WRONG, CS_DNC]
        : [CS_BOOKED, CS_DECLINED, CS_WRONG, CS_DNC];
      if (!dateDriven && terminal.indexOf(status) !== -1) continue;

      // --- call limit: hard ceiling on attempts ---
      // Exempt in the Lapsed mission. Lapsing IS the call limit for the normal
      // queue; if the ceiling also applied here, a contact who lapsed by running
      // out of attempts could never be re-engaged, which defeats the point of
      // having a re-engagement list at all.
      if (mode !== MODE_LAPSED && !dateDriven &&
          Number(r[CL_ATTEMPTS - 1] || 0) >= MAX_ATTEMPTS) continue;

      // --- lapse gate: a Lapsed contact is reachable ONLY in the Lapsed mission ---
      const stage = String(r[CL_STAGE - 1] || ES_NEW);
      if (stage === ES_LAPSED) {
        lapsedCount++;
        if (mode !== MODE_LAPSED && !dateDriven) continue;
      } else if (mode === MODE_LAPSED) {
        continue;                                       // Lapsed mission wants only Lapsed
      }

      // --- no-answer hold: suppressed until 6 AM the morning after ---
      // Skipped for date-driven missions: a shift starting in three hours can't
      // wait until tomorrow morning to be confirmed.
      const hold = r[CL_HOLD_UNTIL - 1];
      if (!dateDriven && hold instanceof Date && hold > now) { heldBack++; continue; }

      // --- follow-up date not yet due ---
      const fu = r[CL_FOLLOWUP_AT - 1];
      if (!dateDriven && fu instanceof Date && fu > now) { heldBack++; continue; }

      // --- time-of-day rotation: don't retry in a slot that already failed ---
      // Exempt in the Lapsed mission, which is a deliberate re-engagement sweep.
      if (mode !== MODE_LAPSED && !dateDriven) {
        const tried = String(r[CL_TRIED_SLOTS - 1] || '');
        if (tried && tried.split(',').indexOf(slot) !== -1) { heldBack++; continue; }
      }

      // Held by somebody else and still fresh → skip.
      const claimedBy = String(r[CL_CLAIMED_BY - 1] || '');
      const claimedAt = r[CL_CLAIMED_AT - 1];
      if (claimedBy && claimedBy !== user.email &&
          claimedAt instanceof Date && claimedAt > staleBefore) continue;

      // Callbacks that are due outrank everything; then new; then retries.
      let score;
      if (dateDriven) {
        score = 5000;
      } else if (status === CS_CALLBACK) {
        const due = r[CL_CALLBACK_AT - 1];
        if (due instanceof Date && due > now) continue;   // not due yet
        score = 3000;
      } else if (mode === MODE_CALLBACK) {
        continue;                                         // callback mission wants callbacks only
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
      let msg;
      if (mode === MODE_LAPSED) {
        msg = 'No lapsed contacts to re-engage right now.';
      } else if (heldBack) {
        msg = 'Nobody callable right now — ' + heldBack + ' contact' +
              (heldBack === 1 ? ' is' : 's are') + ' resting until a different time of day ' +
              'or until tomorrow morning. Try again later, or switch to the Lapsed list.';
      } else {
        msg = 'Nobody left to call right now. Great work!';
      }
      return { done: true, skippedCount: nSkipped, heldBack: heldBack,
               lapsedCount: lapsedCount, mode: mode, message: msg };
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
      segment: r[CL_SEGMENT - 1] || SEG_RECRUIT,
      stage: r[CL_STAGE - 1] || ES_NEW,
      attempts: Number(r[CL_ATTEMPTS - 1] || 0),
      notes: r[CL_NOTES - 1] || '',
      bookedBefore: String(r[CL_BOOKED - 1] || ''),
      // For the confirm / thank-you missions the caller needs to see WHICH shift.
      shiftContext: dateDriven ? (wanted[normPhone_(r[CL_PHONE - 1])] || null) : null,
      mode: mode,
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

  // Double-submit guard: a laggy tap can fire the same outcome several times,
  // which would otherwise inflate the no-answer streak into a false auto-lapse.
  const dedupeKey = 'oc_' + payload.contactId + '_' + outcome + '_' + user.email;
  const cache = CacheService.getScriptCache();
  if (cache.get(dedupeKey)) {
    return { ok: true, duplicate: true, status: String(rec[CL_STATUS - 1] || CS_NEW) };
  }
  cache.put(dedupeKey, '1', Math.ceil(OUTCOME_DEDUPE_MS / 1000));

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
    // They now have a shift, so they belong in the Active volunteer list —
    // future calls are confirm/re-book, not a cold recruitment ask.
    cl.getRange(row, CL_SEGMENT).setValue(SEG_ACTIVE);
  } else if (outcome === 'voicemail'){ status = CS_VOICEMAIL;
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

  // ---- lapse / hold / rotation machinery -----------------------------------
  // A no-answer (or voicemail) is "we tried and got nothing": it burns an
  // attempt, extends the streak, records the time-of-day slot, and rests the
  // contact until 6 AM tomorrow. Anything where a human actually answered
  // clears the whole lot.
  const noAnswerLike = (outcome === 'noanswer' || outcome === 'voicemail');
  let lapsedNow = false;

  if (noAnswerLike) {
    const attempts = Number(rec[CL_ATTEMPTS - 1] || 0) + 1;
    const streak = Number(rec[CL_NA_STREAK - 1] || 0) + 1;
    cl.getRange(row, CL_ATTEMPTS).setValue(attempts);
    cl.getRange(row, CL_NA_STREAK).setValue(streak);

    // remember this slot so the next attempt lands at a different time of day
    const slot = currentSlot_(now);
    const tried = String(rec[CL_TRIED_SLOTS - 1] || '').split(',').filter(String);
    if (tried.indexOf(slot) === -1) tried.push(slot);
    cl.getRange(row, CL_TRIED_SLOTS).setValue(tried.join(','));

    // rest until 6 AM tomorrow
    const hold = new Date(now);
    hold.setDate(hold.getDate() + 1);
    hold.setHours(NO_ANSWER_HOLD_HOUR, 0, 0, 0);
    cl.getRange(row, CL_HOLD_UNTIL).setValue(hold);

    const fu = new Date(now);
    fu.setDate(fu.getDate() + NO_ANSWER_FOLLOWUP_DAYS);
    fu.setHours(NO_ANSWER_HOLD_HOUR, 0, 0, 0);
    cl.getRange(row, CL_FOLLOWUP_AT).setValue(fu);

    if (streak >= NO_ANSWER_LAPSE_THRESHOLD) {
      cl.getRange(row, CL_STAGE).setValue(ES_LAPSED);
      lapsedNow = true;
    } else if (String(rec[CL_STAGE - 1] || ES_NEW) === ES_NEW) {
      cl.getRange(row, CL_STAGE).setValue(ES_ACTIVE);
    }
  } else {
    // We reached them (or learned something definite). Reset the machinery.
    if (outcome === 'callback') {
      cl.getRange(row, CL_ATTEMPTS).setValue(Number(rec[CL_ATTEMPTS - 1] || 0) + 1);
    }
    cl.getRange(row, CL_NA_STREAK).setValue(0);
    cl.getRange(row, CL_TRIED_SLOTS).setValue('');
    cl.getRange(row, CL_HOLD_UNTIL).setValue('');
    if (outcome !== 'callback') cl.getRange(row, CL_FOLLOWUP_AT).setValue('');
    cl.getRange(row, CL_STAGE).setValue(ES_ACTIVE);
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

  return { ok: true, status: status, booked: booked, lapsed: lapsedNow };
}

/**
 * Daily sweep: anything untouched for LAPSE_AFTER_DAYS goes to Lapsed so it
 * leaves the main queue. Install on a daily time trigger (or run by hand).
 * Contacts that have never been called are judged from when they were added,
 * which for a fresh import means they get the full window before lapsing.
 */
function markLapsedContacts() {
  const cl = ensureCallListSheet_(SpreadsheetApp.openById(SHEET_ID));
  if (cl.getLastRow() < 2) return { lapsed: 0 };

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - LAPSE_AFTER_DAYS);
  const n = cl.getLastRow() - 1;
  const rows = cl.getRange(2, 1, n, CL_COLS).getValues();
  let count = 0;

  for (let i = 0; i < n; i++) {
    const r = rows[i];
    if (!r[CL_ID - 1]) continue;
    const status = String(r[CL_STATUS - 1] || CS_NEW);
    if ([CS_BOOKED, CS_DECLINED, CS_WRONG, CS_DNC].indexOf(status) !== -1) continue;
    if (String(r[CL_STAGE - 1] || ES_NEW) === ES_LAPSED) continue;

    const last = r[CL_LAST_CALLED - 1];
    if (!(last instanceof Date)) continue;      // never called — leave in the queue
    if (last > cutoff) continue;

    cl.getRange(i + 2, CL_STAGE).setValue(ES_LAPSED);
    count++;
  }
  return { lapsed: count };
}

/** Bring a lapsed contact back into the normal queue. */
function adminUnlapse(token, contactId) {
  requireAdmin_(token);
  const cl = ensureCallListSheet_(SpreadsheetApp.openById(SHEET_ID));
  const row = findCallRow_(cl, contactId);
  if (row === -1) return { ok: false, error: 'No contact with that ID.' };
  cl.getRange(row, CL_STAGE).setValue(ES_ACTIVE);
  cl.getRange(row, CL_NA_STREAK).setValue(0);
  cl.getRange(row, CL_TRIED_SLOTS).setValue('');
  cl.getRange(row, CL_HOLD_UNTIL).setValue('');
  return { ok: true, message: 'Back in the queue.' };
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
 * Heartbeat: refresh my claim on the contact currently open on my screen.
 *
 * Without this, a claim expires after CLAIM_MINUTES of wall-clock time, so any
 * call that runs longer than that releases the contact mid-conversation and a
 * second caller can be handed the same person. The page pings this every
 * minute, so a call of any length stays held; if the tab closes the pings stop
 * and the claim frees itself within the normal window.
 */
function renewClaim(token, contactId) {
  const user = requireStaff_(token);
  if (!contactId) return { ok: false };
  const cl = ensureCallListSheet_(SpreadsheetApp.openById(SHEET_ID));
  const row = findCallRow_(cl, contactId);
  if (row === -1) return { ok: false };
  const holder = String(cl.getRange(row, CL_CLAIMED_BY).getValue() || '');
  if (holder && holder !== user.email) return { ok: false, takenOver: true };
  cl.getRange(row, CL_CLAIMED_BY).setValue(user.email);
  cl.getRange(row, CL_CLAIMED_AT).setValue(new Date());
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
    const now = Date.now();
    let list = loadSkipList_(user.email);
    // Drop anything past its TTL while we are here, so skips decay instead of
    // permanently shrinking this caller's queue.
    list = list.filter(function (e) { return (now - e.t) < SKIP_TTL_DAYS * 86400000; });
    if (!list.some(function (e) { return e.id === String(contactId); })) {
      list.push({ id: String(contactId), t: now });
    }
    if (list.length > 400) list = list.slice(list.length - 400);
    props.setProperty(key, JSON.stringify(list));
  }
  releaseContact(token, contactId);
  logActivity_(user.email, 'skip', 'contact', contactId, 'skipped for ' + SKIP_TTL_DAYS + 'd');
  return { ok: true };
}

/** Raw skip entries [{id, t}], tolerating the old plain-array format. */
function loadSkipList_(email) {
  try {
    const arr = JSON.parse(
      PropertiesService.getScriptProperties().getProperty(skipKey_(email)) || '[]');
    return arr.map(function (e) {
      return (typeof e === 'string') ? { id: e, t: Date.now() } : e;
    }).filter(function (e) { return e && e.id; });
  } catch (err) { return []; }
}

/** Bring this caller's skipped contacts back into their queue. */
function clearMySkips(token) {
  const user = requireStaff_(token);
  const n = Object.keys(loadSkips_(user.email)).length;
  PropertiesService.getScriptProperties().deleteProperty(skipKey_(user.email));
  logActivity_(user.email, 'skips_cleared', 'contact', '', n + ' restored');
  return { ok: true, cleared: n,
           message: n + ' skipped contact' + (n === 1 ? '' : 's') + ' back in your queue.' };
}

/** Admin: clear the skip list for ANY caller (or everyone). */
function adminClearSkips(token, email) {
  requireAdmin_(token);
  const props = PropertiesService.getScriptProperties();
  let cleared = 0;
  if (email) {
    cleared = Object.keys(loadSkips_(email)).length;
    props.deleteProperty(skipKey_(email));
  } else {
    const all = props.getProperties();
    Object.keys(all).forEach(function (k) {
      if (k.indexOf('skip_') === 0) { props.deleteProperty(k); cleared++; }
    });
  }
  return { ok: true, message: 'Cleared skips for ' + (email || 'all callers') + '.' };
}

function skipKey_(email) { return 'skip_' + String(email).toLowerCase(); }

function loadSkips_(email) {
  const now = Date.now();
  const set = {};
  loadSkipList_(email).forEach(function (e) {
    if ((now - e.t) < SKIP_TTL_DAYS * 86400000) set[String(e.id)] = true;
  });
  return set;
}

/** Counts for the header strip. */
function getPhonebankStats(token) {
  requireStaff_(token);
  const cl = ensureCallListSheet_(SpreadsheetApp.openById(SHEET_ID));
  const out = { total: 0, remaining: 0, booked: 0, calledToday: 0,
                lapsed: 0, resting: 0, maxedOut: 0,
                seg: { Recruit: 0, 'New volunteer': 0, 'Active volunteer': 0 } };
  if (cl.getLastRow() < 2) return out;

  const now = new Date();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const slot = currentSlot_(now);

  cl.getRange(2, 1, cl.getLastRow() - 1, CL_COLS).getValues().forEach(function (r) {
    if (!r[CL_ID - 1]) return;
    out.total++;
    const st = String(r[CL_STATUS - 1] || CS_NEW);
    if (st === CS_BOOKED) out.booked++;
    const lc = r[CL_LAST_CALLED - 1];
    if (lc instanceof Date && lc >= today) out.calledToday++;
    if ([CS_BOOKED, CS_DECLINED, CS_WRONG, CS_DNC].indexOf(st) !== -1) return;

    // Lapsed is checked first: a lapsed contact is still reachable in the
    // re-engagement mission, so counting them as "maxed out" would understate
    // how many people are actually still workable.
    if (String(r[CL_STAGE - 1] || ES_NEW) === ES_LAPSED) { out.lapsed++; return; }
    if (Number(r[CL_ATTEMPTS - 1] || 0) >= MAX_ATTEMPTS) { out.maxedOut++; return; }

    const hold = r[CL_HOLD_UNTIL - 1];
    const fu = r[CL_FOLLOWUP_AT - 1];
    const tried = String(r[CL_TRIED_SLOTS - 1] || '');
    if ((hold instanceof Date && hold > now) ||
        (fu instanceof Date && fu > now) ||
        (tried && tried.split(',').indexOf(slot) !== -1)) { out.resting++; return; }

    out.remaining++;
    const sg = String(r[CL_SEGMENT - 1] || SEG_RECRUIT);
    if (out.seg[sg] === undefined) out.seg[sg] = 0;
    out.seg[sg]++;
  });
  return out;
}

/**
 * Why is the queue empty? Reports, per rejection reason, how many contacts each
 * gate discarded for a given mode. When "there are 3 callable people" and the
 * queue says "nobody left", this is the difference.
 */
function debugQueue(token, mode) {
  requireStaff_(token);
  const user = requireStaff_(token);
  mode = mode || MODE_ELIGIBLE;
  const cl = ensureCallListSheet_(SpreadsheetApp.openById(SHEET_ID));
  if (cl.getLastRow() < 2) return { empty: true };

  const rows = cl.getRange(2, 1, cl.getLastRow() - 1, CL_COLS).getValues();
  const now = new Date();
  const skipped = loadSkips_(user.email);
  const slot = currentSlot_(now);
  const wantSegment = segmentForMode_(mode);
  const staleBefore = new Date(now.getTime() - CLAIM_MINUTES * 60000);
  const why = { candidates: 0, noId: 0, skipped: 0, segment: 0, terminal: 0,
                maxed: 0, lapsed: 0, hold: 0, followup: 0, slot: 0, claimed: 0,
                callbackNotDue: 0, mode: mode, wantSegment: wantSegment,
                slotNow: slot, skipCount: Object.keys(skipped).length };
  const samples = [];

  rows.forEach(function (r) {
    if (!r[CL_ID - 1]) { why.noId++; return; }
    if (skipped[String(r[CL_ID - 1])]) { why.skipped++; return; }
    const seg = String(r[CL_SEGMENT - 1] || SEG_RECRUIT);
    if (wantSegment && seg !== wantSegment) { why.segment++; return; }
    const status = String(r[CL_STATUS - 1] || CS_NEW);
    const terminal = (seg === SEG_ACTIVE)
      ? [CS_DECLINED, CS_WRONG, CS_DNC] : [CS_BOOKED, CS_DECLINED, CS_WRONG, CS_DNC];
    if (terminal.indexOf(status) !== -1) { why.terminal++; return; }
    if (Number(r[CL_ATTEMPTS - 1] || 0) >= MAX_ATTEMPTS) { why.maxed++; return; }
    if (String(r[CL_STAGE - 1] || ES_NEW) === ES_LAPSED) { why.lapsed++; return; }
    const hold = r[CL_HOLD_UNTIL - 1];
    if (hold instanceof Date && hold > now) { why.hold++; return; }
    const fu = r[CL_FOLLOWUP_AT - 1];
    if (fu instanceof Date && fu > now) { why.followup++; return; }
    const tried = String(r[CL_TRIED_SLOTS - 1] || '');
    if (tried && tried.split(',').indexOf(slot) !== -1) { why.slot++; return; }
    const cb = String(r[CL_CLAIMED_BY - 1] || '');
    const ca = r[CL_CLAIMED_AT - 1];
    if (cb && cb !== user.email && ca instanceof Date && ca > staleBefore) { why.claimed++; return; }
    if (status === CS_CALLBACK) {
      const due = r[CL_CALLBACK_AT - 1];
      if (due instanceof Date && due > now) { why.callbackNotDue++; return; }
    }
    why.candidates++;
    if (samples.length < 5) {
      samples.push({ id: r[CL_ID - 1], name: r[CL_FIRST - 1] + ' ' + r[CL_LAST - 1],
                     seg: seg, status: status });
    }
  });
  why.samples = samples;
  return why;
}

/**
 * Confirm or thank a shift signup from the calling screen.
 * kind: 'confirm' → marks the signup Confirmed
 *       'thank'   → stamps the thank-you so they drop off that mission
 *       'cancel'  → they can no longer make it; frees the place
 */
function resolveShiftCall(token, signupId, kind) {
  const user = requireStaff_(token);
  if (!signupId) return { ok: false, error: 'No shift on this contact.' };

  if (kind === 'confirm') return setSignupStatus(token, signupId, SS_CONFIRMED);
  if (kind === 'cancel')  return setSignupStatus(token, signupId, SS_CANCELLED);

  if (kind === 'thank') {
    const su = ensureSignupsSheet_(SpreadsheetApp.openById(SHEET_ID));
    const row = findSignupRow_(su, signupId);
    if (row === -1) return { ok: false, error: 'That signup could not be found.' };
    su.getRange(row, SU_THANKED_AT).setValue(new Date());
    logActivity_(user.email, 'thanked', 'signup', signupId, '');
    return { ok: true };
  }
  return { ok: false, error: 'Unknown action.' };
}

/**
 * Find someone by name, phone, or email — for when a person rings back and the
 * caller needs their record immediately rather than waiting for the queue to
 * offer them. Returns each contact plus their shift history.
 */
function searchContacts(token, query) {
  requireStaff_(token);
  const q = String(query || '').trim().toLowerCase();
  if (q.length < 3) return { results: [], message: 'Type at least 3 characters.' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const cl = ensureCallListSheet_(ss);
  if (cl.getLastRow() < 2) return { results: [] };

  const qDigits = q.replace(/\D/g, '');
  const rows = cl.getRange(2, 1, cl.getLastRow() - 1, CL_COLS).getValues();
  const hits = [];

  for (let i = 0; i < rows.length && hits.length < 25; i++) {
    const r = rows[i];
    if (!r[CL_ID - 1]) continue;
    const name = (String(r[CL_FIRST - 1] || '') + ' ' + String(r[CL_LAST - 1] || '')).toLowerCase();
    const email = String(r[CL_EMAIL - 1] || '').toLowerCase();
    const phone = normPhone_(r[CL_PHONE - 1]);
    const match = (name.indexOf(q) !== -1) ||
                  (email && email.indexOf(q) !== -1) ||
                  (qDigits.length >= 4 && phone && phone.indexOf(qDigits) !== -1);
    if (!match) continue;
    hits.push({
      contactId: r[CL_ID - 1],
      first: r[CL_FIRST - 1] || '', last: r[CL_LAST - 1] || '',
      phone: r[CL_PHONE - 1] || '', email: r[CL_EMAIL - 1] || '',
      postal: r[CL_POSTAL - 1] || '', segment: r[CL_SEGMENT - 1] || '',
      status: r[CL_STATUS - 1] || '', stage: r[CL_STAGE - 1] || '',
      attempts: Number(r[CL_ATTEMPTS - 1] || 0),
      notes: r[CL_NOTES - 1] || '',
      shifts: signupsForPhone_(ss, phone)
    });
  }
  return { results: hits };
}

/** Every shift a phone number has signed up for, newest first. */
function signupsForPhone_(ss, phone) {
  const su = ss.getSheetByName(SHIFT_SIGNUPS_TAB);
  if (!su || su.getLastRow() < 2 || !phone) return [];
  const tz = Session.getScriptTimeZone();
  const out = [];
  su.getRange(2, 1, su.getLastRow() - 1, SU_COLS).getValues().forEach(function (r) {
    if (normPhone_(r[SU_PHONE - 1]) !== phone) return;
    const when = shiftStart_(r[SU_DATE - 1], r[SU_TIME - 1]);
    out.push({ signupId: r[SU_ID - 1],
               label: when ? Utilities.formatDate(when, tz, 'EEE MMM d') + ' · ' +
                             prettyTime_(r[SU_TIME - 1]) : '',
               status: r[SU_STATUS - 1] || SS_SCHEDULED,
               sort: when ? when.getTime() : 0 });
  });
  return out.sort(function (a, b) { return b.sort - a.sort; });
}

/** Open a specific contact found via search (claims it like the queue would). */
function openContact(token, contactId) {
  const user = requireStaff_(token);
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const cl = ensureCallListSheet_(ss);
  const row = findCallRow_(cl, contactId);
  if (row === -1) return { done: true, message: 'That contact could not be found.' };

  const r = cl.getRange(row, 1, 1, CL_COLS).getValues()[0];
  cl.getRange(row, CL_CLAIMED_BY).setValue(user.email);
  cl.getRange(row, CL_CLAIMED_AT).setValue(new Date());

  return {
    done: false, contactId: r[CL_ID - 1],
    first: r[CL_FIRST - 1] || '', last: r[CL_LAST - 1] || '',
    phone: r[CL_PHONE - 1] || '', email: r[CL_EMAIL - 1] || '',
    postal: r[CL_POSTAL - 1] || '', source: r[CL_SOURCE - 1] || '',
    status: r[CL_STATUS - 1] || CS_NEW, segment: r[CL_SEGMENT - 1] || SEG_RECRUIT,
    stage: r[CL_STAGE - 1] || ES_NEW, attempts: Number(r[CL_ATTEMPTS - 1] || 0),
    notes: r[CL_NOTES - 1] || '', bookedBefore: String(r[CL_BOOKED - 1] || ''),
    shiftContext: null, mode: 'search', shifts: listOpenShifts(14)
  };
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
    return cl;
  }
  // Migration: the lapse columns were added after the list was first built.
  // Widen an older sheet in place rather than making anyone rebuild it.
  const have = cl.getLastColumn();
  if (have < CL_COLS) {
    cl.getRange(1, 1, 1, CL_COLS).setValues([CALL_LIST_HEADERS]).setFontWeight('bold');
    const n = cl.getLastRow() - 1;
    if (n > 0) {
      // Default every existing contact to New / no streak / nothing tried.
      const fill = [];
      for (let i = 0; i < n; i++) fill.push([ES_NEW, 0, '', '', '', SEG_RECRUIT]);
      cl.getRange(2, CL_STAGE, n, 6).setValues(fill);
    }
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
