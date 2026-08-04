/**********************************************************************
 * Heino Doessing — staff logins for the shift tracker
 *
 * Per-person accounts, because the tracker exposes volunteer names, emails,
 * and phone numbers and the web app is deployed to "Anyone".
 *
 * Accounts live in a StaffUsers tab. Passwords are never stored — only a
 * salted, stretched SHA-256 hash. Logging in returns a session token that the
 * page holds in sessionStorage and passes with every call.
 *
 * To create the first account, open the Apps Script editor and run:
 *     createStaffUser('name@example.com', 'Their Name', 'admin', 'a-good-password')
 * Then tell them to change it with changeMyPassword(...) or issue a new one.
 *
 * NOTE ON HASHING: Apps Script has no bcrypt/scrypt. This uses SHA-256 with a
 * per-user random salt and 5,000 iterations of stretching — good enough to make
 * a leaked sheet non-trivial to crack, but not equivalent to bcrypt. Treat the
 * StaffUsers tab as sensitive and keep the Sheet restricted to campaign staff.
 **********************************************************************/

const STAFF_TAB = 'StaffUsers';
const ST_EMAIL = 1, ST_NAME = 2, ST_ROLE = 3, ST_HASH = 4, ST_SALT = 5,
      ST_ACTIVE = 6, ST_CREATED = 7, ST_LAST_LOGIN = 8;
const STAFF_COLS = 8;
const STAFF_HEADERS = ['Email', 'Name', 'Role', 'Password hash', 'Salt',
                       'Active', 'Created', 'Last login'];

const SESSION_HOURS = 12;
const HASH_ITERATIONS = 5000;
const MAX_FAILED_LOGINS = 8;         // per email, per hour

/*** ACCOUNT MANAGEMENT (run these from the Apps Script editor) ***/

/**
 * Create a staff login.
 * role: 'admin' (can manage accounts) or 'staff' (can view the tracker).
 */
function createStaffUser(email, name, role, password) {
  email = String(email || '').trim().toLowerCase();
  if (!email) throw new Error('Email is required.');
  if (!password || String(password).length < 8) {
    throw new Error('Password must be at least 8 characters.');
  }
  role = (String(role || 'staff').toLowerCase() === 'admin') ? 'admin' : 'staff';

  const sh = ensureStaffSheet_();
  if (findStaffRow_(sh, email) !== -1) {
    throw new Error('An account already exists for ' + email + '. Use resetStaffPassword() instead.');
  }
  const salt = Utilities.getUuid().replace(/-/g, '');
  sh.appendRow([email, String(name || '').trim(), role,
                hashPassword_(password, salt), salt, true, new Date(), '']);
  return 'Created ' + role + ' login for ' + email;
}

/** Set a new password for an existing account. */
function resetStaffPassword(email, newPassword) {
  email = String(email || '').trim().toLowerCase();
  if (!newPassword || String(newPassword).length < 8) {
    throw new Error('Password must be at least 8 characters.');
  }
  const sh = ensureStaffSheet_();
  const row = findStaffRow_(sh, email);
  if (row === -1) throw new Error('No account for ' + email);
  const salt = Utilities.getUuid().replace(/-/g, '');
  sh.getRange(row, ST_SALT).setValue(salt);
  sh.getRange(row, ST_HASH).setValue(hashPassword_(newPassword, salt));
  return 'Password reset for ' + email;
}

/** Disable a login without deleting the row (keeps the audit trail). */
function deactivateStaffUser(email) {
  const sh = ensureStaffSheet_();
  const row = findStaffRow_(sh, String(email || '').trim().toLowerCase());
  if (row === -1) throw new Error('No account for ' + email);
  sh.getRange(row, ST_ACTIVE).setValue(false);
  return 'Deactivated ' + email;
}

function listStaffUsers() {
  const sh = ensureStaffSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, STAFF_COLS).getValues().map(function (r) {
    return { email: r[ST_EMAIL - 1], name: r[ST_NAME - 1], role: r[ST_ROLE - 1],
             active: r[ST_ACTIVE - 1], lastLogin: r[ST_LAST_LOGIN - 1] };
  });
}

/*** LOGIN (called from the page) ***/

/**
 * Verify credentials and issue a session token.
 * Returns {ok:true, token, name, role} or {ok:false, error}.
 * The error text is deliberately vague — it never reveals whether the email exists.
 */
function staffLogin(email, password) {
  email = String(email || '').trim().toLowerCase();
  const generic = { ok: false, error: 'Incorrect email or password.' };

  if (!email || !password) return generic;
  if (isLockedOut_(email)) {
    return { ok: false, error: 'Too many failed attempts. Try again in an hour.' };
  }

  const sh = ensureStaffSheet_();
  const row = findStaffRow_(sh, email);
  if (row === -1) { recordFailure_(email); return generic; }

  const rec = sh.getRange(row, 1, 1, STAFF_COLS).getValues()[0];
  if (rec[ST_ACTIVE - 1] !== true && String(rec[ST_ACTIVE - 1]).toUpperCase() !== 'TRUE') {
    return { ok: false, error: 'That account has been deactivated.' };
  }

  const expected = String(rec[ST_HASH - 1] || '');
  const actual = hashPassword_(password, String(rec[ST_SALT - 1] || ''));
  if (!expected || !constantTimeEquals_(expected, actual)) {
    recordFailure_(email);
    return generic;
  }

  clearFailures_(email);
  sh.getRange(row, ST_LAST_LOGIN).setValue(new Date());

  const token = Utilities.getUuid().replace(/-/g, '') +
                Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const session = { email: email, name: String(rec[ST_NAME - 1] || email),
                    role: String(rec[ST_ROLE - 1] || 'staff'),
                    exp: Date.now() + SESSION_HOURS * 3600 * 1000 };
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(session),
                                    Math.min(SESSION_HOURS * 3600, 21600));
  PropertiesService.getScriptProperties().setProperty('sess_' + token, JSON.stringify(session));

  return { ok: true, token: token, name: session.name, role: session.role };
}

function staffLogout(token) {
  if (!token) return { ok: true };
  CacheService.getScriptCache().remove('sess_' + token);
  PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
  return { ok: true };
}

/** Change your own password. Requires a valid session and the current password. */
function changeMyPassword(token, currentPassword, newPassword) {
  const user = requireStaff_(token);
  const check = staffLogin(user.email, currentPassword);
  if (!check.ok) return { ok: false, error: 'Your current password is not correct.' };
  if (!newPassword || String(newPassword).length < 8) {
    return { ok: false, error: 'New password must be at least 8 characters.' };
  }
  resetStaffPassword(user.email, newPassword);
  return { ok: true };
}

/**
 * Throw unless the token maps to a live session. Every staff-only function
 * must call this FIRST — the web app is public, so an uncalled guard is an open door.
 */
function requireStaff_(token) {
  if (!token) throw new Error('Please sign in.');
  const key = 'sess_' + token;
  let raw = CacheService.getScriptCache().get(key);
  if (!raw) raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) throw new Error('Your session has expired. Please sign in again.');

  let s;
  try { s = JSON.parse(raw); } catch (err) { throw new Error('Please sign in again.'); }
  if (!s || !s.exp || Date.now() > s.exp) {
    staffLogout(token);
    throw new Error('Your session has expired. Please sign in again.');
  }
  // Refresh the cache copy so an active user isn't logged out mid-session.
  CacheService.getScriptCache().put(key, raw, Math.min(SESSION_HOURS * 3600, 21600));
  return s;
}

function requireAdmin_(token) {
  const user = requireStaff_(token);
  if (String(user.role) !== 'admin') throw new Error('That action needs an admin account.');
  return user;
}

/** Housekeeping: drop expired session properties. Safe to run on a daily trigger. */
function purgeExpiredSessions() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  let removed = 0;
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('sess_') !== 0) return;
    try {
      const s = JSON.parse(all[k]);
      if (!s.exp || Date.now() > s.exp) { props.deleteProperty(k); removed++; }
    } catch (err) { props.deleteProperty(k); removed++; }
  });
  return { removed: removed };
}

/*** INTERNALS ***/

function ensureStaffSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(STAFF_TAB);
  if (!sh) {
    sh = ss.insertSheet(STAFF_TAB);
    sh.getRange(1, 1, 1, STAFF_COLS).setValues([STAFF_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function findStaffRow_(sh, email) {
  if (sh.getLastRow() < 2) return -1;
  const vals = sh.getRange(2, ST_EMAIL, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === email) return i + 2;
  }
  return -1;
}

function hashPassword_(password, salt) {
  let acc = String(salt) + '|' + String(password);
  for (let i = 0; i < HASH_ITERATIONS; i++) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, acc, Utilities.Charset.UTF_8);
    acc = bytes.map(function (b) {
      return ('0' + (b & 0xff).toString(16)).slice(-2);
    }).join('');
  }
  return acc;
}

/** Compare without leaking timing information about how much matched. */
function constantTimeEquals_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= (a.charCodeAt(i) ^ b.charCodeAt(i));
  return diff === 0;
}

function failKey_(email) { return 'fail_' + email; }

function isLockedOut_(email) {
  const n = Number(CacheService.getScriptCache().get(failKey_(email)) || 0);
  return n >= MAX_FAILED_LOGINS;
}

function recordFailure_(email) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(failKey_(email)) || 0) + 1;
  cache.put(failKey_(email), String(n), 3600);
}

function clearFailures_(email) {
  CacheService.getScriptCache().remove(failKey_(email));
}
