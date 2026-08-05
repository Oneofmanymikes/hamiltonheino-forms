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
  revokeSessionsFor_(email);   // a password change must end existing sessions
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

/** Who is this token? Lets a page restore a session on reload without re-login. */
function whoAmI(token) {
  const s = requireStaff_(token);
  return { ok: true, email: s.email, name: s.name, role: s.role, token: token };
}

/**
 * Bootstrap the one super user. Run ONCE from the Apps Script editor:
 *     createFirstAdmin('you@voteheino.ca', 'Your Name', 'a-good-password')
 * Refuses to run if any admin already exists, so it can't be used to sneak in
 * a second super user later — after this, admins are made in the web console.
 */
function createFirstAdmin(email, name, password) {
  const existing = listStaffUsers().filter(function (u) {
    return String(u.role).toLowerCase() === 'admin';
  });
  if (existing.length) {
    throw new Error('An admin already exists (' + existing[0].email +
                    '). Use the Manage accounts page instead.');
  }
  return createStaffUser(email, name, 'admin', password);
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
  // verifyPassword_ rather than staffLogin: logging in again would mint a second
  // session token that nobody ever holds, and would count toward lockout.
  if (!verifyPassword_(user.email, currentPassword)) {
    return { ok: false, error: 'Your current password is not correct.' };
  }
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

/**
 * Kill every live session belonging to an account.
 *
 * Without this, deactivating someone or resetting their password did nothing to
 * anyone already signed in — their token stayed valid for up to 12 hours, so
 * "remove this person's access" did not actually remove it. Called from
 * deactivate and from every password change.
 */
function revokeSessionsFor_(email) {
  email = String(email || '').toLowerCase();
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const cache = CacheService.getScriptCache();
  let killed = 0;
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('sess_') !== 0) return;
    try {
      const s = JSON.parse(all[k]);
      if (s && String(s.email).toLowerCase() === email) {
        props.deleteProperty(k);
        cache.remove(k);
        killed++;
      }
    } catch (err) { /* malformed — leave for the purge sweep */ }
  });
  return killed;
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

/*** SELF-REGISTRATION ***/
/*
 * Volunteers create their own accounts by entering a shared access code, so
 * nobody has to be in the loop. The code lives in Script Properties under
 * SIGNUP_CODE (Project Settings ▸ Script Properties). NO CODE SET = REGISTRATION
 * CLOSED — that's the safe default, and clearing the property shuts it off again.
 *
 * Self-registered accounts are always role 'staff' (can use the tracker and the
 * phone bank). Only an existing admin can mint another admin.
 */

function signupEnabled() {
  return !!PropertiesService.getScriptProperties().getProperty('SIGNUP_CODE');
}

/** Set/replace the shared access code. Run from the editor, or use the admin page. */
function setSignupCode(code) {
  const props = PropertiesService.getScriptProperties();
  if (!code) { props.deleteProperty('SIGNUP_CODE'); return 'Self-registration is now CLOSED.'; }
  if (String(code).length < 6) throw new Error('Use at least 6 characters.');
  props.setProperty('SIGNUP_CODE', String(code));
  return 'Self-registration is OPEN. Access code: ' + code;
}

/** Create your own account. Returns {ok:true, token, ...} so you land signed in. */
function selfRegister(name, email, password, accessCode) {
  const want = PropertiesService.getScriptProperties().getProperty('SIGNUP_CODE');
  if (!want) return { ok: false, error: 'Account sign-up is closed. Ask the campaign for a login.' };

  email = String(email || '').trim().toLowerCase();
  if (!constantTimeEquals_(String(accessCode || ''), want)) {
    return { ok: false, error: 'That access code is not right. Check with the campaign.' };
  }
  if (!email || email.indexOf('@') === -1) return { ok: false, error: 'Please enter a valid email.' };
  if (!String(name || '').trim())          return { ok: false, error: 'Please enter your name.' };
  if (!password || String(password).length < 8) {
    return { ok: false, error: 'Password must be at least 8 characters.' };
  }

  const sh = ensureStaffSheet_();
  if (findStaffRow_(sh, email) !== -1) {
    return { ok: false, error: 'An account already exists for that email. Try signing in instead.' };
  }

  const salt = Utilities.getUuid().replace(/-/g, '');
  sh.appendRow([email, String(name).trim(), 'staff',
                hashPassword_(password, salt), salt, true, new Date(), '']);
  SpreadsheetApp.flush();
  return staffLogin(email, password);      // sign them straight in
}

/*** ADMIN CONSOLE (role=admin only) ***/

function adminListUsers(token) {
  requireAdmin_(token);
  return listStaffUsers();
}

function adminCreateUser(token, email, name, role, password) {
  requireAdmin_(token);
  try { return { ok: true, message: createStaffUser(email, name, role, password) }; }
  catch (err) { return { ok: false, error: String(err.message || err) }; }
}

function adminResetPassword(token, email, newPassword) {
  requireAdmin_(token);
  try { return { ok: true, message: resetStaffPassword(email, newPassword) }; }
  catch (err) { return { ok: false, error: String(err.message || err) }; }
}

function adminSetActive(token, email, active) {
  const me = requireAdmin_(token);
  email = String(email || '').trim().toLowerCase();
  if (email === me.email && !active) {
    return { ok: false, error: 'You cannot deactivate your own account.' };
  }
  const sh = ensureStaffSheet_();
  const row = findStaffRow_(sh, email);
  if (row === -1) return { ok: false, error: 'No account for ' + email };
  sh.getRange(row, ST_ACTIVE).setValue(!!active);
  const killed = active ? 0 : revokeSessionsFor_(email);
  return { ok: true, message: (active ? 'Reactivated ' : 'Deactivated ') + email +
           (killed ? ' (signed out of ' + killed + ' session' + (killed === 1 ? '' : 's') + ')' : '') };
}

function adminSetRole(token, email, role) {
  const me = requireAdmin_(token);
  email = String(email || '').trim().toLowerCase();
  role = (String(role).toLowerCase() === 'admin') ? 'admin' : 'staff';
  if (email === me.email && role !== 'admin') {
    return { ok: false, error: 'You cannot remove your own admin access.' };
  }
  const sh = ensureStaffSheet_();
  const row = findStaffRow_(sh, email);
  if (row === -1) return { ok: false, error: 'No account for ' + email };
  sh.getRange(row, ST_ROLE).setValue(role);
  return { ok: true, message: email + ' is now ' + role };
}

/** Read/replace the shared access code from the admin page. */
function adminGetSignupCode(token) {
  requireAdmin_(token);
  return { code: PropertiesService.getScriptProperties().getProperty('SIGNUP_CODE') || '' };
}

function adminSetSignupCode(token, code) {
  requireAdmin_(token);
  try { return { ok: true, message: setSignupCode(code) }; }
  catch (err) { return { ok: false, error: String(err.message || err) }; }
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

/**
 * Check a password without any of staffLogin's side effects — no session token,
 * no lockout counter, no last-login stamp. For re-confirming identity mid-session.
 */
function verifyPassword_(email, password) {
  if (!email || !password) return false;
  const sh = ensureStaffSheet_();
  const row = findStaffRow_(sh, String(email).trim().toLowerCase());
  if (row === -1) return false;
  const rec = sh.getRange(row, 1, 1, STAFF_COLS).getValues()[0];
  const expected = String(rec[ST_HASH - 1] || '');
  if (!expected) return false;
  return constantTimeEquals_(expected, hashPassword_(password, String(rec[ST_SALT - 1] || '')));
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
