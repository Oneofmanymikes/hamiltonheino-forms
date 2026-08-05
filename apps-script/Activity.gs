/**********************************************************************
 * Heino Doessing — audit trail
 *
 * Every state-changing action gets a row. With volunteers sharing one dataset,
 * "who marked this person as Do Not Call?" and "who cancelled that shift?" are
 * questions that WILL be asked, and without a log the answer is a shrug.
 *
 * Deliberately append-only and best-effort: a logging failure must never break
 * the action it was recording.
 **********************************************************************/

const ACTIVITY_TAB = 'ActivityLog';
const AL_WHEN = 1, AL_WHO = 2, AL_ACTION = 3, AL_TARGET_TYPE = 4,
      AL_TARGET = 5, AL_DETAIL = 6;
const AL_COLS = 6;
const ACTIVITY_HEADERS = ['When', 'Who', 'Action', 'Target type', 'Target', 'Detail'];

// Keep the tab from growing without bound. Oldest rows are trimmed once it
// passes this; a campaign only ever needs the recent past.
const ACTIVITY_MAX_ROWS = 20000;

/**
 * Record an action. Never throws — logging must not be able to fail a write
 * that already succeeded.
 *   who        email of the actor, or 'system' for triggers
 *   action     short verb, e.g. 'call_outcome', 'shift_cancel'
 *   targetType 'contact' | 'signup' | 'shift' | 'account' | 'list'
 */
function logActivity_(who, action, targetType, target, detail) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(ACTIVITY_TAB);
    if (!sh) {
      sh = ss.insertSheet(ACTIVITY_TAB);
      sh.getRange(1, 1, 1, AL_COLS).setValues([ACTIVITY_HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), String(who || 'system'), String(action || ''),
                  String(targetType || ''), String(target || ''),
                  sanitizeCell_(String(detail == null ? '' : detail))]);

    if (sh.getLastRow() - 1 > ACTIVITY_MAX_ROWS) {
      sh.deleteRows(2, 500);            // trim the oldest in blocks
    }
  } catch (err) {
    // swallow — see the comment above
  }
}

/** Recent activity, newest first. Admin only: it names who did what. */
function getActivityLog(token, limit) {
  requireAdmin_(token);
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(ACTIVITY_TAB);
  if (!sh || sh.getLastRow() < 2) return [];

  const tz = Session.getScriptTimeZone();
  const n = Math.min(Number(limit) || 100, sh.getLastRow() - 1);
  const start = sh.getLastRow() - n + 1;
  return sh.getRange(start, 1, n, AL_COLS).getValues().reverse().map(function (r) {
    return {
      when: (r[AL_WHEN - 1] instanceof Date)
        ? Utilities.formatDate(r[AL_WHEN - 1], tz, 'EEE MMM d, h:mm a') : String(r[AL_WHEN - 1] || ''),
      who: r[AL_WHO - 1], action: r[AL_ACTION - 1],
      targetType: r[AL_TARGET_TYPE - 1], target: r[AL_TARGET - 1],
      detail: r[AL_DETAIL - 1]
    };
  });
}
