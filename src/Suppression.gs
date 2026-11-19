/**
 * Lifecycle deliverability add-on: shared config, logging and the suppression list.
 *
 * Everything here is prefixed `lc` / `LC` so it can sit beside the existing
 * Lifecycle script without name clashes.
 */

var LC = {
  QUEUE_SHEET: 'EMAIL_QUEUE',
  SUPPRESSION_SHEET: 'SUPPRESSION',
  LOG_SHEET: 'DIAGNOSTIC_LOG',
  QUEUE_EMAIL_HEADER: 'Email',
  QUEUE_SEND_HEADER: 'Send?',
  BOUNCE_LABEL: 'Lifecycle/Bounce-processed',
  BOUNCE_LOOKBACK_DAYS: 3,
  // Rows addressed to the sender's own domain are test records or placeholders.
  SKIP_INTERNAL_DOMAIN: true,
  SUPPRESSION_HEADERS: ['Email', 'Reason', 'DSN Code', 'Date Added', 'Gmail Thread ID']
};

var lcSuppressedCache_ = null;

function lcNormalizeEmail(email) {
  return String(email == null ? '' : email).trim().toLowerCase();
}

function lcLog_(category, message) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(LC.LOG_SHEET);
  if (sheet) sheet.appendRow([new Date(), category, message]);
  Logger.log(category + ': ' + message);
}

function lcGetSuppressionSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(LC.SUPPRESSION_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LC.SUPPRESSION_SHEET);
    sheet.appendRow(LC.SUPPRESSION_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function lcLoadSuppressed_() {
  if (lcSuppressedCache_) return lcSuppressedCache_;
  var sheet = lcGetSuppressionSheet_();
  var map = {};
  var last = sheet.getLastRow();
  if (last > 1) {
    var values = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var email = lcNormalizeEmail(values[i][0]);
      if (email) map[email] = true;
    }
  }
  lcSuppressedCache_ = map;
  return map;
}

/** True when the address is on the SUPPRESSION tab. Delete its row there to re-enable it. */
function lcIsSuppressed(email) {
  return lcLoadSuppressed_()[lcNormalizeEmail(email)] === true;
}

/** Adds an address to the SUPPRESSION tab. Returns false when it was already there. */
function lcAddSuppression_(email, reason, code, threadId) {
  var normalized = lcNormalizeEmail(email);
  if (!normalized) return false;
  var map = lcLoadSuppressed_();
  if (map[normalized]) return false;
  lcGetSuppressionSheet_().appendRow([normalized, reason || '', code || '', new Date(), threadId || '']);
  map[normalized] = true;
  return true;
}
