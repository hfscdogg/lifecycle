/**
 * Lifecycle deliverability add-on: reads bounce notices from Gmail and
 * suppresses the addresses that permanently failed.
 */

/**
 * Pure. Extracts per-recipient results from the text of a bounce notice
 * (raw MIME or plain body). Returns [{recipient, action, status, diagnostic}].
 */
function lcParseBounce(text) {
  var body = String(text || '').replace(/\r\n/g, '\n');
  var results = [];
  var parts = body.split(/^Final-Recipient:/im);
  for (var i = 1; i < parts.length; i++) {
    var block = parts[i].split(/\n\s*\n/)[0];
    var recipient = (block.match(/^\s*[^;\n]*;\s*<?([^\s<>]+@[^\s<>]+)>?/) || [])[1];
    if (!recipient) continue;
    var diagnostic = (block.match(/^Diagnostic-Code:\s*([\s\S]*?)(?=\n[A-Za-z-]+:|$)/im) || [])[1] || '';
    results.push({
      recipient: recipient.toLowerCase(),
      action: ((block.match(/^Action:\s*(\S+)/im) || [])[1] || '').toLowerCase(),
      status: (block.match(/^Status:\s*(\d\.\d{1,3}\.\d{1,3})/im) || [])[1] || '',
      diagnostic: diagnostic.replace(/\s+/g, ' ').trim()
    });
  }
  if (results.length) return results;

  // Notices with no machine-readable part: fall back to the human wording.
  var m = body.match(/Your message (?:wasn't delivered to|to)\s+<?([^\s<>]+@[^\s<>]+)/i);
  if (m) {
    var delayed = /Delivery incomplete|temporary problem|will retry/i.test(body);
    results.push({
      recipient: m[1].replace(/[.,;:]+$/, '').toLowerCase(),
      action: delayed ? 'delayed' : 'failed',
      status: '',
      diagnostic: ''
    });
  }
  return results;
}

/** Pure. Decides whether a parsed bounce result means "stop mailing this address". */
function lcClassifyBounce(bounce) {
  if (bounce.action !== 'failed') return { suppress: false, reason: 'not a final failure' };
  var d = bounce.diagnostic || '';
  var reason;
  if (/Envelope blocked|blocked by|blocklist|blacklist|only accepts messages from/i.test(d)) {
    reason = 'Blocked by recipient';
  } else if (/domain.*(not|n't) (be )?found|NXDOMAIN|DNS Error/i.test(d)) {
    reason = 'Domain not found';
  } else if (/^4\./.test(bounce.status)) {
    reason = 'Recipient server unreachable';
  } else {
    reason = 'Mailbox not found or rejected';
  }
  return { suppress: true, reason: reason };
}

function lcBounceLabel_() {
  var parent = LC.BOUNCE_LABEL.split('/')[0];
  if (parent !== LC.BOUNCE_LABEL && !GmailApp.getUserLabelByName(parent)) GmailApp.createLabel(parent);
  return GmailApp.getUserLabelByName(LC.BOUNCE_LABEL) || GmailApp.createLabel(LC.BOUNCE_LABEL);
}

/** Maps each queue address to its 1-based row numbers, plus the Send? column index. */
function lcQueueIndex_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(LC.QUEUE_SHEET);
  if (!sheet) throw new Error('Sheet not found: ' + LC.QUEUE_SHEET);
  var values = sheet.getDataRange().getValues();
  var emailCol = values[0].indexOf(LC.QUEUE_EMAIL_HEADER);
  var sendCol = values[0].indexOf(LC.QUEUE_SEND_HEADER);
  if (emailCol < 0) throw new Error('No "' + LC.QUEUE_EMAIL_HEADER + '" column in ' + LC.QUEUE_SHEET);
  var rows = {};
  for (var r = 1; r < values.length; r++) {
    var email = lcNormalizeEmail(values[r][emailCol]);
    if (!email) continue;
    (rows[email] = rows[email] || []).push(r + 1);
  }
  return { sheet: sheet, rows: rows, sendCol: sendCol };
}

/**
 * Scans recent bounce notices, suppresses queue addresses that failed for good,
 * and turns their Send? flag off. Safe to run repeatedly.
 */
function processBounces(lookbackDays) {
  var days = Number(lookbackDays) > 0 ? Number(lookbackDays) : LC.BOUNCE_LOOKBACK_DAYS;
  var label = lcBounceLabel_();
  var labelQuery = LC.BOUNCE_LABEL.toLowerCase().replace(/[\/ ]/g, '-');
  var query = 'from:(mailer-daemon OR postmaster) newer_than:' + days + 'd -label:' + labelQuery;
  var queue = lcQueueIndex_();
  var added = 0, scanned = 0, start = 0, page;

  do {
    page = GmailApp.search(query, start, 100);
    for (var t = 0; t < page.length; t++) {
      var thread = page[t];
      var messages = thread.getMessages();
      var sawFinal = false;
      for (var i = 0; i < messages.length; i++) {
        var message = messages[i];
        if (!/mailer-daemon|postmaster/i.test(message.getFrom())) continue;
        scanned++;
        var bounces = lcParseBounce(message.getRawContent());
        if (!bounces.length) bounces = lcParseBounce(message.getPlainBody());
        for (var b = 0; b < bounces.length; b++) {
          var bounce = bounces[b];
          var verdict = lcClassifyBounce(bounce);
          if (!verdict.suppress) continue;
          sawFinal = true;
          var rowNumbers = queue.rows[bounce.recipient];
          if (!rowNumbers) continue; // not a Lifecycle recipient
          if (lcAddSuppression_(bounce.recipient, verdict.reason, bounce.status, thread.getId())) {
            added++;
            lcLog_('BOUNCE', 'Suppressed ' + bounce.recipient + ' (' + verdict.reason +
              (bounce.status ? ', ' + bounce.status : '') + ')');
          }
          if (queue.sendCol >= 0) {
            for (var r = 0; r < rowNumbers.length; r++) {
              queue.sheet.getRange(rowNumbers[r], queue.sendCol + 1).setValue(false);
            }
          }
        }
      }
      // Leave delay-only threads unlabelled so the eventual failure is still picked up.
      if (sawFinal) thread.addLabel(label);
    }
    start += page.length;
  } while (page.length === 100);

  lcLog_('BOUNCE', 'Scanned ' + scanned + ' bounce notices from the last ' + days +
    ' days; newly suppressed: ' + added);
  return added;
}

/** One-time catch-up over older notices. */
function backfillBounces() {
  return processBounces(45);
}

/** Run once to schedule processBounces daily. */
function installBounceTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processBounces') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('processBounces').timeBased().everyDays(1).atHour(6).create();
}
