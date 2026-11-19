/**
 * Lifecycle deliverability add-on: suppresses customers who reply "unsubscribe"
 * (or similar) to a Lifecycle email. Runs from processBounces each day.
 */

var LC_UNSUB_LABEL = 'Lifecycle/Unsubscribe-processed';

// Subjects of the customer emails a reply can come back on.
var LC_UNSUB_SUBJECTS = ['your tech is showing its age', 'your Lifecycle Report'];

/**
 * Pure. The part of a reply the customer actually wrote: everything before the
 * quoted original, a signature-style "Sent from" line or our own footer text.
 */
function lcReplyText(body) {
  var text = String(body == null ? '' : body).replace(/\r\n?/g, '\n');
  var cuts = [
    /^>/m,
    /^On [\s\S]{0,200}?wrote:\s*$/m,
    /^-{2,}\s*Original Message/im,
    /^_{5,}/m,
    /^From:\s/im,
    /^Sent from my /im,
    /Prefer not to get these emails/i,
    /your tech(nology)? is showing its age/i,
    /You’re receiving this because|You're receiving this because/i
  ];
  var end = text.length;
  for (var i = 0; i < cuts.length; i++) {
    var m = cuts[i].exec(text);
    if (m && m.index < end) end = m.index;
  }
  return text.slice(0, end).trim();
}

/** Pure. True when the customer's own words ask to stop getting these emails. */
function lcIsUnsubscribeRequest(body) {
  var text = lcReplyText(body).slice(0, 500);
  if (!text) return false;
  if (/^\s*(stop|remove|unsubscribe)[\s.!]*$/i.test(text)) return true;
  return /\b(unsubscribe|unsubcribe|opt[\s-]?out|remove (me|my|us|our)|take (me|us) off|stop (e-?mailing|sending|contacting)|no more e-?mails?|(no longer|do not|don'?t|don\u2019t) (wish|want) to (receive|get|hear)|(do not|don'?t|don\u2019t) (e-?mail|contact|send) (me|us))\b/i.test(text);
}

/** Pure. The bare address from a From/To header such as 'Pat <pat@example.com>'. */
function lcHeaderAddresses(header) {
  var out = [];
  var re = /[^\s<>,;"()]+@[^\s<>,;"()]+\.[^\s<>,;"()]+/g;
  var m;
  while ((m = re.exec(String(header == null ? '' : header))) !== null) out.push(lcNormalizeEmail(m[0]));
  return out;
}

function lcGetOrCreateLabel_(name) {
  var parent = name.split('/')[0];
  if (parent !== name && !GmailApp.getUserLabelByName(parent)) GmailApp.createLabel(parent);
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

/**
 * Scans recent replies to Lifecycle emails. When a customer asks to unsubscribe,
 * suppresses both the address that replied and the address we wrote to, and turns
 * their Send? flag off. Safe to run repeatedly.
 */
function processUnsubscribes(lookbackDays) {
  var days = Number(lookbackDays) > 0 ? Number(lookbackDays) : LC.BOUNCE_LOOKBACK_DAYS;
  var label = lcGetOrCreateLabel_(LC_UNSUB_LABEL);
  var labelQuery = LC_UNSUB_LABEL.toLowerCase().replace(/[\/ ]/g, '-');
  var subjects = LC_UNSUB_SUBJECTS.map(function (s) { return '"' + s + '"'; }).join(' OR ');
  var query = 'subject:(' + subjects + ') newer_than:' + days + 'd -label:' + labelQuery;
  var ours = function (address) { return address.split('@')[1] === lcInternalDomain_(); };
  var queue = lcQueueIndex_();
  var added = 0, replies = 0, start = 0, page;

  do {
    page = GmailApp.search(query, start, 100);
    for (var t = 0; t < page.length; t++) {
      var thread = page[t];
      var messages = thread.getMessages();
      var ourRecipients = [];
      var requested = false;
      for (var i = 0; i < messages.length; i++) {
        var message = messages[i];
        var from = lcHeaderAddresses(message.getFrom());
        if (from.length && from.every(ours)) {
          ourRecipients = ourRecipients.concat(lcHeaderAddresses(message.getTo()));
          continue;
        }
        replies++;
        if (!lcIsUnsubscribeRequest(message.getPlainBody())) continue;
        requested = true;
        var targets = from.concat(ourRecipients);
        for (var a = 0; a < targets.length; a++) {
          var address = targets[a];
          if (!address || ours(address)) continue;
          if (lcAddSuppression_(address, 'Unsubscribed (reply)', '', thread.getId())) {
            added++;
            lcLog_('UNSUBSCRIBE', 'Suppressed ' + address + ' (replied to "' + thread.getFirstMessageSubject() + '")');
          }
          var rowNumbers = queue.rows[address];
          if (rowNumbers && queue.sendCol >= 0) {
            for (var r = 0; r < rowNumbers.length; r++) {
              queue.sheet.getRange(rowNumbers[r], queue.sendCol + 1).setValue(false);
            }
          }
        }
      }
      if (requested) thread.addLabel(label);
    }
    start += page.length;
  } while (page.length === 100);

  lcLog_('UNSUBSCRIBE', 'Checked ' + replies + ' customer replies from the last ' + days +
    ' days; newly suppressed: ' + added);
  return added;
}

/** One-time catch-up over older replies. */
function backfillUnsubscribes() {
  return processUnsubscribes(90);
}
