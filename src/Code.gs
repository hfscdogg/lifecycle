/**
 * Livewire Lifecycle™ - Lifecycle Marketing Automation
 * Version: 2.7.9 (Livewire Brand Identity)
 */

const CONFIG = {
  BRAND_NAME: "Livewire Lifecycle™",
  BOOKING_URL: "https://getlivewire.zohobookings.com/#/3987741000011670058",
  MIN_VALUE: 100,
  // true = customer lifecycle emails are logged to DIAGNOSTIC_LOG as DRY_RUN instead of sent,
  // and SentTimestamp is not written. Set to false to send for real.
  DRY_RUN: false,
  // A customer gets at most one lifecycle email in this many months (tracked on SEND_HISTORY).
  REPEAT_MONTHS: 12,
  SHEETS: {
    SALES: "RAW_SALES",
    ITEMS: "RAW_ITEMS",
    CUSTOMERS: "RAW_CUSTOMERS",
    RULES: "LIFECYCLE_RULES",
    METRICS: "AGG_METRICS",
    QUEUE: "EMAIL_QUEUE",
    LOG: "DIAGNOSTIC_LOG"
  }
};

function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'monthlyLifecycleRun') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendReviewDigest').timeBased().onMonthDay(28).atHour(8).create();
  ScriptApp.newTrigger('monthlyLifecycleRun').timeBased().onMonthDay(1).atHour(8).create();
  Logger.log('Triggers registered: sendPreviewDigest on day 28, monthlyLifecycleRun on day 1.');
}

function monthlyLifecycleRun() {
  // Save SentTimestamps to SEND_HISTORY before the rebuild erases them
  lcSyncSendHistory_();
  // Snapshot any exclusions Henry set via the preview digest opt-out links
  var excluded = getExcludedEmails();
  // Rebuild the queue from current data
  processLifecycleData();
  // Re-apply exclusions (processLifecycleData clears the sheet)
  reapplyExclusions(excluded);
  // Put back the send date for anyone emailed within CONFIG.REPEAT_MONTHS
  lcRestoreRecentSends_();
  // Auto-send to everyone not excluded
  var result = sendLifecycleEmails();
  // Email Henry a confirmation
  sendConfirmationDigest(result.sent, result.excluded, result.total);
  // Clear EXCLUDED flags for next month
  clearExcludedFlags();
}

function clearTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert('All triggers cleared.');
}

function sendReviewDigest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qSheet = ss.getSheetByName(CONFIG.SHEETS.QUEUE);
  if (!qSheet) return;
  var data = qSheet.getDataRange().getValues();
  var pending = [];
  var oldestAge = 0;
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r[4] === true && !r[5]) {
      var itemData = r[6] || '';
      var ages = [];
      if (itemData) { itemData.split('; ').forEach(function(item) { var parts = item.split('|'); if (parts[1]) { var age = new Date().getFullYear() - parseInt(parts[1].trim().slice(-4)); if (age > oldestAge) oldestAge = age; ages.push(age + ' yrs'); } }); }
      pending.push({ name: r[0], email: r[1], count: r[2], ages: ages.join(', ') });
    }
  }
  var monthLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');
  var ssUrl = ss.getUrl() + '#gid=' + qSheet.getSheetId();
  if (pending.length === 0) { GmailApp.sendEmail('henry@getlivewire.com', 'Lifecycle Monthly Review — ' + monthLabel + ' — no pending sends', 'No customers are currently marked for sending.'); return; }
  var rows = pending.map(function(p) { return '<tr><td style="padding:6px 10px;border-bottom:1px solid #eee;">' + p.name + '</td><td style="padding:6px 10px;border-bottom:1px solid #eee;">' + p.email + '</td><td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:center;">' + p.count + '</td><td style="padding:6px 10px;border-bottom:1px solid #eee;">' + (p.ages || '—') + '</td></tr>'; }).join('');
  var html = '<div style="font-family:sans-serif;max-width:700px;color:#333;"><h2 style="color:#F47920;">Lifecycle Monthly Review — ' + monthLabel + '</h2><p><strong>' + pending.length + ' customer(s)</strong> in queue. Oldest system age: <strong>' + oldestAge + ' years</strong>.</p><table style="width:100%;border-collapse:collapse;font-size:14px;"><tr style="background:#f8f8f8;"><th style="padding:8px 10px;text-align:left;">Customer</th><th style="padding:8px 10px;text-align:left;">Email</th><th style="padding:8px 10px;text-align:center;">Systems</th><th style="padding:8px 10px;text-align:left;">Ages</th></tr>' + rows + '</table><p style="margin-top:24px;"><a href="' + ssUrl + '" style="background:#F47920;color:#fff;padding:12px 24px;text-decoration:none;border-radius:8px;font-weight:bold;">Open EMAIL_QUEUE</a></p><p style="font-size:12px;color:#999;margin-top:30px;">To send emails: use Livewire Lifecycle menu — Send Pending Emails (Selected).</p></div>';
  GmailApp.sendEmail('henry@getlivewire.com', 'Lifecycle Monthly Review — ' + monthLabel + ' queue ready', 'See HTML version.', { htmlBody: html, name: 'Livewire Lifecycle Engine' });
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Livewire Lifecycle')
    .addItem('1. Initialize/Refresh Data', 'processLifecycleData')
    .addItem('2. Send Pending Emails (Selected)', 'sendLifecycleEmails')
    .addSeparator()
    .addItem('3. Schedule Monthly Automation', 'setupMonthlyTrigger')
    .addItem('4. Clear Monthly Schedule', 'clearTriggers')
    .addSeparator()
    .addSubMenu(ui.createMenu('Lifecycle Visits')
      .addItem('Send Tech Checklist (Selected Row)', 'sendTechChecklistFromMenu')
      .addItem('Send Client Report (Selected Row)', 'sendClientReportFromMenu')
      .addItem('Run Daily Visit Queue Now', 'processVisitQueue')
      .addItem('Schedule Daily Visit Trigger', 'setupVisitTrigger'))
    .addToUi();
}

function processLifecycleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initTabs(ss);
  
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  logSheet.clear().appendRow(["Timestamp", "Category", "Message"]);
  const log = function(cat, msg) { logSheet.appendRow([new Date(), cat, msg]); };

  log("START", "Beginning Lifecycle Data Processing v2.7.8...");

  const salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
  const rulesSheet = ss.getSheetByName(CONFIG.SHEETS.RULES);
  const customerSheet = ss.getSheetByName(CONFIG.SHEETS.CUSTOMERS);
  
  const salesData = salesSheet.getDataRange().getValues();
  if (salesData.length < 2) {
    log("ERROR", "RAW_SALES is empty or too short.");
    return;
  }
  
  // 1. Get Rules
  const rules = rulesSheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] && String(r[0]).trim() !== "")
    .map(r => ({ 
      pattern: new RegExp(String(r[0]).trim(), 'i'), 
      years: parseFloat(r[1]) || 0,
      friendlyName: String(r[2] || "").trim()
    }));
  log("INFO", "Loaded " + rules.length + " rules.");

  // 2. Identify Sales Headers (Scoring System)
  let salesHeadRow = -1;
  let sHead = [];
  const salesKeywords = ["part", "item", "name", "customer", "date", "amount"];
  
  for (let i = 0; i < Math.min(salesData.length, 100); i++) {
    const r = salesData[i];
    const score = r.reduce((acc, cell) => {
      const c = String(cell).toLowerCase();
      return acc + (salesKeywords.some(k => c.indexOf(k) > -1) ? 1 : 0);
    }, 0);
    if (score >= 3) { // Requires at least 3 keyword matches
      salesHeadRow = i;
      sHead = r;
      break;
    }
  }

  if (salesHeadRow === -1) {
    log("ERROR", "Could not find a valid header row in RAW_SALES. Check column names.");
    return;
  }
  log("INFO", "Found Sales Header at row " + (salesHeadRow + 1));

  const findIdx = (patterns, priority) => {
    let idx = sHead.findIndex(h => String(h).toLowerCase() === priority.toLowerCase());
    if (idx === -1) idx = sHead.findIndex(h => patterns.some(p => new RegExp(p, "i").test(String(h).trim())));
    return idx;
  };

  const idx = {
    part: findIdx(["part", "product"], "parts"),
    date: findIdx(["date", "install"], "date"),
    name: findIdx(["name", "customer"], "name"),
    amount: findIdx(["amount", "total", "balance"], "amount")
  };
  log("INFO", "Sales Column Mapping: " + JSON.stringify(idx));

  // Debug first few data rows
  log("DEBUG", "First 5 rows of Sales data (starting at row " + (salesHeadRow + 2) + "):");
  for (let i = salesHeadRow + 1; i < Math.min(salesData.length, salesHeadRow + 6); i++) {
     log("DEBUG", "Row " + i + ": " + JSON.stringify(salesData[i]));
  }

  // 3. Map Customers
  const customerMap = getDeepCustomerMap(customerSheet, log);
  log("INFO", "Customer Email Map: " + Object.keys(customerMap).length + " entries found.");

  const aggregation = {};
  const partCache = {};
  let curName = null;
  let curDate = null;
  let curPart = null;
  const today = new Date();
  const nextYear = new Date();
  nextYear.setDate(today.getDate() + 365);

  for (let i = salesHeadRow + 1; i < salesData.length; i++) {
    const row = salesData[i];
    
    const rName = String(row[idx.name] || "").trim();
    if (rName && rName !== "" && !/total|name|date|type/i.test(rName)) {
      curName = rName;
    }

    const rDate = row[idx.date];
    if (rDate instanceof Date) curDate = rDate;
    else if (rDate && String(rDate).trim() !== "") {
      const d = new Date(rDate);
      if (!isNaN(d.getTime())) curDate = d;
    }

    const rPart = String(row[idx.part] || "").trim();
    if (rPart && rPart !== "" && !/total|undefined|desc|part|item/i.test(rPart)) {
      curPart = rPart;
    }

    if (!curName || !curDate) continue;

    const amtStr = String(row[idx.amount] || "0").replace(/[$,()]/g, "").trim();
    const amt = Math.abs(parseFloat(amtStr) || 0);
    if (amt < CONFIG.MIN_VALUE) continue;

    const rawPart = curPart || rPart;
    if (!rawPart || rawPart === "" || /total|undefined|desc/i.test(rawPart)) continue;

    const customerKey = curName.split(/[:\\-]/)[0].trim();
    if (customerKey === "" || /name|total|customer/i.test(customerKey)) continue;

    if (!aggregation[customerKey]) {
      aggregation[customerKey] = { 
        email: customerMap[customerKey] || "", 
        itemsDue: 0, 
        itemsApproaching: 0, 
        itemsReview: 0, 
        nextRefresh: null,
        itemDetails: {} 
      };
    }

    if (partCache[rawPart] === undefined) {
      const match = rules.find(r => r.pattern.test(rawPart));
      partCache[rawPart] = match ? { years: match.years, friendly: match.friendlyName || rawPart } : null;
    }

    const meta = partCache[rawPart];
    if (!meta) {
      aggregation[customerKey].itemsReview++;
      continue;
    }

    const refDate = new Date(curDate);
    refDate.setFullYear(curDate.getFullYear() + Math.floor(meta.years));

    if (refDate < today || refDate <= nextYear) {
      if (refDate < today) aggregation[customerKey].itemsDue++;
      else aggregation[customerKey].itemsApproaching++;
      
      if (!aggregation[customerKey].itemDetails[meta.friendly] || refDate < aggregation[customerKey].itemDetails[meta.friendly]) {
        aggregation[customerKey].itemDetails[meta.friendly] = refDate;
      }
    }

    if (!aggregation[customerKey].nextRefresh || refDate < aggregation[customerKey].nextRefresh) {
      aggregation[customerKey].nextRefresh = refDate;
    }
  }

  log("INFO", "Aggregation complete. Found " + Object.keys(aggregation).length + " unique customers.");
  writeResults(ss, aggregation);
  log("SUCCESS", "Processing Complete.");
}

function getDeepCustomerMap(sheet, log) {
  const data = sheet.getDataRange().getValues();
  let headIdx = -1;
  let head = [];
  const custKeywords = ["email", "e-mail", "mail", "name", "full name", "customer", "contact", "client"];

  for (let i = 0; i < Math.min(data.length, 100); i++) {
    const r = data[i];
    const score = r.reduce((acc, cell) => {
      const c = String(cell).toLowerCase().replace(/[^a-z0-9]/g, '');
      return acc + (custKeywords.some(k => c.indexOf(k.replace(/[^a-z0-9]/g, '')) > -1) ? 1 : 0);
    }, 0);
    
    // If we find both a name-like and an email-like column, or just a very strong email match
    const hasEmail = r.some(cell => /email|e-mail|mail/i.test(String(cell)));
    if (score >= 2 || (score >= 1 && hasEmail)) {
      headIdx = i;
      head = r;
      break;
    }
  }

  if (headIdx === -1) {
    log("WARN", "Could not find headers in RAW_CUSTOMERS. Search yielded 0 matches.");
    return {};
  }
  log("INFO", "Found Customer Header at row " + (headIdx + 1));

  const nIdx = head.findIndex(h => /full name|name|customer|client/i.test(String(h).trim()));
  const eIdx = head.findIndex(h => /email|e-mail|mail/i.test(String(h).trim()));
  
  const map = {};
  for (let i = headIdx + 1; i < data.length; i++) {
    const row = data[i];
    const n = String(row[nIdx] || "").split(/[:\\-]/)[0].trim();
    const e = String(row[eIdx] || "").trim();
    if (n && e && e.indexOf('@') > -1) map[n] = e;
  }
  return map;
}

function writeResults(ss, aggregation) {
  const mSheet = ss.getSheetByName(CONFIG.SHEETS.METRICS);
  const qSheet = ss.getSheetByName(CONFIG.SHEETS.QUEUE);
  const mRows = [];
  const qRows = [];

  for (const k in aggregation) {
    const e = aggregation[k];
    let status = e.itemsDue > 0 ? "⚠️ ACTION" : (e.itemsApproaching > 0 ? "📅 SOON" : "✅ HEALTHY");
    const dStr = e.nextRefresh ? Utilities.formatDate(e.nextRefresh, Session.getScriptTimeZone(), "yyyy-MM-dd") : "N/A";
    const iStrs = [];
    for (const n in e.itemDetails) {
      iStrs.push(n + "|" + Utilities.formatDate(e.itemDetails[n], Session.getScriptTimeZone(), "MMM yyyy"));
    }
    mRows.push([k, e.email, e.itemsDue, e.itemsApproaching, e.itemsReview, dStr, status, e.nextRefresh ? e.nextRefresh.getTime() : 4102444800000]);
    if (e.itemsDue > 0 || e.itemsApproaching > 0) {
      qRows.push([k, e.email || "MISSING EMAIL", e.itemsDue, dStr, false, "", iStrs.join("; "), e.nextRefresh ? e.nextRefresh.getTime() : 4102444800000]);
    }
  }

  mRows.sort((a, b) => a[7] - b[7]);
  qRows.sort((a, b) => a[7] - b[7]);

  const mFinal = [["Customer", "Email", "Due Assets", "Soon Assets", "Untracked", "Next Review", "Status"]];

  mRows.forEach(r => mFinal.push(r.slice(0, 7)));
  mSheet.clear().getRange(1, 1, Math.max(1, mFinal.length), 7).setValues(mFinal);

  const qFinal = [["Customer", "Email", "Due Count", "Soonest Review", "Send?", "SentTimestamp", "Premium Items & Dates"]];
  qRows.forEach(r => qFinal.push(r.slice(0, 7)));
  qSheet.clear();
  if (qFinal.length > 1) {
    qSheet.getRange(1, 1, qFinal.length, 7).setValues(qFinal);
    qSheet.getRange(2, 5, qFinal.length - 1, 1).insertCheckboxes();
  } else {
    qSheet.getRange(1, 1, 1, 7).setValues(qFinal);
  }
}

function sendLifecycleEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(CONFIG.SHEETS.QUEUE);
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  const log = function(cat, msg) { logSheet.appendRow([new Date(), cat, msg]); };
  lcSyncSendHistory_();
  const data = qSheet.getDataRange().getValues();
  let count = 0;
    let excludedCount = 0;
  let skippedCount = 0;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r[4] === true && !r[5] && r[1].indexOf('@') > -1) {
      const recent = lcRecentSendDate_(r[1]);
      if (recent) {
        lcLog_('SKIP', r[0] + ' <' + r[1] + '>: emailed ' + Utilities.formatDate(recent, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ', within ' + CONFIG.REPEAT_MONTHS + ' months');
        qSheet.getRange(i + 1, 6).setValue(recent);
        skippedCount++;
        continue;
      }
      const check = lcPreSendCheck(r[1]);
      if (!check.ok) {
        lcLog_('SKIP', r[0] + ' <' + r[1] + '>: ' + check.reason);
        skippedCount++;
        continue;
      }
      log("RAW_ITEMS_VALUE", "Customer: " + r[0] + " | Raw Items (r[6]): " + r[6]);
      const html = getEmailTemplate(r[0], r[2], r[3], r[6], log);
      if (lcSendLifecycleEmail_(r[1], lcGreetingName(r[0], ''), html, log)) {
        const sentAt = new Date();
        qSheet.getRange(i + 1, 6).setValue(sentAt);
        lcRecordSend_(r[1], sentAt);
        count++;
      }
    } else if (r[5]) { excludedCount++; }
  }
    Logger.log("Sent " + count + " emails. Skipped " + skippedCount + "." + (CONFIG.DRY_RUN ? " (DRY_RUN: nothing sent)" : ""));
      return { sent: count, excluded: excludedCount, total: data.length - 1, skipped: skippedCount };

}

/**
 * Sends one customer lifecycle email (subject, HTML and plain-text part).
 * firstName may be '' (from lcGreetingName), which gives the no-name subject.
 * Returns true if the email was sent, false when CONFIG.DRY_RUN only logged it.
 */
function lcSendLifecycleEmail_(email, firstName, html, log) {
  var subject = firstName
    ? firstName + ' — your tech is showing its age'
    : 'Your tech is showing its age';
  if (CONFIG.DRY_RUN) {
    log('DRY_RUN', 'Would send to ' + email + ' | Subject: ' + subject);
    return false;
  }
  GmailApp.sendEmail(email, subject, lcHtmlToText(html), { htmlBody: html, name: 'Livewire Concierge' });
  return true;
}

/*
 * SEND_HISTORY: one row per address with the last lifecycle email date. Unlike
 * EMAIL_QUEUE it is never rebuilt, so it is what enforces CONFIG.REPEAT_MONTHS.
 */
var LC_HISTORY_SHEET = 'SEND_HISTORY';
var lcHistoryCache_ = null; // { email: { row: n, date: Date } }

function lcHistorySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LC_HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LC_HISTORY_SHEET);
    sheet.appendRow(['Email', 'Last Sent']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function lcLoadHistory_() {
  if (lcHistoryCache_) return lcHistoryCache_;
  var sheet = lcHistorySheet_();
  var map = {};
  var last = sheet.getLastRow();
  if (last > 1) {
    var values = sheet.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      var email = lcNormalizeEmail(values[i][0]);
      if (email && values[i][1] instanceof Date) map[email] = { row: i + 2, date: values[i][1] };
    }
  }
  lcHistoryCache_ = map;
  return map;
}

/** Saves a send date for the address, keeping whichever date is newer. */
function lcRecordSend_(email, date) {
  var key = lcNormalizeEmail(email);
  if (!key || !(date instanceof Date)) return;
  var map = lcLoadHistory_();
  var entry = map[key];
  if (entry) {
    if (entry.date >= date) return;
    lcHistorySheet_().getRange(entry.row, 2).setValue(date);
    entry.date = date;
  } else {
    var sheet = lcHistorySheet_();
    sheet.appendRow([key, date]);
    map[key] = { row: sheet.getLastRow(), date: date };
  }
}

/** The last send date if it is within CONFIG.REPEAT_MONTHS, otherwise null. */
function lcRecentSendDate_(email) {
  var entry = lcLoadHistory_()[lcNormalizeEmail(email)];
  if (!entry) return null;
  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - CONFIG.REPEAT_MONTHS);
  return entry.date > cutoff ? entry.date : null;
}

/** Copies every SentTimestamp in EMAIL_QUEUE into SEND_HISTORY. */
function lcSyncSendHistory_() {
  var qSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.QUEUE);
  if (!qSheet) return;
  var data = qSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] instanceof Date && data[i][1]) lcRecordSend_(data[i][1], data[i][5]);
  }
}

/** After a rebuild, writes the last send date back into SentTimestamp for recently emailed rows. */
function lcRestoreRecentSends_() {
  var qSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.QUEUE);
  if (!qSheet) return;
  var data = qSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] || !data[i][1]) continue;
    var recent = lcRecentSendDate_(data[i][1]);
    if (recent) qSheet.getRange(i + 1, 6).setValue(recent);
  }
}

function initTabs(ss) {
  Object.values(CONFIG.SHEETS).forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
}

function getEmailTemplate(name, count, date, itemData, log) {
  var firstName = lcGreetingName(name, '');
  var headline = firstName
    ? firstName + ' — your technology is showing its age.'
    : 'Your technology is showing its age.';
  var itemsListHtml = "";
  var THUMBNAILS = {
    "Networking": "https://images.unsplash.com/photo-1544197150-b99a580bb7a8?w=400",
    "Network": "https://images.unsplash.com/photo-1544197150-b99a580bb7a8?w=400",
    "Router": "https://images.unsplash.com/photo-1544197150-b99a580bb7a8?w=400",
    "WiFi": "https://images.unsplash.com/photo-1544197150-b99a580bb7a8?w=400",
    "Home Automation": "https://images.unsplash.com/photo-1558002038-1055907df827?w=400",
    "Automation": "https://images.unsplash.com/photo-1558002038-1055907df827?w=400",
    "Control": "https://images.unsplash.com/photo-1558002038-1055907df827?w=400",
    "Power Protection": "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=400",
    "Power": "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=400",
    "UPS": "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=400",
    "Surge": "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=400",
    "TV & Displays": "https://images.unsplash.com/photo-1593784991095-a205069470b6?w=400",
    "Video": "https://images.unsplash.com/photo-1593784991095-a205069470b6?w=400",
    "TV": "https://images.unsplash.com/photo-1593784991095-a205069470b6?w=400",
    "Display": "https://images.unsplash.com/photo-1593784991095-a205069470b6?w=400",
    "Displays": "https://images.unsplash.com/photo-1593784991095-a205069470b6?w=400",
    "Audio": "https://images.unsplash.com/photo-1545454675-3531b543be5d?w=400",
    "AV": "https://images.unsplash.com/photo-1545454675-3531b543be5d?w=400",
    "Speaker": "https://images.unsplash.com/photo-1545454675-3531b543be5d?w=400",
    "Speakers": "https://images.unsplash.com/photo-1545454675-3531b543be5d?w=400",
    "Sound": "https://images.unsplash.com/photo-1545454675-3531b543be5d?w=400",
    "Security Cameras": "https://images.unsplash.com/photo-1557597774-9d273605dfa9?w=400",
    "Camera": "https://images.unsplash.com/photo-1557597774-9d273605dfa9?w=400",
    "Cameras": "https://images.unsplash.com/photo-1557597774-9d273605dfa9?w=400",
    "Security": "https://images.unsplash.com/photo-1557597774-9d273605dfa9?w=400",
    "Surveillance": "https://images.unsplash.com/photo-1557597774-9d273605dfa9?w=400",
    "NVR": "https://images.unsplash.com/photo-1557597774-9d273605dfa9?w=400",
    "Lighting": "https://images.unsplash.com/photo-1565814329452-e1efa11c5b89?w=400",
    "Lights": "https://images.unsplash.com/photo-1565814329452-e1efa11c5b89?w=400",
    "Light": "https://images.unsplash.com/photo-1565814329452-e1efa11c5b89?w=400",
    "Shade": "https://images.unsplash.com/photo-1565814329452-e1efa11c5b89?w=400",
    "Shades": "https://images.unsplash.com/photo-1565814329452-e1efa11c5b89?w=400",
    "Shade Control": "https://images.unsplash.com/photo-1565814329452-e1efa11c5b89?w=400",
    "Projectors": "https://images.unsplash.com/photo-1478720568477-152d9b164e26?w=400",
    "Projector": "https://images.unsplash.com/photo-1478720568477-152d9b164e26?w=400",
    "Touch Panels": "https://images.unsplash.com/photo-1726607424562-62cf93236dd8?w=400",
    "Touch Panel": "https://images.unsplash.com/photo-1726607424562-62cf93236dd8?w=400",
    "Remote": "https://images.unsplash.com/photo-1726607424562-62cf93236dd8?w=400",
    "Keypad": "https://images.unsplash.com/photo-1726607424562-62cf93236dd8?w=400",
    "Climate Control": "https://images.unsplash.com/photo-1655194827229-a1d3192b533e?w=400",
    "Climate": "https://images.unsplash.com/photo-1655194827229-a1d3192b533e?w=400",
    "HVAC": "https://images.unsplash.com/photo-1655194827229-a1d3192b533e?w=400",
    "Thermostat": "https://images.unsplash.com/photo-1655194827229-a1d3192b533e?w=400"
  };

  var STATUS_COLORS = {
    "Upgrade": "#c0392b",
    "Update": "#d4850a",
    "Review": "#e8a020",
    "Good": "#27ae60"
  };

  if (itemData) {
    var itemsArr = itemData.split("; ");
    var cards = [];
    itemsArr.slice(0, 8).forEach(function(item) {
      var parts = item.split("|");
      var categoryName = parts[0] ? parts[0].trim() : "";
      var installYear = parts[1] ? parts[1].trim() : "";
      var statusLabel = parts[2] ? parts[2].trim() : "Review";
      var imgUrl = THUMBNAILS[categoryName];
      var matchType = "EXACT";
      if (!imgUrl) {
        var keys = Object.keys(THUMBNAILS);
        for (var k = 0; k < keys.length; k++) {
          var key = keys[k];
          if (categoryName.toLowerCase().indexOf(key.toLowerCase()) > -1 || key.toLowerCase().indexOf(categoryName.toLowerCase()) > -1) {
            imgUrl = THUMBNAILS[key];
            matchType = "FUZZY (" + key + ")";
            break;
          }
        }
      }
      if (!imgUrl) {
        imgUrl = "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=520&h=200&fit=crop&auto=format&q=80";
        matchType = "NO_MATCH";
      } else {
        imgUrl = imgUrl.replace(/w=\d+/, "w=520");
        if (imgUrl.indexOf("&h=") === -1 && imgUrl.indexOf("?") > -1) {
          imgUrl += "&h=200&fit=crop&auto=format&q=80";
        }
      }
      if (log) log("IMAGE_LOOKUP", "Item: '" + categoryName + "' | Match: " + matchType);
      var badgeColor = STATUS_COLORS[statusLabel] || "#e8a020";
      var cardHtml = '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff;border-radius:3px;overflow:hidden">' +
        '<tbody><tr><td style="padding:0">' +
        '<img src="' + imgUrl + '" width="100%" height="120" alt="' + categoryName + '" style="display:block;object-fit:cover;border:0">' +
        '</td></tr>' +
        '<tr><td style="padding:14px 14px 12px 14px">' +
        '<p style="margin:0 0 2px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:bold;color:#111111">' + categoryName + '</p>' +
        '<p style="margin:0 0 10px 0;font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#888888">Installed ' + installYear + '</p>' +
        '<table cellpadding="0" cellspacing="0" border="0"><tbody><tr>' +
        '<td style="background-color:' + badgeColor + ';border-radius:2px;padding:4px 10px">' +
        '<span style="font-family:Arial,Helvetica,sans-serif;font-size:10px;font-weight:bold;color:#ffffff;letter-spacing:0.08em;text-transform:uppercase">' + statusLabel + '</span>' +
        '</td></tr></tbody></table>' +
        '</td></tr></tbody></table>';
      cards.push(cardHtml);
    });
    for (var r = 0; r < cards.length; r += 2) {
      var padding = r === 0 ? "8px 28px 6px 28px" : "6px 28px " + (r + 2 >= cards.length ? "20px" : "6px") + " 28px";
      itemsListHtml += '<tr><td style="background-color:#e8e7e2;padding:' + padding + '">' +
        '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tbody><tr>' +
        '<td width="48%" style="vertical-align:top;padding-right:8px">' + cards[r] + '</td>';
      if (r + 1 < cards.length) {
        itemsListHtml += '<td width="48%" style="vertical-align:top;padding-left:8px">' + cards[r + 1] + '</td>';
      } else {
        itemsListHtml += '<td width="48%"></td>';
      }
      itemsListHtml += '</tr></tbody></table></td></tr>';
    }
  }

  return '<div style="margin:0;padding:0;background-color:#f4f4f0;font-family:Arial,Helvetica,sans-serif">' +
    '<div style="display:none;max-height:0;overflow:hidden;font-size:1px;line-height:1px;color:#f4f4f0">Your installed technology</div>' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f4f4f0"><tbody><tr><td align="center" style="padding:20px 0">' +
    '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%"><tbody>' +
    '<tr><td style="background-color:#111111;padding:18px 28px;border-radius:4px 4px 0 0">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tbody><tr>' +
    '<td style="vertical-align:middle"><img src="https://www.getlivewire.com/wp-content/uploads/2023/01/logosec.png" width="140" height="auto" alt="Livewire" style="display:block;border:0;outline:none"></td>' +
    '<td align="right" style="vertical-align:middle"><span style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#555555;font-variant:small-caps;letter-spacing:0.08em;white-space:nowrap">Lifecycle\u2122</span></td>' +
    '</tr></tbody></table></td></tr>' +
    '<tr><td style="background-color:#111111;padding:0 28px 14px 28px;border-bottom:1px solid #2a2a2a">' +
    '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#666666;letter-spacing:0.06em;text-transform:uppercase">Your installed technology</p></td></tr>' +
    '<tr><td style="background-color:#111111;padding:40px 28px 44px 28px">' +
    '<h1 style="margin:0;font-family:Georgia,serif;font-size:30px;line-height:1.25;color:#f5f5f0;font-weight:normal">' + headline + '</h1></td></tr>' +
    '<tr><td style="background-color:#ffffff;padding:32px 28px 24px 28px">' +
    '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.65;color:#333333">We installed and have maintained your systems for years. We wanted to give you an honest picture of where things stand today.</p></td></tr>' +
    '<tr><td style="background-color:#e8e7e2;padding:22px 28px 10px 28px">' +
    '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#888888;letter-spacing:0.12em;text-transform:uppercase;font-weight:bold">Your Installed Systems</p></td></tr>' +
    itemsListHtml +
    '<tr><td style="background-color:#ffffff;padding:32px 28px 8px 28px">' +
    '<p style="margin:0 0 18px 0;font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.65;color:#333333">Technology ages faster than most people realize. The average lifespan of residential AV and control systems is 7\u201310 years \u2014 and several of your systems are approaching or past that threshold.</p>' +
    '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.65;color:#333333">The good news: a single consultation is all it takes to understand your options. We\u2019ll walk you through what\u2019s worth keeping, what should be updated, and what the investment looks like.</p></td></tr>' +
    '<tr><td style="background-color:#ffffff;padding:24px 28px 32px 28px">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tbody><tr>' +
    '<td style="border-left:3px solid #e8a020;padding:6px 0 6px 20px">' +
    '<p style="margin:0;font-family:Georgia,serif;font-size:16px;line-height:1.55;color:#555555;font-style:italic">\u201cMost of our clients who do a Lifecycle Review end up spending less than they expected \u2014 because we catch things early.\u201d</p>' +
    '</td></tr></tbody></table></td></tr>' +
    '<tr><td style="background-color:#111111;padding:38px 28px 40px 28px;text-align:center">' +
    '<p style="margin:0 0 8px 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#888888;letter-spacing:0.08em;text-transform:uppercase">Ready to see where things stand?</p>' +
    '<p style="margin:0 0 26px 0;font-family:Georgia,serif;font-size:20px;color:#f5f5f0;line-height:1.4">A Lifecycle Review takes about an hour. We\u2019ll walk through every system in your home or business.</p>' +
    '<table cellpadding="0" cellspacing="0" border="0" align="center"><tbody><tr>' +
    '<td style="background-color:#e8a020;border-radius:3px">' +
    '<a href="https://getlivewire.zohobookings.com/#/3987741000011670058" style="display:inline-block;font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;color:#ffffff;text-decoration:none;padding:14px 30px;letter-spacing:0.04em" target="_blank">Reserve a Consultation \u2192</a>' +
    '</td></tr></tbody></table></td></tr>' +
    '<tr><td style="background-color:#ffffff;padding:28px 28px 24px 28px;border-top:1px solid #eeeeee">' +
    '<p style="margin:0 0 4px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#333333;font-weight:bold">Henry Clifford</p>' +
    '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#666666;line-height:1.6">Livewire \u00b7 Richmond, VA \u00b7 804.937.9001</p></td></tr>' +
    '<tr><td style="background-color:#f4f4f0;padding:18px 28px;border-radius:0 0 4px 4px;border-top:1px solid #e0e0da">' +
    '<p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#999999;line-height:1.5;text-align:center">You\u2019re receiving this because Livewire services your home technology. Questions? Reply to this email.</p>' +
    '<p style="margin:6px 0 0 0;font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#999999;line-height:1.5;text-align:center">Prefer not to get these emails? Reply with \u201cunsubscribe\u201d and we\u2019ll take you off the list.</p></td></tr>' +
    '</tbody></table></td></tr></tbody></table></div>';
}

function getExcludedEmails() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
      var qSheet = ss.getSheetByName(CONFIG.SHEETS.QUEUE);
        if (!qSheet) return [];
          var data = qSheet.getDataRange().getValues();
            var excluded = [];
              for (var i = 1; i < data.length; i++) {
                  // Send dates are not exclusions; they are kept on SEND_HISTORY instead.
                  if (data[i][5] && !(data[i][5] instanceof Date)) {
                        excluded.push(data[i][1]);
                            }
                              }
                                return excluded;
                                }


function reapplyExclusions(excluded) {
    if (!excluded || excluded.length === 0) return;
      var ss = SpreadsheetApp.getActiveSpreadsheet();
        var qSheet = ss.getSheetByName(CONFIG.SHEETS.QUEUE);
          if (!qSheet) return;
            var data = qSheet.getDataRange().getValues();
              for (var i = 1; i < data.length; i++) {
                  if (excluded.indexOf(data[i][1]) !== -1) {
                        qSheet.getRange(i + 1, 6).setValue(true);
                            }
                              }
                              }

                              function sendConfirmationDigest(sent, excluded, total) {
                                var monthLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');
                                  var subject = 'Lifecycle Run Complete — ' + monthLabel + ' — ' + sent + ' sent, ' + excluded + ' excluded';
                                    var body = 'Monthly lifecycle run complete for ' + monthLabel + '.\n\nTotal in queue: ' + total + '\nEmails sent: ' + sent + '\nExcluded: ' + excluded + '\n\nCheck the EMAIL_QUEUE sheet for details.';
                                      GmailApp.sendEmail('henry@getlivewire.com', subject, body, { name: 'Livewire Lifecycle Engine' });
                                      }

                                      function clearExcludedFlags() {
                                        var ss = SpreadsheetApp.getActiveSpreadsheet();
                                          var qSheet = ss.getSheetByName(CONFIG.SHEETS.QUEUE);
                                            if (!qSheet) return;
                                              var data = qSheet.getDataRange().getValues();
                                                for (var i = 1; i < data.length; i++) {
                                                    // Clear exclusion flags only; leave send dates in place.
                                                    if (data[i][5] && !(data[i][5] instanceof Date)) {
                                                          qSheet.getRange(i + 1, 6).setValue(false);
                                                              }
                                                                }
                                                                }

/**
 * sendCatchupBatch - sends a daily batch of catch-up lifecycle emails
 * to customers whose Soonest Review date predates 2026-04-01 and
 * who have not yet been sent an email (ExclusionFlag is empty).
 * Automatically disables its own daily trigger once caught up,
 * defined as: no unsent customer has a Soonest Review date older
 * than 30 days from today.
 */
function sendCatchupBatch() {
  var BATCH_SIZE = 20;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qSheet = ss.getSheetByName(CONFIG.SHEETS.QUEUE);
  var logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  var log = function(cat, msg) { logSheet.appendRow([new Date(), cat, msg]); };

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var catchupCutoff = new Date('2026-04-01');
  var caughtUpThreshold = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);

  lcSyncSendHistory_();
  var data = qSheet.getDataRange().getValues();

  // Collect all rows still needing a catch-up send
  var pending = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var soonestReview = r[3] ? new Date(r[3]) : null;
    var hasEmail = r[1] && typeof r[1] === 'string' && r[1].indexOf('@') > -1;
    var notExcluded = !r[5];
    if (soonestReview && soonestReview < catchupCutoff && hasEmail && notExcluded) {
      // Emailed within CONFIG.REPEAT_MONTHS: restore the send date and leave it out.
      var recent = lcRecentSendDate_(r[1]);
      if (recent) {
        qSheet.getRange(i + 1, 6).setValue(recent);
        continue;
      }
      // Undeliverable rows are left out so they cannot block the batch or the caught-up check.
      var check = lcPreSendCheck(r[1]);
      if (!check.ok) {
        lcLog_('SKIP', r[0] + ' <' + r[1] + '>: ' + check.reason);
        continue;
      }
      pending.push({ idx: i, r: r, date: soonestReview });
    }
  }

  // Sort oldest Soonest Review first
  pending.sort(function(a, b) { return a.date - b.date; });

  // Check if caught up: nothing overdue by more than 30 days remains
  var stillOverdue = pending.filter(function(p) { return p.date < caughtUpThreshold; });
  if (pending.length === 0 || stillOverdue.length === 0) {
    log('CATCHUP', 'Caught up! Remaining unsent (' + pending.length + ') all within 30 days. Disabling trigger.');
    Logger.log('Catchup complete, trigger will be removed.');
    stopCatchupTrigger_();
    return;
  }

  // Send this batch
  var batch = pending.slice(0, BATCH_SIZE);
  var sent = 0;
  for (var j = 0; j < batch.length; j++) {
    var item = batch[j];
    var r = item.r;
    var rowIndex = item.idx;
    try {
      var html = getEmailTemplate(r[0], r[2], r[3], r[6], log);
      if (lcSendLifecycleEmail_(r[1], lcGreetingName(r[0], ''), html, log)) {
        var sentAt = new Date();
        qSheet.getRange(rowIndex + 1, 6).setValue(sentAt);
        lcRecordSend_(r[1], sentAt);
        sent++;
      }
    } catch (e) {
      log('CATCHUP_ERROR', 'Row ' + (rowIndex + 1) + ' (' + r[0] + '): ' + e.message);
    }
  }

  var remaining = pending.length - sent;
  log('CATCHUP', 'Batch sent: ' + sent + ' | Remaining backlog: ' + remaining);
  Logger.log('Catchup batch done. Sent: ' + sent + ', Remaining: ' + remaining);
}

/** Remove any existing daily catchup triggers. */
function stopCatchupTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendCatchupBatch') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log('Catchup trigger deleted.');
    }
  }
}

/**
 * Call once to register the daily 9am catch-up trigger.
 * Safe to re-run: removes any existing catchup trigger first.
 */
function setupCatchupTrigger() {
  stopCatchupTrigger_();
  ScriptApp.newTrigger('sendCatchupBatch')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  Logger.log('Daily catchup trigger created (9am local time).');
}


function updateApolloEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('EMAIL_QUEUE');
  var emailMap = {
  "fwallace@mrpapts.com": "fwallace@mccannrealty.com",
  "Dave.DeGonia@aftonchemical.com": "degonia@aftonchemical.com",
  "mrfalcone@stephengould.com": "mfalcone@aeglobal.com",
  "bobby.fauntleroy@parksterlingbank.com": "bobby.fauntleroy@southstatebank.com",
  "wisorfamily@mac.com": "dwisor@treatingpain.com",
  "jisaacs@mcvh-vcu.edu": "jonathan@jonathanisaacs.com",
  "waynekshepherd@gmail.com": "wayne.shepherd@ivwatch.com",
  "Edmyrtetus@goodmanallen.com": "edmyrtetus@eckertseamans.com",
  "richardnlloyd@hotmail.com": "rick.lloyd@perdue.com",
  "ericsvedberg@gmail.com": "eric@europeanautowerks.com",
  "slove@midatlanticins.com": "steve.love@novatae.com",
  "ScottM@NormanWF.com": "scottm@normanusa.com",
  "James.Pulsifer@altria.com": "james.r.pulsifer@altria.com",
  "ansolon.ebj@gmail.com": "ejones@cleverdevices.com",
  "jtuck@c21option1.com": "jtuck@o1realty.com",
  "djdomster@gmail.com": "ddomster@coremgmt.com",
  "cshaff@u.washington.edu": "cshaff@gastonengineering.com",
  "annem_srikanth@yahoo.com": "srikanth.annem@vdot.virginia.gov",
  "Timothy_Stevens@bose.com": "tims@justvideowalls.com",
  "frederickthompson2@gmail.com": "thompsonf@darden.virginia.edu",
  "nipermd@hotmail.com": "niraj.patel@carecentrix.com",
  "aharris@gmail.com": "austin@firehook.com",
  "glee2179@gmail.com": "glee@arubanetworks.com",
  "pbliley@comcast.net": "mark.bliley@ferguson.com",
  "darrynappleton@gmail.com": "dappleton@vacardio.com",
  "damisnyder@yahoo.com": "dami.snyder@chiesiusa.com",
  "devinlipawsky@hotmail.com": "devin.lipawsky@wellsfargo.com",
  "sis63@comcast.net": "sisiskys@uww.edu",
  "joestepp@comcast.net": "joe.stepp@atlfa.org",
  "locke@lockepartin.com": "richardl@lockepartin.com",
  "fishqueen2007@gmail.com": "kristi.turner@vcuhealth.org",
  "srs609@comcast.net": "susansnyder@srmfre.com",
  "rleggett@anchor-financial.com": "ryan@buoyhub.com",
  "rminnick@manusmed.com": "bm@apiarymedical.com",
  "mmccrimmon@reynolds.edu": "mmccrimmon@pamunkeylibrary.org",
  "megan.baynes@hotmail.com": "megan.baynes@ezcorp.com",
  "jim@schneidercb.com": "jim@schncons.com",
  "james.w.beamer@dom.com": "james.b@dominionenergy.com",
  "bradcummings@yahoo.com": "bradc@bradfordcummingsinc.com",
  "john.watt@verizon.net": "johna@chesterfield.gov",
  "jon@legacynavigator.com": "jon.stanley@cseonline.net",
  "diane@getchristensenlaw.com": "diane@ccplawyers.com",
  "jsteele@hf-law.com": "jsteele@hirschlerlaw.com",
  "dfontaine5@cox.net": "doug.fountaine@hii.com",
  "RKaplan@KV-legal.com": "bkaplan@kapadvisors.com",
  "andre@getlivewire.com": "asaunders@vicom-corp.com",
  "khoi.b.do@gmail.com": "khoi@eagleoffshore.org",
  "paige.wolk@spinnakerconsultinggroup.com": "probert@kbsgc.com",
  "ryan.ganley@swedishmatch.com": "ryan.ganley@smpmi.com",
  "theoabel@gmail.com": "theo.abel@martinagency.com",
  "krmontague@aol.com": "kmontague@montaguecustomhomes.com",
  "Lanky.Hicks@signaturefamilywealth.com": "ivor.massey@brownadvisory.com",
  "mlkrouse@comcast.net": "michael.krouse@dell.com",
  "susanhockman@rrpcusa.org": "askus@rrpcusa.org",
  "rcostanz@vcu.edu": "richard.costanzo@vcuhealth.org",
  "egregory2072@gmail.com": "eric@heftywiley.com",
  "jonathen.bingham@gmail.com": "jonathen.bingham@dla.mil",
  "courtneyworrell1@gmail.com": "cworrell@vaetlaw.com",
  "sirjeffington@gmail.com": "jmason@repirva.com",
  "eddavis94@gmail.com": "evan.davis@vita.virginia.gov",
  "chutchens@holladaycorp.com": "chutchens@thirdlinecapital.com",
  "pokdoc@gmail.com": "andrew.poklepovic@vcuhealth.org",
  "jonathan@hlc.law": "jonathan@jonathanjhalperin.com",
  "sandigrivat@gmail.com": "sandi@sandigrivat.com",
  "Moyer_me@yahoo.com": "mamoyer@eqt.com",
  "ashlee.randall@live.com": "ashlee.randall@richmondnational.com",
  "jedvan@gmail.com": "jed_vanichkachorn@bshsi.org",
  "michael.mccullough@genworth.com": "mccullough.michael@genworth.com",
  "brad@getlivewire.com": "brad@onevisionresources.com",
  "ron.villacarillo@gmail.com": "ron@dottedline.agency",
  "ski6buoys@gmail.com": "jmiller@cvwltd.com",
  "bakerda1@gmail.com": "david.baker@medtronic.com",
  "mike@outdoordreamsva.com": "greg@outdoordreamsva.com",
  "TOBryan@serviceins.com": "tobryan@acrisure.com",
  "pradhanr82@gmail.com": "rohit.pradhan@indivior.com",
  "Hagee90@gmail.com": "kmiller@odec.com",
  "leegarner@yahoo.com": "lee_garner@carmax.com",
  "Patrick.Duke@cbre.com": "patrick.duke@turntown.com",
  "agropper@northbeachvascular.com": "adam.gropper@vividvascular.com",
  "BRoyals317@aol.com": "john.royals@connectamerica.com",
  "allanhamby@gmail.com": "ashamby@sentara.com",
  "ralphgiffin@hotmail.com": "ralph.giffin@pronghornresort.com",
  "ds120869@gmail.com": "dave.starr@tributebaking.com",
  "mark@buckheads.com": "mark@rooseveltrva.com",
  "jud@getlivewire.com": "judson.wilton@kaeser.com",
  "roberttadams@gmail.com": "rob@10next.biz",
  "ricksorce@yahoo.com": "rick.thompson@gettyimages.com",
  "deads@liverangewater.com": "deads@galleryresidential.com",
  "jamie.w.morris@gmail.com": "jamie.morris@markel.com",
  "john.gavin@gmail.com": "john.cmfc@nuveen.com",
  "eric.edwards@mac.com": "eric.edwards@kaleo.com",
  "cheryn.clark@gmail.com": "cclark@nvrinc.com",
  "jayryan30@hotmail.com": "jcox@tandemadvisors.com",
  "marclwise@aol.com": "mwise@rxante.com",
  "hicks480@gmail.com": "stacy@getlivewire.com",
  "SJMAHughes@gmail.com": "scott.hughes2@bd.com",
  "teague.avey@gmail.com": "tavey@brownadvisory.com",
  "dkurzawa@hotmail.com": "david.kurzawa@capitalone.com",
  "jhmpsu99@gmail.com": "john.h.milligan@truist.com",
  "murraywright@mac.com": "murray@murraywright.com",
  "dave.daniele@gmail.com": "david.daniele@vcuhealth.org",
  "rheaume11@comcast.net": "lrheaume@eaglebankcorp.com",
  "pmriley12@verizon.net": "paul.riley@sharecare.com",
  "karen.berkness@joynerfineproperties.com": "karen.berkness@thesteelegroupsir.com",
  "mattkipp10@gmail.com": "matt.kipp@wellsfargo.com",
  "jjkascsak@hotmail.com": "john.kascsak@qtsdatacenters.com",
  "ansell@usa.net": "jon.ansell@merlinio.com",
  "pl_smmns@yahoo.com": "simmonsp@rushenterprises.com",
  "Dyochelson@yahoo.com": "dan.yochelson@richmondfed.org",
  "crmaze5@gmail.com": "cmaze@pescanovausa.com",
  "hillegassdev@gmail.com": "john.hillegass@capitalone.com",
  "christina@7hillscre.com": "cjeffries@cresa.com",
  "gsimos@barbermartin.com": "greg@elevationadvertising.com",
  "astraebel@cormanconstruction.com": "astraebel@kokosing.biz",
  "whuffman@capitalsquare1031.com": "whuffman@capitalsq.com",
  "QBBurgess@yahoo.com": "quinton.burgess@goarmy.com",
  "alexandra.casagrande@gmail.com": "acasagrande@balanceuw.com",
  "matty_dolan@yahoo.com": "matt.dolan@milehighdatasolutions.com",
  "scm0104@gmail.com": "jeffhammer@northwesternmutual.com"
  };
  
  var data = sheet.getRange('B2:B816').getValues();
  var updated = 0;
  
  for (var i = 0; i < data.length; i++) {
    var currentEmail = (data[i][0] || '').toString().trim();
    if (emailMap.hasOwnProperty(currentEmail.toLowerCase())) {
      var newEmail = emailMap[currentEmail.toLowerCase()];
      sheet.getRange(i + 2, 2).setValue(newEmail);
      updated++;
    }
  }
  
  SpreadsheetApp.getUi().alert('Done! Updated ' + updated + ' email addresses via Apollo enrichment.');
}


// === VISIT MODULE START — paste the rest at the bottom of your script ===

const VISITS_TAB = "LIFECYCLE_VISITS";
const VISIT_LEAD_HOURS = 24;    // send checklist this many hours BEFORE visit
const VISIT_FOLLOW_HOURS = 24;  // send report this many hours AFTER visit

/**
 * Ensure the LIFECYCLE_VISITS tab exists with the correct header row.
 * Columns: Customer, Email, Address, VisitDate, Tech Name, Tech Email,
 *          Tech Notes, Checklist Sent, Report Sent, Status, Booking Notes
 */
function ensureVisitsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(VISITS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(VISITS_TAB);
    sheet.getRange(1, 1, 1, 11).setValues([[
      "Customer", "Email", "Address", "VisitDate", "Tech Name", "Tech Email",
      "Tech Notes", "Checklist Sent", "Report Sent", "Status", "Booking Notes"
    ]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff");
    sheet.setColumnWidths(1, 11, 140);
    sheet.setColumnWidth(7, 320);  // Tech Notes wider
    sheet.setColumnWidth(11, 320); // Booking Notes wider
  }
  return sheet;
}

/**
 * Build a map of friendlyName -> lifespanYears from LIFECYCLE_RULES.
 * Used to compute the actual install year from the end-of-useful-life date
 * stored in EMAIL_QUEUE column G.
 */
function getRulesByFriendlyName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rSheet = ss.getSheetByName(CONFIG.SHEETS.RULES);
  if (!rSheet) return {};
  var data = rSheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var friendly = String(data[i][2] || "").trim();
    var years = parseFloat(data[i][1]) || 0;
    if (friendly && years > 0 && !map[friendly]) {
      map[friendly] = years;
    }
  }
  return map;
}

/**
 * Normalize a name to "first last" lowercase, handling both the booking-form
 * format ("John Smith") and the EMAIL_QUEUE format ("Smith, John").
 */
function normalizeName(name) {
  name = String(name || "").trim();
  if (!name) return "";
  if (name.indexOf(",") > -1) {
    var parts = name.split(",");
    name = (parts[1] || "").trim() + " " + (parts[0] || "").trim();
  }
  return name.toLowerCase().replace(/\s+/g, " ");
}

/**
 * Look up a customer's lifecycle data from EMAIL_QUEUE.
 * Matches on EMAIL first (most reliable join key, since the booking form
 * captures email and a single free-text Name field that won't match the
 * "Last, First" format in EMAIL_QUEUE). Falls back to a normalized name match.
 * Returns { name, email, items: [{ category, installYear, status, refDate }] }
 * or null if not found. The installYear is the ACTUAL year of install,
 * derived by subtracting the rule's lifespan from the end-of-life date.
 */
function getCustomerLifecycleData(customerName, customerEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qSheet = ss.getSheetByName(CONFIG.SHEETS.QUEUE);
  if (!qSheet) return null;
  var rulesMap = getRulesByFriendlyName();
  var data = qSheet.getDataRange().getValues();
  var emailTarget = String(customerEmail || "").trim().toLowerCase();
  var nameTarget = normalizeName(customerName);

  var matchRow = -1;
  // Pass 1: email match
  if (emailTarget) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1] || "").trim().toLowerCase() === emailTarget) { matchRow = i; break; }
    }
  }
  // Pass 2: normalized name match
  if (matchRow === -1 && nameTarget) {
    for (var j = 1; j < data.length; j++) {
      if (normalizeName(data[j][0]) === nameTarget) { matchRow = j; break; }
    }
  }
  if (matchRow === -1) return null;

  var itemData = data[matchRow][6] || "";
  var items = [];
  var today = new Date();
  itemData.split("; ").forEach(function(s) {
    if (!s) return;
    var parts = s.split("|");
    if (!parts[0] || !parts[1]) return;
    var category = parts[0].trim();
    var dateStr = parts[1].trim();
    var refDate = new Date(dateStr); // "MMM yyyy" parses
    if (isNaN(refDate.getTime())) return;
    var status = (refDate < today) ? "Upgrade" : "Update";
    var ruleYears = rulesMap[category] || 7; // 7-year default if no rule found
    var installDate = new Date(refDate.getTime());
    installDate.setFullYear(refDate.getFullYear() - Math.floor(ruleYears));
    var installYear = installDate.getFullYear().toString();
    items.push({ category: category, installYear: installYear, status: status, refDate: refDate });
  });
  return { name: data[matchRow][0], email: data[matchRow][1], items: items };
}

/**
 * Send the dynamic tech checklist for a given LIFECYCLE_VISITS row.
 */
function sendTechChecklist(rowIndex) {
  var sheet = ensureVisitsTab();
  var row = sheet.getRange(rowIndex, 1, 1, 11).getValues()[0];
  var customer = row[0], clientEmail = row[1], address = row[2], visitDate = row[3];
  var techName = row[4], techEmail = row[5], bookingNotes = row[10];
  if (!customer || !techEmail) return { ok: false, error: "Missing customer or tech email." };
  var custData = getCustomerLifecycleData(customer, clientEmail);
  if (!custData) return { ok: false, error: "Customer not found in EMAIL_QUEUE: " + customer + " / " + clientEmail };
  if (custData.items.length === 0) return { ok: false, error: "No flagged systems for: " + customer };

  var html = buildTechChecklistHtml(custData, customer, address, visitDate, techName, bookingNotes);
  var subject = "Lifecycle Visit · " + customer + " · " +
    (visitDate ? Utilities.formatDate(new Date(visitDate), Session.getScriptTimeZone(), "EEE MMM d, h:mm a") : "TBD");
  if (CONFIG.DRY_RUN) {
    lcLog_('DRY_RUN', 'Would send tech checklist to ' + techEmail + ' | Subject: ' + subject);
    return { ok: false, error: "DRY_RUN is on: checklist logged to DIAGNOSTIC_LOG, not sent." };
  }
  GmailApp.sendEmail(techEmail, subject, lcHtmlToText(html), {
    htmlBody: html,
    name: "Livewire Lifecycle Engine",
    bcc: "henry@getlivewire.com"
  });
  sheet.getRange(rowIndex, 8).setValue(new Date());
  sheet.getRange(rowIndex, 10).setValue("Checklist Sent");
  return { ok: true };
}

/**
 * Send the client-facing Lifecycle Report for a given LIFECYCLE_VISITS row.
 */
function sendClientReport(rowIndex) {
  var sheet = ensureVisitsTab();
  var row = sheet.getRange(rowIndex, 1, 1, 11).getValues()[0];
  var customer = row[0], clientEmail = row[1], visitDate = row[3];
  var techName = row[4], techNotes = row[6];
  if (!customer || !clientEmail) return { ok: false, error: "Missing customer or client email." };
  var custData = getCustomerLifecycleData(customer, clientEmail);
  if (!custData) return { ok: false, error: "Customer not found in EMAIL_QUEUE: " + customer + " / " + clientEmail };
  var check = lcPreSendCheck(clientEmail);
  if (!check.ok) {
    lcLog_('SKIP', 'Client report for ' + customer + ' <' + clientEmail + '>: ' + check.reason);
    return { ok: false, error: "Not sent to " + clientEmail + ": " + check.reason };
  }

  var html = buildClientReportHtml(custData, customer, visitDate, techName, techNotes);
  var firstName = lcGreetingName(customer, '');
  var subject = firstName ? firstName + " — your Lifecycle Report" : "Your Lifecycle Report";
  if (CONFIG.DRY_RUN) {
    lcLog_('DRY_RUN', 'Would send client report to ' + clientEmail + ' | Subject: ' + subject);
    return { ok: false, error: "DRY_RUN is on: report logged to DIAGNOSTIC_LOG, not sent." };
  }
  GmailApp.sendEmail(clientEmail, subject, lcHtmlToText(html), {
    htmlBody: html,
    name: "Livewire Concierge",
    bcc: "henry@getlivewire.com,3sk9a8_4c69hy5@mails3.zohocrm.com" // second address = Zoho CRM dropbox, auto-files report to the contact record
  });
  sheet.getRange(rowIndex, 9).setValue(new Date());
  sheet.getRange(rowIndex, 10).setValue("Complete");
  return { ok: true };
}

/**
 * Daily 7am sweep. For every row in LIFECYCLE_VISITS:
 *   - Send tech checklist if visit is within VISIT_LEAD_HOURS and not yet sent
 *   - Send client report if visit was VISIT_FOLLOW_HOURS ago and not yet sent
 */
function processVisitQueue() {
  var sheet = ensureVisitsTab();
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var sent = { checklist: 0, report: 0, skipped: 0 };
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || !row[3]) continue;
    var visitDate = new Date(row[3]);
    if (isNaN(visitDate.getTime())) { sent.skipped++; continue; }
    var hoursToVisit = (visitDate - now) / (1000 * 60 * 60);
    var hoursSinceVisit = (now - visitDate) / (1000 * 60 * 60);
    if (!row[7] && hoursToVisit > 0 && hoursToVisit <= VISIT_LEAD_HOURS) {
      var r1 = sendTechChecklist(i + 1);
      if (r1.ok) sent.checklist++;
    }
    if (!row[8] && hoursSinceVisit >= VISIT_FOLLOW_HOURS) {
      var r2 = sendClientReport(i + 1);
      if (r2.ok) sent.report++;
    }
  }
  Logger.log("Visit queue: " + sent.checklist + " checklists, " + sent.report + " reports, " + sent.skipped + " skipped.");
  return sent;
}

/**
 * Register the daily 7am trigger. Safe to re-run (removes existing first).
 */
function setupVisitTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "processVisitQueue") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("processVisitQueue").timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert("Daily visit queue trigger registered (runs at 7am).");
}

/**
 * Menu wrappers for manual override on a selected row.
 */
function sendTechChecklistFromMenu() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== VISITS_TAB) {
    SpreadsheetApp.getUi().alert("Switch to the " + VISITS_TAB + " tab and select a row first.");
    return;
  }
  var rowIndex = sheet.getActiveRange().getRow();
  if (rowIndex < 2) { SpreadsheetApp.getUi().alert("Select a data row (not the header)."); return; }
  var result = sendTechChecklist(rowIndex);
  SpreadsheetApp.getUi().alert(result.ok ? "Checklist sent." : "Error: " + result.error);
}

function sendClientReportFromMenu() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== VISITS_TAB) {
    SpreadsheetApp.getUi().alert("Switch to the " + VISITS_TAB + " tab and select a row first.");
    return;
  }
  var rowIndex = sheet.getActiveRange().getRow();
  if (rowIndex < 2) { SpreadsheetApp.getUi().alert("Select a data row (not the header)."); return; }
  var result = sendClientReport(rowIndex);
  SpreadsheetApp.getUi().alert(result.ok ? "Report sent." : "Error: " + result.error);
}

// ============================================================================
// HTML BUILDERS — Tech Checklist and Client Report
// ============================================================================

function buildTechChecklistHtml(custData, customer, address, visitDate, techName, bookingNotes) {
  var dateLabel = visitDate
    ? Utilities.formatDate(new Date(visitDate), Session.getScriptTimeZone(), "EEE MMM d, h:mm a")
    : "TBD";
  var systemSections = custData.items.map(function(item) {
    return buildChecklistSystemSection(item);
  }).join("");
  var bookingBlock = bookingNotes
    ? '<div style="background:#fff8e8;border-left:3px solid #e8a020;padding:14px 18px;margin:0 28px 0 28px;">' +
        '<p style="margin:0 0 4px 0;font-size:11px;letter-spacing:0.1em;text-transform:uppercase;color:#b8860b;font-weight:bold;">What the customer told us when booking</p>' +
        '<p style="margin:0;font-size:14px;line-height:1.6;color:#333;font-style:italic;">' + bookingNotes + '</p>' +
      '</div>'
    : '';
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:680px;color:#222;background:#f4f4f0;">' +
    '<div style="background:#111;color:#fff;padding:24px 28px;">' +
      '<div style="font-size:11px;letter-spacing:0.1em;color:#888;text-transform:uppercase;">Livewire Lifecycle™ · Field Checklist</div>' +
      '<h1 style="font-family:Georgia,serif;font-size:24px;margin:8px 0 0 0;font-weight:normal;color:#f5f5f0;">' + customer + '</h1>' +
      '<p style="margin:4px 0 0 0;color:#aaa;font-size:13px;">' + dateLabel + (address ? ' · ' + address : '') + '</p>' +
    '</div>' +
    '<div style="background:#ffffff;padding:20px 28px;font-size:14px;line-height:1.6;color:#333;">' +
      'Hi ' + (techName || 'team') + ', here is the visit-specific checklist for this Lifecycle visit. Only the systems this client owns are listed below. Work the list and close out the visit in Zoho the same way you would any other service call.' +
    '</div>' +
    bookingBlock +
    '<div style="background:#fff;padding:4px 28px 28px 28px;">' +
      buildBoilerplateSection("Before You Arrive", [
        "Review the client's Lifecycle email in Zoho",
        "Reference the client folder or the original potential for scope of work or system context",
        "Confirm model numbers, firmware versions, and warranty status",
        "Check for known issues or open service tickets",
        "Load spare parts, test cables, network meter, label maker, camera"
      ]) +
      buildBoilerplateSection("On Arrival", [
        "Introduce yourself and reference the Lifecycle report by name",
        "Set a 60 to 90 minute expectation",
        "Walk the home with the client and ask what is working, what is frustrating, what has changed",
        "Take before photos of every rack, head end, and key touchpoint"
      ]) +
      '<h2 style="font-family:Arial;font-size:13px;text-transform:uppercase;letter-spacing:0.1em;color:#111;border-bottom:2px solid #e8a020;padding-bottom:6px;margin:32px 0 16px 0;">System-by-System Inspection</h2>' +
      systemSections +
      buildBoilerplateSection("Document Findings", [
        "Photo every issue and every piece of gear with visible model and serial",
        "Note firmware version and install year for every major component",
        "Flag anything beyond the 7 to 10 year lifecycle window",
        "Capture client quotes about pain points verbatim"
      ]) +
      buildBoilerplateSection("Before Leaving the Home", [
        "Walk the client through findings in plain language",
        "Handle anything you can fix on the spot: firmware, reboots, app logins, small cables",
        "Create a referral in Zoho for every upsell flagged, tagged Lifecycle, with rough dollar range",
        "Close out the visit in Zoho the same as any service call: time, parts, notes, billable vs warranty",
        "Ask for the Google review if it went well (per Raving Fan gating rules)"
      ]) +
      buildBoilerplateSection("Back at the Shop", [
        "Confirm each referral routed in properly to the correct party (sales or service)",
        "Get the referral written and sent out"
      ]) +
    '</div></div>';
}

function buildChecklistSystemSection(item) {
  var taskMap = {
    "Network": ["Speed test at router and three far points", "Wi-Fi coverage walk, note dead zones", "Confirm firmware on router, switches, APs", "Note any devices offline or unrecognized"],
    "Networking": ["Speed test at router and three far points", "Wi-Fi coverage walk, note dead zones", "Confirm firmware on router, switches, APs", "Note any devices offline or unrecognized"],
    "Home Automation": ["Verify every scene, remote, and app login", "Check processor uptime and firmware", "Test voice control integrations"],
    "Automation": ["Verify every scene, remote, and app login", "Check processor uptime and firmware", "Test voice control integrations"],
    "Control": ["Verify every scene, remote, and app login", "Check processor uptime and firmware", "Test voice control integrations"],
    "Touch Panels": ["Test every touch panel response and screen condition", "Confirm latest firmware", "Check mounting and wake behavior"],
    "TV & Displays": ["Power cycle and test every display", "Confirm HDMI handshakes on all sources", "Verify streaming app logins"],
    "Video": ["Power cycle and test every display", "Confirm HDMI handshakes on all sources", "Verify streaming app logins"],
    "Audio": ["Test every audio zone", "Confirm DSP and amp settings", "Verify streaming service logins"],
    "AV": ["Power cycle and test every zone and display", "Confirm HDMI handshakes", "Verify streaming app logins"],
    "Projectors": ["Lamp or laser hours check", "Test image quality and convergence", "Verify lens and filter cleanliness"],
    "Lighting": ["Test every keypad and scene", "Check for dead bulbs or drifted color temperature", "Verify timeclock and astronomic scheduling"],
    "Shades": ["Cycle every shade, listen for motor noise", "Check battery levels on battery-powered shades"],
    "Security Cameras": ["Verify recording on every camera", "Test remote view and push notifications", "Check lens cleanliness and night IR"],
    "Security": ["Verify alarm panel status and cellular signal", "Test sensors and contacts", "Confirm monitoring connection"],
    "Power Protection": ["Check UPS battery age and self-test", "Verify surge unit indicators", "Thermal check"],
    "Power": ["Check UPS battery age and self-test", "Verify surge unit indicators", "Thermal check"],
    "Climate Control": ["Test thermostat schedules and modes", "Verify sensor calibration", "Check integration with control system"]
  };
  var tasks = taskMap[item.category];
  if (!tasks) {
    // Fuzzy match fallback
    var keys = Object.keys(taskMap);
    for (var k = 0; k < keys.length; k++) {
      if (item.category.toLowerCase().indexOf(keys[k].toLowerCase()) > -1 ||
          keys[k].toLowerCase().indexOf(item.category.toLowerCase()) > -1) {
        tasks = taskMap[keys[k]];
        break;
      }
    }
  }
  if (!tasks) tasks = ["Visual inspection", "Power cycle and verify function", "Note firmware version and any anomalies"];
  var badgeColor = item.status === "Upgrade" ? "#c0392b" : "#d4850a";
  return '<div style="margin:0 0 24px 0;">' +
    '<table cellpadding="0" cellspacing="0" border="0"><tr>' +
      '<td><h3 style="font-family:Arial;font-size:15px;margin:0;color:#111;">' + item.category + '</h3>' +
        '<p style="margin:2px 0 0 0;font-size:11px;color:#888;">Installed ' + item.installYear + '</p></td>' +
      '<td style="padding-left:14px;"><span style="background:' + badgeColor + ';color:#fff;padding:3px 9px;border-radius:2px;font-size:10px;font-weight:bold;letter-spacing:0.08em;text-transform:uppercase;">' + item.status + '</span></td>' +
    '</tr></table>' +
    '<ul style="margin:10px 0 0 0;padding-left:22px;font-size:13px;line-height:1.7;color:#333;list-style:none;">' +
      tasks.map(function(t) { return '<li>☐ ' + t + '</li>'; }).join('') +
    '</ul></div>';
}

function buildBoilerplateSection(title, tasks) {
  return '<h2 style="font-family:Arial;font-size:13px;text-transform:uppercase;letter-spacing:0.1em;color:#111;border-bottom:2px solid #e8a020;padding-bottom:6px;margin:24px 0 12px 0;">' + title + '</h2>' +
    '<ul style="margin:0;padding-left:22px;font-size:13px;line-height:1.8;color:#333;list-style:none;">' +
    tasks.map(function(t) { return '<li>☐ ' + t + '</li>'; }).join('') + '</ul>';
}

function buildClientReportHtml(custData, customer, visitDate, techName, techNotes) {
  var firstName = lcGreetingName(customer, '');
  var headline = firstName ? firstName + ', here is where things stand.' : 'Here is where things stand.';
  var dateLabel = visitDate
    ? Utilities.formatDate(new Date(visitDate), Session.getScriptTimeZone(), "MMMM d, yyyy")
    : "your recent visit";
  var STATUS_COLORS = { "Upgrade": "#c0392b", "Update": "#d4850a", "Review": "#e8a020", "Good": "#27ae60" };
  var systemRows = custData.items.map(function(item) {
    var color = STATUS_COLORS[item.status] || "#e8a020";
    return '<tr>' +
      '<td style="padding:14px 12px;border-bottom:1px solid #eee;font-size:14px;color:#111;font-weight:bold;">' + item.category + '</td>' +
      '<td style="padding:14px 12px;border-bottom:1px solid #eee;font-size:13px;color:#666;">' + item.installYear + '</td>' +
      '<td style="padding:14px 12px;border-bottom:1px solid #eee;"><span style="background:' + color + ';color:#fff;padding:4px 10px;border-radius:2px;font-size:10px;font-weight:bold;letter-spacing:0.08em;text-transform:uppercase;">' + item.status + '</span></td>' +
      '</tr>';
  }).join("");
  var upgradeItems = custData.items.filter(function(i) { return i.status === "Upgrade"; });
  var updateItems = custData.items.filter(function(i) { return i.status === "Update"; });
  var recommendations = [];
  if (upgradeItems.length) {
    recommendations.push('<strong>Replace soon:</strong> ' +
      upgradeItems.map(function(i) { return i.category; }).join(", ") +
      '. These systems are past their expected lifespan.');
  }
  if (updateItems.length) {
    recommendations.push('<strong>Plan within 12 months:</strong> ' +
      updateItems.map(function(i) { return i.category; }).join(", ") +
      '. Healthy today, approaching end of useful life.');
  }
  recommendations.push('If you would like to tackle any of these, just reply to this email or give us a call and we will put together the details.');
  var notesSection = techNotes
    ? '<tr><td style="background:#fff;padding:24px 28px;border-top:1px solid #eee;">' +
        '<h2 style="font-family:Arial;font-size:11px;letter-spacing:0.1em;text-transform:uppercase;color:#888;margin:0 0 12px 0;">Notes from your visit</h2>' +
        '<p style="margin:0;font-family:Georgia,serif;font-size:15px;line-height:1.6;color:#333;font-style:italic;">' + techNotes + '</p>' +
        '<p style="margin:14px 0 0 0;font-size:12px;color:#888;">— ' + (techName || "Livewire Field Team") + '</p>' +
      '</td></tr>'
    : '';
  return '<div style="margin:0;padding:0;background:#f4f4f0;font-family:Arial,Helvetica,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f4f4f0;"><tr><td align="center" style="padding:20px 0;">' +
    '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;">' +
      '<tr><td style="background:#111;padding:18px 28px;border-radius:4px 4px 0 0;">' +
        '<table width="100%"><tr>' +
          '<td><img src="https://www.getlivewire.com/wp-content/uploads/2023/01/logosec.png" width="140" alt="Livewire" style="display:block;border:0;"></td>' +
          '<td align="right"><span style="font-size:11px;color:#555;letter-spacing:0.08em;text-transform:uppercase;">Lifecycle™ Report</span></td>' +
        '</tr></table>' +
      '</td></tr>' +
      '<tr><td style="background:#111;padding:32px 28px 36px 28px;">' +
        '<p style="margin:0 0 6px 0;font-size:11px;color:#888;letter-spacing:0.1em;text-transform:uppercase;">Your home technology · ' + dateLabel + '</p>' +
        '<h1 style="margin:0;font-family:Georgia,serif;font-size:28px;color:#f5f5f0;font-weight:normal;line-height:1.3;">' + headline + '</h1>' +
      '</td></tr>' +
      '<tr><td style="background:#fff;padding:32px 28px 8px 28px;">' +
        '<p style="margin:0 0 16px 0;font-size:15px;line-height:1.65;color:#333;">Thank you for having us in your home. Here is a plain-English look at every system we installed and where each one stands against its expected lifespan.</p>' +
        '<p style="margin:0;font-size:15px;line-height:1.65;color:#333;font-style:italic;">Think of it as an annual physical for your home technology. Most of what we found is healthy. A few things deserve attention, and we have flagged those clearly.</p>' +
      '</td></tr>' +
      '<tr><td style="background:#fff;padding:24px 28px 8px 28px;">' +
        '<h2 style="font-family:Arial;font-size:11px;letter-spacing:0.1em;text-transform:uppercase;color:#888;margin:0 0 12px 0;">Your Systems</h2>' +
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-top:1px solid #eee;">' + systemRows + '</table>' +
      '</td></tr>' +
      notesSection +
      '<tr><td style="background:#fff;padding:28px 28px 12px 28px;">' +
        '<h2 style="font-family:Arial;font-size:11px;letter-spacing:0.1em;text-transform:uppercase;color:#888;margin:0 0 14px 0;">Recommended Next Steps</h2>' +
        recommendations.map(function(r) {
          return '<p style="margin:0 0 12px 0;font-size:14px;line-height:1.6;color:#333;">' + r + '</p>';
        }).join('') +
      '</td></tr>' +
      '<tr><td style="background:#111;padding:36px 28px;text-align:center;">' +
        '<p style="margin:0 0 8px 0;font-size:12px;color:#888;letter-spacing:0.08em;text-transform:uppercase;">Have questions or want to move on something?</p>' +
        '<p style="margin:0 0 18px 0;font-family:Georgia,serif;font-size:18px;color:#f5f5f0;line-height:1.4;">Just reply to this email or give us a call. A real person on our team will help.</p>' +
        '<p style="margin:0;font-size:16px;color:#e8a020;font-weight:bold;letter-spacing:0.04em;">804.937.9001</p>' +
      '</td></tr>' +
      '<tr><td style="background:#fff;padding:24px 28px;border-top:1px solid #eee;">' +
        '<p style="margin:0 0 4px 0;font-size:13px;color:#333;font-weight:bold;">Henry Clifford</p>' +
        '<p style="margin:0;font-size:13px;color:#666;line-height:1.6;">Livewire · Richmond, VA · 804.937.9001</p>' +
        '<p style="margin:14px 0 0 0;font-family:Georgia,serif;font-size:13px;color:#333;font-style:italic;">Stay frosty, and see you in the field.</p>' +
      '</td></tr>' +
      '<tr><td style="background:#f4f4f0;padding:18px 28px;border-radius:0 0 4px 4px;text-align:center;">' +
        '<p style="margin:0;font-size:11px;color:#999;line-height:1.5;">Questions? Reply to this email.</p>' +
      '</td></tr>' +
    '</table></td></tr></table></div>';
}

// === VISIT MODULE END ===