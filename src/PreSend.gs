/**
 * Lifecycle deliverability add-on: checks to run before each send, plus
 * helpers for the greeting and the plain-text part.
 */

var LC_TITLES = { dr: 1, mr: 1, mrs: 1, ms: 1, miss: 1, prof: 1, rev: 1, hon: 1, sir: 1, the: 1 };

var LC_SUFFIXES = { JR: 1, SR: 1, II: 1, III: 1, IV: 1, COD: 1, DDS: 1, DMD: 1, MD: 1, DVM: 1, PHD: 1, CPA: 1, ESQ: 1 };

var LC_BUSINESS_WORDS = ('inc llc ltd lp llp pc pllc corp corporation co company companies group partners ' +
  'associates association holdings enterprises ventures capital management consulting advisors advisory ' +
  'strategies benefits investments asset assets financial insurance bank credit realty real estate properties ' +
  'homes home homecrafters builders builder building construction contracting custom landscaping nurseries ' +
  'design designs studio advertising marketing media brewing restaurant cafe center centre church club school ' +
  'university college academy foundation medical health dental clinic hospital law legal services solutions ' +
  'systems technologies technology navigator residence of son sons').split(' ');

var lcBusinessWordMap_ = (function () {
  var map = {};
  for (var i = 0; i < LC_BUSINESS_WORDS.length; i++) map[LC_BUSINESS_WORDS[i]] = true;
  return map;
})();

/** Pure. Basic shape check; catches blanks, spaces, missing dots and multiple addresses in one cell. */
function lcIsValidEmailSyntax(email) {
  return /^[^\s@,;<>()"]+@[a-z0-9-]+(\.[a-z0-9-]+)+$/i.test(String(email == null ? '' : email).trim());
}

function lcInternalDomain_() {
  return String(LC.INTERNAL_DOMAIN || '').toLowerCase();
}

function lcDnsHasRecords_(domain, type) {
  var response = UrlFetchApp.fetch(
    'https://dns.google/resolve?name=' + encodeURIComponent(domain) + '&type=' + type,
    { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return null; // unknown
  var data = JSON.parse(response.getContentText());
  if (data.Status === 3) return false; // domain does not exist
  if (data.Status !== 0) return null;
  return !!(data.Answer && data.Answer.length);
}

/** True unless DNS says the domain cannot receive mail. Lookup problems count as "can". */
function lcDomainAcceptsMail_(domain) {
  var cache = CacheService.getScriptCache();
  var key = 'lc-mx-' + domain;
  var cached = cache.get(key);
  if (cached !== null) return cached === '1';
  var ok = true;
  try {
    var mx = lcDnsHasRecords_(domain, 'MX');
    if (mx === false) {
      var a = lcDnsHasRecords_(domain, 'A');
      ok = a !== false;
    }
  } catch (e) {
    return true; // do not cache a failed lookup
  }
  cache.put(key, ok ? '1' : '0', 21600);
  return ok;
}

/**
 * Call for every queue row before sending.
 * Returns {ok: true} or {ok: false, reason: '...'}.
 */
function lcPreSendCheck(email) {
  var address = lcNormalizeEmail(email);
  if (!lcIsValidEmailSyntax(address)) return { ok: false, reason: 'malformed address' };
  if (lcIsSuppressed(address)) return { ok: false, reason: 'on suppression list' };
  var domain = address.split('@')[1];
  if (LC.SKIP_INTERNAL_DOMAIN && domain === lcInternalDomain_()) {
    return { ok: false, reason: 'internal address' };
  }
  if (!lcDomainAcceptsMail_(domain)) return { ok: false, reason: 'domain has no mail server' };
  return { ok: true, reason: '' };
}

/**
 * Pure. Picks the first name to greet with, or '' when there is no safe one
 * (business names, titles only, initials). Callers should drop the name from
 * the subject and headline when this returns ''.
 */
function lcGreetingName(customerName, primaryContact) {
  var contact = String(primaryContact == null ? '' : primaryContact).trim();
  var source = contact || String(customerName == null ? '' : customerName).trim();
  if (!source) return '';
  var tokens = source.replace(/\([^)]*\)/g, ' ').split(/[\s,]+/).filter(function (t) { return t; });

  if (!contact) {
    if (tokens.length < 2) return '';
    for (var i = 0; i < tokens.length; i++) {
      var raw = tokens[i].replace(/\.+$/, '');
      var bare = raw.replace(/\./g, '');
      if (lcBusinessWordMap_[bare.toLowerCase()] && i > 0) return '';
      if (/\d/.test(bare) && i === 0) return '';
      if (/^[A-Z]{3,}$/.test(bare) && !LC_SUFFIXES[bare]) return '';
    }
  }

  var index = 0;
  while (index < tokens.length && LC_TITLES[tokens[index].replace(/\./g, '').toLowerCase()]) index++;
  if (!contact && index > 0 && tokens[0].replace(/\./g, '').toLowerCase() === 'the') return '';
  var first = (tokens[index] || '').replace(/[^A-Za-z'À-ɏ-]/g, '');
  if (first.length < 2) return '';
  if (first === first.toUpperCase()) return ''; // initials or an acronym
  if (lcBusinessWordMap_[first.toLowerCase()]) return '';
  return first;
}

/** Pure. Turns the HTML email into a readable plain-text alternative. */
function lcHtmlToText(html) {
  var entities = { nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', '#39': "'", rsquo: '’',
    lsquo: '‘', ldquo: '“', rdquo: '”', mdash: '—', ndash: '–',
    middot: '·', rarr: '→', trade: '™' };
  return String(html || '')
    .replace(/<(style|script|head)[\s\S]*?<\/\1>/gi, '')
    .replace(/<[^>]*display\s*:\s*none[^>]*>[\s\S]*?<\/[^>]+>/gi, '')
    .replace(/<a\b[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, function (m, href, label) {
      var text = label.replace(/<[^>]+>/g, '').trim();
      return text && text !== href ? text + ' (' + href + ')' : href;
    })
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|h[1-6]|tr|table|li)>/gi, '\n\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&(#39|[a-z]+);/gi, function (m, name) {
      var key = name.toLowerCase();
      return entities.hasOwnProperty(key) ? entities[key] : m;
    })
    .replace(/[ \t]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
