// Runs the pure helpers under Node: `node tests/run.js`
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const ctx = vm.createContext({});
for (const file of ['Suppression.gs', 'Bounces.gs', 'PreSend.gs', 'Unsubscribe.gs']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', file), 'utf8'), ctx, { filename: file });
}
const call = (expr) => vm.runInContext(expr, ctx);
const plain = (value) => JSON.parse(JSON.stringify(value));
let passed = 0;
const test = (name, fn) => { fn(); passed++; };

const gmailDsn = (recipient, action, status, diagnostic) => [
  '** Message blocked **', '',
  'The response from the remote server was:', diagnostic,
  'Final-Recipient: rfc822; ' + recipient,
  'Action: ' + action,
  'Status: ' + status,
  'Remote-MTA: dns; mx.example.net. (203.0.113.5, the server',
  ' for the domain example.net.)',
  'Diagnostic-Code: smtp; ' + diagnostic,
  'Last-Attempt-Date: Thu, 19 Nov 2026 06:56:34 -0800 (PST)'
].join('\r\n');

test('parses a Gmail failure notice', () => {
  ctx.text = gmailDsn('Pat@Example.net', 'failed', '5.4.1',
    '550 5.4.1 Recipient address rejected: Access denied. [X.prod.outlook.com]');
  const [b] = plain(call('lcParseBounce(text)'));
  assert.strictEqual(b.recipient, 'pat@example.net');
  assert.strictEqual(b.action, 'failed');
  assert.strictEqual(b.status, '5.4.1');
  assert.ok(/Access denied/.test(b.diagnostic));
  ctx.b = b;
  assert.deepStrictEqual(plain(call('lcClassifyBounce(b)')),
    { suppress: true, reason: 'Mailbox not found or rejected' });
});

test('classifies a recipient-side block', () => {
  ctx.text = gmailDsn('a@example.net', 'failed', '5.7.0', '550 Envelope blocked - User Entry - https://x [id]');
  ctx.b = plain(call('lcParseBounce(text)'))[0];
  assert.strictEqual(call('lcClassifyBounce(b).reason'), 'Blocked by recipient');
});

test('does not suppress on a delay notice', () => {
  ctx.text = gmailDsn('a@example.net', 'delayed', '4.4.1', 'The recipient server did not accept our requests');
  ctx.b = plain(call('lcParseBounce(text)'))[0];
  assert.strictEqual(call('lcClassifyBounce(b).suppress'), false);
});

test('suppresses once retries are exhausted', () => {
  ctx.text = gmailDsn('a@example.net', 'failed', '4.4.1', 'The recipient server did not accept our requests');
  ctx.b = plain(call('lcParseBounce(text)'))[0];
  assert.deepStrictEqual(plain(call('lcClassifyBounce(b)')),
    { suppress: true, reason: 'Recipient server unreachable' });
});

test('classifies a dead domain', () => {
  ctx.text = gmailDsn('a@gone.example', 'failed', '4.0.0', 'DNS Error: DNS type mx lookup of gone.example responded with code NXDOMAIN');
  ctx.b = plain(call('lcParseBounce(text)'))[0];
  assert.strictEqual(call('lcClassifyBounce(b).reason'), 'Domain not found');
});

test('handles several recipients in one notice', () => {
  ctx.text = gmailDsn('a@example.net', 'failed', '5.1.1', '550 no such user') + '\r\n\r\n' +
    gmailDsn('b@example.net', 'failed', '5.1.1', '550 no such user');
  assert.deepStrictEqual(plain(call('lcParseBounce(text)')).map((b) => b.recipient),
    ['a@example.net', 'b@example.net']);
});

test('falls back to the human wording', () => {
  ctx.text = "Address not found\nYour message wasn't delivered to pat@example.net because the address couldn't be found.";
  assert.deepStrictEqual(plain(call('lcParseBounce(text)')),
    [{ recipient: 'pat@example.net', action: 'failed', status: '', diagnostic: '' }]);
  ctx.text = 'Delivery incomplete\nThere was a temporary problem delivering your message to pat@example.net. Gmail will retry for 44 more hours.';
  assert.strictEqual(plain(call('lcParseBounce(text)'))[0].action, 'delayed');
  ctx.text = 'Your message to pat@example.net has been blocked. See technical details below.';
  assert.strictEqual(plain(call('lcParseBounce(text)'))[0].action, 'failed');
});

test('ignores ordinary mail', () => {
  ctx.text = 'Hi, thanks for the note. Talk soon.';
  assert.deepStrictEqual(plain(call('lcParseBounce(text)')), []);
});

test('email syntax', () => {
  const ok = (e) => { ctx.e = e; return call('lcIsValidEmailSyntax(e)'); };
  assert.ok(ok('pat@example.com'));
  assert.ok(ok('  Pat.Smith+x@mail.example.co.uk '));
  for (const bad of ['', null, 'pat@example', 'pat example@x.com', 'a@x.com; b@x.com', 'pat@@example.com', '@example.com']) {
    assert.ok(!ok(bad), String(bad));
  }
});

test('greeting name', () => {
  const g = (customer, contact) => { ctx.c = customer; ctx.p = contact; return call('lcGreetingName(c, p)'); };
  assert.strictEqual(g('Pat Example'), 'Pat');
  assert.strictEqual(g('Pat and Sam Example'), 'Pat');
  assert.strictEqual(g("Jim O'Example"), 'Jim');
  assert.strictEqual(g('Jo Anne Example'), 'Jo');
  assert.strictEqual(g('Pat Example JR'), 'Pat');
  assert.strictEqual(g('Pat Example (2)'), 'Pat');
  assert.strictEqual(g('Pat C. Example'), 'Pat');
  assert.strictEqual(g('Dr. Pat Example D.D.S.'), 'Pat');
  for (const business of ['ExampleFirm, P.C.', 'ACME of Virginia', 'Acme Asset Management', 'Acme Consulting Group',
    'Acme Health', 'Acme Custom Homes', 'Example Builders, Inc.', 'H. J. Example and Son, Inc.', 'The Example Center',
    'Estate of Pat Example', 'Pat Example Construction', 'Acme', 'System13', 'Apple REIT', 'CT Example',
    'VA Medical Center (Richmond)', 'Pat Example Residence', '']) {
    assert.strictEqual(g(business), '', business);
  }
  assert.strictEqual(g('Acme Custom Homes', 'Sam Builder'), 'Sam');
  assert.strictEqual(g('Acme Custom Homes', 'Dr. Sam Builder'), 'Sam');
});

test('html to text', () => {
  ctx.h = '<div><div style="display:none;max-height:0">preheader</div><h1>Pat &mdash; hello</h1>' +
    '<p>We&rsquo;ll walk   through it.</p><a href="https://example.com/book" style="x">Reserve &rarr;</a></div>';
  assert.strictEqual(call('lcHtmlToText(h)'),
    'Pat — hello\n\nWe’ll walk through it.\n\nReserve → (https://example.com/book)');
});

test('unsubscribe replies', () => {
  const u = (body) => { ctx.b = body; return call('lcIsUnsubscribeRequest(b)'); };
  const quoted = '\n\nOn Thu, Nov 19, 2026 at 9:56 AM Livewire Concierge <x@example.com> wrote:\n' +
    '> Pat \u2014 your technology is showing its age.\n> Prefer not to get these emails? Reply with \u201cunsubscribe\u201d';
  for (const yes of ['unsubscribe' + quoted, 'Unsubscribe.', 'STOP', 'Please remove me from your list.' + quoted,
    'Please take me off this list, thanks' + quoted, 'opt out',
    'Hello, I am retiring and no longer wish to receive email from livewire.' + quoted,
    'Please remove my email from your list', 'Do not email me again', 'I don\u2019t want to receive these']) {
    assert.ok(u(yes), yes);
  }
  for (const no of ['Sounds good, can we book for next Tuesday?' + quoted,
    'Please remove the old TV from the bill and call me' + quoted,
    'Thanks! Who should I call?\n\nSent from my iPhone\n\n> unsubscribe',
    'Thanks\r\n\r\n-----Original Message-----\r\nFrom: Livewire\r\nPrefer not to get these emails? Reply with unsubscribe',
    'Thanks, interested.\n________________________________\nFrom: Livewire\nPrefer not to get these emails? Reply with unsubscribe',
    '']) {
    assert.ok(!u(no), no);
  }
  ctx.h = '"Smith, Pat" <Pat.Smith@Example.com>, lee@example.org';
  assert.deepStrictEqual(plain(call('lcHeaderAddresses(h)')), ['pat.smith@example.com', 'lee@example.org']);
});

console.log(passed + ' tests passed');
