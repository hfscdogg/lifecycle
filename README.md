# Lifecycle deliverability add-on

Apps Script files that sit beside the existing Lifecycle sender (the script bound to the
Lifecycle Engine Google Sheet). They stop the sender from mailing addresses that cannot
receive mail, and fix the greeting for business customers.

The main sender script is not in this repo yet. These files are self-contained and are
wired in with the few lines shown under "Wire into the send loop".

## What is here

| File | Purpose |
|---|---|
| `src/Suppression.gs` | `SUPPRESSION` tab (created on first use), `lcIsSuppressed`, logging to `DIAGNOSTIC_LOG` |
| `src/Bounces.gs` | `processBounces` reads Gmail bounce notices, suppresses failed queue addresses and turns off their `Send?` flag |
| `src/PreSend.gs` | `lcPreSendCheck`, `lcGreetingName`, `lcHtmlToText` |
| `tests/run.js` | Node tests for the pure helpers: `node tests/run.js` |

## Install

1. Open the sheet, then Extensions > Apps Script.
2. Add three script files named `Suppression`, `Bounces` and `PreSend`, and paste in the matching file from `src/`.
3. Run `backfillBounces` once from the editor. Approve the permission prompt (it adds "connect to an external service", used for the DNS check). This fills the `SUPPRESSION` tab from the last 45 days of bounce notices.
4. Run `installBounceTrigger` once. `processBounces` then runs daily.

If the project has an `appsscript.json` with an explicit `oauthScopes` list, add
`https://www.googleapis.com/auth/script.external_request` to it.

## Wire into the send loop

In the existing function that walks `EMAIL_QUEUE` and sends, before each send:

```js
var check = lcPreSendCheck(email);
if (!check.ok) {
  lcLog_('SKIP', customer + ' <' + email + '>: ' + check.reason);
  continue;
}

var firstName = lcGreetingName(customer, primaryContact); // primaryContact may be ''
var subject = firstName
  ? firstName + ' — your tech is showing its age'
  : 'Your tech is showing its age';
// Use the same rule for the headline inside the email.
```

When sending, pass a plain-text body next to the HTML:

```js
GmailApp.sendEmail(email, subject, lcHtmlToText(html), { htmlBody: html, name: senderName });
```

## Behaviour notes

- `lcPreSendCheck` skips an address when it is malformed, on the `SUPPRESSION` tab, on the
  sender's own domain (set `LC.SKIP_INTERNAL_DOMAIN = false` to allow), or its domain has
  no mail server. If the DNS lookup itself fails, the address is allowed.
- `processBounces` suppresses on a final failure only. "Delivery incomplete" delay notices
  are ignored until Gmail gives up. Only addresses present in `EMAIL_QUEUE` are suppressed.
  Processed threads get the Gmail label `Lifecycle/Bounce-processed`.
- To mail a suppressed address again, delete its row from the `SUPPRESSION` tab.
- `lcGreetingName` returns `''` for business names, single-word names, initials and
  acronyms. A person whose surname is a business word (for example "Church" or "Law")
  also gets `''`, which falls back to the no-name subject.

## Main sender (`src/Code.gs`)

`src/Code.gs` and `src/appsscript.json` are the live project's files. The first commit that
adds them is an unchanged copy; the next commit wires in the add-on:

- `sendLifecycleEmails` and `sendCatchupBatch` call `lcPreSendCheck` before each send, use
  `lcGreetingName` for the subject and headline, and send `lcHtmlToText(html)` as the plain-text part.
- `CONFIG.DRY_RUN` (live: `false` since 2026-11-19) when `true` logs each customer email to `DIAGNOSTIC_LOG` as `DRY_RUN`
  instead of sending it, and leaves `SentTimestamp` empty. Set it to `false` to send.
- The lifecycle email footer has a visible opt-out line (reply "unsubscribe").
- Visit emails: `sendClientReport` runs `lcPreSendCheck` on the client address, greets with
  `lcGreetingName` (subject and headline) and adds a plain-text part. `sendTechChecklist`
  adds a plain-text part. Both are covered by `DRY_RUN`, which leaves their sent columns empty.

## Not done yet

- A `List-Unsubscribe` header. `GmailApp` cannot set custom headers; it needs the Gmail API.
- Replacing the stock-photo image URLs with images hosted on the company site.
