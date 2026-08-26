const assert = require("node:assert");
const { _test } = require("./haloAuth");

const email = {
  bodyHtml: "<p>Confidential email body</p>",
  bodyText: "Confidential email body",
  cc: ["Copied User <cc@example.com>"],
  dateTimeCreated: "2026-08-25T10:00:00.000Z",
  from: "Sender User <sender@example.com>",
  internetMessageId: "<private@example.com>",
  normalizedSubject: "Confidential subject",
  subject: "Confidential subject",
  timeZone: "Europe/London",
  to: ["Support User <support@example.com>"],
};

const publicPayload = _test.buildEmailActionPayload(1001, email, { actionMode: "email" });
assert.strictEqual(publicPayload.outcome, "Email");
assert.strictEqual(publicPayload.hiddenfromuser, undefined);
assert.strictEqual(publicPayload.emailsubject, email.subject);
assert.strictEqual(publicPayload.email_message_id, email.internetMessageId);

const privateAttachment = _test.toHaloActionAttachment(
  { id: "7001", filename: "confidential.pdf", filesize: 42 },
  "private-note"
);
const privatePayload = _test.buildEmailActionPayload(1001, email, {
  actionMode: "private-note",
  attachments: [privateAttachment],
});
assert.strictEqual(privatePayload.outcome, "Private Note");
assert.strictEqual(privatePayload.hiddenfromuser, true);
assert.strictEqual(privatePayload.sendemail, false);
assert.strictEqual(privatePayload.emailbody_html, undefined);
assert.strictEqual(privatePayload.emailsubject, undefined);
assert.strictEqual(privatePayload.email_message_id, undefined);
assert.strictEqual(privatePayload.actioninternetmessageid, undefined);
assert.strictEqual(privatePayload.emailtolistall, undefined);
assert.match(privatePayload.note_html, /Confidential email body/);
assert.strictEqual(privatePayload.attachments[0].showforusers, false);

console.log("Private-note action payload tests passed");
