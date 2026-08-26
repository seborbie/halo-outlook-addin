const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const marker = {
  version: 1,
  composeAttachId: "compose-operation-1",
  ticketId: "1001",
  ticketNumber: "T1001",
  ticketSummary: "Example ticket",
  draftItemId: "draft-item-id",
};

const creationMarker = {
  version: 4,
  destinationKind: "create-ticket",
  composeAttachId: "create-operation-1",
  creationOperationId: "create-operation-1",
  ticketId: "",
  ticketNumber: "New Project Engineer",
  ticketSummary: "Example project email",
  ticketTypeId: "12",
  ticketTypeName: "Project Engineer",
  draftItemId: "draft-item-id",
};

function run() {
  const isBuiltRuntime = Boolean(process.argv[2]);
  const runtimePath = isBuiltRuntime
    ? path.resolve(__dirname, "..", process.argv[2])
    : path.join(__dirname, "..", "src", "commands", "classic-send-runtime.js");
  const commandsHtmlPath = isBuiltRuntime
    ? path.join(__dirname, "..", "dist", "commands.html")
    : path.join(__dirname, "..", "src", "commands", "commands.html");
  const runtime = fs.readFileSync(runtimePath, "utf8");
  const commandsHtml = fs.readFileSync(commandsHtmlPath, "utf8");

  const officeReadyIndex = commandsHtml.indexOf("Office.onReady");
  const runtimeScriptIndex = commandsHtml.indexOf("/public/classic-send-runtime.js");
  assert(officeReadyIndex >= 0, "The HTML runtime must initialize Office.js.");
  assert(
    runtimeScriptIndex > officeReadyIndex,
    "The HTML runtime must initialize Office.js before loading the send handlers."
  );

  if (isBuiltRuntime) {
    assert.doesNotMatch(
      runtime,
      /\?\?|\?\./,
      "The classic Outlook runtime must not contain optional chaining or nullish coalescing."
    );
    assert.doesNotMatch(
      runtime,
      /__HALO_PUBLIC_BASE_URL__/,
      "The built event runtime must not retain the public-origin placeholder."
    );
    assert.match(
      runtime,
      /https:\/\/[^"']+\/api\/halo\/email\/send-auto-attach/,
      "The built event runtime must use an absolute HTTPS API URL."
    );
    assert.match(
      commandsHtml,
      /classic-send-runtime\.js\?v=2026\.8\.24\.(?:0|2-references-1)/,
      "The built commands runtime must use the release or diagnostics cache token."
    );
  }

  testUnselectedWatchdog(runtime, runtimePath);
  testSessionSelection(runtime, runtimePath);
  testPrivateNoteSelection(runtime, runtimePath);
  testSessionSelectionCarriesRuntimeAuthContext(runtime, runtimePath);
  testExplicitSendDiagnostics(runtime, runtimePath);
  testCreationSelection(runtime, runtimePath);
  testCreationFailureUsesCreationWarning(runtime, runtimePath);
  testCustomPropertyFallback(runtime, runtimePath);
  testInvalidJsonMarkerIsIgnored(runtime, runtimePath);
  testFinalComposeRecipients(runtime, runtimePath);
  testStaleInheritedSelection(runtime, runtimePath);
  testExplicitNoSessionWarning(runtime, runtimePath);
  testExplicitUnexpectedResponseWarning(runtime, runtimePath);
  testExplicitRequestTimeout(runtime, runtimePath);
  testPrefetchedInlineImagesAvoidContentReads(runtime, runtimePath);
  testLegacyWebCryptoPrefetchReuse(runtime, runtimePath);
  testEncodedContentIdPrefetchReuse(runtime, runtimePath);
  testChangedDraftReadsFourAtATime(runtime, runtimePath);
  testTimedOutReadsDoNotSpawnMore(runtime, runtimePath);
  testOversizedInlineImageIsNotRead(runtime, runtimePath);
  testPrefetchedEmailAttachmentsAvoidContentReads(runtime, runtimePath);
  testExplicitEmailOnlyAvoidsContentReads(runtime, runtimePath);
  testStaleEmailAttachmentStateIsNotReused(runtime, runtimePath);
  testAutomaticPreparedAttachmentsBlockWithoutSession(runtime, runtimePath);
  testAutomaticPreparedAttachmentsBlockOnWatchdog(runtime, runtimePath);
  testMissingEmailAttachmentStateBlocksWithoutReads(runtime, runtimePath);
  testAttachmentsChangedPrefetchesAutomaticReply(runtime, runtimePath);
  testAttachmentsChangedKeepsUnreadEligibleFilePending(runtime, runtimePath);
  testAttachmentsChangedBoundsHungContentReads(runtime, runtimePath);
  testAttachmentsChangedRespectsExplicitExclusion(runtime, runtimePath);
  testAttachmentsChangedKeepsBackgroundSessionOutOfCustomProperties(runtime, runtimePath);
  testReferencedAttachmentIdIsNotOrdinary(runtime, runtimePath);
  testDuplicateOrdinaryMetadataIsReadOnce(runtime, runtimePath);
  testEmlIcsAndCloudAttachmentConversion(runtime, runtimePath);
  testRecoveryHeaderCandidateWins(runtime, runtimePath);
  testRecoveryWaitsForSentPropertyCandidate(runtime, runtimePath);
  testRecoveryTraversesDraftReferencesWhenDirectParentIsIncoming(runtime, runtimePath);
  testRecoveryNoMatchAllowsSend(runtime, runtimePath);

  console.log("Classic Outlook send runtime tests passed");
}

function testRecoveryHeaderCandidateWins(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    conversationId: "new-outlook-conversation",
    headerComposeId: "compose-operation-1",
    responseForRecovery(request) {
      const body = JSON.parse(request.body);
      return body.composeAttachIds.length
        ? { ok: true, status: "matched", candidateIndex: 0, ticketId: "1001" }
        : { ok: true, status: "no-match" };
    },
    response: { ok: true, status: "attached" },
  });
  harness.send();
  assert.strictEqual(harness.completions[0].allowEvent, true);
  const request = harness.requests.find((entry) => /send-auto-attach$/.test(entry.url));
  assert(request);
  assert.strictEqual(JSON.parse(request.body).recoveredComposeAttachId, "compose-operation-1");
}

function testRecoveryWaitsForSentPropertyCandidate(runtime, runtimePath) {
  const customValue = JSON.stringify({
    "halo.composeAttach.v1": JSON.stringify({ composeAttachId: "sent-compose-id" }),
  })
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  const harness = createHarness(runtime, runtimePath, {
    inReplyTo: "<parent@example.com>",
    ewsResponder(request) {
      if (request.includes("FindItem")) {
        return '<m:FindItemResponse><t:ItemId Id="sent-item-id"/></m:FindItemResponse>';
      }
      if (request.includes('PropertyName="cecp-')) {
        return (
          '<t:ExtendedProperty><t:ExtendedFieldURI PropertyName="cecp-55bbcff2-8191-4411-aec6-f9d2f9b4b5e8"/><t:Value>' +
          customValue +
          "</t:Value></t:ExtendedProperty>"
        );
      }
      return "<m:GetItemResponse/>";
    },
    responseForRecovery(request) {
      const body = JSON.parse(request.body);
      return body.composeAttachIds.includes("sent-compose-id")
        ? { ok: true, status: "matched", candidateIndex: 0, ticketId: "1001" }
        : { ok: true, status: "no-match" };
    },
    response: { ok: true, status: "attached" },
  });
  harness.send();
  const request = harness.requests.find((entry) => /send-auto-attach$/.test(entry.url));
  assert(request, "The slower Sent Items branch should still be allowed to establish the mapping.");
  assert.strictEqual(JSON.parse(request.body).recoveredComposeAttachId, "sent-compose-id");
}

function testRecoveryTraversesDraftReferencesWhenDirectParentIsIncoming(runtime, runtimePath) {
  const customValue = JSON.stringify({
    "halo.composeAttach.v1": JSON.stringify({ composeAttachId: "original-compose-id" }),
  })
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  const harness = createHarness(runtime, runtimePath, {
    inReplyTo: "<incoming-reply@example.com>",
    ewsResponder(request) {
      if (request.includes('PropertyTag="0x1039"')) {
        return (
          '<t:ExtendedProperty><t:Value>&lt;original-sent@example.com&gt;</t:Value>' +
          "</t:ExtendedProperty>"
        );
      }
      if (request.includes("FindItem") && request.includes("original-sent@example.com")) {
        return '<m:FindItemResponse><t:ItemId Id="original-sent-item"/></m:FindItemResponse>';
      }
      if (request.includes("FindItem")) {
        return "<m:FindItemResponse/>";
      }
      if (request.includes('PropertyName="cecp-')) {
        return (
          '<t:ExtendedProperty><t:ExtendedFieldURI PropertyName="cecp-55bbcff2-8191-4411-aec6-f9d2f9b4b5e8"/><t:Value>' +
          customValue +
          "</t:Value></t:ExtendedProperty>"
        );
      }
      return "<m:GetItemResponse/>";
    },
    responseForRecovery(request) {
      const body = JSON.parse(request.body);
      return body.composeAttachIds.includes("original-compose-id")
        ? { ok: true, status: "matched", candidateIndex: 0, ticketId: "1001" }
        : { ok: true, status: "no-match" };
    },
    response: { ok: true, status: "attached" },
  });
  harness.send();
  const request = harness.requests.find((entry) => /send-auto-attach$/.test(entry.url));
  assert(request, "Recovery must traverse References to the originally sent mapped message.");
  assert.strictEqual(JSON.parse(request.body).recoveredComposeAttachId, "original-compose-id");
}

function testRecoveryNoMatchAllowsSend(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    attachments: [createOrdinaryAttachment(1)],
    conversationId: "unmapped-conversation",
    recoveryResponse: { ok: true, status: "no-match" },
  });
  harness.send();
  assert.strictEqual(harness.completions[0].allowEvent, true);
  assert.strictEqual(harness.attachmentContentCalls, 0);
  assert.strictEqual(harness.requests.length, 0);
}

function testFinalComposeRecipients(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    cc: [{ displayName: "Copied", emailAddress: "copied@example.com" }],
    sessionMarker: marker,
    response: { ok: true, status: "attached" },
  });
  harness.send();

  const body = JSON.parse(harness.requests[0].body);
  assert.deepStrictEqual(body.to, [
    { displayName: "Recipient", emailAddress: "recipient@example.com" },
  ]);
  assert.deepStrictEqual(body.cc, [{ displayName: "Copied", emailAddress: "copied@example.com" }]);
  assert.strictEqual(body.bcc, undefined, "Bcc must not be archived in the Halo action payload.");
}

function testUnselectedWatchdog(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    conversationId: "mapped-conversation",
    hangBodyRead: true,
  });
  harness.send();
  assert.strictEqual(harness.completions.length, 0);
  harness.fireTimer(25000);
  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, true);
}

function testCreationSelection(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    sessionMarker: creationMarker,
    response: { ok: true, status: "attached", ticketNumber: "T2001" },
  });
  harness.send();

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, true);
  assert.strictEqual(harness.requests.length, 1);
  assert.match(
    harness.requests[0].url,
    /\/api\/halo\/ticket-creation\/intents\/create-operation-1\/send$/
  );
  const body = JSON.parse(harness.requests[0].body);
  assert.strictEqual(body.composeAttachId, creationMarker.creationOperationId);
  assert.strictEqual(body.subject, "New subject");
}

function testCreationFailureUsesCreationWarning(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    sessionMarker: creationMarker,
    response: { error: "Halo validation failed", ok: false, status: "failed" },
  });
  harness.send();

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, false);
  assert.match(harness.completions[0].errorMessage, /Could not create/i);
  assert.doesNotMatch(harness.completions[0].errorMessage, /mapped Halo ticket/i);
}

function testSessionSelection(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    sessionMarker: marker,
    response: { ok: true, status: "attached", ticketNumber: "T1001" },
  });
  harness.send();

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, true);
  assert.strictEqual(harness.requests.length, 1);
  assert.match(harness.requests[0].url, /\/api\/halo\/tickets\/1001\/sent-email$/);
  const body = JSON.parse(harness.requests[0].body);
  assert.strictEqual(body.composeAttachId, marker.composeAttachId);
  assert.strictEqual(body.ticketNumber, marker.ticketNumber);
  assert.strictEqual(body.conversationId, "");
  assert.strictEqual(body.bodyHtml, "<p>Composed email</p>");
}

function testPrivateNoteSelection(runtime, runtimePath) {
  const privateMarker = { ...marker, actionMode: "private-note" };
  const harness = createHarness(runtime, runtimePath, {
    sessionMarker: privateMarker,
    response: { ok: true, status: "attached", ticketNumber: "T1001" },
  });
  harness.send();

  const body = JSON.parse(harness.requests[0].body);
  assert.strictEqual(body.actionMode, "private-note");
}

function testSessionSelectionCarriesRuntimeAuthContext(runtime, runtimePath) {
  const runtimeMarker = {
    ...marker,
    backgroundSessionId: "marker-background-session",
    mailboxEmail: "marker-mailbox@example.com",
  };
  const harness = createHarness(runtime, runtimePath, {
    roamingBackgroundSession: "",
    sessionMarker: runtimeMarker,
    userProfileEmail: "",
    response: { ok: true, status: "attached", ticketNumber: "T1001" },
  });
  harness.send();

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, true);
  const body = JSON.parse(harness.requests[0].body);
  assert.strictEqual(body.backgroundSessionId, runtimeMarker.backgroundSessionId);
  assert.strictEqual(body.mailboxEmail, runtimeMarker.mailboxEmail);
}

function testExplicitSendDiagnostics(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    sessionMarker: marker,
    response: { ok: true, status: "attached", ticketNumber: "T1001" },
  });
  harness.send();

  const diagnostics = harness.diagnosticRequests.map((request) => JSON.parse(request.body));
  const stages = diagnostics.map((entry) => entry.stage);
  assert(stages.includes("runtime-loaded"));
  assert(stages.includes("event-start"));
  assert(stages.includes("marker-read"));
  assert(stages.includes("draft-save-start"));
  assert(stages.includes("compose-read-complete"));
  assert(stages.includes("assets-complete"));
  assert(stages.includes("request-start"));
  assert(stages.includes("request-complete"));
  assert(stages.includes("event-complete"));
  const serialized = JSON.stringify(diagnostics);
  assert.doesNotMatch(serialized, /T1001|New subject|recipient@example\.com|Composed email/);
}

function testAttachmentsChangedKeepsBackgroundSessionOutOfCustomProperties(runtime, runtimePath) {
  const runtimeMarker = {
    ...marker,
    backgroundSessionId: "marker-background-session",
    mailboxEmail: "marker-mailbox@example.com",
  };
  const harness = createHarness(runtime, runtimePath, {
    sessionMarker: runtimeMarker,
  });
  harness.attachmentsChanged();

  assert.strictEqual(harness.attachmentChangeCompletions, 1);
  const sessionMarker = JSON.parse(harness.sessionValues["halo.composeAttach.v1"]);
  const customMarker = JSON.parse(harness.customPropertyValues["halo.composeAttach.v1"]);
  assert.strictEqual(sessionMarker.backgroundSessionId, runtimeMarker.backgroundSessionId);
  assert.strictEqual(customMarker.backgroundSessionId, undefined);
  assert.strictEqual(customMarker.mailboxEmail, runtimeMarker.mailboxEmail);
}

function testCustomPropertyFallback(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    customMarker: marker,
    response: { ok: true, status: "already-attached", ticketNumber: "T1001" },
  });
  harness.send();

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, true);
  assert.strictEqual(harness.requests.length, 1);
  assert.match(harness.requests[0].url, /\/sent-email$/);
}

function testInvalidJsonMarkerIsIgnored(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    sessionMarkerRaw: "null",
  });
  harness.send();

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, true);
  assert.strictEqual(harness.requests.length, 0);
}

function testStaleInheritedSelection(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    customMarker: marker,
    itemId: "different-draft-id",
  });
  harness.send();

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, true);
  assert.strictEqual(harness.requests.length, 0, "A stale inherited marker must be ignored.");
}

function testExplicitNoSessionWarning(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    sessionMarker: marker,
    response: { ok: true, status: "no-session", ticketNumber: "T1001" },
  });
  harness.send();

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, false);
  assert.match(harness.completions[0].errorMessage, /T1001/);
}

function testExplicitUnexpectedResponseWarning(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    sessionMarker: marker,
    response: { ok: true, status: "no-match" },
  });
  harness.send();

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, false);
  assert.match(harness.completions[0].errorMessage, /T1001/);
}

function testExplicitRequestTimeout(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    hangRequest: true,
    sessionMarker: marker,
  });
  harness.send();
  assert.strictEqual(harness.completions.length, 0);
  harness.fireTimer(20000);

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, false);
  assert.match(harness.completions[0].errorMessage, /T1001/);
}

function testPrefetchedInlineImagesAvoidContentReads(runtime, runtimePath) {
  const attachment = createInlineAttachment(1);
  const fingerprint = sha256Hex(
    ["logo1@example.com", attachment.contentId, attachment.name, attachment.size, ""].join("\u0000")
  );
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    bodyHtml: '<p>Composed email</p><img src="cid:logo1@example.com">',
    sessionMarker: {
      ...marker,
      version: 2,
      inlineImageFingerprint: fingerprint,
      inlineImagePrefetchKey: "opaque-prefetch-key",
    },
    response: { ok: true, status: "attached-with-image-warnings" },
  });
  harness.send();

  assert.strictEqual(harness.attachmentContentCalls, 0);
  assert.strictEqual(harness.requests.length, 1);
  const body = JSON.parse(harness.requests[0].body);
  assert.strictEqual(body.inlineImagePrefetchKey, "opaque-prefetch-key");
  assert.strictEqual(body.inlineImageFingerprint, fingerprint);
  assert.strictEqual(body.inlineImageUploads, undefined);
  assert.strictEqual(harness.completions[0].allowEvent, true);
}

function testChangedDraftReadsFourAtATime(runtime, runtimePath) {
  const attachments = Array.from({ length: 6 }, (_, index) => createInlineAttachment(index + 1));
  const bodyHtml = attachments
    .map((attachment) => `<img src="cid:${attachment.contentId}">`)
    .join("");
  const harness = createHarness(runtime, runtimePath, {
    attachments,
    bodyHtml,
    deferAttachmentContent: true,
    sessionMarker: {
      ...marker,
      version: 2,
      inlineImageFingerprint: "draft-changed",
      inlineImagePrefetchKey: "stale-prefetch-key",
    },
    response: { ok: true, status: "attached" },
  });
  harness.send();
  assert.strictEqual(harness.attachmentContentCalls, 4);
  assert.strictEqual(harness.maxConcurrentAttachmentReads, 4);

  harness.completeAllAttachmentReads();
  assert.strictEqual(harness.attachmentContentCalls, 6);
  assert.strictEqual(harness.requests.length, 1);
  const body = JSON.parse(harness.requests[0].body);
  assert.strictEqual(body.inlineImageRefs.length, 6);
  assert.strictEqual(body.inlineImageUploads.length, 6);
  assert.strictEqual(body.inlineImagePrefetchKey, undefined);
  assert.strictEqual(harness.completions[0].allowEvent, true);
}

function testLegacyWebCryptoPrefetchReuse(runtime, runtimePath) {
  const attachment = createInlineAttachment(1);
  const fingerprint = sha256Hex(
    ["logo1@example.com", attachment.contentId, attachment.name, attachment.size, ""].join("\u0000")
  );
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    bodyHtml: '<img src="cid:logo1@example.com">',
    legacyCrypto: true,
    sessionMarker: {
      ...marker,
      version: 2,
      inlineImageFingerprint: fingerprint,
      inlineImagePrefetchKey: "legacy-prefetch-key",
    },
    response: { ok: true, status: "attached" },
  });
  harness.send();

  assert.strictEqual(harness.attachmentContentCalls, 0);
  assert.strictEqual(harness.requests.length, 1);
  assert.strictEqual(
    JSON.parse(harness.requests[0].body).inlineImagePrefetchKey,
    "legacy-prefetch-key"
  );
  assert.strictEqual(harness.completions[0].allowEvent, true);
}

function testEncodedContentIdPrefetchReuse(runtime, runtimePath) {
  const attachment = createInlineAttachment(1);
  attachment.contentId = "<Logo@Example.com>";
  const fingerprint = sha256Hex(
    ["logo@example.com", "logo@example.com", attachment.name, attachment.size, ""].join("\u0000")
  );
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    bodyHtml: '<img originalsrc="CID:&lt;Logo%40Example.COM&gt;">',
    sessionMarker: {
      ...marker,
      version: 2,
      inlineImageFingerprint: fingerprint,
      inlineImagePrefetchKey: "encoded-cid-prefetch-key",
    },
    response: { ok: true, status: "attached" },
  });
  harness.send();

  assert.strictEqual(harness.attachmentContentCalls, 0);
  assert.strictEqual(
    JSON.parse(harness.requests[0].body).inlineImagePrefetchKey,
    "encoded-cid-prefetch-key"
  );
  assert.strictEqual(harness.completions[0].allowEvent, true);
}

function testOversizedInlineImageIsNotRead(runtime, runtimePath) {
  const attachment = createInlineAttachment(1);
  attachment.size = 2 * 1024 * 1024 + 1;
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    bodyHtml: `<img src="cid:${attachment.contentId}">`,
    sessionMarker: marker,
    response: { ok: true, status: "attached-with-image-warnings" },
  });
  harness.send();

  assert.strictEqual(harness.attachmentContentCalls, 0);
  assert.strictEqual(harness.requests.length, 1);
  const body = JSON.parse(harness.requests[0].body);
  assert.deepStrictEqual(body.inlineImageRefs, []);
  assert.deepStrictEqual(body.inlineImageUploads, []);
  assert.strictEqual(harness.completions[0].allowEvent, true);
}

function testTimedOutReadsDoNotSpawnMore(runtime, runtimePath) {
  const attachments = Array.from({ length: 6 }, (_, index) => createInlineAttachment(index + 1));
  const harness = createHarness(runtime, runtimePath, {
    attachments,
    bodyHtml: attachments.map((attachment) => `<img src="cid:${attachment.contentId}">`).join(""),
    deferAttachmentContent: true,
    sessionMarker: marker,
    response: { ok: true, status: "attached-with-image-warnings" },
  });
  harness.send();
  assert.strictEqual(harness.attachmentContentCalls, 4);

  harness.fireTimer(2000);
  assert.strictEqual(harness.requests.length, 1);
  harness.completeAllAttachmentReads();
  assert.strictEqual(
    harness.attachmentContentCalls,
    4,
    "Timed-out Outlook reads must not start additional background attachment requests."
  );
  assert.strictEqual(harness.completions[0].allowEvent, true);
}

function testPrefetchedEmailAttachmentsAvoidContentReads(runtime, runtimePath) {
  const attachment = createOrdinaryAttachment(1);
  const fingerprint = getOrdinaryAttachmentFingerprint([attachment]);
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    sessionMarker: {
      ...marker,
      version: 3,
      emailAttachmentDecision: "include",
    },
    emailAttachmentState: {
      version: 2,
      draftItemId: marker.draftItemId,
      ticketId: marker.ticketId,
      operationId: marker.composeAttachId,
      emailAttachmentDecision: "include",
      emailAttachmentFingerprint: fingerprint,
      emailAttachmentPrefetchKey: "prefetched-email-attachments",
      emailAttachmentStagingVersion: 2,
      emailAttachmentSummary: {
        detected: 1,
        selected: 1,
        prepared: 1,
        attached: 0,
        skipped: 0,
        failed: 0,
        warnings: [],
      },
    },
    response: { ok: true, status: "attached" },
  });
  harness.send();
  assert.strictEqual(harness.attachmentContentCalls, 0);
  const body = JSON.parse(harness.requests[0].body);
  assert.strictEqual(body.includeEmailAttachments, true);
  assert.strictEqual(body.emailAttachmentPrefetchKey, "prefetched-email-attachments");
  assert.strictEqual(body.emailAttachmentFingerprint, fingerprint);
  assert.strictEqual(body.emailAttachmentStagingVersion, 2);
  assert.strictEqual(body.emailAttachmentOperationId, marker.composeAttachId);
}

function testExplicitEmailOnlyAvoidsContentReads(runtime, runtimePath) {
  const attachment = createOrdinaryAttachment(1);
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    sessionMarker: {
      ...marker,
      version: 3,
      emailAttachmentDecision: "exclude",
      emailAttachmentFingerprint: "stale-fingerprint",
    },
    response: { ok: true, status: "attached" },
  });
  harness.send();
  assert.strictEqual(harness.attachmentContentCalls, 0);
  const body = JSON.parse(harness.requests[0].body);
  assert.strictEqual(body.includeEmailAttachments, false);
}

function testStaleEmailAttachmentStateIsNotReused(runtime, runtimePath) {
  const attachment = createOrdinaryAttachment(1);
  const fingerprint = getOrdinaryAttachmentFingerprint([attachment]);
  const attachmentKey = getOrdinaryAttachmentKey(attachment);
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    conversationId: "mapped-conversation",
    emailAttachmentState: {
      version: 2,
      draftItemId: "copied-from-another-draft",
      emailAttachmentDecision: "include",
      emailAttachmentFingerprint: fingerprint,
      emailAttachmentPrefetchKey: "stale-prefetch-key",
      operationId: "stale-operation-id",
      ticketId: "1001",
    },
    response: { ok: false, status: "attachments-not-ready" },
  });
  harness.send();

  assert.strictEqual(harness.attachmentContentCalls, 0);
  assert.strictEqual(
    harness.requests.filter((request) => /email-attachments\/prefetch/.test(request.url)).length,
    0
  );
  const actionRequest = harness.requests.find((request) => /\/send-auto-attach$/.test(request.url));
  assert(actionRequest);
  assert.strictEqual(JSON.parse(actionRequest.body).emailAttachmentPrefetchKey, "");
  assert.strictEqual(harness.completions[0].allowEvent, false);
  const diagnostics = harness.diagnosticRequests.map((request) => JSON.parse(request.body));
  assert(
    diagnostics.some(
      (entry) => entry.stage === "attachment-state" && entry.outcome === "state-mismatch"
    )
  );
  assert(diagnostics.some((entry) => entry.stage === "request-complete"));
  assert.doesNotMatch(
    JSON.stringify(diagnostics),
    /copied-from-another-draft|stale-prefetch-key|document-1\.pdf/
  );
}

function testAutomaticPreparedAttachmentsBlockWithoutSession(runtime, runtimePath) {
  const attachment = createOrdinaryAttachment(1);
  const fingerprint = getOrdinaryAttachmentFingerprint([attachment]);
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    conversationId: "mapped-conversation",
    emailAttachmentState: {
      version: 2,
      draftItemId: marker.draftItemId,
      ticketId: "1001",
      operationId: "automatic-operation",
      emailAttachmentDecision: "include",
      emailAttachmentFingerprint: fingerprint,
      emailAttachmentPrefetchKey: "automatic-prefetch",
      emailAttachmentStagingVersion: 2,
      emailAttachmentSummary: {
        detected: 1,
        selected: 1,
        prepared: 1,
        attached: 0,
        skipped: 0,
        failed: 0,
        warnings: [],
      },
    },
    response: { ok: true, status: "no-session" },
  });
  harness.send();

  assert.strictEqual(harness.attachmentContentCalls, 0);
  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, false);
}

function testAutomaticPreparedAttachmentsBlockOnWatchdog(runtime, runtimePath) {
  const attachment = createOrdinaryAttachment(1);
  const fingerprint = getOrdinaryAttachmentFingerprint([attachment]);
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    conversationId: "mapped-conversation",
    emailAttachmentState: {
      version: 2,
      draftItemId: marker.draftItemId,
      ticketId: "1001",
      operationId: "automatic-timeout-operation",
      emailAttachmentDecision: "include",
      emailAttachmentFingerprint: fingerprint,
      emailAttachmentPrefetchKey: "automatic-timeout-prefetch",
      emailAttachmentStagingVersion: 2,
      emailAttachmentSummary: {
        detected: 1,
        selected: 1,
        prepared: 1,
        attached: 0,
        skipped: 0,
        failed: 0,
        warnings: [],
      },
    },
    hangRequest: true,
  });
  harness.send();
  harness.fireTimer(25000);

  assert.strictEqual(harness.completions.length, 1);
  assert.strictEqual(harness.completions[0].allowEvent, false);
}

function testMissingEmailAttachmentStateBlocksWithoutReads(runtime, runtimePath) {
  const attachments = Array.from({ length: 5 }, (_, index) => createOrdinaryAttachment(index + 1));
  const harness = createHarness(runtime, runtimePath, {
    attachments,
    sessionMarker: { ...marker, version: 3, emailAttachmentDecision: "include" },
    response: { ok: false, status: "attachments-not-ready", ticketNumber: "T1001" },
  });
  harness.send();
  assert.strictEqual(harness.attachmentContentCalls, 0);
  assert.strictEqual(harness.requests.filter((request) => /\/items$/.test(request.url)).length, 0);
  const actionRequest = harness.requests.find((request) => /\/sent-email$/.test(request.url));
  assert(actionRequest);
  assert.strictEqual(JSON.parse(actionRequest.body).emailAttachmentPrefetchKey, "");
  assert.strictEqual(harness.completions[0].allowEvent, false);
}

function testAttachmentsChangedPrefetchesAutomaticReply(runtime, runtimePath) {
  const attachment = createOrdinaryAttachment(1);
  const key = getOrdinaryAttachmentKey(attachment);
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    conversationId: "mapped-conversation",
    responseForRequest(request) {
      if (/\/prefetch\/start$/.test(request.url)) {
        return {
          ok: true,
          status: "ready",
          prefetchKey: "event-prefetch",
          pendingAttachmentKeys: [key],
          ticketId: "1001",
        };
      }
      return { ok: true, status: "prepared" };
    },
  });
  harness.attachmentsChanged();
  assert.strictEqual(harness.attachmentChangeCompletions, 1);
  assert.strictEqual(harness.attachmentContentCalls, 1);
  assert.strictEqual(harness.requests.length, 2);
  assert.match(harness.requests[0].url, /\/prefetch\/start$/);
  assert.match(harness.requests[1].url, /\/items$/);
  const diagnostics = harness.diagnosticRequests.map((request) => JSON.parse(request.body));
  assert(diagnostics.some((entry) => entry.stage === "attachment-change-event"));
  assert(diagnostics.some((entry) => entry.stage === "attachment-inventory"));
  assert(
    diagnostics.some((entry) => entry.stage === "attachment-staging" && entry.outcome === "ready")
  );
  assert(diagnostics.some((entry) => entry.stage === "attachment-prefetch-complete"));
  assert.doesNotMatch(
    JSON.stringify(diagnostics),
    /event-prefetch|document-1\.pdf|mapped-conversation/
  );
}

function testAttachmentsChangedKeepsUnreadEligibleFilePending(runtime, runtimePath) {
  const attachments = [createOrdinaryAttachment(1), createOrdinaryAttachment(2)];
  const readableKey = getOrdinaryAttachmentKey(attachments[0]);
  const harness = createHarness(runtime, runtimePath, {
    attachments,
    failAttachmentContentIds: [attachments[1].id],
    sessionMarker: { ...marker, version: 3, emailAttachmentDecision: "include" },
    responseForRequest(request) {
      if (/\/prefetch\/start$/.test(request.url)) {
        const body = JSON.parse(request.body);
        assert.strictEqual(
          body.emailAttachments.length,
          2,
          "The unread eligible file must remain in the server inventory as pending."
        );
        return {
          aggregate: { failed: 0, pending: 2, prepared: 0, selected: 2 },
          ok: true,
          status: "pending",
          prefetchKey: "unread-prefetch",
          pendingAttachmentKeys: body.emailAttachments.map((entry) => entry.attachmentKey),
          ticketId: "1001",
        };
      }
      const body = JSON.parse(request.body);
      assert.strictEqual(body.attachmentKey, readableKey);
      return { ok: true, status: "prepared" };
    },
  });
  harness.attachmentsChanged();
  harness.fireTimer(150);
  harness.fireTimer(350);
  assert.strictEqual(harness.attachmentChangeCompletions, 1);
  assert.strictEqual(harness.attachmentContentCalls, 4);
  assert.strictEqual(harness.requests.filter((request) => /\/items$/.test(request.url)).length, 1);
  const state = JSON.parse(harness.customPropertyValues["halo.emailAttachmentPrefetch.v2"]);
  assert.strictEqual(state.emailAttachmentSummary.selected, 2);
  assert.strictEqual(state.emailAttachmentSummary.prepared, 1);
  assert.strictEqual(state.emailAttachmentSummary.failed, 1);
}

function testAttachmentsChangedBoundsHungContentReads(runtime, runtimePath) {
  const attachment = createOrdinaryAttachment(1);
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    conversationId: "mapped-conversation",
    deferAttachmentContent: true,
    responseForRequest(request) {
      if (/\/prefetch\/start$/.test(request.url)) {
        const body = JSON.parse(request.body);
        return {
          aggregate: { failed: 0, pending: 1, prepared: 0, selected: 1 },
          ok: true,
          status: "pending",
          prefetchKey: "hung-content-prefetch",
          pendingAttachmentKeys: body.emailAttachments.map((entry) => entry.attachmentKey),
          ticketId: "1001",
        };
      }
      return { ok: true, status: "prepared" };
    },
  });
  harness.attachmentsChanged();
  harness.fireTimer(15000);
  harness.fireTimer(150);
  harness.fireTimer(15000);
  harness.fireTimer(350);
  harness.fireTimer(15000);

  assert.strictEqual(harness.attachmentContentCalls, 3);
  assert.strictEqual(harness.attachmentChangeCompletions, 1);
  assert.strictEqual(harness.requests.filter((request) => /\/items$/.test(request.url)).length, 0);
  harness.completeAllAttachmentReads();
  assert.strictEqual(
    harness.attachmentChangeCompletions,
    1,
    "Late Outlook callbacks must be ignored after the bounded read has completed."
  );
}

function testAttachmentsChangedRespectsExplicitExclusion(runtime, runtimePath) {
  const harness = createHarness(runtime, runtimePath, {
    attachments: [createOrdinaryAttachment(1)],
    sessionMarker: { ...marker, version: 3, emailAttachmentDecision: "exclude" },
  });
  harness.attachmentsChanged();
  assert.strictEqual(harness.attachmentChangeCompletions, 1);
  assert.strictEqual(harness.attachmentContentCalls, 0);
  assert.strictEqual(harness.itemSaveCalls, 2);
  assert.strictEqual(harness.requests.length, 0);
}

function testReferencedAttachmentIdIsNotOrdinary(runtime, runtimePath) {
  const attachment = createInlineAttachment(1);
  attachment.id = "fallback-logo@example.com";
  attachment.contentId = "";
  delete attachment.isInline;
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment],
    bodyHtml: '<img src="cid:fallback-logo@example.com">',
    sessionMarker: marker,
    response: { ok: true, status: "attached" },
  });
  harness.send();
  assert.strictEqual(harness.attachmentContentCalls, 1);
  assert.strictEqual(
    harness.requests.filter((request) => /email-attachments\/prefetch/.test(request.url)).length,
    0,
    "A CID represented by Outlook's attachment ID must not trigger the ordinary prompt pipeline."
  );
}

function testDuplicateOrdinaryMetadataIsReadOnce(runtime, runtimePath) {
  const attachment = createOrdinaryAttachment(1);
  const key = getOrdinaryAttachmentKey(attachment);
  const harness = createHarness(runtime, runtimePath, {
    attachments: [attachment, { ...attachment }],
    sessionMarker: { ...marker, version: 3, emailAttachmentDecision: "include" },
    responseForRequest(request) {
      if (/\/prefetch\/start$/.test(request.url)) {
        const body = JSON.parse(request.body);
        assert.strictEqual(body.emailAttachments.length, 1);
        return {
          ok: true,
          status: "ready",
          prefetchKey: "duplicate-prefetch",
          pendingAttachmentKeys: [key],
          ticketId: "1001",
        };
      }
      if (/\/items$/.test(request.url)) {
        return { ok: true, status: "prepared" };
      }
      return { ok: true, status: "attached" };
    },
  });
  harness.attachmentsChanged();
  assert.strictEqual(harness.attachmentContentCalls, 1);
}

function testEmlIcsAndCloudAttachmentConversion(runtime, runtimePath) {
  const eml = {
    ...createOrdinaryAttachment(1),
    content: "From: test@example.com\r\nSubject: Nested\r\n\r\nHello",
    contentFormat: "eml",
    contentType: "message/rfc822",
    name: "nested-message",
  };
  const calendar = {
    ...createOrdinaryAttachment(2),
    content: "BEGIN:VCALENDAR\r\nEND:VCALENDAR",
    contentFormat: "icalendar",
    contentType: "text/calendar",
    name: "meeting",
  };
  const cloud = {
    ...createOrdinaryAttachment(3),
    content: "https://sharepoint.example/file",
    contentFormat: "url",
    name: "cloud-link.docx",
  };
  const descriptorsByKey = new Map(
    [eml, calendar].map((attachment) => [getOrdinaryAttachmentKey(attachment), attachment])
  );
  const harness = createHarness(runtime, runtimePath, {
    attachments: [eml, calendar, cloud],
    sessionMarker: { ...marker, version: 3, emailAttachmentDecision: "include" },
    responseForRequest(request) {
      if (/\/prefetch\/start$/.test(request.url)) {
        const body = JSON.parse(request.body);
        assert.strictEqual(body.emailAttachments.length, 2);
        assert(body.emailAttachments.some((value) => value.name === "nested-message.eml"));
        assert(body.emailAttachments.some((value) => value.name === "meeting.ics"));
        return {
          ok: true,
          status: "ready",
          prefetchKey: "format-prefetch",
          pendingAttachmentKeys: Array.from(descriptorsByKey.keys()),
          ticketId: "1001",
        };
      }
      if (/\/items$/.test(request.url)) {
        const body = JSON.parse(request.body);
        const original = descriptorsByKey.get(body.attachmentKey);
        assert(original);
        assert.strictEqual(
          Buffer.from(body.contentBase64, "base64").toString("utf8"),
          original.content
        );
        return { ok: true, status: "prepared" };
      }
      return { ok: true, status: "attached-with-attachment-warnings" };
    },
  });
  harness.attachmentsChanged();
  assert.strictEqual(harness.attachmentContentCalls, 3);
  assert.strictEqual(harness.requests.filter((request) => /\/items$/.test(request.url)).length, 2);
}

function createInlineAttachment(index) {
  const bytes = Buffer.concat([
    Buffer.from("89504e470d0a1a0a", "hex"),
    Buffer.from(`image-${index}`),
  ]);
  return {
    contentBase64: bytes.toString("base64"),
    contentId: `logo${index}@example.com`,
    id: `attachment-${index}`,
    isInline: true,
    name: `logo-${index}.png`,
    size: bytes.length,
    attachmentType: "file",
  };
}

function createOrdinaryAttachment(index) {
  const bytes = Buffer.from(`ordinary-attachment-${index}`);
  return {
    attachmentType: "file",
    contentBase64: bytes.toString("base64"),
    contentId: "",
    contentType: "application/pdf",
    id: `ordinary-${index}`,
    isInline: false,
    name: `document-${index}.pdf`,
    size: bytes.length,
  };
}

function getOrdinaryAttachmentKey(attachment, duplicateIndex = 1) {
  const descriptor = getOrdinaryAttachmentDescriptor(attachment);
  return sha256Hex(`${descriptor}\u0000${duplicateIndex}`);
}

function getOrdinaryAttachmentFingerprint(attachments) {
  const counts = new Map();
  attachments.forEach((attachment) => {
    const descriptor = getOrdinaryAttachmentDescriptor(attachment);
    counts.set(descriptor, (counts.get(descriptor) || 0) + 1);
  });
  return sha256Hex(
    Array.from(counts.entries())
      .map(([descriptor, count]) => `${descriptor}\u0000${count}`)
      .sort()
      .join("\u0001")
  );
}

function getOrdinaryAttachmentDescriptor(attachment) {
  return [
    String(attachment.name || "").toLowerCase(),
    attachment.size,
    String(attachment.contentType || "").toLowerCase(),
    String(attachment.attachmentType || "").toLowerCase(),
  ].join("\u0000");
}

function sha256Hex(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

function createHarness(runtime, runtimePath, options = {}) {
  const completions = [];
  const diagnosticRequests = [];
  let attachmentChangeCompletions = 0;
  const requests = [];
  const timers = [];
  let sendHandler = null;
  let attachmentsChangedHandler = null;
  let attachmentContentCalls = 0;
  let activeAttachmentReads = 0;
  let maxConcurrentAttachmentReads = 0;
  let itemSaveCalls = 0;
  const pendingAttachmentReads = [];
  const attachmentValues = (options.attachments || []).map((attachment) => ({ ...attachment }));
  const customPropertyValues = {
    "halo.composeAttach.v1": options.customMarker ? JSON.stringify(options.customMarker) : "",
    "halo.emailAttachmentPrefetch.v2": options.emailAttachmentState
      ? JSON.stringify(options.emailAttachmentState)
      : "",
  };
  const sessionValues = {};
  const itemId = options.itemId || marker.draftItemId;
  const cryptoWindow = options.legacyCrypto
    ? {
        msCrypto: {
          subtle: {
            digest(_algorithm, bytes) {
              const digest = crypto.createHash("sha256").update(Buffer.from(bytes)).digest();
              const operation = {};
              Object.defineProperty(operation, "oncomplete", {
                set(handler) {
                  handler({ target: { result: new Uint8Array(digest).buffer } });
                },
              });
              return operation;
            },
          },
        },
      }
    : {
        crypto: {
          subtle: {
            digest(_algorithm, bytes) {
              const digest = crypto.createHash("sha256").update(Buffer.from(bytes)).digest();
              return {
                then(resolve) {
                  resolve(new Uint8Array(digest).buffer);
                },
              };
            },
          },
        },
      };

  const success = (value) => ({ status: "succeeded", value });
  const item = {
    attachments: attachmentValues,
    body: {
      getAsync(coercionType, callback) {
        if (!options.hangBodyRead) {
          callback(
            success(
              coercionType === "html"
                ? options.bodyHtml || "<p>Composed email</p>"
                : "Composed email"
            )
          );
        }
      },
    },
    cc: {
      getAsync(callback) {
        callback(success(options.cc || []));
      },
    },
    conversationId: options.conversationId || "",
    inReplyTo: options.inReplyTo || "",
    internetHeaders: options.headerComposeId
      ? {
          getAsync(names, callback) {
            callback(success({ [names[0]]: options.headerComposeId }));
          },
        }
      : undefined,
    from: {
      getAsync(callback) {
        callback(success({ displayName: "Sender", emailAddress: "sender@example.com" }));
      },
    },
    getItemIdAsync(callback) {
      callback(success(itemId));
    },
    getAttachmentsAsync(callback) {
      callback(success(item.attachments));
    },
    getAttachmentContentAsync(attachmentId, callback) {
      attachmentContentCalls += 1;
      activeAttachmentReads += 1;
      maxConcurrentAttachmentReads = Math.max(maxConcurrentAttachmentReads, activeAttachmentReads);
      if (
        Array.isArray(options.failAttachmentContentIds) &&
        options.failAttachmentContentIds.includes(attachmentId)
      ) {
        activeAttachmentReads -= 1;
        callback({
          status: "failed",
          error: { code: "AttachmentNotReady", message: "Attachment content is not ready." },
        });
        return;
      }
      if (options.replaceAttachmentIdOnFirstRead && attachmentContentCalls === 1) {
        attachmentValues[0] = {
          ...attachmentValues[0],
          id: options.replaceAttachmentIdOnFirstRead,
        };
        item.attachments = attachmentValues;
        activeAttachmentReads -= 1;
        callback({
          status: "failed",
          error: { code: "InvalidAttachmentId", message: "Attachment identifier changed." },
        });
        return;
      }
      const attachment = attachmentValues.find((entry) => entry.id === attachmentId);
      if (!attachment) {
        activeAttachmentReads -= 1;
        callback({
          status: "failed",
          error: { code: "InvalidAttachmentId", message: "Attachment identifier was not found." },
        });
        return;
      }
      const complete = () => {
        activeAttachmentReads -= 1;
        callback(
          success({
            content:
              attachment.content === undefined ? attachment.contentBase64 : attachment.content,
            format: attachment.contentFormat || "base64",
          })
        );
      };
      if (options.deferAttachmentContent) {
        pendingAttachmentReads.push(complete);
      } else {
        complete();
      }
    },
    itemType: "message",
    loadCustomPropertiesAsync(callback) {
      callback(
        success({
          get(name) {
            return customPropertyValues[name] || "";
          },
          set(name, value) {
            customPropertyValues[name] = value;
          },
          saveAsync(saveCallback) {
            saveCallback(success());
          },
        })
      );
    },
    saveAsync(callback) {
      itemSaveCalls += 1;
      callback(success(itemId));
    },
    sessionData: {
      getAsync(name, callback) {
        if (sessionValues[name] !== undefined) {
          callback(success(sessionValues[name]));
          return;
        }
        if (name === "halo.composeAttach.v1") {
          callback(
            success(
              options.sessionMarkerRaw !== undefined
                ? options.sessionMarkerRaw
                : options.sessionMarker
                  ? JSON.stringify(options.sessionMarker)
                  : ""
            )
          );
          return;
        }
        callback(
          success(options.emailAttachmentState ? JSON.stringify(options.emailAttachmentState) : "")
        );
      },
      setAsync(name, value, callback) {
        sessionValues[name] = value;
        callback(success());
      },
    },
    subject: {
      getAsync(callback) {
        callback(success("New subject"));
      },
    },
    to: {
      getAsync(callback) {
        callback(success([{ displayName: "Recipient", emailAddress: "recipient@example.com" }]));
      },
    },
  };

  function XMLHttpRequest() {
    this.headers = {};
    this.open = (method, url) => {
      this.method = method;
      this.url = url;
    };
    this.setRequestHeader = (name, value) => {
      this.headers[name] = value;
    };
    this.send = (body) => {
      if (this.url.includes("/api/diagnostics/send-event")) {
        diagnosticRequests.push({ body, method: this.method, url: this.url });
        this.status = 200;
        this.readyState = 4;
        this.responseText = JSON.stringify({ ok: true });
        this.onreadystatechange();
        return;
      }
      if (this.url.includes("/api/halo/email/recover-mapping")) {
        this.status = 200;
        this.readyState = 4;
        this.responseText = JSON.stringify(
          options.responseForRecovery
            ? options.responseForRecovery({ body, method: this.method, url: this.url })
            : options.recoveryResponse || {
                ok: true,
                status: options.conversationId || options.inReplyTo ? "matched" : "no-match",
                candidateIndex: -1,
                ticketId: "1001",
                ticketNumber: "T1001",
                actionMode: "email",
              }
        );
        this.onreadystatechange();
        return;
      }
      requests.push({ body, method: this.method, url: this.url });
      if (options.hangRequest) {
        return;
      }

      this.status = 200;
      this.readyState = 4;
      this.responseText = JSON.stringify(
        options.responseForRequest
          ? options.responseForRequest(requests[requests.length - 1])
          : options.response || { ok: true, status: "no-match" }
      );
      this.onreadystatechange();
    };
  }

  const Office = {
    actions: {
      associate(name, handler) {
        if (name === "onHaloMessageSend") {
          sendHandler = handler;
        } else if (name === "onHaloMessageAttachmentsChanged") {
          attachmentsChangedHandler = handler;
        } else {
          assert.fail(`Unexpected event association: ${name}`);
        }
      },
    },
    AsyncResultStatus: { Succeeded: "succeeded" },
    CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: { ItemType: { Message: "message" } },
    context: {
      mailbox: {
        item,
        makeEwsRequestAsync: options.ewsResponder
          ? (request, callback) => callback(success(options.ewsResponder(request)))
          : undefined,
        userProfile: {
          displayName: "Sender",
          emailAddress:
            options.userProfileEmail === undefined
              ? "sender@example.com"
              : options.userProfileEmail,
        },
      },
      roamingSettings: {
        get() {
          return options.roamingBackgroundSession === undefined
            ? "background-session"
            : options.roamingBackgroundSession;
        },
      },
    },
  };

  const context = vm.createContext({
    Array,
    Date,
    Error,
    Intl,
    JSON,
    Object,
    Office,
    String,
    Uint8Array,
    XMLHttpRequest,
    clearTimeout(id) {
      const timer = timers.find((entry) => entry.id === id);
      if (timer) {
        timer.cleared = true;
      }
    },
    encodeURIComponent,
    decodeURIComponent,
    unescape,
    atob(value) {
      return Buffer.from(value, "base64").toString("binary");
    },
    btoa(value) {
      return Buffer.from(value, "binary").toString("base64");
    },
    window: cryptoWindow,
    setTimeout(callback, delay) {
      const timer = { callback, cleared: false, delay, id: timers.length + 1 };
      timers.push(timer);
      return timer.id;
    },
  });

  vm.runInContext(runtime, context, { filename: runtimePath });
  assert.strictEqual(typeof sendHandler, "function");
  assert.strictEqual(typeof attachmentsChangedHandler, "function");

  return {
    get attachmentContentCalls() {
      return attachmentContentCalls;
    },
    completions,
    get attachmentChangeCompletions() {
      return attachmentChangeCompletions;
    },
    get itemSaveCalls() {
      return itemSaveCalls;
    },
    attachmentsChanged() {
      attachmentsChangedHandler({
        completed() {
          attachmentChangeCompletions += 1;
        },
      });
    },
    completeAllAttachmentReads() {
      while (pendingAttachmentReads.length) {
        pendingAttachmentReads.shift()();
      }
    },
    fireTimer(delay) {
      const timer = timers.find((entry) => entry.delay === delay && !entry.cleared);
      assert(timer, `Expected an active ${delay}ms timer.`);
      timer.callback();
    },
    requests,
    diagnosticRequests,
    customPropertyValues,
    get maxConcurrentAttachmentReads() {
      return maxConcurrentAttachmentReads;
    },
    sessionValues,
    send() {
      sendHandler({
        completed(options) {
          completions.push(options);
        },
      });
    },
  };
}

run();
