const assert = require("node:assert");
const crypto = require("node:crypto");
const { createTestStore } = require("./testDatabase");
const { createTokenCrypto } = require("./tokenCrypto");
const {
  EMAIL_ATTACHMENT_PREFETCH_TTL_MS,
  MAX_EMAIL_ATTACHMENTS,
  MAX_EMAIL_ATTACHMENT_BYTES,
  MAX_EMAIL_ATTACHMENT_TOTAL_BYTES,
  decryptStagedAttachment,
  getEmailAttachmentPreparationStatus,
  resolveEmailAttachments,
  sanitizeAttachmentName,
  stageEmailAttachmentItem,
  startEmailAttachmentPrefetch,
} = require("./emailAttachments");

function hash(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

function descriptor(index, size = 8) {
  return {
    attachmentKey: hash(`attachment-${index}`),
    attachmentType: "file",
    contentType: "application/octet-stream",
    name: `document-${index}.bin`,
    reportedSize: size,
  };
}

function descriptorWithContent(index, content) {
  return {
    ...descriptor(index, content.length),
    contentSha256: hash(content),
  };
}

function finalInput(record, selected) {
  return {
    includeEmailAttachments: true,
    emailAttachmentDraftItemId: record.draftItemId,
    emailAttachmentFingerprint: record.attachmentFingerprint,
    emailAttachmentOperationId: record.operationId,
    emailAttachmentPrefetchKey: record.prefetchKey,
    emailAttachmentStagingVersion: 2,
    emailAttachmentSummary: {
      attached: 0,
      detected: selected,
      failed: 0,
      prepared: selected,
      selected,
      skipped: 0,
      warnings: [],
    },
  };
}

async function run() {
  assert.strictEqual(sanitizeAttachmentName("../folder/report.pdf"), "report.pdf");
  assert.strictEqual(sanitizeAttachmentName("..\\folder\\\u0000name.txt"), "name.txt");
  assert.strictEqual(sanitizeAttachmentName(".."), "email-attachment.bin");
  assert.strictEqual(EMAIL_ATTACHMENT_PREFETCH_TTL_MS, 7 * 24 * 60 * 60 * 1000);

  const store = await createTestStore();
  const tokenCrypto = createTokenCrypto({ NODE_ENV: "test" });
  const user = await store.upsertUser({
    tenantId: "microsoft-tenant",
    objectId: "user-one",
    email: "one@example.com",
  });
  const secondUser = await store.upsertUser({
    tenantId: "microsoft-tenant",
    objectId: "user-two",
    email: "two@example.com",
  });
  const haloTenant = "https://customer.halopsa.com";
  const contents = [Buffer.from("business attachment"), Buffer.from("second attachment")];
  const descriptors = [
    descriptorWithContent(1, contents[0]),
    descriptorWithContent(2, contents[1]),
  ];
  descriptors[0].contentType = "Application/PDF; charset=binary";
  const first = await startEmailAttachmentPrefetch({
    attachmentFingerprint: hash("fingerprint-one"),
    descriptors,
    draftItemId: "saved-draft-one",
    haloTenant,
    operationId: "operation-one",
    store,
    ticketId: 1001,
    userId: user.id,
  });
  assert.strictEqual(first.stagingVersion, 2);
  assert.strictEqual(first.pendingAttachmentKeys.length, 2);

  for (let index = 0; index < descriptors.length; index += 1) {
    const contentBase64 = contents[index].toString("base64");
    const prepared = await stageEmailAttachmentItem({
      attachmentKey: descriptors[index].attachmentKey,
      contentBase64,
      contentFormat: "base64",
      contentSha256: hash(contents[index]),
      haloTenant,
      prefetchKey: first.prefetchKey,
      store,
      ticketId: 1001,
      tokenCrypto,
      userId: user.id,
    });
    assert.strictEqual(prepared.status, "prepared");
  }

  const duplicate = await stageEmailAttachmentItem({
    attachmentKey: descriptors[0].attachmentKey,
    contentBase64: contents[0].toString("base64"),
    contentFormat: "base64",
    contentSha256: hash(contents[0]),
    haloTenant,
    prefetchKey: first.prefetchKey,
    store,
    ticketId: 1001,
    tokenCrypto,
    userId: user.id,
  });
  assert.strictEqual(duplicate.status, "already-prepared");

  const stagedRecord = await store.getEmailAttachmentPrefetch(first.prefetchKey, {
    haloTenant,
    ticketId: 1001,
    userId: user.id,
  });
  assert.strictEqual(stagedRecord.stagingVersion, 2);
  assert.strictEqual(
    stagedRecord.items.every((item) => item.status === "prepared"),
    true
  );
  assert.strictEqual(
    stagedRecord.items.every((item) => Buffer.isBuffer(item.contentCiphertext)),
    true
  );
  assert.strictEqual(
    stagedRecord.items.every((item) => !item.haloAttachmentId),
    true
  );
  assert.strictEqual(
    stagedRecord.items[0].contentCiphertext.includes(contents[0]),
    false,
    "PostgreSQL staging must not contain plaintext bytes."
  );

  const status = await getEmailAttachmentPreparationStatus({
    haloTenant,
    prefetchKey: first.prefetchKey,
    store,
    ticketId: 1001,
    userId: user.id,
  });
  assert.deepStrictEqual(status.aggregate, { failed: 0, pending: 0, prepared: 2, selected: 2 });

  await assert.rejects(
    () =>
      resolveEmailAttachments({
        haloTenant,
        input: finalInput(stagedRecord, 2),
        store,
        ticketId: 1001,
        userId: secondUser.id,
      }),
    (error) =>
      error.code === "attachments-not-ready" &&
      error.attachmentDiagnosticReason === "stage-unavailable"
  );

  const copiedDraftInput = finalInput(stagedRecord, 2);
  copiedDraftInput.emailAttachmentDraftItemId = "a-different-saved-draft";
  await assert.rejects(
    () =>
      resolveEmailAttachments({
        haloTenant,
        input: copiedDraftInput,
        store,
        ticketId: 1001,
        userId: user.id,
      }),
    (error) =>
      error.code === "attachments-not-ready" &&
      error.attachmentDiagnosticReason === "inventory-mismatch"
  );

  const legacyInput = finalInput(stagedRecord, 2);
  legacyInput.emailAttachmentStagingVersion = 1;
  await assert.rejects(
    () =>
      resolveEmailAttachments({
        haloTenant,
        input: legacyInput,
        store,
        ticketId: 1001,
        userId: user.id,
      }),
    (error) =>
      error.code === "attachments-not-ready" && error.attachmentDiagnosticReason === "legacy-state"
  );

  const resolved = await resolveEmailAttachments({
    haloTenant,
    input: finalInput(stagedRecord, 2),
    store,
    ticketId: 1001,
    userId: user.id,
  });
  assert.strictEqual(resolved.stagedItems.length, 2);
  assert.strictEqual(resolved.summary.prepared, 2);
  const decrypted = decryptStagedAttachment(tokenCrypto, resolved.record, resolved.stagedItems[0]);
  assert(decrypted.equals(contents[0]) || decrypted.equals(contents[1]));
  decrypted.fill(0);
  await store.releaseEmailAttachmentPrefetchCommit(first.prefetchKey);

  const mismatchInput = finalInput(stagedRecord, 2);
  mismatchInput.emailAttachmentFingerprint = hash("different-inventory");
  await assert.rejects(
    () =>
      resolveEmailAttachments({
        haloTenant,
        input: mismatchInput,
        store,
        ticketId: 1001,
        userId: user.id,
      }),
    (error) =>
      error.code === "attachments-not-ready" &&
      error.attachmentDiagnosticReason === "inventory-mismatch"
  );

  assert.strictEqual(await store.consumeEmailAttachmentPrefetch(first.prefetchKey), 1);
  const consumed = await store.getEmailAttachmentPrefetch(first.prefetchKey, {
    haloTenant,
    ticketId: 1001,
    userId: user.id,
  });
  assert.strictEqual(consumed.status, "consumed");
  assert.strictEqual(
    consumed.items.every((item) => item.contentCiphertext === null),
    true
  );

  const second = await startEmailAttachmentPrefetch({
    attachmentFingerprint: hash("fingerprint-two"),
    descriptors: [descriptors[0]],
    draftItemId: "saved-draft-two",
    haloTenant,
    operationId: "operation-two",
    store,
    ticketId: 1002,
    userId: user.id,
  });
  await stageEmailAttachmentItem({
    attachmentKey: descriptors[0].attachmentKey,
    contentBase64: contents[0].toString("base64"),
    contentFormat: "base64",
    contentSha256: hash(contents[0]),
    haloTenant,
    prefetchKey: second.prefetchKey,
    store,
    ticketId: 1002,
    tokenCrypto,
    userId: user.id,
  });
  const reconciled = await startEmailAttachmentPrefetch({
    attachmentFingerprint: hash("fingerprint-two-with-file-removed"),
    descriptors: [],
    draftItemId: "saved-draft-two",
    haloTenant,
    operationId: "operation-two",
    store,
    ticketId: 1002,
    userId: user.id,
  });
  assert.strictEqual(reconciled.prefetchKey, second.prefetchKey);
  const removedRecord = await store.getEmailAttachmentPrefetch(second.prefetchKey, {
    haloTenant,
    ticketId: 1002,
    userId: user.id,
  });
  assert.strictEqual(removedRecord.items[0].status, "removed");
  assert.strictEqual(removedRecord.items[0].contentCiphertext, null);

  const changingDescriptor = {
    ...descriptor(150),
    contentSha256: hash(contents[0]),
  };
  const changing = await startEmailAttachmentPrefetch({
    attachmentFingerprint: hash("same-metadata-fingerprint"),
    descriptors: [changingDescriptor],
    draftItemId: "saved-changing-draft",
    haloTenant,
    operationId: "changing-operation",
    store,
    ticketId: 1004,
    userId: user.id,
  });
  await stageEmailAttachmentItem({
    attachmentKey: changingDescriptor.attachmentKey,
    contentBase64: contents[0].toString("base64"),
    contentFormat: "base64",
    contentSha256: hash(contents[0]),
    haloTenant,
    prefetchKey: changing.prefetchKey,
    store,
    ticketId: 1004,
    tokenCrypto,
    userId: user.id,
  });
  const unchangedContent = await startEmailAttachmentPrefetch({
    attachmentFingerprint: hash("same-metadata-fingerprint"),
    descriptors: [changingDescriptor],
    draftItemId: "saved-changing-draft",
    haloTenant,
    operationId: "changing-operation",
    store,
    ticketId: 1004,
    userId: user.id,
  });
  assert.deepStrictEqual(unchangedContent.pendingAttachmentKeys, []);

  const changedContent = await startEmailAttachmentPrefetch({
    attachmentFingerprint: hash("same-metadata-fingerprint"),
    descriptors: [{ ...changingDescriptor, contentSha256: hash(contents[1]) }],
    draftItemId: "saved-changing-draft",
    haloTenant,
    operationId: "changing-operation",
    store,
    ticketId: 1004,
    userId: user.id,
  });
  assert.deepStrictEqual(changedContent.pendingAttachmentKeys, [changingDescriptor.attachmentKey]);
  const changedPendingRecord = await store.getEmailAttachmentPrefetch(changing.prefetchKey, {
    haloTenant,
    ticketId: 1004,
    userId: user.id,
  });
  assert.strictEqual(changedPendingRecord.items[0].status, "pending");
  assert.strictEqual(changedPendingRecord.items[0].contentCiphertext, null);

  await assert.rejects(
    () =>
      stageEmailAttachmentItem({
        attachmentKey: changingDescriptor.attachmentKey,
        contentBase64: contents[0].toString("base64"),
        contentFormat: "base64",
        contentSha256: hash(contents[0]),
        haloTenant,
        prefetchKey: changing.prefetchKey,
        store,
        ticketId: 1004,
        tokenCrypto,
        userId: user.id,
      }),
    (error) =>
      error.code === "attachments-not-ready" &&
      error.attachmentDiagnosticReason === "content-changed" &&
      /content changed/.test(error.message)
  );
  const delayedOldUploadRecord = await store.getEmailAttachmentPrefetch(changing.prefetchKey, {
    haloTenant,
    ticketId: 1004,
    userId: user.id,
  });
  assert.strictEqual(delayedOldUploadRecord.items[0].status, "pending");
  assert.strictEqual(delayedOldUploadRecord.items[0].contentCiphertext, null);

  const currentClaim = await store.claimEmailAttachmentPrefetchItem(
    changing.prefetchKey,
    changingDescriptor.attachmentKey,
    contents[1].length,
    hash(contents[1]),
    MAX_EMAIL_ATTACHMENT_TOTAL_BYTES
  );
  assert.strictEqual(currentClaim.status, "preparing");
  assert.strictEqual(
    await store.saveEmailAttachmentPrefetchItemFailure(
      changing.prefetchKey,
      changingDescriptor.attachmentKey,
      hash(contents[0]),
      "stale-upload-failure"
    ),
    0
  );
  const afterStaleFailure = await store.getEmailAttachmentPrefetch(changing.prefetchKey, {
    haloTenant,
    ticketId: 1004,
    userId: user.id,
  });
  assert.strictEqual(afterStaleFailure.items[0].status, "preparing");
  assert.strictEqual(
    await store.saveEmailAttachmentPrefetchItemFailure(
      changing.prefetchKey,
      changingDescriptor.attachmentKey,
      hash(contents[1]),
      "current-upload-retry"
    ),
    1
  );

  await stageEmailAttachmentItem({
    attachmentKey: changingDescriptor.attachmentKey,
    contentBase64: contents[1].toString("base64"),
    contentFormat: "base64",
    contentSha256: hash(contents[1]),
    haloTenant,
    prefetchKey: changing.prefetchKey,
    store,
    ticketId: 1004,
    tokenCrypto,
    userId: user.id,
  });
  const changedPreparedRecord = await store.getEmailAttachmentPrefetch(changing.prefetchKey, {
    haloTenant,
    ticketId: 1004,
    userId: user.id,
  });
  assert.strictEqual(changedPreparedRecord.items[0].status, "prepared");
  assert.strictEqual(changedPreparedRecord.items[0].contentSha256, hash(contents[1]));

  const expiringDescriptor = descriptorWithContent(200, contents[0]);
  const expiring = await startEmailAttachmentPrefetch({
    attachmentFingerprint: hash("expiring-fingerprint"),
    descriptors: [expiringDescriptor],
    draftItemId: "saved-expiring-draft",
    haloTenant,
    operationId: "expiring-operation",
    store,
    ticketId: 1003,
    userId: user.id,
  });
  await stageEmailAttachmentItem({
    attachmentKey: expiringDescriptor.attachmentKey,
    contentBase64: contents[0].toString("base64"),
    contentFormat: "base64",
    contentSha256: hash(contents[0]),
    haloTenant,
    prefetchKey: expiring.prefetchKey,
    store,
    ticketId: 1003,
    tokenCrypto,
    userId: user.id,
  });
  await store.upsertEmailAttachmentPrefetch(
    {
      attachmentFingerprint: hash("expiring-fingerprint"),
      draftItemId: "saved-expiring-draft",
      expectedBytes: expiringDescriptor.reportedSize,
      expectedCount: 1,
      expiresAt: Date.now() - 1,
      haloTenant,
      operationId: "expiring-operation",
      prefetchKey: expiring.prefetchKey,
      ticketId: 1003,
      userId: user.id,
    },
    [expiringDescriptor]
  );
  const restagedExpired = await startEmailAttachmentPrefetch({
    attachmentFingerprint: hash("expiring-fingerprint"),
    descriptors: [expiringDescriptor],
    draftItemId: "saved-expiring-draft",
    haloTenant,
    operationId: "expiring-operation",
    store,
    ticketId: 1003,
    userId: user.id,
  });
  assert.deepStrictEqual(restagedExpired.pendingAttachmentKeys, [expiringDescriptor.attachmentKey]);
  const expiredRestartRecord = await store.getEmailAttachmentPrefetch(expiring.prefetchKey, {
    haloTenant,
    ticketId: 1003,
    userId: user.id,
  });
  assert.strictEqual(expiredRestartRecord.items[0].status, "pending");
  assert.strictEqual(expiredRestartRecord.items[0].contentCiphertext, null);

  const expiredCommitDescriptor = descriptorWithContent(201, contents[0]);
  const expiredCommitKey = "expired-commit-prefetch-key";
  const expiredCommitOperation = "expired-commit-operation";
  const expiredAt = Date.now() - 1;
  await store.upsertEmailAttachmentPrefetch(
    {
      attachmentFingerprint: hash("expired-commit-fingerprint"),
      draftItemId: "saved-expired-commit-draft",
      expectedBytes: expiredCommitDescriptor.reportedSize,
      expectedCount: 1,
      expiresAt: expiredAt,
      haloTenant,
      operationId: expiredCommitOperation,
      prefetchKey: expiredCommitKey,
      ticketId: 1005,
      userId: user.id,
    },
    [expiredCommitDescriptor]
  );
  const expiredCommitClaim = await store.claimEmailAttachmentPrefetchCommit(
    expiredCommitKey,
    { haloTenant, ticketId: 1005, userId: user.id },
    expiredAt - 1
  );
  assert.strictEqual(expiredCommitClaim.status, "committing");
  const restartedExpiredCommit = await startEmailAttachmentPrefetch({
    attachmentFingerprint: hash("expired-commit-fingerprint"),
    descriptors: [expiredCommitDescriptor],
    draftItemId: "saved-expired-commit-draft",
    haloTenant,
    operationId: expiredCommitOperation,
    store,
    ticketId: 1005,
    userId: user.id,
  });
  assert.strictEqual(restartedExpiredCommit.prefetchKey, expiredCommitKey);
  assert.deepStrictEqual(restartedExpiredCommit.pendingAttachmentKeys, [
    expiredCommitDescriptor.attachmentKey,
  ]);
  const restartedExpiredCommitRecord = await store.getEmailAttachmentPrefetch(expiredCommitKey, {
    haloTenant,
    ticketId: 1005,
    userId: user.id,
  });
  assert.strictEqual(restartedExpiredCommitRecord.status, "active");

  const limitInputs = [
    {
      fingerprint: "too-many",
      descriptors: Array.from({ length: MAX_EMAIL_ATTACHMENTS + 1 }, (_, index) =>
        descriptor(index)
      ),
      operationId: "too-many-operation",
      pattern: /No more than 20/,
    },
    {
      fingerprint: "too-large",
      descriptors: [descriptor(99, MAX_EMAIL_ATTACHMENT_BYTES + 1)],
      operationId: "too-large-operation",
      pattern: /25 MiB/,
    },
    {
      fingerprint: "too-large-total",
      descriptors: [
        descriptor(100, MAX_EMAIL_ATTACHMENT_BYTES),
        descriptor(101, MAX_EMAIL_ATTACHMENT_BYTES),
        descriptor(102, 1),
      ],
      operationId: "too-large-total-operation",
      pattern: /50 MiB/,
    },
  ];
  for (const limit of limitInputs) {
    await assert.rejects(
      () =>
        startEmailAttachmentPrefetch({
          attachmentFingerprint: hash(limit.fingerprint),
          descriptors: limit.descriptors,
          draftItemId: `draft-${limit.operationId}`,
          haloTenant,
          operationId: limit.operationId,
          store,
          ticketId: 1001,
          userId: user.id,
        }),
      limit.pattern
    );
  }
  assert.strictEqual(MAX_EMAIL_ATTACHMENT_TOTAL_BYTES, 50 * 1024 * 1024);

  await store.close();
  console.log("Email attachment staging tests passed");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
