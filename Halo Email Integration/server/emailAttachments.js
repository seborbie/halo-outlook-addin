const crypto = require("crypto");

const MAX_EMAIL_ATTACHMENTS = 20;
const MAX_EMAIL_ATTACHMENT_BYTES = 25 * 1024 * 1024;
const MAX_EMAIL_ATTACHMENT_TOTAL_BYTES = 50 * 1024 * 1024;
const EMAIL_ATTACHMENT_PREFETCH_TTL_MS = 7 * 24 * 60 * 60 * 1000;
const EMAIL_ATTACHMENT_STAGING_VERSION = 2;

async function startEmailAttachmentPrefetch({
  attachmentFingerprint,
  descriptors,
  draftItemId,
  haloTenant,
  operationId,
  store,
  ticketId,
  userId,
}) {
  const normalizedFingerprint = normalizeHash(attachmentFingerprint, "attachment fingerprint");
  const normalizedOperationId = normalizeOpaqueIdentifier(operationId, "operation ID");
  const normalizedDraftItemId = normalizeDraftItemId(draftItemId);
  const normalizedDescriptors = normalizeDescriptors(descriptors);
  const expectedBytes = normalizedDescriptors.reduce(
    (total, descriptor) => total + descriptor.reportedSize,
    0
  );
  const record = await store.upsertEmailAttachmentPrefetch(
    {
      attachmentFingerprint: normalizedFingerprint,
      draftItemId: normalizedDraftItemId,
      expectedBytes,
      expectedCount: normalizedDescriptors.length,
      expiresAt: Date.now() + EMAIL_ATTACHMENT_PREFETCH_TTL_MS,
      haloTenant,
      operationId: normalizedOperationId,
      prefetchKey: randomBase64Url(32),
      ticketId,
      userId,
    },
    normalizedDescriptors
  );
  const loaded = await store.getEmailAttachmentPrefetch(record.prefetchKey, {
    haloTenant,
    ticketId,
    userId,
  });
  if (!loaded || loaded.status !== "active" || loaded.stagingVersion !== 2) {
    throw requestError("The email attachment operation has already been finalized.", 409);
  }
  return {
    aggregate: summarizeRecord(loaded),
    pendingAttachmentKeys: loaded.items
      .filter((item) => ["pending", "failed"].includes(item.status))
      .map((item) => item.attachmentKey),
    prefetchKey: loaded.prefetchKey,
    removedUploads: loaded.items
      .filter((item) => item.haloAttachmentId)
      .map((item) => ({ ...toLegacyHaloAttachment(item), attachmentKey: item.attachmentKey })),
    stagingVersion: EMAIL_ATTACHMENT_STAGING_VERSION,
  };
}

async function stageEmailAttachmentItem({
  attachmentKey,
  contentBase64,
  contentFormat,
  contentSha256,
  haloTenant,
  prefetchKey,
  store,
  ticketId,
  tokenCrypto,
  userId,
}) {
  const record = await store.getEmailAttachmentPrefetch(prefetchKey, {
    haloTenant,
    ticketId,
    userId,
  });
  if (!record || record.status !== "active" || record.stagingVersion !== 2) {
    throw attachmentsNotReady(
      "The email attachment preparation is unavailable or expired.",
      "stage-unavailable"
    );
  }
  const normalizedAttachmentKey = normalizeHash(attachmentKey, "attachment key");
  const item = record.items.find(
    (candidate) => candidate.attachmentKey === normalizedAttachmentKey
  );
  if (!item || item.status === "removed") {
    throw requestError("The email attachment is not part of this preparation.", 400);
  }

  normalizeContentFormat(contentFormat);
  const normalizedBase64 = normalizeBase64(contentBase64);
  const decodedSize = getDecodedBase64Length(normalizedBase64);
  if (decodedSize > MAX_EMAIL_ATTACHMENT_BYTES) {
    throw requestError("The email attachment exceeds the 25 MiB attachment limit.", 413);
  }
  const bytes = Buffer.from(normalizedBase64, "base64");
  const calculatedHash = crypto.createHash("sha256").update(bytes).digest("hex");
  const suppliedHash = normalizeHash(contentSha256, "attachment content hash");
  if (calculatedHash !== suppliedHash || bytes.length !== decodedSize) {
    bytes.fill(0);
    throw requestError("The email attachment content hash did not match.", 400);
  }
  if (!item.expectedContentSha256 || item.expectedContentSha256 !== calculatedHash) {
    bytes.fill(0);
    throw attachmentsNotReady(
      "The attachment content changed after this preparation started. Reconcile the draft and retry.",
      "content-changed"
    );
  }

  if (item.status === "prepared" && item.contentSha256 === calculatedHash) {
    bytes.fill(0);
    return { attachmentKey: normalizedAttachmentKey, status: "already-prepared" };
  }
  if (item.status === "prepared") {
    await store.resetEmailAttachmentPrefetchItem(prefetchKey, normalizedAttachmentKey);
  }

  const claimed = await store.claimEmailAttachmentPrefetchItem(
    prefetchKey,
    normalizedAttachmentKey,
    decodedSize,
    calculatedHash,
    MAX_EMAIL_ATTACHMENT_TOTAL_BYTES
  );
  if (
    !claimed ||
    claimed.status !== "preparing" ||
    Number(claimed.decoded_size) !== decodedSize ||
    claimed.expected_content_sha256 !== calculatedHash
  ) {
    bytes.fill(0);
    if (claimed && claimed.status === "prepared") {
      return { attachmentKey: normalizedAttachmentKey, status: "already-prepared" };
    }
    throw attachmentsNotReady(
      "The email attachment could not be reserved for preparation.",
      "reservation-failed"
    );
  }

  try {
    const encrypted = tokenCrypto.encryptBytes(bytes, createAttachmentAdditionalData(record, item));
    const saved = await store.saveEmailAttachmentPrefetchItemPrepared(
      prefetchKey,
      normalizedAttachmentKey,
      {
        ...encrypted,
        contentSha256: calculatedHash,
        decodedSize,
        preparedAt: Date.now(),
      }
    );
    if (!saved) {
      throw attachmentsNotReady(
        "The email attachment operation finalized during preparation.",
        "operation-finalized"
      );
    }
    return { attachmentKey: normalizedAttachmentKey, status: "prepared" };
  } catch (error) {
    await store.saveEmailAttachmentPrefetchItemFailure(
      prefetchKey,
      normalizedAttachmentKey,
      calculatedHash,
      getSafeFailureCode(error)
    );
    throw error;
  } finally {
    bytes.fill(0);
  }
}

async function getEmailAttachmentPreparationStatus({
  haloTenant,
  prefetchKey,
  store,
  ticketId,
  userId,
}) {
  const record = await store.getEmailAttachmentPrefetch(prefetchKey, {
    haloTenant,
    ticketId,
    userId,
  });
  if (!record || record.stagingVersion !== 2) {
    throw attachmentsNotReady(
      "The email attachment preparation is unavailable or expired.",
      "stage-unavailable"
    );
  }
  return {
    aggregate: summarizeRecord(record),
    attachmentFingerprint: record.attachmentFingerprint,
    items: record.items.map((item) => ({
      attachmentKey: item.attachmentKey,
      failureCode: item.failureCode || "",
      status: normalizeItemStatus(item.status),
    })),
    operationId: record.operationId,
    prefetchKey: record.prefetchKey,
    stagingVersion: record.stagingVersion,
    status: record.status,
  };
}

async function resolveEmailAttachments({ haloTenant, input, store, ticketId, userId }) {
  const clientSummary = normalizeClientSummary(input && input.emailAttachmentSummary);
  if (!input || input.includeEmailAttachments !== true) {
    return {
      prefetchKey: "",
      stagedItems: [],
      summary: { ...clientSummary, attached: 0, prepared: 0 },
    };
  }

  if (clientSummary.selected === 0) {
    return { prefetchKey: "", stagedItems: [], summary: clientSummary };
  }
  if (Number(input.emailAttachmentStagingVersion) !== EMAIL_ATTACHMENT_STAGING_VERSION) {
    throw attachmentsNotReady(
      "Open the Halo pane to prepare this draft's attachments, then retry Send.",
      "legacy-state"
    );
  }
  const prefetchKey = String(input.emailAttachmentPrefetchKey || "");
  const fingerprint = normalizeHash(input.emailAttachmentFingerprint, "attachment fingerprint");
  const operationId = normalizeOpaqueIdentifier(
    input.emailAttachmentOperationId,
    "attachment operation ID"
  );
  const draftItemId = normalizeDraftItemId(input.emailAttachmentDraftItemId);
  if (!prefetchKey) {
    throw attachmentsNotReady(
      "Open the Halo pane to finish preparing attachments, then retry Send.",
      "stage-missing"
    );
  }
  const record = await store.getEmailAttachmentPrefetch(prefetchKey, {
    haloTenant,
    ticketId,
    userId,
  });
  validatePreparedRecord(record, { clientSummary, draftItemId, fingerprint, operationId });

  const claimed = await store.claimEmailAttachmentPrefetchCommit(prefetchKey, {
    haloTenant,
    ticketId,
    userId,
  });
  if (!claimed) {
    throw attachmentsNotReady(
      "Attachment preparation is already being committed or has expired.",
      "commit-unavailable"
    );
  }
  try {
    validatePreparedRecord(claimed, { clientSummary, draftItemId, fingerprint, operationId });
    return {
      prefetchKey,
      record: claimed,
      stagedItems: claimed.items.filter((item) => item.status === "prepared"),
      summary: {
        ...clientSummary,
        attached: 0,
        failed: 0,
        prepared: claimed.expectedCount,
        selected: claimed.expectedCount,
      },
    };
  } catch (error) {
    await store.releaseEmailAttachmentPrefetchCommit(prefetchKey);
    throw error;
  }
}

function validatePreparedRecord(record, expected) {
  if (!record) {
    throw attachmentsNotReady(
      "The email attachment preparation is unavailable or expired.",
      "stage-unavailable"
    );
  }
  if (record.stagingVersion !== 2) {
    throw attachmentsNotReady(
      "Open the Halo pane to prepare this draft's attachments, then retry Send.",
      "legacy-state"
    );
  }
  if (!["active", "committing"].includes(record.status)) {
    throw attachmentsNotReady(
      "The email attachment preparation is unavailable or expired.",
      "stage-unavailable"
    );
  }
  if (
    record.attachmentFingerprint !== expected.fingerprint ||
    record.operationId !== expected.operationId ||
    record.draftItemId !== expected.draftItemId ||
    record.expectedCount !== expected.clientSummary.selected
  ) {
    throw attachmentsNotReady(
      "The live attachment list does not match the prepared draft.",
      "inventory-mismatch"
    );
  }
  const activeItems = record.items.filter((item) => item.status !== "removed");
  const allPrepared =
    activeItems.length === record.expectedCount &&
    activeItems.every(
      (item) =>
        item.status === "prepared" &&
        item.contentCiphertext &&
        item.contentIv &&
        item.contentTag &&
        item.contentKeyId &&
        item.contentSha256
    );
  if (!allPrepared) {
    throw attachmentsNotReady(
      "Attachment preparation is incomplete. Open the Halo pane and retry.",
      "preparation-incomplete"
    );
  }
}

function decryptStagedAttachment(tokenCrypto, record, item) {
  const bytes = tokenCrypto.decryptBytes(
    {
      ciphertext: item.contentCiphertext,
      iv: item.contentIv,
      keyId: item.contentKeyId,
      tag: item.contentTag,
    },
    createAttachmentAdditionalData(record, item)
  );
  const hash = crypto.createHash("sha256").update(bytes).digest("hex");
  if (hash !== item.contentSha256 || bytes.length !== Number(item.decodedSize)) {
    bytes.fill(0);
    throw attachmentsNotReady(
      "A prepared email attachment failed integrity validation.",
      "integrity-failed"
    );
  }
  return bytes;
}

function createAttachmentAdditionalData(record, item) {
  return [
    "email-attachment:v2",
    record.prefetchKey,
    item.attachmentKey,
    record.userId,
    record.haloTenant,
  ].join("\u0000");
}

function summarizeRecord(record) {
  const items = record.items.filter((item) => item.status !== "removed");
  return {
    failed: items.filter((item) => item.status === "failed").length,
    pending: items.filter((item) => ["pending", "preparing"].includes(item.status)).length,
    prepared: items.filter((item) => item.status === "prepared").length,
    selected: items.length,
  };
}

function normalizeItemStatus(value) {
  return ["pending", "prepared", "failed", "skipped", "removed"].includes(value)
    ? value
    : value === "preparing"
      ? "pending"
      : "failed";
}

function toLegacyHaloAttachment(item) {
  return {
    filename: item.filename || item.haloFilename,
    filesize: Number(item.haloFilesize || item.decodedSize || 0),
    id: item.haloAttachmentId,
  };
}

function normalizeDescriptors(values) {
  if (!Array.isArray(values)) {
    throw requestError("Email attachment descriptors are required.", 400);
  }
  if (values.length > MAX_EMAIL_ATTACHMENTS) {
    throw requestError("No more than 20 email attachments can be prepared.", 400);
  }
  const seen = new Set();
  let reportedTotal = 0;
  return values.map((value) => {
    const descriptor = value && typeof value === "object" ? value : {};
    const attachmentKey = normalizeHash(descriptor.attachmentKey, "attachment key");
    if (seen.has(attachmentKey)) {
      throw requestError("Duplicate email attachment descriptors are not allowed.", 400);
    }
    seen.add(attachmentKey);
    const reportedSize = normalizeByteLength(descriptor.reportedSize);
    if (reportedSize > MAX_EMAIL_ATTACHMENT_BYTES) {
      throw requestError("An email attachment exceeds the 25 MiB attachment limit.", 413);
    }
    reportedTotal += reportedSize;
    if (reportedTotal > MAX_EMAIL_ATTACHMENT_TOTAL_BYTES) {
      throw requestError("Email attachments exceed the 50 MiB email limit.", 413);
    }
    return {
      attachmentKey,
      attachmentType: normalizeShortText(descriptor.attachmentType, "file", 40),
      contentSha256: descriptor.contentSha256
        ? normalizeHash(descriptor.contentSha256, "attachment content hash")
        : "",
      contentType: normalizeContentType(descriptor.contentType),
      name: sanitizeAttachmentName(descriptor.name),
      reportedSize,
    };
  });
}

function normalizeClientSummary(value) {
  const input = value && typeof value === "object" ? value : {};
  return {
    attached: 0,
    detected: normalizeCount(input.detected),
    failed: normalizeCount(input.failed),
    prepared: normalizeCount(input.prepared !== undefined ? input.prepared : input.uploaded),
    selected: normalizeCount(input.selected),
    skipped: normalizeCount(input.skipped),
    warnings: Array.isArray(input.warnings)
      ? uniqueWarnings(input.warnings.map(() => "An email attachment is unsupported.").slice(0, 5))
      : [],
  };
}

function normalizeCount(value) {
  const count = Number(value);
  return Number.isInteger(count) && count >= 0 ? Math.min(count, 1000) : 0;
}

function sanitizeAttachmentName(value) {
  const basename = String(value || "")
    .replace(/\\/g, "/")
    .split("/")
    .pop()
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .trim();
  if (!basename || basename === "." || basename === "..") {
    return "email-attachment.bin";
  }
  if (basename.length <= 255) {
    return basename;
  }
  const extensionIndex = basename.lastIndexOf(".");
  const extension = extensionIndex > 0 ? basename.slice(extensionIndex, extensionIndex + 32) : "";
  return `${basename.slice(0, Math.max(1, 255 - extension.length))}${extension}`;
}

function normalizeBase64(value) {
  const normalized = String(value || "").replace(/\s/g, "");
  if (!normalized || normalized.length % 4 !== 0 || !/^[a-z0-9+/]*={0,2}$/i.test(normalized)) {
    throw requestError("Email attachment content must be valid Base64.", 400);
  }
  return normalized;
}

function getDecodedBase64Length(value) {
  const padding = value.endsWith("==") ? 2 : value.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor((value.length * 3) / 4) - padding);
}

function normalizeContentFormat(value) {
  const format = String(value || "base64").toLowerCase();
  if (!["base64", "eml", "icalendar"].includes(format)) {
    throw requestError("The email attachment content format is unsupported.", 400);
  }
  return format;
}

function normalizeHash(value, label) {
  const hash = String(value || "").toLowerCase();
  if (!/^[a-f0-9]{64}$/.test(hash)) {
    throw requestError(`A valid ${label} is required.`, 400);
  }
  return hash;
}

function normalizeOpaqueIdentifier(value, label) {
  const identifier = String(value || "");
  if (!/^[a-z0-9._~-]{8,200}$/i.test(identifier)) {
    throw requestError(`A valid ${label} is required.`, 400);
  }
  return identifier;
}

function normalizeDraftItemId(value) {
  const identifier = String(value || "").trim();
  if (!identifier || identifier.length > 2048) {
    throw requestError("A valid saved draft ID is required.", 400);
  }
  return identifier;
}

function normalizeByteLength(value) {
  const length = Number(value);
  return Number.isInteger(length) && length >= 0 ? length : 0;
}

function normalizeShortText(value, fallback, maxLength) {
  const normalized = String(value || fallback).trim();
  return (normalized || fallback).slice(0, maxLength);
}

function normalizeContentType(value) {
  const contentType = String(value || "")
    .split(";", 1)[0]
    .trim()
    .toLowerCase();
  const mimeToken = /^[!#$%&'*+.^_`|~0-9a-z-]+\/[!#$%&'*+.^_`|~0-9a-z-]+$/;
  return contentType.length <= 200 && mimeToken.test(contentType)
    ? contentType
    : "application/octet-stream";
}

function randomBase64Url(byteLength) {
  return crypto.randomBytes(byteLength).toString("base64url");
}

function uniqueWarnings(values) {
  return Array.from(new Set(values.filter(Boolean))).slice(0, 10);
}

function requestError(message, status, code) {
  const error = new Error(message);
  error.status = status;
  error.code = code || "invalid-attachment-request";
  error.isRequestError = true;
  return error;
}

function attachmentsNotReady(message, diagnosticReason = "not-ready") {
  const error = requestError(message, 409, "attachments-not-ready");
  error.attachmentDiagnosticReason = diagnosticReason;
  return error;
}

function getSafeFailureCode(error) {
  if (error && Number(error.status) === 413) {
    return "size-limit";
  }
  if (error && Number(error.status) >= 400 && Number(error.status) < 500) {
    return "rejected";
  }
  return "preparation-failed";
}

module.exports = {
  EMAIL_ATTACHMENT_PREFETCH_TTL_MS,
  EMAIL_ATTACHMENT_STAGING_VERSION,
  MAX_EMAIL_ATTACHMENTS,
  MAX_EMAIL_ATTACHMENT_BYTES,
  MAX_EMAIL_ATTACHMENT_TOTAL_BYTES,
  createAttachmentAdditionalData,
  decryptStagedAttachment,
  getEmailAttachmentPreparationStatus,
  resolveEmailAttachments,
  sanitizeAttachmentName,
  stageEmailAttachmentItem,
  startEmailAttachmentPrefetch,
};
