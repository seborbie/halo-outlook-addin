const crypto = require("crypto");
const parse5 = require("parse5");

const MAX_INLINE_IMAGES = 20;
const MAX_INLINE_IMAGE_BYTES = 2 * 1024 * 1024;
const MAX_INLINE_IMAGE_TOTAL_BYTES = 5 * 1024 * 1024;
const COMPOSE_PREFETCH_TTL_MS = 30 * 24 * 60 * 60 * 1000;
const SUPPORTED_MEDIA_TYPES = new Set(["image/png", "image/jpeg", "image/gif", "image/webp"]);
const inFlightUploadsByStore = new WeakMap();

function normalizeContentId(value) {
  let normalized = String(value || "").trim();
  if (/^cid:/i.test(normalized)) {
    normalized = normalized.slice(4);
  }

  try {
    normalized = decodeURIComponent(normalized);
  } catch {
    // Outlook can expose a literal percent sign in a content ID. Keep it unchanged.
  }

  normalized = normalized
    .trim()
    .replace(/^<+|>+$/g, "")
    .trim();
  return normalized.toLowerCase();
}

function normalizeHaloTenant(value) {
  const url = new URL(String(value || ""));
  return url.origin.toLowerCase();
}

function extractReferencedContentIds(html) {
  const document = parse5.parseFragment(String(html || ""));
  const contentIds = [];
  const seen = new Set();

  visitNodes(document, (node) => {
    if (node.tagName !== "img") {
      return;
    }

    for (const attributeName of ["src", "originalsrc"]) {
      const attribute = getAttribute(node, attributeName);
      if (!attribute || !/^cid:/i.test(attribute.value.trim())) {
        continue;
      }

      const contentId = normalizeContentId(attribute.value);
      if (contentId && !seen.has(contentId)) {
        seen.add(contentId);
        contentIds.push(contentId);
      }
    }
  });

  return contentIds;
}

function rewriteInlineImages(html, resolutions) {
  const document = parse5.parseFragment(String(html || ""));
  let failed = 0;

  replaceNodes(document, (node) => {
    if (node.tagName !== "img") {
      return null;
    }

    const cidAttributes = ["src", "originalsrc"]
      .map((name) => getAttribute(node, name))
      .filter((attribute) => attribute && /^cid:/i.test(attribute.value.trim()));
    if (!cidAttributes.length) {
      return null;
    }

    const contentId = normalizeContentId(cidAttributes[0].value);
    const resolution = resolutions.get(contentId);
    if (resolution && resolution.renderableUrl) {
      setAttribute(node, "src", resolution.renderableUrl);
      if (getAttribute(node, "originalsrc")) {
        setAttribute(node, "originalsrc", resolution.renderableUrl);
      }
      return null;
    }

    failed += 1;
    const alt = (getAttribute(node, "alt") || {}).value || "image";
    return {
      nodeName: "#text",
      value: `[Inline image unavailable: ${String(alt).trim() || "image"}]`,
      parentNode: node.parentNode,
    };
  });

  return {
    bodyHtml: parse5.serialize(document),
    failed,
  };
}

function normalizeInlineImageInput(value) {
  const input = value && typeof value === "object" ? value : {};
  const refs = [];
  const refContentIds = new Set();

  for (const candidate of Array.isArray(input.inlineImageRefs) ? input.inlineImageRefs : []) {
    const contentId = normalizeContentId(candidate && candidate.contentId);
    const sha256 = normalizeSha256(candidate && candidate.sha256);
    if (contentId && sha256 && !refContentIds.has(contentId)) {
      refContentIds.add(contentId);
      refs.push({ contentId, sha256 });
    }
  }

  const uploads = [];
  const uploadHashes = new Set();
  for (const candidate of Array.isArray(input.inlineImageUploads) ? input.inlineImageUploads : []) {
    const sha256 = normalizeSha256(candidate && candidate.sha256);
    if (!sha256 || uploadHashes.has(sha256)) {
      continue;
    }

    uploadHashes.add(sha256);
    uploads.push({
      contentBase64: String((candidate && candidate.contentBase64) || ""),
      contentType: String((candidate && candidate.contentType) || "").toLowerCase(),
      name: sanitizeFilename(candidate && candidate.name),
      sha256,
    });
  }

  return {
    attachmentFingerprint: String(input.inlineImageFingerprint || "").slice(0, 128),
    composeOperationId: String(input.composeAttachId || "").slice(0, 200),
    prefetchKey: String(input.inlineImagePrefetchKey || "").slice(0, 200),
    refs,
    uploads,
  };
}

async function lookupInlineImages(store, haloTenant, descriptors) {
  const normalizedTenant = normalizeHaloTenant(haloTenant);
  const hashes = Array.from(
    new Set(
      (Array.isArray(descriptors) ? descriptors : [])
        .map((descriptor) => normalizeSha256(descriptor && descriptor.sha256))
        .filter(Boolean)
    )
  ).slice(0, MAX_INLINE_IMAGES);
  const entries = await store.getInlineImageCacheEntries(normalizedTenant, hashes);
  const hitsByHash = new Map(
    entries
      .filter((entry) => {
        try {
          validateCachedInlineImage(entry, normalizedTenant);
          return true;
        } catch {
          return false;
        }
      })
      .map((entry) => [entry.sha256, entry])
  );

  return {
    hits: hashes.filter((hash) => hitsByHash.has(hash)),
    misses: hashes.filter((hash) => !hitsByHash.has(hash)),
  };
}

async function prefetchInlineImages(options) {
  const start = Date.now();
  const input = normalizeInlineImageInput(options.input);
  const haloTenant = normalizeHaloTenant(options.haloTenant);
  const requestedRefs = input.refs.slice(0, MAX_INLINE_IMAGES);
  const warnings = [];

  if (input.refs.length > MAX_INLINE_IMAGES) {
    warnings.push("Some inline images were skipped because the 20-image limit was exceeded.");
  }

  const resolution = await resolveHashes({
    hashes: requestedRefs.map((ref) => ref.sha256),
    haloTenant,
    store: options.store,
    ticketId: options.ticketId,
    uploadImage: options.uploadImage,
    uploads: input.uploads,
    warnings,
    showForUsers: options.showForUsers !== false,
  });
  const cidHash = {};
  requestedRefs.forEach((ref) => {
    if (resolution.byHash.has(ref.sha256)) {
      cidHash[ref.contentId] = ref.sha256;
    }
  });

  const prefetchKey = crypto.randomBytes(32).toString("base64url");
  await options.store.saveComposeInlineImagePrefetch({
    attachmentFingerprint: input.attachmentFingerprint,
    cidHash,
    composeOperationId: String(options.composeOperationId || "").slice(0, 200),
    expiresAt: Date.now() + COMPOSE_PREFETCH_TTL_MS,
    haloTenant,
    prefetchKey,
  });

  const result = {
    inlineImagePrefetchKey: prefetchKey,
    summary: {
      referenced: requestedRefs.length,
      cacheHits: resolution.cacheHits,
      uploaded: resolution.uploaded,
      failed: requestedRefs.length - Object.keys(cidHash).length,
      warnings,
    },
    timings: {
      totalMs: Date.now() - start,
      cacheLookupMs: resolution.timings.cacheLookupMs,
      haloUploadMs: resolution.timings.haloUploadMs,
    },
  };
  emitDiagnostic(options, "prefetch-completed", {
    cacheHits: result.summary.cacheHits,
    failed: result.summary.failed,
    referenced: result.summary.referenced,
    totalMs: result.timings.totalMs,
    uploaded: result.summary.uploaded,
  });
  return result;
}

async function resolveInlineImages(options) {
  const start = Date.now();
  const input = normalizeInlineImageInput(options.input);
  const haloTenant = normalizeHaloTenant(options.haloTenant);
  const referencedContentIds = extractReferencedContentIds(options.bodyHtml);
  if (!referencedContentIds.length) {
    return {
      bodyHtml: String(options.bodyHtml || ""),
      summary: { referenced: 0, cacheHits: 0, uploaded: 0, failed: 0, warnings: [] },
      timings: { totalMs: Date.now() - start, cacheLookupMs: 0, haloUploadMs: 0, rewritingMs: 0 },
    };
  }
  const retainedContentIds = referencedContentIds.slice(0, MAX_INLINE_IMAGES);
  const warnings = [];
  const refHashByCid = new Map(input.refs.map((ref) => [ref.contentId, ref.sha256]));

  if (referencedContentIds.length > MAX_INLINE_IMAGES) {
    warnings.push("Some inline images were skipped because the 20-image limit was exceeded.");
  }

  if (input.prefetchKey && input.attachmentFingerprint) {
    const prefetch = await options.store.getComposeInlineImagePrefetch(
      input.prefetchKey,
      haloTenant
    );
    if (
      prefetch &&
      prefetch.attachmentFingerprint === input.attachmentFingerprint &&
      prefetch.composeOperationId === input.composeOperationId
    ) {
      Object.entries(prefetch.cidHash).forEach(([contentId, sha256]) => {
        refHashByCid.set(normalizeContentId(contentId), normalizeSha256(sha256));
      });
    } else if (input.prefetchKey) {
      warnings.push("The prepared inline images no longer matched the current draft.");
    }
  }

  const hashes = retainedContentIds.map((contentId) => refHashByCid.get(contentId)).filter(Boolean);
  const resolution = await resolveHashes({
    hashes,
    haloTenant,
    store: options.store,
    ticketId: options.ticketId,
    uploadImage: options.uploadImage,
    uploads: input.uploads,
    warnings,
    showForUsers: options.showForUsers !== false,
  });

  const byCid = new Map();
  retainedContentIds.forEach((contentId) => {
    const sha256 = refHashByCid.get(contentId);
    const resolved = sha256 && resolution.byHash.get(sha256);
    if (resolved) {
      byCid.set(contentId, resolved);
    }
  });

  const rewriteStart = Date.now();
  const rewritten = rewriteInlineImages(options.bodyHtml, byCid);
  let bodyHtml = rewritten.bodyHtml;
  const failed = referencedContentIds.filter((contentId) => !byCid.has(contentId)).length;
  if (failed && options.addWarningFooter) {
    bodyHtml +=
      '<p style="color:#8a5a00"><em>One or more inline email images could not be imported.</em></p>';
  }
  if (failed && !warnings.length) {
    warnings.push("One or more inline images could not be imported.");
  }

  const result = {
    bodyHtml,
    summary: {
      referenced: referencedContentIds.length,
      cacheHits: resolution.cacheHits,
      uploaded: resolution.uploaded,
      failed,
      warnings,
    },
    timings: {
      totalMs: Date.now() - start,
      cacheLookupMs: resolution.timings.cacheLookupMs,
      haloUploadMs: resolution.timings.haloUploadMs,
      rewritingMs: Date.now() - rewriteStart,
    },
  };
  emitDiagnostic(options, "action-prepared", {
    cacheHits: result.summary.cacheHits,
    failed: result.summary.failed,
    referenced: result.summary.referenced,
    totalMs: result.timings.totalMs,
    uploaded: result.summary.uploaded,
  });
  return result;
}

async function resolveHashes(options) {
  const hashes = Array.from(new Set((options.hashes || []).map(normalizeSha256).filter(Boolean)));
  const lookupStart = Date.now();
  const showForUsers = options.showForUsers !== false;
  const entries = await options.store.getInlineImageCacheEntries(
    options.haloTenant,
    hashes,
    showForUsers
  );
  const cacheLookupMs = Date.now() - lookupStart;
  const cachedByHash = new Map(entries.map((entry) => [entry.sha256, entry]));
  const byHash = new Map();
  const uploadByHash = new Map((options.uploads || []).map((upload) => [upload.sha256, upload]));
  let cacheHits = 0;
  let totalDecodedBytes = 0;

  const pending = [];
  for (const sha256 of hashes) {
    const cached = cachedByHash.get(sha256);
    if (cached) {
      try {
        validateCachedInlineImage(cached, options.haloTenant);
        if (totalDecodedBytes + cached.byteLength > MAX_INLINE_IMAGE_TOTAL_BYTES) {
          options.warnings.push(
            "Some inline images were skipped because the 5 MiB email limit was exceeded."
          );
          continue;
        }
        totalDecodedBytes += cached.byteLength;
        cacheHits += 1;
        byHash.set(sha256, cached);
        await options.store.touchInlineImageCacheEntry(
          options.haloTenant,
          sha256,
          showForUsers
        );
        continue;
      } catch {
        options.warnings.push("A cached inline image was invalid and could not be reused.");
      }
    }

    const upload = uploadByHash.get(sha256);
    if (!upload) {
      continue;
    }

    try {
      const validated = validateInlineImageUpload(upload);
      if (totalDecodedBytes + validated.bytes.length > MAX_INLINE_IMAGE_TOTAL_BYTES) {
        options.warnings.push(
          "Some inline images were skipped because the 5 MiB email limit was exceeded."
        );
        continue;
      }
      totalDecodedBytes += validated.bytes.length;
      pending.push({ ...validated, sha256 });
    } catch (error) {
      options.warnings.push(error.message);
    }
  }

  const uploadStart = Date.now();
  const results = await mapWithConcurrency(pending, 4, async (image) => {
    try {
      return await resolveUploadMiss(options, image);
    } catch (error) {
      emitDiagnostic(options, "upload-failed", {
        errorCode: getInlineImageErrorCode(error),
      });
      options.warnings.push("An inline image could not be uploaded to Halo.");
      return null;
    }
  });
  results.filter(Boolean).forEach((result) => byHash.set(result.sha256, result.entry));
  cacheHits += results.filter((result) => result && !result.uploaded).length;

  return {
    byHash,
    cacheHits,
    uploaded: results.filter((result) => result && result.uploaded).length,
    timings: {
      cacheLookupMs,
      haloUploadMs: Date.now() - uploadStart,
    },
  };
}

function validateCachedInlineImage(entry, haloTenant) {
  if (
    !entry ||
    !entry.haloAttachmentId ||
    !SUPPORTED_MEDIA_TYPES.has(entry.mediaType) ||
    !Number.isInteger(entry.byteLength) ||
    entry.byteLength <= 0 ||
    entry.byteLength > MAX_INLINE_IMAGE_BYTES
  ) {
    throw new Error("Cached inline image metadata was invalid.");
  }
  validateRenderableUrl(entry.renderableUrl, haloTenant);
}

async function resolveUploadMiss(options, image) {
  const showForUsers = options.showForUsers !== false;
  const key = `${options.haloTenant}\u0000${image.sha256}\u0000${showForUsers ? "public" : "private"}`;
  let inFlightUploads = inFlightUploadsByStore.get(options.store);
  if (!inFlightUploads) {
    inFlightUploads = new Map();
    inFlightUploadsByStore.set(options.store, inFlightUploads);
  }
  let uploadPromise = inFlightUploads.get(key);
  const joinedExistingUpload = Boolean(uploadPromise);

  if (!uploadPromise) {
    uploadPromise = (async () => {
      const existing = (
        await options.store.getInlineImageCacheEntries(
          options.haloTenant,
          [image.sha256],
          showForUsers
        )
      )[0];
      if (existing) {
        try {
          validateCachedInlineImage(existing, options.haloTenant);
          return { entry: existing, didUpload: false };
        } catch {
          // Replace invalid cache metadata with the verified upload below.
        }
      }

      const uploaded = await options.uploadImage({
        ...image,
        ticketId: options.ticketId,
      });
      const renderableUrl = validateRenderableUrl(
        uploaded && uploaded.renderableUrl,
        options.haloTenant
      );
      const entry = await options.store.upsertInlineImageCacheEntry({
        byteLength: image.bytes.length,
        filename: image.name,
        haloAttachmentId: uploaded && uploaded.attachmentId,
        haloTenant: options.haloTenant,
        mediaType: image.mediaType,
        renderableUrl,
        sha256: image.sha256,
        showForUsers,
      });
      return { entry, didUpload: true };
    })();
    inFlightUploads.set(key, uploadPromise);
  }

  try {
    const result = await uploadPromise;
    const uploaded = result.didUpload && !joinedExistingUpload;
    if (!uploaded) {
      await options.store.touchInlineImageCacheEntry(
        options.haloTenant,
        image.sha256,
        showForUsers
      );
    }
    return { entry: result.entry, sha256: image.sha256, uploaded };
  } finally {
    if (inFlightUploads.get(key) === uploadPromise) {
      inFlightUploads.delete(key);
    }
  }
}

function validateInlineImageUpload(upload) {
  if (
    !upload.contentBase64 ||
    upload.contentBase64.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]+={0,2}$/.test(upload.contentBase64)
  ) {
    throw new Error("An inline image contained invalid Base64 data.");
  }

  const bytes = Buffer.from(upload.contentBase64, "base64");
  if (bytes.toString("base64") !== upload.contentBase64) {
    throw new Error("An inline image contained invalid Base64 data.");
  }
  if (!bytes.length || bytes.length > MAX_INLINE_IMAGE_BYTES) {
    throw new Error("An inline image exceeded the 2 MiB image limit.");
  }

  const actualHash = crypto.createHash("sha256").update(bytes).digest("hex");
  if (!crypto.timingSafeEqual(Buffer.from(actualHash), Buffer.from(upload.sha256))) {
    throw new Error("An inline image failed SHA-256 verification.");
  }

  const mediaType = detectMediaType(bytes);
  if (!mediaType || !SUPPORTED_MEDIA_TYPES.has(mediaType)) {
    throw new Error("An inline image used an unsupported format.");
  }
  const declaredMediaType = upload.contentType === "image/jpg" ? "image/jpeg" : upload.contentType;
  if (
    declaredMediaType &&
    declaredMediaType !== "application/octet-stream" &&
    declaredMediaType !== mediaType
  ) {
    throw new Error("An inline image content type did not match its data.");
  }

  return { bytes, mediaType, name: ensureFilenameExtension(upload.name, mediaType) };
}

function detectMediaType(bytes) {
  if (bytes.length >= 8 && bytes.subarray(0, 8).equals(Buffer.from("89504e470d0a1a0a", "hex"))) {
    return "image/png";
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return "image/jpeg";
  }
  if (bytes.length >= 6 && /GIF8[79]a/.test(bytes.subarray(0, 6).toString("ascii"))) {
    return "image/gif";
  }
  if (
    bytes.length >= 12 &&
    bytes.subarray(0, 4).toString("ascii") === "RIFF" &&
    bytes.subarray(8, 12).toString("ascii") === "WEBP"
  ) {
    return "image/webp";
  }
  return "";
}

function validateRenderableUrl(value, haloTenant) {
  const url = new URL(String(value || ""), `${haloTenant}/`);
  const normalizedPath = url.pathname.replace(/\/+$/, "").toLowerCase();
  if (
    url.protocol !== "https:" ||
    url.origin.toLowerCase() !== haloTenant.toLowerCase() ||
    normalizedPath !== "/api/attachment/image" ||
    !url.searchParams.get("token") ||
    url.username ||
    url.password ||
    url.hash
  ) {
    throw new Error("Halo did not return a native inline-image URL for this tenant.");
  }
  return url.toString();
}

function emitDiagnostic(options, event, details) {
  if (typeof options.onDiagnostic === "function") {
    options.onDiagnostic(event, details);
  }
}

function getInlineImageErrorCode(error) {
  const message = String((error && error.message) || "").toLowerCase();
  if (
    message.includes("renderable") ||
    message.includes("render url") ||
    message.includes("native inline-image")
  ) {
    return "render-url-invalid";
  }
  if (message.includes("attachment id")) {
    return "attachment-id-missing";
  }
  return "halo-upload-failed";
}

function normalizeSha256(value) {
  const normalized = String(value || "")
    .trim()
    .toLowerCase();
  return /^[a-f0-9]{64}$/.test(normalized) ? normalized : "";
}

function sanitizeFilename(value) {
  const filename = String(value || "inline-image")
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "-")
    .trim()
    .slice(0, 180);
  return filename || "inline-image";
}

function ensureFilenameExtension(filename, mediaType) {
  if (/\.[a-z0-9]{2,5}$/i.test(filename)) {
    return filename;
  }
  const extension = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/webp": ".webp",
  }[mediaType];
  return `${filename}${extension}`;
}

async function mapWithConcurrency(values, limit, mapper) {
  const results = new Array(values.length);
  let nextIndex = 0;

  async function worker() {
    while (nextIndex < values.length) {
      const index = nextIndex;
      nextIndex += 1;
      results[index] = await mapper(values[index], index);
    }
  }

  await Promise.all(Array.from({ length: Math.min(limit, values.length) }, () => worker()));
  return results;
}

function getAttribute(node, name) {
  return Array.isArray(node.attrs)
    ? node.attrs.find((attribute) => attribute.name.toLowerCase() === name)
    : null;
}

function setAttribute(node, name, value) {
  const existing = getAttribute(node, name);
  if (existing) {
    existing.value = value;
    return;
  }
  node.attrs = Array.isArray(node.attrs) ? node.attrs : [];
  node.attrs.push({ name, value });
}

function visitNodes(node, visitor) {
  visitor(node);
  (node.childNodes || []).forEach((child) => visitNodes(child, visitor));
}

function replaceNodes(node, replacer) {
  if (!Array.isArray(node.childNodes)) {
    return;
  }

  node.childNodes = node.childNodes.map((child) => {
    const replacement = replacer(child);
    const selected = replacement || child;
    selected.parentNode = node;
    replaceNodes(selected, replacer);
    return selected;
  });
}

module.exports = {
  COMPOSE_PREFETCH_TTL_MS,
  MAX_INLINE_IMAGES,
  MAX_INLINE_IMAGE_BYTES,
  MAX_INLINE_IMAGE_TOTAL_BYTES,
  detectMediaType,
  extractReferencedContentIds,
  lookupInlineImages,
  normalizeContentId,
  normalizeInlineImageInput,
  prefetchInlineImages,
  resolveInlineImages,
  rewriteInlineImages,
};
