/* global Office, DOMParser, crypto, window, setTimeout, clearTimeout, atob */

export type InlineImageRef = {
  contentId: string;
  sha256: string;
};

export type InlineImageUpload = {
  sha256: string;
  name: string;
  contentType: string;
  contentBase64: string;
};

export type PreparedInlineImages = {
  inlineImageRefs: InlineImageRef[];
  inlineImageUploads: InlineImageUpload[];
  inlineImageFingerprint: string;
  inlineImageWarnings: string[];
  inlineImageTimings: {
    outlookReadMs: number;
    hashingMs: number;
  };
};

type AttachmentMetadata = {
  id: string;
  name: string;
  size: number;
  isInline: boolean;
  contentId: string;
  contentType: string;
  attachmentType: string;
};

type AttachmentItem = {
  itemId?: string;
  attachments?: unknown[];
  getAttachmentsAsync?: (callback: (result: Office.AsyncResult<unknown[]>) => void) => void;
  getAttachmentContentAsync?: (
    attachmentId: string,
    callback: (result: Office.AsyncResult<{ content: string; format: unknown }>) => void
  ) => void;
};

const MAX_INLINE_IMAGES = 20;
const MAX_INLINE_IMAGE_BYTES = 2 * 1024 * 1024;
const MAX_INLINE_IMAGE_TOTAL_BYTES = 5 * 1024 * 1024;
const preparationCache = new Map<string, Promise<PreparedInlineImages>>();

export function clearInlineImagePreparationCache() {
  preparationCache.clear();
}

export function extractCidReferences(html: string): string[] {
  const document = new DOMParser().parseFromString(String(html || ""), "text/html");
  const seen = new Set<string>();
  const contentIds: string[] = [];
  Array.from(document.querySelectorAll("img")).forEach((image) => {
    [image.getAttribute("src"), image.getAttribute("originalsrc")].forEach((value) => {
      if (!value || !/^cid:/i.test(value.trim())) {
        return;
      }
      const contentId = normalizeContentId(value);
      if (contentId && !seen.has(contentId)) {
        seen.add(contentId);
        contentIds.push(contentId);
      }
    });
  });
  return contentIds;
}

export async function getInlineImageFingerprint(
  item: AttachmentItem,
  html: string,
  timeoutMs = 0
): Promise<string> {
  const references = extractCidReferences(html).slice(0, MAX_INLINE_IMAGES);
  const attachments = await withOptionalTimeout(
    readAttachmentMetadata(item),
    timeoutMs,
    "Inline image metadata retrieval timed out."
  );
  return hashText(buildFingerprintSource(references, attachments));
}

export function prepareOutlookInlineImages(
  item: AttachmentItem,
  html: string,
  itemKey: string,
  timeoutMs = 0
): Promise<PreparedInlineImages> {
  const references = extractCidReferences(html).slice(0, MAX_INLINE_IMAGES);
  if (!references.length) {
    return Promise.resolve(emptyPreparation());
  }

  const cacheKey = `${itemKey || "unsaved"}:${references.join("|")}`;
  let preparation = preparationCache.get(cacheKey);
  if (!preparation) {
    preparation = collectInlineImages(item, references, timeoutMs).catch((error) => {
      if (preparationCache.get(cacheKey) === preparation) {
        preparationCache.delete(cacheKey);
      }
      throw error;
    });
    preparationCache.set(cacheKey, preparation);
  }
  return preparation;
}

async function collectInlineImages(
  item: AttachmentItem,
  references: string[],
  timeoutMs: number
): Promise<PreparedInlineImages> {
  const collectionStart = Date.now();
  const deadline = timeoutMs ? collectionStart + timeoutMs : 0;
  const attachments = await withOptionalTimeout(
    readAttachmentMetadata(item),
    deadline ? Math.max(1, deadline - Date.now()) : 0,
    "Inline image metadata retrieval timed out."
  );
  const fingerprintSource = buildFingerprintSource(references, attachments);
  const matched = references
    .map((contentId) => ({ contentId, attachment: matchAttachment(contentId, attachments) }))
    .filter((entry): entry is { contentId: string; attachment: AttachmentMetadata } =>
      Boolean(entry.attachment)
    );
  const warnings: string[] = [];
  let selectedMetadataBytes = 0;
  const selected = matched.filter((entry) => {
    const byteLength = entry.attachment.size;
    if (byteLength > MAX_INLINE_IMAGE_BYTES) {
      warnings.push("An inline image exceeded the 2 MiB image limit.");
      return false;
    }
    if (byteLength > 0 && selectedMetadataBytes + byteLength > MAX_INLINE_IMAGE_TOTAL_BYTES) {
      warnings.push("Some inline images exceeded the 5 MiB email limit.");
      return false;
    }
    selectedMetadataBytes += byteLength;
    return true;
  });
  const retrieved = await mapWithConcurrency(selected, 4, async (entry) => {
    try {
      if (deadline && Date.now() >= deadline) {
        throw new Error("Inline image retrieval timed out.");
      }
      const content = await withOptionalTimeout(
        readAttachmentContent(item, entry.attachment.id),
        deadline ? Math.max(1, deadline - Date.now()) : 0,
        "Inline image retrieval timed out."
      );
      return { ...entry, content };
    } catch {
      warnings.push("An inline image could not be read from Outlook.");
      return null;
    }
  });
  const outlookReadMs = Date.now() - collectionStart;
  const hashingStart = Date.now();
  let decodedTotalBytes = 0;
  const withinDecodedLimits = retrieved.filter(isPresent).filter((entry) => {
    const byteLength = getDecodedBase64Length(entry.content);
    if (byteLength > MAX_INLINE_IMAGE_BYTES) {
      warnings.push("An inline image exceeded the 2 MiB image limit.");
      return false;
    }
    if (decodedTotalBytes + byteLength > MAX_INLINE_IMAGE_TOTAL_BYTES) {
      warnings.push("Some inline images exceeded the 5 MiB email limit.");
      return false;
    }
    decodedTotalBytes += byteLength;
    return true;
  });
  const [fingerprint, hashResults] = await Promise.all([
    hashText(fingerprintSource),
    Promise.all(
      withinDecodedLimits.map(async (entry) => {
        try {
          return { ...entry, sha256: await hashBase64(entry.content) };
        } catch {
          warnings.push("An inline image could not be hashed.");
          return null;
        }
      })
    ),
  ]);
  const hashed = hashResults.filter(isPresent);

  return {
    inlineImageFingerprint: fingerprint,
    inlineImageRefs: hashed.map((entry) => ({
      contentId: entry.contentId,
      sha256: entry.sha256,
    })),
    inlineImageUploads: hashed.map((entry) => ({
      contentBase64: entry.content,
      contentType: entry.attachment.contentType || inferContentType(entry.attachment.name),
      name: entry.attachment.name || "inline-image",
      sha256: entry.sha256,
    })),
    inlineImageWarnings: warnings,
    inlineImageTimings: {
      outlookReadMs,
      hashingMs: Date.now() - hashingStart,
    },
  };
}

async function readAttachmentMetadata(item: AttachmentItem): Promise<AttachmentMetadata[]> {
  let rawAttachments: unknown[] = [];
  if (Array.isArray(item.attachments)) {
    rawAttachments = item.attachments;
  } else if (item.getAttachmentsAsync) {
    rawAttachments = await new Promise((resolve) => {
      item.getAttachmentsAsync!((result) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || [] : []);
      });
    });
  }

  return rawAttachments.map(normalizeAttachment).filter(isPresent);
}

function normalizeAttachment(value: unknown): AttachmentMetadata | null {
  const attachment = (value || {}) as {
    id?: unknown;
    name?: unknown;
    size?: unknown;
    isInline?: unknown;
    contentId?: unknown;
    contentType?: unknown;
    attachmentType?: unknown;
  };
  const id = String(attachment.id || "");
  const attachmentType = String(attachment.attachmentType || "file");
  if (!id || !/file/i.test(attachmentType)) {
    return null;
  }
  return {
    attachmentType,
    contentId: normalizeContentId(attachment.contentId),
    contentType: String(attachment.contentType || "").toLowerCase(),
    id,
    isInline: attachment.isInline !== false,
    name: String(attachment.name || "inline-image"),
    size: Number(attachment.size) || 0,
  };
}

function matchAttachment(
  contentId: string,
  attachments: AttachmentMetadata[]
): AttachmentMetadata | null {
  return (
    attachments.find(
      (attachment) =>
        attachment.isInline &&
        (attachment.contentId === contentId || normalizeContentId(attachment.id) === contentId)
    ) || null
  );
}

function readAttachmentContent(item: AttachmentItem, attachmentId: string): Promise<string> {
  return new Promise((resolve, reject) => {
    if (!item.getAttachmentContentAsync) {
      reject(new Error("Outlook does not support inline attachment content retrieval."));
      return;
    }
    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
        reject(new Error(result.error.message || "Outlook could not read an inline image."));
        return;
      }
      const format = String(result.value.format || "").toLowerCase();
      if (format && !format.includes("base64")) {
        reject(new Error("Outlook returned a non-Base64 inline image."));
        return;
      }
      resolve(String(result.value.content || "").replace(/\s/g, ""));
    });
  });
}

function buildFingerprintSource(references: string[], attachments: AttachmentMetadata[]): string {
  const matching = references.map((contentId) => {
    const attachment = matchAttachment(contentId, attachments);
    return attachment
      ? [
          contentId,
          attachment.contentId || normalizeContentId(attachment.id),
          attachment.name,
          attachment.size,
          attachment.contentType,
        ].join("\u0000")
      : `${contentId}\u0000missing`;
  });
  return matching.sort().join("\u0001");
}

async function hashBase64(value: string): Promise<string> {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return hashBytes(bytes);
}

function getDecodedBase64Length(value: string): number {
  const normalized = String(value || "").replace(/\s/g, "");
  const padding = normalized.endsWith("==") ? 2 : normalized.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor((normalized.length * 3) / 4) - padding);
}

async function hashText(value: string): Promise<string> {
  return hashBytes(encodeUtf8(String(value || "")));
}

function encodeUtf8(value: string): Uint8Array {
  const bytes: number[] = [];
  for (let index = 0; index < value.length; index += 1) {
    let codePoint = value.charCodeAt(index);
    if (codePoint >= 0xd800 && codePoint <= 0xdbff) {
      const low = value.charCodeAt(index + 1);
      if (low >= 0xdc00 && low <= 0xdfff) {
        codePoint = 0x10000 + ((codePoint - 0xd800) << 10) + (low - 0xdc00);
        index += 1;
      } else {
        codePoint = 0xfffd;
      }
    } else if (codePoint >= 0xdc00 && codePoint <= 0xdfff) {
      codePoint = 0xfffd;
    }

    if (codePoint <= 0x7f) {
      bytes.push(codePoint);
    } else if (codePoint <= 0x7ff) {
      bytes.push(0xc0 | (codePoint >> 6), 0x80 | (codePoint & 0x3f));
    } else if (codePoint <= 0xffff) {
      bytes.push(
        0xe0 | (codePoint >> 12),
        0x80 | ((codePoint >> 6) & 0x3f),
        0x80 | (codePoint & 0x3f)
      );
    } else {
      bytes.push(
        0xf0 | (codePoint >> 18),
        0x80 | ((codePoint >> 12) & 0x3f),
        0x80 | ((codePoint >> 6) & 0x3f),
        0x80 | (codePoint & 0x3f)
      );
    }
  }
  return new Uint8Array(bytes);
}

async function hashBytes(value: Uint8Array | ArrayBuffer): Promise<string> {
  const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
  const digestInput = new Uint8Array(bytes.length);
  digestInput.set(bytes);
  const provider = getCryptoProvider();
  if (!provider || !provider.subtle || !provider.subtle.digest) {
    throw new Error("SHA-256 is unavailable in this Outlook webview.");
  }

  let operation: unknown;
  try {
    operation = provider.subtle.digest("SHA-256", digestInput.buffer);
  } catch {
    operation = provider.subtle.digest({ name: "SHA-256" }, digestInput.buffer);
  }
  const digest = await normalizeDigestOperation(operation);
  return Array.from(new Uint8Array(digest))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

type CryptoProvider = {
  subtle?: {
    digest: (algorithm: string | { name: string }, data: ArrayBuffer) => unknown;
  };
};

function getCryptoProvider(): CryptoProvider | null {
  const standardCrypto =
    typeof crypto !== "undefined" ? (crypto as unknown as CryptoProvider) : null;
  if (standardCrypto && standardCrypto.subtle) {
    return standardCrypto;
  }
  const legacyWindow = window as unknown as { msCrypto?: CryptoProvider };
  return legacyWindow.msCrypto || null;
}

function normalizeDigestOperation(operation: unknown): Promise<ArrayBuffer> {
  const promiseOperation = operation as { then?: unknown };
  if (promiseOperation && typeof promiseOperation.then === "function") {
    return operation as Promise<ArrayBuffer>;
  }

  return new Promise((resolve, reject) => {
    const legacyOperation = operation as {
      oncomplete?: (event: { target?: { result?: ArrayBuffer } }) => void;
      onerror?: () => void;
    };
    if (!legacyOperation || typeof legacyOperation !== "object") {
      reject(new Error("SHA-256 failed in this Outlook webview."));
      return;
    }
    legacyOperation.oncomplete = (event) => {
      const result = event && event.target && event.target.result;
      if (result) {
        resolve(result);
      } else {
        reject(new Error("SHA-256 returned no result."));
      }
    };
    legacyOperation.onerror = () => reject(new Error("SHA-256 failed in this Outlook webview."));
  });
}

function inferContentType(filename: string): string {
  const extension = String(filename || "")
    .split(".")
    .pop()!
    .toLowerCase();
  return (
    {
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      webp: "image/webp",
    }[extension] || "application/octet-stream"
  );
}

function normalizeContentId(value: unknown): string {
  let normalized = String(value || "")
    .trim()
    .replace(/^cid:/i, "");
  try {
    normalized = decodeURIComponent(normalized);
  } catch {
    // Keep literal percent signs from Outlook attachment metadata.
  }
  return normalized
    .trim()
    .replace(/^<+|>+$/g, "")
    .trim()
    .toLowerCase();
}

function emptyPreparation(): PreparedInlineImages {
  return {
    inlineImageFingerprint: "",
    inlineImageRefs: [],
    inlineImageUploads: [],
    inlineImageWarnings: [],
    inlineImageTimings: { hashingMs: 0, outlookReadMs: 0 },
  };
}

function withOptionalTimeout<T>(
  promise: Promise<T>,
  timeoutMs: number,
  message: string
): Promise<T> {
  if (!timeoutMs) {
    return promise;
  }
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error(message)), timeoutMs);
    promise.then(
      (value) => {
        clearTimeout(timeout);
        resolve(value);
      },
      (error) => {
        clearTimeout(timeout);
        reject(error);
      }
    );
  });
}

async function mapWithConcurrency<T, R>(
  values: T[],
  limit: number,
  mapper: (value: T) => Promise<R>
): Promise<R[]> {
  const results = new Array<R>(values.length);
  let nextIndex = 0;
  async function worker() {
    while (nextIndex < values.length) {
      const index = nextIndex;
      nextIndex += 1;
      results[index] = await mapper(values[index]);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, values.length) }, () => worker()));
  return results;
}

function isPresent<T>(value: T | null): value is T {
  return Boolean(value);
}
