/* global Office, crypto, window, setTimeout, clearTimeout, atob, btoa */

import { extractCidReferences } from "./inlineImages";

export const MAX_EMAIL_ATTACHMENTS = 20;
export const MAX_EMAIL_ATTACHMENT_BYTES = 25 * 1024 * 1024;
export const MAX_EMAIL_ATTACHMENT_TOTAL_BYTES = 50 * 1024 * 1024;
export const EMAIL_ATTACHMENT_READ_CONCURRENCY = 3;

export type EmailAttachmentDescriptor = {
  attachmentKey: string;
  name: string;
  contentType: string;
  attachmentType: string;
  reportedSize: number;
  contentSha256?: string;
};

export type EmailAttachmentUpload = EmailAttachmentDescriptor & {
  contentBase64: string;
  contentSha256: string;
  contentFormat: "base64" | "eml" | "icalendar";
  decodedSize: number;
};

export type EmailAttachmentSummary = {
  detected: number;
  selected: number;
  prepared: number;
  attached: number;
  skipped: number;
  failed: number;
  warnings: string[];
};

type AttachmentItem = {
  attachments?: unknown[];
  getAttachmentsAsync?: (callback: (result: Office.AsyncResult<unknown[]>) => void) => void;
  getAttachmentContentAsync?: (
    attachmentId: string,
    callback: (result: Office.AsyncResult<{ content: string; format: unknown }>) => void
  ) => void;
};

type LocalEmailAttachment = EmailAttachmentDescriptor & {
  outlookAttachmentId: string;
};

export type PreparedEmailAttachmentMetadata = {
  emailAttachmentFingerprint: string;
  attachments: LocalEmailAttachment[];
  reportedTotalBytes: number;
  reportedOversized: number;
  overCount: number;
};

export type PreparedEmailAttachments = {
  emailAttachmentFingerprint: string;
  descriptors: EmailAttachmentDescriptor[];
  uploads: EmailAttachmentUpload[];
  failureCodes: string[];
  warnings: string[];
  skipped: number;
  failed: number;
  timings: {
    outlookReadMs: number;
  };
};

export async function collectEmailAttachmentMetadata(
  item: AttachmentItem,
  html: string
): Promise<PreparedEmailAttachmentMetadata> {
  const references = new Set(extractCidReferences(html));
  const rawAttachments = await readAttachmentMetadata(item);
  const uniqueAttachments = Array.from(
    new Map(rawAttachments.map((attachment) => [attachment.id, attachment])).values()
  );
  const ordinary = uniqueAttachments.filter((attachment) => {
    if (attachment.isInline) {
      return false;
    }
    const fallbackContentId = normalizeContentId(attachment.id);
    return !(
      (attachment.contentId && references.has(attachment.contentId)) ||
      (fallbackContentId && references.has(fallbackContentId))
    );
  });
  const keyed = ordinary
    .map((attachment, originalIndex) => ({
      attachment,
      descriptor: [
        sanitizeAttachmentName(attachment.name).toLowerCase(),
        attachment.size,
        attachment.contentType || "application/octet-stream",
        attachment.attachmentType,
      ].join("\u0000"),
      originalIndex,
    }))
    .sort((left, right) =>
      left.descriptor === right.descriptor
        ? left.originalIndex - right.originalIndex
        : left.descriptor.localeCompare(right.descriptor)
    );
  const duplicateCounts = new Map<string, number>();
  const attachments = await Promise.all(
    keyed.map(async ({ attachment, descriptor }) => {
      const duplicateIndex = (duplicateCounts.get(descriptor) || 0) + 1;
      duplicateCounts.set(descriptor, duplicateIndex);
      return {
        attachmentKey: await hashText(`${descriptor}\u0000${duplicateIndex}`),
        attachmentType: attachment.attachmentType,
        contentType: attachment.contentType || "application/octet-stream",
        name: sanitizeAttachmentName(attachment.name),
        outlookAttachmentId: attachment.id,
        reportedSize: attachment.size,
      };
    })
  );
  const fingerprint = await hashText(
    Array.from(duplicateCounts.entries())
      .map(([descriptor, count]) => `${descriptor}\u0000${count}`)
      .sort()
      .join("\u0001")
  );

  return {
    attachments,
    emailAttachmentFingerprint: attachments.length ? fingerprint : "",
    overCount: Math.max(0, attachments.length - MAX_EMAIL_ATTACHMENTS),
    reportedOversized: attachments.filter(
      (attachment) => attachment.reportedSize > MAX_EMAIL_ATTACHMENT_BYTES
    ).length,
    reportedTotalBytes: attachments.reduce(
      (total, attachment) => total + Math.max(0, attachment.reportedSize),
      0
    ),
  };
}

export async function prepareOutlookEmailAttachments(
  item: AttachmentItem,
  metadata: PreparedEmailAttachmentMetadata,
  timeoutMs = 0
): Promise<PreparedEmailAttachments> {
  const start = Date.now();
  const deadline = timeoutMs ? start + timeoutMs : 0;
  const failureCodes: string[] = [];
  const warnings: string[] = [];
  const skippedAttachmentKeys = new Set<string>();
  let skipped = metadata.overCount;
  let failed = 0;
  let reportedEligibleBytes = 0;
  const candidates = metadata.attachments.slice(0, MAX_EMAIL_ATTACHMENTS).filter((attachment) => {
    if (attachment.reportedSize > MAX_EMAIL_ATTACHMENT_BYTES) {
      skipped += 1;
      warnings.push("An email attachment exceeded the 25 MiB attachment limit.");
      return false;
    }
    if (isUnsupportedAttachmentType(attachment.attachmentType)) {
      skipped += 1;
      warnings.push("A cloud-link attachment is unsupported and will not be copied to Halo.");
      return false;
    }
    if (reportedEligibleBytes + attachment.reportedSize > MAX_EMAIL_ATTACHMENT_TOTAL_BYTES) {
      skipped += 1;
      warnings.push("Some email attachments exceeded the 50 MiB email limit.");
      return false;
    }
    reportedEligibleBytes += attachment.reportedSize;
    return true;
  });
  if (metadata.overCount) {
    warnings.push("Some email attachments exceeded the 20 attachment limit.");
  }

  const retrieved = await mapWithConcurrency(
    candidates,
    EMAIL_ATTACHMENT_READ_CONCURRENCY,
    async (attachment) => {
      try {
        const remaining = deadline ? Math.max(0, deadline - Date.now()) : 0;
        if (deadline && !remaining) {
          throw new Error("Email attachment retrieval timed out.");
        }
        const content = await withOptionalTimeout(
          readAttachmentContent(item, attachment.outlookAttachmentId),
          remaining,
          "Email attachment retrieval timed out."
        );
        const encoded = encodeAttachmentContent(content.content, content.format, attachment);
        const contentSha256 = encoded.unsupported ? "" : await hashBase64(encoded.contentBase64);
        return { attachment, contentSha256, ...encoded };
      } catch (error) {
        failed += 1;
        const failureCode = normalizeAttachmentReadFailureCode(error);
        if (!failureCodes.includes(failureCode)) {
          failureCodes.push(failureCode);
        }
        warnings.push("An email attachment could not be read from Outlook.");
        return null;
      }
    }
  );

  let decodedTotal = 0;
  const uploads: EmailAttachmentUpload[] = [];
  retrieved.filter(isPresent).forEach((entry) => {
    if (entry.unsupported) {
      skipped += 1;
      skippedAttachmentKeys.add(entry.attachment.attachmentKey);
      warnings.push("A cloud attachment could not be copied to Halo.");
      return;
    }
    if (entry.decodedSize > MAX_EMAIL_ATTACHMENT_BYTES) {
      skipped += 1;
      skippedAttachmentKeys.add(entry.attachment.attachmentKey);
      warnings.push("An email attachment exceeded the 25 MiB attachment limit.");
      return;
    }
    if (decodedTotal + entry.decodedSize > MAX_EMAIL_ATTACHMENT_TOTAL_BYTES) {
      skipped += 1;
      skippedAttachmentKeys.add(entry.attachment.attachmentKey);
      warnings.push("Some email attachments exceeded the 50 MiB email limit.");
      return;
    }
    decodedTotal += entry.decodedSize;
    uploads.push({
      attachmentKey: entry.attachment.attachmentKey,
      attachmentType: entry.attachment.attachmentType,
      contentBase64: entry.contentBase64,
      contentFormat: entry.contentFormat,
      contentSha256: entry.contentSha256,
      contentType: entry.contentType,
      decodedSize: entry.decodedSize,
      name: entry.name,
      reportedSize: entry.attachment.reportedSize,
    });
  });

  return {
    // A transient Outlook content-read failure is still an eligible file. Keep
    // its descriptor in the server inventory so the stage remains pending and
    // can be retried instead of incorrectly appearing complete.
    descriptors: candidates
      .filter((attachment) => !skippedAttachmentKeys.has(attachment.attachmentKey))
      .map((attachment) => ({
        ...toDescriptor(attachment),
        contentSha256:
          uploads.find((upload) => upload.attachmentKey === attachment.attachmentKey)
            ?.contentSha256 || "",
      })),
    emailAttachmentFingerprint: metadata.emailAttachmentFingerprint,
    failureCodes,
    failed,
    skipped,
    timings: { outlookReadMs: Date.now() - start },
    uploads,
    warnings: uniqueWarnings(warnings),
  };
}

export function sanitizeAttachmentName(value: string, fallback = "email-attachment.bin"): string {
  const basename = Array.from(
    String(value || "")
      .replace(/\\/g, "/")
      .split("/")
      .pop()!
  )
    .filter((character) => {
      const code = character.charCodeAt(0);
      return code >= 32 && code !== 127;
    })
    .join("")
    .trim();
  if (!basename || basename === "." || basename === "..") {
    return fallback;
  }
  if (basename.length <= 255) {
    return basename;
  }
  const extensionIndex = basename.lastIndexOf(".");
  const extension = extensionIndex > 0 ? basename.slice(extensionIndex, extensionIndex + 32) : "";
  return `${basename.slice(0, Math.max(1, 255 - extension.length))}${extension}`;
}

function toDescriptor(attachment: EmailAttachmentDescriptor): EmailAttachmentDescriptor {
  return {
    attachmentKey: attachment.attachmentKey,
    attachmentType: attachment.attachmentType,
    contentType: attachment.contentType,
    name: attachment.name,
    reportedSize: attachment.reportedSize,
  };
}

function isUnsupportedAttachmentType(value: string): boolean {
  return ["cloud", "link", "reference", "url"].includes(String(value || "").toLowerCase());
}

type NormalizedAttachment = {
  id: string;
  name: string;
  size: number;
  isInline: boolean;
  contentId: string;
  contentType: string;
  attachmentType: string;
};

async function readAttachmentMetadata(item: AttachmentItem): Promise<NormalizedAttachment[]> {
  let rawAttachments: unknown[] = [];
  if (item.getAttachmentsAsync) {
    rawAttachments = await new Promise((resolve) => {
      item.getAttachmentsAsync!((result) => {
        resolve(
          result.status === Office.AsyncResultStatus.Succeeded
            ? result.value || []
            : Array.isArray(item.attachments)
              ? item.attachments
              : []
        );
      });
    });
  } else if (Array.isArray(item.attachments)) {
    rawAttachments = item.attachments;
  }

  return rawAttachments.map(normalizeAttachment).filter(isPresent);
}

function normalizeAttachment(value: unknown): NormalizedAttachment | null {
  const attachment = (value || {}) as Record<string, unknown>;
  const id = String(attachment.id || "");
  if (!id) {
    return null;
  }
  return {
    attachmentType: String(attachment.attachmentType || "file").toLowerCase(),
    contentId: normalizeContentId(attachment.contentId),
    contentType: String(attachment.contentType || "").toLowerCase(),
    id,
    isInline: attachment.isInline === true,
    name: String(attachment.name || "email-attachment.bin"),
    size: Math.max(0, Number(attachment.size) || 0),
  };
}

function normalizeContentId(value: unknown): string {
  let normalized = String(value || "")
    .replace(/^\s*cid:/i, "")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&amp;/gi, "&")
    .trim();
  try {
    normalized = decodeURIComponent(normalized);
  } catch {
    // Keep literal percent signs from Outlook attachment metadata.
  }
  return normalized
    .replace(/^<+|>+$/g, "")
    .trim()
    .toLowerCase();
}

function readAttachmentContent(
  item: AttachmentItem,
  attachmentId: string
): Promise<{ content: string; format: string }> {
  return new Promise((resolve, reject) => {
    if (!item.getAttachmentContentAsync) {
      reject(new Error("Outlook does not support attachment content retrieval."));
      return;
    }
    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
        const resultError = result.error as
          { code?: string | number; message?: string } | undefined;
        const error = new Error(
          resultError?.message || "Outlook could not read an email attachment."
        ) as Error & { code?: string };
        if (resultError?.code !== undefined && resultError.code !== null) {
          error.code = String(resultError.code);
        }
        reject(error);
        return;
      }
      resolve({
        content: String(result.value.content || ""),
        format: String(result.value.format || "").toLowerCase(),
      });
    });
  });
}

function normalizeAttachmentReadFailureCode(error: unknown): string {
  const value = error as { code?: unknown; message?: unknown };
  const code = String(value?.code || "").toLowerCase();
  const message = String(value?.message || "").toLowerCase();
  const combined = `${code} ${message}`;
  if (/invalid[\s_-]*attachment[\s_-]*id/.test(combined)) {
    return "invalid-attachment-id";
  }
  if (/timed?\s*out|timeout/.test(combined)) {
    return "timeout";
  }
  if (/not\s*supported|unsupported|not-implemented/.test(combined)) {
    return "not-supported";
  }
  if (/invalid\s*(base64|content)/.test(combined)) {
    return "invalid-content";
  }
  return "read-failed";
}

function encodeAttachmentContent(
  content: string,
  format: string,
  attachment: LocalEmailAttachment
): {
  contentBase64: string;
  contentFormat: "base64" | "eml" | "icalendar";
  contentType: string;
  decodedSize: number;
  name: string;
  unsupported: boolean;
} {
  if (format.includes("url")) {
    return {
      contentBase64: "",
      contentFormat: "base64",
      contentType: attachment.contentType,
      decodedSize: 0,
      name: attachment.name,
      unsupported: true,
    };
  }
  if (format.includes("eml")) {
    const contentBase64 = utf8ToBase64(content);
    return {
      contentBase64,
      contentFormat: "eml",
      contentType: "message/rfc822",
      decodedSize: getDecodedBase64Length(contentBase64),
      name: ensureExtension(attachment.name, ".eml"),
      unsupported: false,
    };
  }
  if (format.includes("icalendar") || format.includes("calendar")) {
    const contentBase64 = utf8ToBase64(content);
    return {
      contentBase64,
      contentFormat: "icalendar",
      contentType: "text/calendar",
      decodedSize: getDecodedBase64Length(contentBase64),
      name: ensureExtension(attachment.name, ".ics"),
      unsupported: false,
    };
  }
  const contentBase64 = String(content || "").replace(/\s/g, "");
  if (!isValidBase64(contentBase64)) {
    throw new Error("Outlook returned invalid Base64 attachment content.");
  }
  return {
    contentBase64,
    contentFormat: "base64",
    contentType: attachment.contentType || "application/octet-stream",
    decodedSize: getDecodedBase64Length(contentBase64),
    name: attachment.name,
    unsupported: false,
  };
}

function ensureExtension(name: string, extension: string): string {
  const sanitized = sanitizeAttachmentName(name);
  if (sanitized.toLowerCase().endsWith(extension)) {
    return sanitized;
  }
  return `${sanitized.slice(0, Math.max(1, 255 - extension.length))}${extension}`;
}

function isValidBase64(value: string): boolean {
  if (!value || value.length % 4 !== 0 || !/^[a-z0-9+/]*={0,2}$/i.test(value)) {
    return false;
  }
  try {
    atob(value.slice(0, Math.min(value.length, 4096)));
    return true;
  } catch {
    return false;
  }
}

function getDecodedBase64Length(value: string): number {
  const padding = value.endsWith("==") ? 2 : value.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor((value.length * 3) / 4) - padding);
}

function utf8ToBase64(value: string): string {
  const bytes = encodeUtf8(value);
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    const chunk = bytes.slice(offset, offset + chunkSize);
    binary += String.fromCharCode.apply(null, Array.from(chunk));
  }
  return btoa(binary);
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

async function hashText(value: string): Promise<string> {
  const bytes = encodeUtf8(String(value || ""));
  let cryptoProvider = crypto;
  if (!cryptoProvider || !cryptoProvider.subtle) {
    cryptoProvider = (window as unknown as { msCrypto: typeof crypto }).msCrypto;
  }
  if (!cryptoProvider || !cryptoProvider.subtle) {
    throw new Error("Web Crypto is unavailable.");
  }
  const digestInput = new Uint8Array(bytes.length);
  digestInput.set(bytes);
  const digest = await cryptoProvider.subtle.digest("SHA-256", digestInput.buffer);
  return Array.from(new Uint8Array(digest))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

async function hashBase64(value: string): Promise<string> {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  let cryptoProvider = crypto;
  if (!cryptoProvider || !cryptoProvider.subtle) {
    cryptoProvider = (window as unknown as { msCrypto: typeof crypto }).msCrypto;
  }
  if (!cryptoProvider || !cryptoProvider.subtle) {
    throw new Error("Web Crypto is unavailable.");
  }
  const digest = await cryptoProvider.subtle.digest("SHA-256", bytes.buffer);
  bytes.fill(0);
  return Array.from(new Uint8Array(digest))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
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

async function mapWithConcurrency<T, TResult>(
  values: T[],
  limit: number,
  worker: (value: T) => Promise<TResult>
): Promise<TResult[]> {
  const results = new Array<TResult>(values.length);
  let nextIndex = 0;
  async function run() {
    while (nextIndex < values.length) {
      const index = nextIndex++;
      results[index] = await worker(values[index]);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, values.length) }, () => run()));
  return results;
}

function uniqueWarnings(warnings: string[]): string[] {
  return Array.from(new Set(warnings));
}

function isPresent<T>(value: T | null | undefined): value is T {
  return value !== null && value !== undefined;
}
