const assert = require("assert");
const crypto = require("crypto");
const { createTestStore } = require("./testDatabase");
const {
  extractReferencedContentIds,
  lookupInlineImages,
  normalizeContentId,
  prefetchInlineImages,
  resolveInlineImages,
  rewriteInlineImages,
} = require("./inlineImages");

function makeUpload(bytes, overrides = {}) {
  return {
    contentBase64: bytes.toString("base64"),
    contentType: "image/png",
    name: "signature.png",
    sha256: crypto.createHash("sha256").update(bytes).digest("hex"),
    ...overrides,
  };
}

async function run() {
  assert.strictEqual(normalizeContentId(" CID:<Logo%40Example.COM> "), "logo@example.com");
  assert.deepStrictEqual(
    extractReferencedContentIds(
      '<img src="CID:&lt;Logo@Example.com&gt;"><img originalsrc="cid:logo@example.com">' +
        '<img src="https://cdn.example.com/logo.png"><img src="data:image/png;base64,AA==">'
    ),
    ["logo@example.com"]
  );

  const rewrite = rewriteInlineImages(
    '<img alt="Public" src="https://cdn.example.com/logo.png">' +
      '<img alt="CID logo" src="https://outlook.example/trace" originalsrc="cid:logo">',
    new Map([
      ["logo", { renderableUrl: "https://customer.halopsa.com/api/attachment/image?token=safe" }],
    ])
  );
  assert.match(rewrite.bodyHtml, /https:\/\/cdn\.example\.com\/logo\.png/);
  assert.match(rewrite.bodyHtml, /src="https:\/\/customer\.halopsa\.com\/api\/attachment/);
  assert.match(rewrite.bodyHtml, /originalsrc="https:\/\/customer\.halopsa\.com\/api\/attachment/);

  const store = await createTestStore();
  const png = Buffer.concat([Buffer.from("89504e470d0a1a0a", "hex"), Buffer.from("same footer")]);
  const upload = makeUpload(png);
  const input = {
    inlineImageRefs: [
      { contentId: "<Logo@Example.com>", sha256: upload.sha256 },
      { contentId: "logo@example.com", sha256: upload.sha256 },
    ],
    inlineImageUploads: [upload],
  };
  let uploadCalls = 0;
  const uploadImage = async (image) => {
    uploadCalls += 1;
    assert.strictEqual(image.ticketId, 1001);
    return {
      attachmentId: "halo-attachment-1",
      renderableUrl: "/api/attachment/image?token=footer",
    };
  };

  const first = await resolveInlineImages({
    bodyHtml: '<p>Hello</p><img src="cid:logo@example.com">',
    haloTenant: "https://customer.halopsa.com/path",
    input,
    store,
    ticketId: 1001,
    uploadImage,
  });
  assert.strictEqual(first.summary.referenced, 1);
  assert.strictEqual(first.summary.uploaded, 1);
  assert.strictEqual(first.summary.failed, 0);
  assert.match(first.bodyHtml, /token=footer/);
  assert.strictEqual(uploadCalls, 1);

  const second = await resolveInlineImages({
    bodyHtml: '<img src="CID:<LOGO@EXAMPLE.COM>">',
    haloTenant: "https://customer.halopsa.com",
    input: { inlineImageRefs: input.inlineImageRefs },
    store,
    ticketId: 2002,
    uploadImage: async () => {
      throw new Error("Cache hits must not upload");
    },
  });
  assert.strictEqual(second.summary.cacheHits, 1);
  assert.strictEqual(second.summary.uploaded, 0);
  assert.strictEqual(second.summary.failed, 0);
  assert.strictEqual(uploadCalls, 1);

  let privateUploadCalls = 0;
  const privateFirst = await resolveInlineImages({
    bodyHtml: '<img src="cid:logo@example.com">',
    haloTenant: "https://customer.halopsa.com",
    input,
    showForUsers: false,
    store,
    ticketId: 1001,
    uploadImage: async () => {
      privateUploadCalls += 1;
      return {
        attachmentId: "halo-private-attachment-1",
        renderableUrl: "/api/attachment/image?token=private-footer",
      };
    },
  });
  assert.strictEqual(privateFirst.summary.uploaded, 1);
  assert.strictEqual(privateFirst.summary.cacheHits, 0);
  assert.match(privateFirst.bodyHtml, /token=private-footer/);
  assert.strictEqual(privateUploadCalls, 1);
  assert.strictEqual(
    (await store.getInlineImageCacheEntries("https://customer.halopsa.com", [upload.sha256]))[0]
      .showForUsers,
    true
  );
  assert.strictEqual(
    (
      await store.getInlineImageCacheEntries(
        "https://customer.halopsa.com",
        [upload.sha256],
        false
      )
    )[0].showForUsers,
    false
  );

  const privateSecond = await resolveInlineImages({
    bodyHtml: '<img src="cid:logo@example.com">',
    haloTenant: "https://customer.halopsa.com",
    input: { inlineImageRefs: input.inlineImageRefs },
    showForUsers: false,
    store,
    ticketId: 1001,
    uploadImage: async () => {
      throw new Error("Private cache hits must not reuse public entries or upload again");
    },
  });
  assert.strictEqual(privateSecond.summary.cacheHits, 1);
  assert.match(privateSecond.bodyHtml, /token=private-footer/);

  const genericAttachmentBytes = Buffer.concat([
    Buffer.from("89504e470d0a1a0a", "hex"),
    Buffer.from("temporary attachment URL"),
  ]);
  const genericAttachmentUpload = makeUpload(genericAttachmentBytes);
  const rejectedGenericAttachment = await resolveInlineImages({
    bodyHtml: '<img alt="Generic attachment" src="cid:generic-attachment">',
    haloTenant: "https://customer.halopsa.com",
    input: {
      inlineImageRefs: [
        { contentId: "generic-attachment", sha256: genericAttachmentUpload.sha256 },
      ],
      inlineImageUploads: [genericAttachmentUpload],
    },
    store,
    ticketId: 1001,
    uploadImage: async () => ({
      attachmentId: "not-an-inline-image",
      renderableUrl: "/api/Attachment/99",
    }),
  });
  assert.strictEqual(rejectedGenericAttachment.summary.uploaded, 0);
  assert.strictEqual(rejectedGenericAttachment.summary.failed, 1);
  assert.match(rejectedGenericAttachment.bodyHtml, /Inline image unavailable: Generic attachment/);
  assert.deepStrictEqual(
    (
      await lookupInlineImages(store, "https://customer.halopsa.com", [
        { sha256: genericAttachmentUpload.sha256 },
      ])
    ).hits,
    []
  );

  const cachedHashes = ["a".repeat(64), "b".repeat(64), "c".repeat(64)];
  for (const [index, sha256] of cachedHashes.entries()) {
    await store.upsertInlineImageCacheEntry({
      byteLength: 2 * 1024 * 1024,
      filename: `large-${index}.png`,
      haloAttachmentId: `large-${index}`,
      haloTenant: "https://customer.halopsa.com",
      mediaType: "image/png",
      renderableUrl: `https://customer.halopsa.com/api/attachment/image?token=large-${index}`,
      sha256,
    });
  }
  const cachedTotalLimit = await resolveInlineImages({
    bodyHtml: cachedHashes.map((_, index) => `<img src="cid:large-${index}">`).join(""),
    haloTenant: "https://customer.halopsa.com",
    input: {
      inlineImageRefs: cachedHashes.map((sha256, index) => ({
        contentId: `large-${index}`,
        sha256,
      })),
    },
    store,
    ticketId: 1001,
    uploadImage,
  });
  assert.strictEqual(cachedTotalLimit.summary.cacheHits, 2);
  assert.strictEqual(cachedTotalLimit.summary.failed, 1);
  assert.match(cachedTotalLimit.summary.warnings.join(" "), /5 MiB/);

  const isolated = await lookupInlineImages(store, "https://other.halopsa.com", [
    { sha256: upload.sha256 },
  ]);
  assert.deepStrictEqual(isolated.hits, []);
  assert.deepStrictEqual(isolated.misses, [upload.sha256]);

  const refreshBytes = Buffer.concat([
    Buffer.from("89504e470d0a1a0a", "hex"),
    Buffer.from("refresh invalid cache metadata"),
  ]);
  const refreshUpload = makeUpload(refreshBytes);
  await store.upsertInlineImageCacheEntry({
    byteLength: refreshBytes.length,
    filename: "refresh.png",
    haloAttachmentId: "stale-attachment",
    haloTenant: "https://customer.halopsa.com",
    mediaType: "image/png",
    renderableUrl: "https://wrong-tenant.example/api/attachment/image?token=stale",
    sha256: refreshUpload.sha256,
  });
  const refreshedCache = await resolveInlineImages({
    bodyHtml: '<img src="cid:refresh-cache">',
    haloTenant: "https://customer.halopsa.com",
    input: {
      inlineImageRefs: [{ contentId: "refresh-cache", sha256: refreshUpload.sha256 }],
      inlineImageUploads: [refreshUpload],
    },
    store,
    ticketId: 1001,
    uploadImage,
  });
  assert.strictEqual(refreshedCache.summary.uploaded, 1);
  assert.strictEqual(refreshedCache.summary.failed, 0);
  assert.match(refreshedCache.summary.warnings.join(" "), /cached inline image was invalid/i);

  const unavailable = await resolveInlineImages({
    addWarningFooter: true,
    bodyHtml: '<p>Before</p><img alt="Company logo &amp; banner" src="cid:missing">',
    haloTenant: "https://customer.halopsa.com",
    input: {},
    store,
    ticketId: 1001,
    uploadImage,
  });
  assert.strictEqual(unavailable.summary.failed, 1);
  assert.match(unavailable.bodyHtml, /Inline image unavailable: Company logo &amp; banner/);
  assert.match(unavailable.bodyHtml, /could not be imported/);

  const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"></svg>');
  const rejected = await resolveInlineImages({
    bodyHtml: '<img src="cid:unsafe">',
    haloTenant: "https://customer.halopsa.com",
    input: {
      inlineImageRefs: [{ contentId: "unsafe", sha256: makeUpload(svg).sha256 }],
      inlineImageUploads: [makeUpload(svg, { contentType: "image/svg+xml", name: "unsafe.svg" })],
    },
    store,
    ticketId: 1001,
    uploadImage,
  });
  assert.strictEqual(rejected.summary.failed, 1);
  assert.match(rejected.summary.warnings.join(" "), /unsupported format/);

  const invalidBase64Bytes = Buffer.from("invalid-base64-test");
  const invalidBase64 = await resolveInlineImages({
    bodyHtml: '<img src="cid:invalid-base64">',
    haloTenant: "https://customer.halopsa.com",
    input: {
      inlineImageRefs: [
        {
          contentId: "invalid-base64",
          sha256: crypto.createHash("sha256").update(invalidBase64Bytes).digest("hex"),
        },
      ],
      inlineImageUploads: [
        makeUpload(invalidBase64Bytes, {
          contentBase64: "AAAAA===",
          contentType: "image/png",
        }),
      ],
    },
    store,
    ticketId: 1001,
    uploadImage,
  });
  assert.strictEqual(invalidBase64.summary.failed, 1);
  assert.match(invalidBase64.summary.warnings.join(" "), /invalid Base64/);

  const octetStreamPng = Buffer.concat([
    Buffer.from("89504e470d0a1a0a", "hex"),
    Buffer.from("unknown Outlook content type"),
  ]);
  const octetStreamUpload = makeUpload(octetStreamPng, {
    contentType: "application/octet-stream",
    name: "inline-image",
  });
  const octetStreamAccepted = await resolveInlineImages({
    bodyHtml: '<img src="cid:octet-stream">',
    haloTenant: "https://customer.halopsa.com",
    input: {
      inlineImageRefs: [{ contentId: "octet-stream", sha256: octetStreamUpload.sha256 }],
      inlineImageUploads: [octetStreamUpload],
    },
    store,
    ticketId: 1001,
    uploadImage,
  });
  assert.strictEqual(octetStreamAccepted.summary.failed, 0);

  const oversizedBytes = Buffer.concat([
    Buffer.from("89504e470d0a1a0a", "hex"),
    Buffer.alloc(2 * 1024 * 1024),
  ]);
  const oversizedUpload = makeUpload(oversizedBytes);
  const oversized = await resolveInlineImages({
    bodyHtml: '<img src="cid:oversized">',
    haloTenant: "https://customer.halopsa.com",
    input: {
      inlineImageRefs: [{ contentId: "oversized", sha256: oversizedUpload.sha256 }],
      inlineImageUploads: [oversizedUpload],
    },
    store,
    ticketId: 1001,
    uploadImage,
  });
  assert.strictEqual(oversized.summary.failed, 1);
  assert.match(oversized.summary.warnings.join(" "), /2 MiB/);

  const prefetchInput = {
    ...input,
    inlineImageFingerprint: "fingerprint-1",
  };
  const prefetched = await prefetchInlineImages({
    composeOperationId: "compose-1",
    haloTenant: "https://customer.halopsa.com",
    input: prefetchInput,
    store,
    ticketId: 1001,
    uploadImage,
  });
  assert(prefetched.inlineImagePrefetchKey);
  const reused = await resolveInlineImages({
    bodyHtml: '<img src="cid:logo@example.com">',
    haloTenant: "https://customer.halopsa.com",
    input: {
      composeAttachId: "compose-1",
      inlineImageFingerprint: "fingerprint-1",
      inlineImagePrefetchKey: prefetched.inlineImagePrefetchKey,
    },
    store,
    ticketId: 1001,
    uploadImage: async () => {
      throw new Error("Prefetched sends must not upload");
    },
  });
  assert.strictEqual(reused.summary.failed, 0);
  assert.strictEqual(reused.summary.uploaded, 0);

  const wrongComposeOperation = await resolveInlineImages({
    bodyHtml: '<img src="cid:logo@example.com">',
    haloTenant: "https://customer.halopsa.com",
    input: {
      composeAttachId: "different-compose-operation",
      inlineImageFingerprint: "fingerprint-1",
      inlineImagePrefetchKey: prefetched.inlineImagePrefetchKey,
    },
    store,
    ticketId: 1001,
    uploadImage,
  });
  assert.strictEqual(wrongComposeOperation.summary.failed, 1);
  assert.match(wrongComposeOperation.summary.warnings.join(" "), /no longer matched/);

  const concurrentStore = await createTestStore();
  const concurrentPng = Buffer.concat([
    Buffer.from("89504e470d0a1a0a", "hex"),
    Buffer.from("concurrent signature"),
  ]);
  const concurrentUpload = makeUpload(concurrentPng);
  const concurrentInput = {
    inlineImageRefs: [{ contentId: "concurrent", sha256: concurrentUpload.sha256 }],
    inlineImageUploads: [concurrentUpload],
  };
  let concurrentUploadCalls = 0;
  let releaseConcurrentUpload;
  let signalConcurrentUploadStarted;
  const uploadGate = new Promise((resolve) => {
    releaseConcurrentUpload = resolve;
  });
  const uploadStarted = new Promise((resolve) => {
    signalConcurrentUploadStarted = resolve;
  });
  const concurrentUploadImage = async () => {
    concurrentUploadCalls += 1;
    signalConcurrentUploadStarted();
    await uploadGate;
    return {
      attachmentId: "concurrent-attachment",
      renderableUrl: "/api/attachment/image?token=concurrent",
    };
  };
  const concurrentOptions = {
    bodyHtml: '<img src="cid:concurrent">',
    haloTenant: "https://concurrent.halopsa.com",
    input: concurrentInput,
    store: concurrentStore,
    ticketId: 3003,
    uploadImage: concurrentUploadImage,
  };
  const concurrentResultsPromise = Promise.all([
    resolveInlineImages(concurrentOptions),
    resolveInlineImages(concurrentOptions),
  ]);
  await uploadStarted;
  assert.strictEqual(concurrentUploadCalls, 1);
  releaseConcurrentUpload();
  const concurrentResults = await concurrentResultsPromise;
  assert.strictEqual(concurrentResults[0].summary.failed, 0);
  assert.strictEqual(concurrentResults[1].summary.failed, 0);
  assert.strictEqual(
    concurrentResults[0].summary.uploaded + concurrentResults[1].summary.uploaded,
    1
  );
  await concurrentStore.close();

  await store.saveComposeInlineImagePrefetch({
    attachmentFingerprint: "expired",
    cidHash: {},
    composeOperationId: "compose-expired",
    expiresAt: Date.now() - 1,
    haloTenant: "https://customer.halopsa.com",
    prefetchKey: "expired-key",
  });
  assert.strictEqual(
    await store.getComposeInlineImagePrefetch(
      "expired-key",
      "https://customer.halopsa.com"
    ),
    null
  );

  await store.close();
  console.log("Inline image tests passed");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
