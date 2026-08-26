const assert = require("assert");
const { _test } = require("./haloAuth");

function createHaloImageToken(attachmentId) {
  const header = Buffer.from(JSON.stringify({ alg: "HS256", typ: "JWT" })).toString("base64url");
  const payload = Buffer.from(JSON.stringify({ id: attachmentId })).toString("base64url");
  return `${header}.${payload}.test-signature`;
}

async function run() {
  const originalFetch = global.fetch;
  const png = Buffer.concat([Buffer.from("89504e470d0a1a0a", "hex"), Buffer.from("halo")]);
  const attachmentId = "825458da-7545-42eb-bb16-a57737b8d821";
  const token = createHaloImageToken(attachmentId);
  let requests = 0;
  let expectedVisibility = "true";

  try {
    global.fetch = async (url, options) => {
      requests += 1;
      assert.strictEqual(String(url), "https://customer.halopsa.com/api/attachment/image");
      assert.strictEqual(options.method, "POST");
      assert.strictEqual(options.headers.Authorization, "Bearer test-token");
      assert(options.body instanceof FormData);
      assert.strictEqual(options.body.get("ticket_id"), "1001");
      assert.strictEqual(options.body.get("image_upload_id"), "0");
      assert.strictEqual(options.body.get("image_upload_key"), "");
      assert.strictEqual(options.body.get("showforusers"), expectedVisibility);
      assert.strictEqual(options.body.get("type"), null);
      assert.strictEqual(options.body.get("unique_id"), null);

      const file = options.body.get("file");
      assert.strictEqual(file.name, "signature.png");
      assert.strictEqual(file.type, "image/png");
      assert.deepStrictEqual(Buffer.from(await file.arrayBuffer()), png);

      return new Response(
        JSON.stringify({ link: `/api/attachment/image?token=${encodeURIComponent(token)}` }),
        { status: 200, headers: { "Content-Type": "application/json" } }
      );
    };

    const payload = await _test.fetchHaloAttachment(
      { haloUrl: "https://customer.halopsa.com" },
      { access_token: "test-token" },
      {
        bytes: png,
        mediaType: "image/png",
        name: "signature.png",
        ticketId: 1001,
      }
    );
    expectedVisibility = "false";
    await _test.fetchHaloAttachment(
      { haloUrl: "https://customer.halopsa.com" },
      { access_token: "test-token" },
      {
        bytes: png,
        mediaType: "image/png",
        name: "signature.png",
        ticketId: 1001,
      },
      "private-note"
    );

    const renderableUrl = _test.extractHaloAttachmentUrl(
      payload,
      "https://customer.halopsa.com"
    );
    assert.strictEqual(requests, 2);
    assert.strictEqual(
      renderableUrl,
      `https://customer.halopsa.com/api/attachment/image?token=${encodeURIComponent(token)}`
    );
    assert.strictEqual(_test.extractHaloAttachmentId(payload, renderableUrl), attachmentId);

    assert.strictEqual(
      _test.extractHaloAttachmentUrl({ token }, "https://customer.halopsa.com"),
      `https://customer.halopsa.com/api/attachment/image?token=${encodeURIComponent(token)}`
    );
    assert.strictEqual(
      _test.extractHaloAttachmentId({ ticket: { id: 999 }, attachment_id: 73 }),
      73
    );
    assert.strictEqual(
      _test.extractHaloAttachmentId({ id: 42 }, renderableUrl),
      attachmentId,
      "The attachment ID carried by Halo's signed image URL is authoritative"
    );
    assert.strictEqual(_test.extractHaloAttachmentId({ id: 42 }), 42);

    assert.strictEqual(
      _test.extractHaloAttachmentUrl(
        { link: `https://evil.example/api/attachment/image?token=${token}` },
        "https://customer.halopsa.com"
      ),
      ""
    );
    assert.strictEqual(
      _test.extractHaloAttachmentUrl(
        { link: "https://customer.halopsa.com/api/Attachment/42" },
        "https://customer.halopsa.com"
      ),
      ""
    );
    assert.strictEqual(
      _test.extractHaloAttachmentUrl(
        { link: "https://eu-cdn.haloservicedesk.com/file.png?Expires=60" },
        "https://customer.halopsa.com"
      ),
      ""
    );
    assert.strictEqual(
      _test.validateHaloInlineImageUrl(
        `/api/attachment/image?token=${encodeURIComponent(token)}`,
        "http://customer.halopsa.com"
      ),
      ""
    );

    const diagnostic = _test.sanitizeEmailAttachmentDiagnosticMessage(
      `Halo failed for secret-report.docx at ticket 9514, attachment ${attachmentId}, ` +
        `user@example.com, https://example.com/file and data_base64: ${Buffer.alloc(96, 7).toString("base64")}`,
      ["secret-report.docx", "9514", attachmentId]
    );
    assert.doesNotMatch(diagnostic, /secret-report\.docx/);
    assert.doesNotMatch(diagnostic, /9514/);
    assert.doesNotMatch(diagnostic, new RegExp(attachmentId));
    assert.doesNotMatch(diagnostic, /user@example\.com/);
    assert.doesNotMatch(diagnostic, /https:\/\//);
    assert.doesNotMatch(diagnostic, new RegExp(Buffer.alloc(96, 7).toString("base64")));
    assert.match(diagnostic, /\[redacted/);
  } finally {
    global.fetch = originalFetch;
  }

  console.log("Halo attachment tests passed");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
