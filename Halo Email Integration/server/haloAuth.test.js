const assert = require("assert");
process.env.NODE_ENV = "test";
process.env.HALO_TOKEN_ENCRYPTION_KEY = Buffer.alloc(32, 3).toString("base64url");

const fs = require("fs");
const crypto = require("crypto");
const os = require("os");
const path = require("path");
const { registerHaloAuthRoutes } = require("./haloAuth");
const { createHaloStore } = require("./haloStore");
const {
  getApiAudience,
  getAuthScopes,
  getTokenAudiences,
  validateMicrosoftClaims,
} = require("./microsoftAuth");
const { decodeEncryptionKey } = require("./tokenCrypto");

const TEST_AUTH_HEADER = "Bearer test-microsoft-token";

const microsoftAuthVerifier = {
  async verify(token) {
    if (token !== "test-microsoft-token") {
      throw new Error("Invalid test Microsoft token.");
    }

    return {
      aud: "test-addin-client-id",
      email: "support@example.com",
      name: "Support User",
      oid: "test-object-id",
      preferred_username: "support@example.com",
      tid: "test-tenant-id",
    };
  },
};

function createMockApp() {
  const routes = {
    GET: new Map(),
    POST: new Map(),
  };

  return {
    locals: {},
    routes,
    get(path, handler) {
      routes.GET.set(path, handler);
    },
    post(path, handler) {
      routes.POST.set(path, handler);
    },
  };
}

function createMockReq({ url, body, cookie, headers, params } = {}) {
  return {
    body,
    headers: {
      host: "localhost:3000",
      authorization: TEST_AUTH_HEADER,
      ...(headers || {}),
      ...(cookie ? { cookie } : {}),
    },
    originalUrl: url,
    params: params || {},
    url,
  };
}

function createMockRes() {
  return {
    body: undefined,
    headers: {},
    redirectedTo: "",
    statusCode: 200,
    status(code) {
      this.statusCode = code;
      return this;
    },
    json(body) {
      this.body = body;
      return this;
    },
    redirect(url) {
      this.statusCode = 302;
      this.redirectedTo = url;
      return this;
    },
    set(name, value) {
      this.headers[name.toLowerCase()] = value;
      return this;
    },
    setHeader(name, value) {
      this.headers[name.toLowerCase()] = value;
      return this;
    },
    send(body) {
      this.body = body;
      return this;
    },
  };
}

function jsonResponse(body, status = 200, statusText = "") {
  return new Response(JSON.stringify(body), {
    status,
    statusText,
    headers: {
      "Content-Type": "application/json",
    },
  });
}

async function invoke(app, method, path, request = {}) {
  const handler = app.routes[method].get(path);
  assert(handler, `Expected ${method} ${path} to be registered`);

  const response = createMockRes();
  await handler(createMockReq(request), response);
  return response;
}

function registerTestRoutes(app, store = createHaloStore({ dbPath: ":memory:" })) {
  registerHaloAuthRoutes(app, {
    env: {
      ...process.env,
      HALO_CLIENT_ID: "test-client-id",
      HALO_URL: "https://customer.halopsa.com/some/path",
    },
    microsoftAuth: {
      clientId: "test-addin-client-id",
    },
    microsoftAuthVerifier,
    store,
  });
  return store;
}

function createTempDbPath(name) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "halo-auth-test-"));
  return path.join(dir, `${name}.sqlite`);
}

function getCookieValue(cookieHeader, name) {
  const part = String(cookieHeader || "")
    .split(";")
    .find((entry) => entry.trim().startsWith(`${name}=`));
  return part ? decodeURIComponent(part.split("=")[1]) : "";
}

function sha256Hex(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

async function loginAndGetCookie(app) {
  const start = await invoke(app, "POST", "/api/auth/start", {
    url: "/api/auth/start",
    headers: { authorization: TEST_AUTH_HEADER },
  });
  assert.strictEqual(start.statusCode, 200, start.body && start.body.error);
  const dialogUrl = new URL(start.body.dialogUrl);
  const callback = await invoke(app, "GET", "/auth/callback", {
    url: `/auth/callback?code=test-code&state=${encodeURIComponent(
      dialogUrl.searchParams.get("state")
    )}`,
  });
  const handoffMatch = callback.body.match(/"handoffCode":"([^"]+)"/);
  assert(handoffMatch, "Expected callback page to include a handoff code");

  const complete = await invoke(app, "POST", "/api/auth/complete", {
    url: "/api/auth/complete",
    headers: { authorization: TEST_AUTH_HEADER },
    body: { handoffCode: handoffMatch[1] },
  });
  assert.strictEqual(complete.statusCode, 200);
  return complete.headers["set-cookie"];
}

function createEmailPayload(overrides = {}) {
  return {
    bodyHtml: "<p>Hello from Outlook</p>",
    bodyText: "",
    cc: [{ displayName: "Copied User", emailAddress: "cc@example.com" }],
    conversationId: "conversation-id",
    dateTimeCreated: "2026-07-07T10:00:00.000Z",
    from: { displayName: "Sender User", emailAddress: "sender@example.com" },
    inReplyToMessageIds: [],
    internetHeaders: "",
    internetMessageId: "<message@example.com>",
    itemId: "outlook-item-id",
    mailboxEmail: "support@example.com",
    normalizedSubject: "Example subject",
    referenceMessageIds: [],
    subject: "RE: Example subject",
    timeZone: "Europe/London",
    to: [{ displayName: "Support User", emailAddress: "support@example.com" }],
    ...overrides,
  };
}

function createSendPayload(overrides = {}) {
  return createEmailPayload({
    bodyHtml: "<p>Sent reply from Outlook</p>",
    bodyText: "",
    internetMessageId: "",
    itemId: "outgoing-draft-id",
    subject: "RE: Example subject",
    ...overrides,
  });
}

async function run() {
  assert.strictEqual(getApiAudience("test-client-id"), "api://test-client-id");
  assert.deepStrictEqual(getTokenAudiences("test-client-id", "api://test-client-id"), [
    "test-client-id",
    "api://test-client-id",
  ]);
  assert.deepStrictEqual(
    getTokenAudiences("spa-client-id", "api://custom-api", "api-client-id"),
    ["api-client-id", "api://custom-api"]
  );
  assert.deepStrictEqual(getAuthScopes("api://test-client-id"), [
    "api://test-client-id/access_as_user",
  ]);
  assert.deepStrictEqual(
    getAuthScopes("api://test-client-id", "api://test-client-id/custom.scope openid"),
    ["api://test-client-id/custom.scope", "openid"]
  );
  assert.throws(
    () => getAuthScopes("api://test-client-id", "openid profile email User.Read"),
    /delegated scope for the add-in API/
  );
  assert.doesNotThrow(() =>
    validateMicrosoftClaims({
      iss: "https://login.microsoftonline.com/test-tenant-id/v2.0",
      oid: "test-object-id",
      scp: "access_as_user",
      tid: "test-tenant-id",
    })
  );
  assert.throws(
    () => validateMicrosoftClaims({ oid: "test-object-id", tid: "test-tenant-id" }),
    /access_as_user scope/
  );
  assert.throws(
    () =>
      registerHaloAuthRoutes(createMockApp(), {
        env: { ...process.env, HALO_CLIENT_ID: "test-client-id", HALO_URL: "" },
      }),
    /HALO_URL must be set/
  );
  assert.throws(
    () =>
      registerHaloAuthRoutes(createMockApp(), {
        env: {
          ...process.env,
          HALO_CLIENT_ID: "test-client-id",
          HALO_URL: "http://customer.halopsa.com",
        },
      }),
    /HALO_URL must use https/
  );
  assert.throws(
    () =>
      registerHaloAuthRoutes(createMockApp(), {
        env: { ...process.env, HALO_CLIENT_ID: "", HALO_URL: "https://customer.halopsa.com" },
      }),
    /HALO_CLIENT_ID must be set/
  );
  assert.throws(
    () => decodeEncryptionKey("", { NODE_ENV: "production" }),
    /HALO_TOKEN_ENCRYPTION_KEY/
  );
  assert.throws(
    () => decodeEncryptionKey(Buffer.alloc(31, 1).toString("base64"), { NODE_ENV: "production" }),
    /exactly 32 bytes/
  );

  const schemaStore = createHaloStore({ dbPath: ":memory:" });
  const schemaUser = schemaStore.upsertUser({
    objectId: "schema-object-id",
    tenantId: "schema-tenant-id",
  });
  assert(schemaUser.id);
  schemaStore.close();

  const invalidAuthApp = createMockApp();
  registerTestRoutes(invalidAuthApp);
  const invalidAuthStatus = await invoke(invalidAuthApp, "GET", "/api/auth/status", {
    url: "/api/auth/status",
    headers: { authorization: "Bearer invalid-token" },
  });
  assert.strictEqual(invalidAuthStatus.statusCode, 401);
  assert.strictEqual(invalidAuthStatus.body.authenticated, false);
  assert.match(invalidAuthStatus.body.error, /Microsoft add-in authentication failed/);

  const app = createMockApp();
  const store = registerTestRoutes(app);

  const start = await invoke(app, "POST", "/api/auth/start", {
    url: "/api/auth/start",
  });
  assert.strictEqual(start.statusCode, 200);
  assert.match(start.body.dialogUrl, /^https:\/\/localhost:3000\/auth\/start\?state=/);

  const dialogUrl = new URL(start.body.dialogUrl);
  const redirect = await invoke(app, "GET", "/auth/start", {
    url: `/auth/start?state=${encodeURIComponent(dialogUrl.searchParams.get("state"))}`,
  });
  assert.strictEqual(redirect.statusCode, 302);

  const haloAuthUrl = new URL(redirect.redirectedTo);
  assert.strictEqual(haloAuthUrl.origin, "https://customer.halopsa.com");
  assert.strictEqual(haloAuthUrl.pathname, "/auth/authorize");
  assert.strictEqual(haloAuthUrl.searchParams.get("response_type"), "code");
  assert.strictEqual(haloAuthUrl.searchParams.get("client_id"), "test-client-id");
  assert.strictEqual(
    haloAuthUrl.searchParams.get("redirect_uri"),
    "https://localhost:3000/auth/callback"
  );
  assert.strictEqual(haloAuthUrl.searchParams.get("scope"), "all");
  assert.strictEqual(haloAuthUrl.searchParams.get("code_challenge_method"), "S256");
  assert(haloAuthUrl.searchParams.get("code_challenge"));

  const status = await invoke(app, "GET", "/api/auth/status", {
    url: "/api/auth/status",
  });
  assert.deepStrictEqual(status.body, {
    authenticated: false,
    haloUrl: null,
    expiresAt: null,
  });

  const originalFetch = global.fetch;
  let tokenFetchCount = 0;
  let apiFetchCount = 0;

  global.fetch = async (requestUrl, options = {}) => {
    const url = String(requestUrl);

    if (url === "https://customer.halopsa.com/auth/token") {
      tokenFetchCount += 1;
      const form = new URLSearchParams(options.body);
      assert.strictEqual(form.get("scope"), "all");
      return jsonResponse({
        access_token: "test-access-token",
        expires_in: 3600,
        refresh_token: "test-refresh-token",
      });
    }

    if (url === "https://customer.halopsa.com/api/Tickets?count=1") {
      apiFetchCount += 1;
      return jsonResponse({ message: "Tickets permission missing" }, 403, "Forbidden");
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  };

  try {
    const callback = await invoke(app, "GET", "/auth/callback", {
      url: `/auth/callback?code=test-code&state=${encodeURIComponent(
        dialogUrl.searchParams.get("state")
      )}`,
    });
    assert.strictEqual(callback.statusCode, 200);
    assert.match(callback.body, /Halo API Auth works/);
    assert.strictEqual(tokenFetchCount, 1);
    assert.strictEqual(apiFetchCount, 0);

    const handoffMatch = callback.body.match(/"handoffCode":"([^"]+)"/);
    assert(handoffMatch, "Expected callback page to include a handoff code");

    const complete = await invoke(app, "POST", "/api/auth/complete", {
      url: "/api/auth/complete",
      body: { handoffCode: handoffMatch[1] },
    });
    assert.strictEqual(complete.statusCode, 200);
    assert.strictEqual(complete.body.authenticated, true);

    const ping = await invoke(app, "GET", "/api/halo/ping", {
      url: "/api/halo/ping",
      cookie: complete.headers["set-cookie"],
    });
    assert.strictEqual(ping.statusCode, 502);
    assert.strictEqual(ping.body.ok, false);
    assert.match(ping.body.error, /HTTP 403 Forbidden/);
    assert.match(ping.body.error, /Tickets permission missing/);
    assert.strictEqual(ping.body.debug.phase, "api-test");
    assert.strictEqual(ping.body.debug.status, 403);
    assert.strictEqual(
      ping.body.debug.endpoint,
      "https://customer.halopsa.com/api/Tickets?count=1"
    );
    assert.strictEqual(ping.body.debug.bodyExcerpt, '{"message":"Tickets permission missing"}');
    assert.strictEqual(ping.body.debug.requestedScope, "all");
  } finally {
    global.fetch = originalFetch;
  }

  let ticketTokenFetchCount = 0;
  let ticketListFetchCount = 0;
  let ticketSearchFetchCount = 0;

  global.fetch = async (requestUrl, options = {}) => {
    const url = String(requestUrl);

    if (url === "https://customer.halopsa.com/auth/token") {
      ticketTokenFetchCount += 1;
      const form = new URLSearchParams(options.body);
      assert.strictEqual(form.get("scope"), "all");
      return jsonResponse({
        access_token: "test-access-token",
        expires_in: 3600,
        refresh_token: "test-refresh-token",
      });
    }

    if (
      url ===
      "https://customer.halopsa.com/api/Tickets?count=50&open_only=true&mine=true&includeagent=true&includestatus=true"
    ) {
      ticketListFetchCount += 1;
      return jsonResponse({
        tickets: [
          {
            id: 1001,
            ticketnumber: "T1001",
            summary: "Open assigned ticket",
            status: "In Progress",
            agent_id: 123,
            client_name: "Digital Origin",
            agent_name: "Sebastian",
          },
          {
            id: 1002,
            ticketnumber: "T1002",
            summary: "Closed assigned ticket",
            status: "Closed",
            agent_id: 123,
          },
          {
            id: 1003,
            ticketnumber: "T1003",
            summary: "Second open mine ticket",
            status: "Open",
            agent_id: 456,
          },
        ],
      });
    }

    if (
      url ===
      "https://customer.halopsa.com/api/Tickets?count=20&search=2200&includeagent=true&includestatus=true"
    ) {
      ticketSearchFetchCount += 1;
      return jsonResponse({
        tickets: [
          {
            id: 2200,
            ticketnumber: "0002200",
            summary: "Closed ticket assigned to another agent",
            status: "Closed",
            agent_id: 999,
            client_name: "Another Customer",
            agent_name: "Another Agent",
          },
          {
            id: 7777,
            ticketnumber: "0007777",
            summary: "Reference to 2200 in the summary",
            status: "Open",
          },
        ],
      });
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  };

  try {
    const ticketsStart = await invoke(app, "POST", "/api/auth/start", {
      url: "/api/auth/start",
    });
    const ticketsDialogUrl = new URL(ticketsStart.body.dialogUrl);
    const ticketsCallback = await invoke(app, "GET", "/auth/callback", {
      url: `/auth/callback?code=test-code&state=${encodeURIComponent(
        ticketsDialogUrl.searchParams.get("state")
      )}`,
    });
    const ticketsHandoffMatch = ticketsCallback.body.match(/"handoffCode":"([^"]+)"/);
    assert(ticketsHandoffMatch, "Expected ticket callback page to include a handoff code");

    const ticketsComplete = await invoke(app, "POST", "/api/auth/complete", {
      url: "/api/auth/complete",
      body: { handoffCode: ticketsHandoffMatch[1] },
    });
    const tickets = await invoke(app, "GET", "/api/halo/tickets", {
      url: "/api/halo/tickets",
      cookie: ticketsComplete.headers["set-cookie"],
    });

    assert.strictEqual(tickets.statusCode, 200);
    assert.strictEqual(tickets.body.ok, true);
    assert.strictEqual(ticketTokenFetchCount, 1);
    assert.strictEqual(ticketListFetchCount, 1);
    assert.deepStrictEqual(tickets.body.tickets, [
      {
        id: "1001",
        ticketNumber: "T1001",
        summary: "Open assigned ticket",
        status: "In Progress",
        client: "Digital Origin",
        agent: "Sebastian",
      },
      {
        id: "1003",
        ticketNumber: "T1003",
        summary: "Second open mine ticket",
        status: "Open",
        client: "",
        agent: "",
      },
    ]);

    const ticketSearch = await invoke(app, "GET", "/api/halo/tickets/search", {
      url: "/api/halo/tickets/search?ticketNumber=%5BID%3A%202200%5D",
      cookie: ticketsComplete.headers["set-cookie"],
    });

    assert.strictEqual(ticketSearch.statusCode, 200);
    assert.strictEqual(ticketSearch.body.ok, true);
    assert.strictEqual(ticketSearchFetchCount, 1);
    assert.deepStrictEqual(ticketSearch.body.tickets, [
      {
        id: "2200",
        ticketNumber: "0002200",
        summary: "Closed ticket assigned to another agent",
        status: "Closed",
        client: "Another Customer",
        agent: "Another Agent",
      },
    ]);

    const emptyTicketSearch = await invoke(app, "GET", "/api/halo/tickets/search", {
      url: "/api/halo/tickets/search?ticketNumber=",
      cookie: ticketsComplete.headers["set-cookie"],
    });

    assert.strictEqual(emptyTicketSearch.statusCode, 400);
    assert.strictEqual(emptyTicketSearch.body.ok, false);
    assert.match(emptyTicketSearch.body.error, /ticket number/i);
  } finally {
    global.fetch = originalFetch;
  }

  const unauthenticatedAttach = await invoke(app, "POST", "/api/halo/tickets/:ticketId/email", {
    url: "/api/halo/tickets/1001/email",
    params: { ticketId: "1001" },
    headers: { authorization: "" },
    body: createEmailPayload(),
  });
  assert.strictEqual(unauthenticatedAttach.statusCode, 401);
  assert.strictEqual(unauthenticatedAttach.body.ok, false);

  const unauthenticatedAutoAttach = await invoke(app, "POST", "/api/halo/email/auto-attach", {
    url: "/api/halo/email/auto-attach",
    headers: { authorization: "" },
    body: createEmailPayload(),
  });
  assert.strictEqual(unauthenticatedAutoAttach.statusCode, 401);
  assert.strictEqual(unauthenticatedAutoAttach.body.ok, false);

  const noSessionSendAutoAttach = await invoke(app, "POST", "/api/halo/email/send-auto-attach", {
    url: "/api/halo/email/send-auto-attach",
    headers: { authorization: "" },
    body: createSendPayload(),
  });
  assert.strictEqual(noSessionSendAutoAttach.statusCode, 200);
  assert.strictEqual(noSessionSendAutoAttach.body.ok, true);
  assert.strictEqual(noSessionSendAutoAttach.body.status, "no-session");

  let attachTokenFetchCount = 0;
  let attachActionFetchCount = 0;
  let failNextAutoAttach = true;
  let failNextSentAutoAttach = true;
  const attachActions = [];

  global.fetch = async (requestUrl, options = {}) => {
    const url = String(requestUrl);

    if (url === "https://customer.halopsa.com/auth/token") {
      attachTokenFetchCount += 1;
      return jsonResponse({
        access_token: "attach-access-token",
        expires_in: 3600,
        refresh_token: "attach-refresh-token",
      });
    }

    if (url === "https://customer.halopsa.com/api/Actions") {
      attachActionFetchCount += 1;
      assert.strictEqual(options.method, "POST");
      const actions = JSON.parse(options.body);
      assert(Array.isArray(actions));
      assert.strictEqual(actions.length, 1);
      attachActions.push(actions[0]);
      assert.strictEqual(actions[0].ticket_id, 1001);
      assert.strictEqual(actions[0].outcome, "Email");
      assert.strictEqual(actions[0].sendemail, false);
      assert(Number.isFinite(Date.parse(actions[0].datetime)));
      assert.match(actions[0].note, /Outlook email attached to ticket/);
      assert.match(actions[0].note_html, /Sender User &lt;sender@example.com&gt;/);

      if (actions[0].email_message_id === "<failing-reply@example.com>" && failNextAutoAttach) {
        failNextAutoAttach = false;
        return jsonResponse({ message: "Temporary action failure" }, 403, "Forbidden");
      }

      if (/Fail sent reply/.test(actions[0].note_html) && failNextSentAutoAttach) {
        failNextSentAutoAttach = false;
        return jsonResponse({ message: "Temporary sent action failure" }, 403, "Forbidden");
      }

      return jsonResponse({ id: 9000 + attachActionFetchCount }, 201, "Created");
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  };

  try {
    const attachCookie = await loginAndGetCookie(app);
    const attach = await invoke(app, "POST", "/api/halo/tickets/:ticketId/email", {
      url: "/api/halo/tickets/1001/email",
      params: { ticketId: "1001" },
      cookie: attachCookie,
      body: createEmailPayload({
        bodyHtml:
          "<p>Hello from Outlook</p><blockquote><p>Prior thread content should stay for first attach</p></blockquote>",
        inReplyToMessageIds: ["<prior-reply@example.com>"],
        referenceMessageIds: ["<original-message@example.com>", "<prior-reply@example.com>"],
        ticketNumber: "T1001",
      }),
    });

    assert.strictEqual(attach.statusCode, 200);
    assert.strictEqual(attach.body.ok, true);
    assert.strictEqual(attach.body.attachMode, "full-chain");
    assert.strictEqual(attach.body.message, "Full email chain attached to Halo ticket");
    assert.strictEqual(attach.body.actionId, "9001");
    assert(attach.body.backgroundSessionId);
    assert.strictEqual(attachTokenFetchCount, 1);
    assert.strictEqual(attachActionFetchCount, 1);
    assert.strictEqual(attachActions[0].emailsubject, "RE: Example subject");
    assert.strictEqual(attachActions[0].email_message_id, "<message@example.com>");
    assert.strictEqual(attachActions[0].actioninternetmessageid, "<message@example.com>");
    assert.notStrictEqual(attachActions[0].datetime, "2026-07-07T10:00:00.000Z");
    assert.match(attachActions[0].note_html, /<p>Hello from Outlook<\/p>/);
    assert.match(attachActions[0].note_html, /Prior thread content should stay for first attach/);
    assert.match(attachActions[0].note_html, /Email date/);
    assert.match(attachActions[0].note_html, /07\/07\/2026, 11:00:00 BST/);
    assert.doesNotMatch(attachActions[0].note_html, /Internet Message ID/);
    assert.doesNotMatch(attachActions[0].note_html, /<message@example\.com>/);

    const alreadyAttached = await invoke(app, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: attachCookie,
      body: createEmailPayload(),
    });
    assert.strictEqual(alreadyAttached.statusCode, 200);
    assert.strictEqual(alreadyAttached.body.ok, true);
    assert.strictEqual(alreadyAttached.body.status, "already-attached");
    assert.strictEqual(alreadyAttached.body.ticketId, 1001);
    assert.strictEqual(alreadyAttached.body.ticketNumber, "T1001");
    assert.strictEqual(attachActionFetchCount, 1);

    const initialReferenceAlreadyCovered = await invoke(
      app,
      "POST",
      "/api/halo/email/auto-attach",
      {
        url: "/api/halo/email/auto-attach",
        cookie: attachCookie,
        body: createEmailPayload({
          bodyHtml: "<p>Earlier email already covered by first full-chain attach</p>",
          conversationId: "different-older-email-conversation-id",
          internetMessageId: "<original-message@example.com>",
          itemId: "original-message-item-id",
        }),
      }
    );
    assert.strictEqual(initialReferenceAlreadyCovered.statusCode, 200);
    assert.strictEqual(initialReferenceAlreadyCovered.body.ok, true);
    assert.strictEqual(initialReferenceAlreadyCovered.body.status, "already-attached");
    assert.strictEqual(attachActionFetchCount, 1);

    const inReplyToAttach = await invoke(app, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: attachCookie,
      body: createEmailPayload({
        bodyHtml: "<p>New reply content</p><blockquote><p>Old thread content</p></blockquote>",
        inReplyToMessageIds: ["<message@example.com>"],
        internetMessageId: "<reply@example.com>",
        itemId: "reply-item-id",
      }),
    });
    assert.strictEqual(inReplyToAttach.statusCode, 200);
    assert.strictEqual(inReplyToAttach.body.ok, true);
    assert.strictEqual(inReplyToAttach.body.status, "attached");
    assert.strictEqual(inReplyToAttach.body.ticketNumber, "T1001");
    assert.strictEqual(inReplyToAttach.body.actionId, "9002");
    assert.strictEqual(attachActionFetchCount, 2);
    assert.strictEqual(attachActions[1].email_message_id, "<reply@example.com>");
    assert.match(attachActions[1].note_html, /<p>New reply content<\/p>/);
    assert.doesNotMatch(attachActions[1].note_html, /Old thread content/);

    const referencesAttach = await invoke(app, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: attachCookie,
      body: createEmailPayload({
        bodyHtml:
          '<p>Reference reply content</p><div class="gmail_quote"><p>Quoted history</p></div>',
        internetMessageId: "<references-reply@example.com>",
        itemId: "references-reply-item-id",
        referenceMessageIds: ["<message@example.com>", "<reply@example.com>"],
      }),
    });
    assert.strictEqual(referencesAttach.statusCode, 200);
    assert.strictEqual(referencesAttach.body.ok, true);
    assert.strictEqual(referencesAttach.body.status, "attached");
    assert.strictEqual(attachActionFetchCount, 3);
    assert.match(attachActions[2].note_html, /Reference reply content/);
    assert.doesNotMatch(attachActions[2].note_html, /Quoted history/);

    const conversationAttach = await invoke(app, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: attachCookie,
      body: createEmailPayload({
        bodyHtml: "<p>Conversation match content</p>",
        inReplyToMessageIds: [],
        internetMessageId: "<conversation-reply@example.com>",
        itemId: "conversation-reply-item-id",
        referenceMessageIds: [],
      }),
    });
    assert.strictEqual(conversationAttach.statusCode, 200);
    assert.strictEqual(conversationAttach.body.ok, true);
    assert.strictEqual(conversationAttach.body.status, "attached");
    assert.strictEqual(attachActionFetchCount, 4);

    const unrelated = await invoke(app, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: attachCookie,
      body: createEmailPayload({
        conversationId: "unrelated-conversation-id",
        internetMessageId: "<unrelated@example.com>",
        itemId: "unrelated-item-id",
      }),
    });
    assert.strictEqual(unrelated.statusCode, 200);
    assert.strictEqual(unrelated.body.ok, true);
    assert.strictEqual(unrelated.body.status, "no-match");
    assert.strictEqual(attachActionFetchCount, 4);

    const differentMailbox = await invoke(app, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: attachCookie,
      body: createEmailPayload({
        inReplyToMessageIds: ["<message@example.com>"],
        internetMessageId: "<other-mailbox-reply@example.com>",
        itemId: "other-mailbox-reply-item-id",
        mailboxEmail: "other@example.com",
      }),
    });
    assert.strictEqual(differentMailbox.statusCode, 200);
    assert.strictEqual(differentMailbox.body.ok, true);
    assert.strictEqual(differentMailbox.body.status, "no-match");
    assert.strictEqual(attachActionFetchCount, 4);

    const failedAutoAttach = await invoke(app, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: attachCookie,
      body: createEmailPayload({
        bodyHtml: "<p>Retry me later</p>",
        inReplyToMessageIds: ["<message@example.com>"],
        internetMessageId: "<failing-reply@example.com>",
        itemId: "failing-reply-item-id",
      }),
    });
    assert.strictEqual(failedAutoAttach.statusCode, 502);
    assert.strictEqual(failedAutoAttach.body.ok, false);
    assert.match(failedAutoAttach.body.error, /Temporary action failure/);
    assert.strictEqual(failedAutoAttach.body.debug.phase, "email-auto-attach");
    assert.strictEqual(attachActionFetchCount, 5);

    const retryAutoAttach = await invoke(app, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: attachCookie,
      body: createEmailPayload({
        bodyHtml: "<p>Retry me later</p>",
        inReplyToMessageIds: ["<message@example.com>"],
        internetMessageId: "<failing-reply@example.com>",
        itemId: "failing-reply-item-id",
      }),
    });
    assert.strictEqual(retryAutoAttach.statusCode, 200);
    assert.strictEqual(retryAutoAttach.body.ok, true);
    assert.strictEqual(retryAutoAttach.body.status, "attached");
    assert.strictEqual(attachActionFetchCount, 6);

    const sentInReplyToAttach = await invoke(app, "POST", "/api/halo/email/send-auto-attach", {
      url: "/api/halo/email/send-auto-attach",
      cookie: attachCookie,
      body: createSendPayload({
        bodyHtml: "<p>Sent reply from Outlook</p><blockquote>Old quoted content</blockquote>",
        inReplyToMessageIds: ["<message@example.com>"],
        itemId: "sent-draft-in-reply-to-id",
      }),
    });
    assert.strictEqual(sentInReplyToAttach.statusCode, 200);
    assert.strictEqual(sentInReplyToAttach.body.ok, true);
    assert.strictEqual(sentInReplyToAttach.body.status, "attached");
    assert.strictEqual(sentInReplyToAttach.body.ticketNumber, "T1001");
    assert.strictEqual(attachActionFetchCount, 7);
    assert.match(attachActions[6].email_message_id, /^<halo-outlook-[a-f0-9]{32}@local>$/);
    assert.match(attachActions[6].note_html, /Sent reply from Outlook/);
    assert.doesNotMatch(attachActions[6].note_html, /Old quoted content/);

    const duplicateSentAttach = await invoke(app, "POST", "/api/halo/email/send-auto-attach", {
      url: "/api/halo/email/send-auto-attach",
      cookie: attachCookie,
      body: createSendPayload({
        bodyHtml: "<p>Sent reply from Outlook</p><blockquote>Old quoted content</blockquote>",
        inReplyToMessageIds: ["<message@example.com>"],
        itemId: "sent-draft-in-reply-to-id",
      }),
    });
    assert.strictEqual(duplicateSentAttach.statusCode, 200);
    assert.strictEqual(duplicateSentAttach.body.ok, true);
    assert.strictEqual(duplicateSentAttach.body.status, "already-attached");
    assert.strictEqual(attachActionFetchCount, 7);

    const sentConversationAttach = await invoke(app, "POST", "/api/halo/email/send-auto-attach", {
      url: "/api/halo/email/send-auto-attach",
      body: createSendPayload({
        backgroundSessionId: attach.body.backgroundSessionId,
        bodyHtml: "<p>Sent conversation fallback</p>",
        inReplyToMessageIds: [],
        itemId: "sent-draft-conversation-id",
      }),
    });
    assert.strictEqual(sentConversationAttach.statusCode, 200);
    assert.strictEqual(sentConversationAttach.body.ok, true);
    assert.strictEqual(sentConversationAttach.body.status, "attached");
    assert.strictEqual(attachActionFetchCount, 8);

    const expiredBackgroundSessionId = "expired-background-session-id";
    store.createBackgroundSession({
      backgroundSessionHash: sha256Hex(expiredBackgroundSessionId),
      sessionHash: sha256Hex(getCookieValue(attachCookie, "halo_session")),
      expiresAt: Date.now() - 1000,
    });
    const expiredBackgroundAttach = await invoke(app, "POST", "/api/halo/email/send-auto-attach", {
      url: "/api/halo/email/send-auto-attach",
      headers: { authorization: "" },
      body: createSendPayload({
        backgroundSessionId: expiredBackgroundSessionId,
        bodyHtml: "<p>Expired background handle</p>",
        inReplyToMessageIds: ["<message@example.com>"],
        itemId: "expired-background-draft-id",
      }),
    });
    assert.strictEqual(expiredBackgroundAttach.statusCode, 200);
    assert.strictEqual(expiredBackgroundAttach.body.ok, true);
    assert.strictEqual(expiredBackgroundAttach.body.status, "no-session");
    assert.strictEqual(attachActionFetchCount, 8);

    const unrelatedSent = await invoke(app, "POST", "/api/halo/email/send-auto-attach", {
      url: "/api/halo/email/send-auto-attach",
      cookie: attachCookie,
      body: createSendPayload({
        conversationId: "unrelated-sent-conversation-id",
        inReplyToMessageIds: [],
        itemId: "unrelated-sent-draft-id",
      }),
    });
    assert.strictEqual(unrelatedSent.statusCode, 200);
    assert.strictEqual(unrelatedSent.body.ok, true);
    assert.strictEqual(unrelatedSent.body.status, "no-match");
    assert.strictEqual(attachActionFetchCount, 8);

    const failedSentAttach = await invoke(app, "POST", "/api/halo/email/send-auto-attach", {
      url: "/api/halo/email/send-auto-attach",
      cookie: attachCookie,
      body: createSendPayload({
        bodyHtml: "<p>Fail sent reply</p>",
        inReplyToMessageIds: ["<message@example.com>"],
        itemId: "failing-sent-draft-id",
      }),
    });
    assert.strictEqual(failedSentAttach.statusCode, 502);
    assert.strictEqual(failedSentAttach.body.ok, false);
    assert.strictEqual(failedSentAttach.body.status, "failed");
    assert.match(failedSentAttach.body.error, /Temporary sent action failure/);
    assert.strictEqual(failedSentAttach.body.debug.phase, "email-send-auto-attach");
    assert.strictEqual(failedSentAttach.body.ticketNumber, "T1001");
    assert.strictEqual(attachActionFetchCount, 9);

    const retrySentAttach = await invoke(app, "POST", "/api/halo/email/send-auto-attach", {
      url: "/api/halo/email/send-auto-attach",
      cookie: attachCookie,
      body: createSendPayload({
        bodyHtml: "<p>Fail sent reply</p>",
        inReplyToMessageIds: ["<message@example.com>"],
        itemId: "failing-sent-draft-id",
      }),
    });
    assert.strictEqual(retrySentAttach.statusCode, 200);
    assert.strictEqual(retrySentAttach.body.ok, true);
    assert.strictEqual(retrySentAttach.body.status, "attached");
    assert.strictEqual(attachActionFetchCount, 10);

    const invalidTicket = await invoke(app, "POST", "/api/halo/tickets/:ticketId/email", {
      url: "/api/halo/tickets/not-a-ticket/email",
      params: { ticketId: "not-a-ticket" },
      cookie: attachCookie,
      body: createEmailPayload(),
    });
    assert.strictEqual(invalidTicket.statusCode, 400);
    assert.match(invalidTicket.body.error, /ticket ID/);

    const missingBody = await invoke(app, "POST", "/api/halo/tickets/:ticketId/email", {
      url: "/api/halo/tickets/1001/email",
      params: { ticketId: "1001" },
      cookie: attachCookie,
      body: createEmailPayload({ bodyHtml: "", bodyText: "" }),
    });
    assert.strictEqual(missingBody.statusCode, 400);
    assert.match(missingBody.body.error, /email body/);

    const oversized = await invoke(app, "POST", "/api/halo/tickets/:ticketId/email", {
      url: "/api/halo/tickets/1001/email",
      params: { ticketId: "1001" },
      cookie: attachCookie,
      body: createEmailPayload({ bodyHtml: "x".repeat(2 * 1024 * 1024) }),
    });
    assert.strictEqual(oversized.statusCode, 413);
    assert.match(oversized.body.error, /too large/);
  } finally {
    global.fetch = originalFetch;
  }

  let attachFailureTokenFetchCount = 0;

  global.fetch = async (requestUrl) => {
    const url = String(requestUrl);

    if (url === "https://customer.halopsa.com/auth/token") {
      attachFailureTokenFetchCount += 1;
      return jsonResponse({
        access_token: "attach-failure-token",
        expires_in: 3600,
        refresh_token: "attach-failure-refresh-token",
      });
    }

    if (url === "https://customer.halopsa.com/api/Actions") {
      return jsonResponse({ message: "No action permission" }, 403, "Forbidden");
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  };

  try {
    const attachFailureCookie = await loginAndGetCookie(app);
    const attachFailure = await invoke(app, "POST", "/api/halo/tickets/:ticketId/email", {
      url: "/api/halo/tickets/1001/email",
      params: { ticketId: "1001" },
      cookie: attachFailureCookie,
      body: createEmailPayload(),
    });

    assert.strictEqual(attachFailure.statusCode, 502);
    assert.strictEqual(attachFailure.body.ok, false);
    assert.match(attachFailure.body.error, /HTTP 403 Forbidden/);
    assert.match(attachFailure.body.error, /No action permission/);
    assert.strictEqual(attachFailure.body.debug.phase, "email-attach");
    assert.strictEqual(attachFailure.body.debug.method, "POST");
    assert.strictEqual(attachFailureTokenFetchCount, 1);
  } finally {
    global.fetch = originalFetch;
  }

  let refreshTokenFetchCount = 0;
  let refreshActionFetchCount = 0;

  global.fetch = async (requestUrl, options = {}) => {
    const url = String(requestUrl);

    if (url === "https://customer.halopsa.com/auth/token") {
      const form = new URLSearchParams(options.body);

      if (form.get("grant_type") === "refresh_token") {
        refreshTokenFetchCount += 1;
        return jsonResponse({
          access_token: "fresh-post-token",
          expires_in: 3600,
          refresh_token: "refresh-token",
        });
      }

      return jsonResponse({
        access_token: "expired-post-token",
        expires_in: 3600,
        refresh_token: "refresh-token",
      });
    }

    if (url === "https://customer.halopsa.com/api/Actions") {
      refreshActionFetchCount += 1;
      if (options.headers.Authorization === "Bearer expired-post-token") {
        return jsonResponse({ message: "Expired token" }, 401, "Unauthorized");
      }

      assert.strictEqual(options.headers.Authorization, "Bearer fresh-post-token");
      return jsonResponse({ id: 9002 }, 201, "Created");
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  };

  try {
    const refreshCookie = await loginAndGetCookie(app);
    const refreshAttach = await invoke(app, "POST", "/api/halo/tickets/:ticketId/email", {
      url: "/api/halo/tickets/1001/email",
      params: { ticketId: "1001" },
      cookie: refreshCookie,
      body: createEmailPayload(),
    });

    assert.strictEqual(refreshAttach.statusCode, 200);
    assert.strictEqual(refreshAttach.body.ok, true);
    assert.strictEqual(refreshAttach.body.actionId, "9002");
    assert.strictEqual(refreshTokenFetchCount, 1);
    assert.strictEqual(refreshActionFetchCount, 2);
  } finally {
    global.fetch = originalFetch;
  }

  let logoutTokenFetchCount = 0;
  global.fetch = async (requestUrl) => {
    const url = String(requestUrl);

    if (url === "https://customer.halopsa.com/auth/token") {
      logoutTokenFetchCount += 1;
      return jsonResponse({
        access_token: "logout-access-token",
        expires_in: 3600,
        refresh_token: "logout-refresh-token",
      });
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  };

  try {
    const logoutCookie = await loginAndGetCookie(app);
    const logout = await invoke(app, "POST", "/api/auth/logout", {
      url: "/api/auth/logout",
      cookie: logoutCookie,
    });
    assert.strictEqual(logout.statusCode, 200);
    assert.strictEqual(logout.body.authenticated, false);
    assert.match(logout.headers["set-cookie"], /Max-Age=0/);

    const postLogoutStatus = await invoke(app, "GET", "/api/auth/status", {
      url: "/api/auth/status",
    });
    assert.strictEqual(postLogoutStatus.statusCode, 200);
    assert.strictEqual(postLogoutStatus.body.authenticated, true);
    assert.strictEqual(postLogoutStatus.body.haloUrl, "https://customer.halopsa.com");
    assert(postLogoutStatus.body.backgroundSessionId);
    assert.strictEqual(logoutTokenFetchCount, 1);
  } finally {
    global.fetch = originalFetch;
  }

  const persistentDbPath = createTempDbPath("persistent-auth");
  const persistentStore = createHaloStore({ dbPath: persistentDbPath });
  const persistentApp = createMockApp();
  registerTestRoutes(persistentApp, persistentStore);

  let persistentTokenFetchCount = 0;
  let persistentActionFetchCount = 0;
  global.fetch = async (requestUrl) => {
    const url = String(requestUrl);

    if (url === "https://customer.halopsa.com/auth/token") {
      persistentTokenFetchCount += 1;
      return jsonResponse({
        access_token: "persistent-access-token",
        expires_in: 3600,
        refresh_token: "persistent-refresh-token",
      });
    }

    if (url === "https://customer.halopsa.com/api/Actions") {
      persistentActionFetchCount += 1;
      return jsonResponse({ id: 9900 + persistentActionFetchCount }, 201, "Created");
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  };

  let persistentCookie;
  try {
    persistentCookie = await loginAndGetCookie(persistentApp);
    const persistentAttach = await invoke(
      persistentApp,
      "POST",
      "/api/halo/tickets/:ticketId/email",
      {
        url: "/api/halo/tickets/1001/email",
        params: { ticketId: "1001" },
        cookie: persistentCookie,
        body: createEmailPayload({ ticketNumber: "T1001" }),
      }
    );

    assert.strictEqual(persistentAttach.statusCode, 200);
    assert.strictEqual(persistentAttach.body.ok, true);
    assert.strictEqual(persistentTokenFetchCount, 1);
    assert.strictEqual(persistentActionFetchCount, 1);
  } finally {
    global.fetch = originalFetch;
    persistentStore.close();
  }

  const restartedStore = createHaloStore({ dbPath: persistentDbPath });
  const restartedApp = createMockApp();
  registerTestRoutes(restartedApp, restartedStore);

  try {
    const restartedStatus = await invoke(restartedApp, "GET", "/api/auth/status", {
      url: "/api/auth/status",
      cookie: persistentCookie,
      headers: { authorization: "" },
    });
    assert.strictEqual(restartedStatus.statusCode, 200);
    assert.strictEqual(restartedStatus.body.authenticated, true);
    assert.strictEqual(restartedStatus.body.haloUrl, "https://customer.halopsa.com");
    assert(restartedStatus.body.backgroundSessionId);

    const restartedAlreadyAttached = await invoke(
      restartedApp,
      "POST",
      "/api/halo/email/auto-attach",
      {
        url: "/api/halo/email/auto-attach",
        cookie: persistentCookie,
        headers: { authorization: "" },
        body: createEmailPayload(),
      }
    );
    assert.strictEqual(restartedAlreadyAttached.statusCode, 200);
    assert.strictEqual(restartedAlreadyAttached.body.ok, true);
    assert.strictEqual(restartedAlreadyAttached.body.status, "already-attached");
    assert.strictEqual(restartedAlreadyAttached.body.ticketNumber, "T1001");
  } finally {
    restartedStore.close();
  }
}

run()
  .then(() => {
    console.log("Halo auth smoke tests passed");
  })
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
