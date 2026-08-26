const assert = require("assert");
process.env.NODE_ENV = "test";
process.env.HALO_TOKEN_ENCRYPTION_KEY = Buffer.alloc(32, 3).toString("base64url");

const crypto = require("crypto");
const { registerHaloAuthRoutes, _test: haloAuthTest } = require("./haloAuth");
const { createTestDatabase, createTestStore } = require("./testDatabase");
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
    DELETE: new Map(),
    GET: new Map(),
    POST: new Map(),
  };

  return {
    locals: {},
    routes,
    get(path, handler) {
      routes.GET.set(path, handler);
    },
    delete(path, handler) {
      routes.DELETE.set(path, handler);
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

async function registerTestRoutes(app, suppliedStore, routeOptions = {}) {
  const store = suppliedStore || (await createTestStore());
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
    logger: routeOptions.logger,
    store,
  });
  return store;
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
  const sendDiagnosticLines = [];
  const sendDiagnosticLogger = haloAuthTest.createSendDiagnosticLogger(
    { info: (line) => sendDiagnosticLines.push(line) },
    { NODE_ENV: "test" }
  );
  sendDiagnosticLogger("client", {
    attachmentCount: 2,
    body: "private body",
    elapsedMs: 15.2,
    outcome: "ok",
    recipient: "private@example.com",
    stage: "assets-complete",
    ticketId: "9514",
  });
  assert.strictEqual(sendDiagnosticLines.length, 1);
  assert.match(sendDiagnosticLines[0], /\[halo-send\]/);
  assert.match(sendDiagnosticLines[0], /assets-complete/);
  assert.match(sendDiagnosticLines[0], /"attachmentCount":2/);
  assert.doesNotMatch(sendDiagnosticLines[0], /private body|private@example\.com|9514/);
  sendDiagnosticLogger("client", {
    attachmentCount: 1,
    attachmentError: "invalid-attachment-id",
    attachmentId: "private-outlook-id",
    attemptCount: 2,
    failedCount: 0,
    outcome: "ok",
    stage: "attachment-read-complete",
    uploadedCount: 1,
  });
  assert.strictEqual(sendDiagnosticLines.length, 2);
  assert.match(sendDiagnosticLines[1], /attachment-read-complete/);
  assert.match(sendDiagnosticLines[1], /invalid-attachment-id/);
  assert.match(sendDiagnosticLines[1], /"attemptCount":2/);
  assert.doesNotMatch(sendDiagnosticLines[1], /private-outlook-id/);

  const attachmentDiagnosticLines = [];
  const attachmentDiagnosticLogger = haloAuthTest.createEmailAttachmentDiagnosticLogger(
    (line) => attachmentDiagnosticLines.push(line),
    { NODE_ENV: "test" }
  );
  attachmentDiagnosticLogger("commit-validation-failed", {
    attachmentKey: "private-attachment-key",
    count: 1,
    durationMs: 12,
    filename: "private-document.docx",
    outcome: "inventory-mismatch",
    stage: "commit-validation",
    ticketId: "9514",
  });
  assert.strictEqual(attachmentDiagnosticLines.length, 1);
  assert.match(attachmentDiagnosticLines[0], /\[email-attachments\]/);
  assert.match(attachmentDiagnosticLines[0], /commit-validation-failed/);
  assert.match(attachmentDiagnosticLines[0], /inventory-mismatch/);
  assert.match(attachmentDiagnosticLines[0], /"count":1/);
  assert.doesNotMatch(
    attachmentDiagnosticLines[0],
    /private-attachment-key|private-document\.docx|9514/
  );

  assert.strictEqual(getApiAudience("test-client-id"), "api://test-client-id");
  assert.deepStrictEqual(getTokenAudiences("test-client-id", "api://test-client-id"), [
    "test-client-id",
    "api://test-client-id",
  ]);
  assert.deepStrictEqual(getTokenAudiences("spa-client-id", "api://custom-api", "api-client-id"), [
    "api-client-id",
    "api://custom-api",
  ]);
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

  const schemaStore = await createTestStore();
  const schemaUser = await schemaStore.upsertUser({
    objectId: "schema-object-id",
    tenantId: "schema-tenant-id",
  });
  assert(schemaUser.id);
  await schemaStore.close();

  const invalidAuthApp = createMockApp();
  const invalidAuthStore = await registerTestRoutes(invalidAuthApp);
  const invalidAuthStatus = await invoke(invalidAuthApp, "GET", "/api/auth/status", {
    url: "/api/auth/status",
    headers: { authorization: "Bearer invalid-token" },
  });
  assert.strictEqual(invalidAuthStatus.statusCode, 401);
  assert.strictEqual(invalidAuthStatus.body.authenticated, false);
  assert.match(invalidAuthStatus.body.error, /Microsoft add-in authentication failed/);
  await invalidAuthStore.close();

  const routeDiagnosticLines = [];
  const app = createMockApp();
  const store = await registerTestRoutes(app, undefined, {
    logger: (line) => routeDiagnosticLines.push(line),
  });

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
  let subjectSearchFetchCount = 0;
  let closedTicketListFetchCount = 0;

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
          {
            id: 1004,
            ticketnumber: "T1004",
            summary: "Incomplete configuration",
            status: "Incomplete",
            client: { name: "Nested Customer" },
            assigned_agent: { name: "Nested Agent" },
          },
        ],
      });
    }

    if (
      url ===
      "https://customer.halopsa.com/api/Tickets?count=50&search=2200&includeagent=true&includestatus=true"
    ) {
      ticketSearchFetchCount += 1;
      return jsonResponse({
        tickets: [
          {
            id: 7777,
            ticketnumber: "0007777",
            summary: "Reference to 2200 in the summary",
            status: "Open",
          },
          {
            id: 2200,
            ticketnumber: "0002200",
            summary: "Closed ticket assigned to another agent",
            status: "Closed",
            agent_id: 999,
            client_name: "Another Customer",
            agent_name: "Another Agent",
          },
        ],
      });
    }

    if (
      url ===
      "https://customer.halopsa.com/api/Tickets?count=50&open_only=true&search=printer+%5Bwarehouse%5D&includeagent=true&includestatus=true"
    ) {
      subjectSearchFetchCount += 1;
      return jsonResponse({
        tickets: [
          {
            id: 3100,
            ticketnumber: "T3100",
            summary: "Warehouse printer will not connect",
            status: "In Progress",
            client_name: "Warehouse Ltd",
            agent_name: "Taylor",
          },
        ],
      });
    }

    if (
      url ===
      "https://customer.halopsa.com/api/Tickets?count=50&includeagent=true&includestatus=true"
    ) {
      closedTicketListFetchCount += 1;
      return jsonResponse({
        tickets: [
          {
            id: 4100,
            ticketnumber: "T4100",
            summary: "Completed onboarding",
            status: "In Progress",
            closed: "1",
            client_name: "Example Co",
            agent_name: "Morgan",
          },
          {
            id: 4101,
            ticketnumber: "T4101",
            summary: "Active onboarding",
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
        lifecycle: "open",
        client: "Digital Origin",
        agent: "Sebastian",
      },
      {
        id: "1003",
        ticketNumber: "T1003",
        summary: "Second open mine ticket",
        status: "Open",
        lifecycle: "open",
        client: "",
        agent: "",
      },
      {
        id: "1004",
        ticketNumber: "T1004",
        summary: "Incomplete configuration",
        status: "Incomplete",
        lifecycle: "open",
        client: "Nested Customer",
        agent: "Nested Agent",
      },
    ]);

    const ticketSearch = await invoke(app, "GET", "/api/halo/tickets/search", {
      url: "/api/halo/tickets/search?ticketNumber=%5BID%3A%202200%5D&ownership=all&lifecycle=all",
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
        lifecycle: "closed",
        client: "Another Customer",
        agent: "Another Agent",
      },
      {
        id: "7777",
        ticketNumber: "0007777",
        summary: "Reference to 2200 in the summary",
        status: "Open",
        lifecycle: "open",
        client: "",
        agent: "",
      },
    ]);

    const subjectSearch = await invoke(app, "GET", "/api/halo/tickets/search", {
      url: "/api/halo/tickets/search?query=printer%20%5Bwarehouse%5D&ownership=all",
      cookie: ticketsComplete.headers["set-cookie"],
    });

    assert.strictEqual(subjectSearch.statusCode, 200);
    assert.strictEqual(subjectSearchFetchCount, 1);
    assert.deepStrictEqual(subjectSearch.body.tickets, [
      {
        id: "3100",
        ticketNumber: "T3100",
        summary: "Warehouse printer will not connect",
        status: "In Progress",
        lifecycle: "open",
        client: "Warehouse Ltd",
        agent: "Taylor",
      },
    ]);

    const closedTickets = await invoke(app, "GET", "/api/halo/tickets", {
      url: "/api/halo/tickets?ownership=all&lifecycle=closed",
      cookie: ticketsComplete.headers["set-cookie"],
    });

    assert.strictEqual(closedTickets.statusCode, 200);
    assert.strictEqual(closedTicketListFetchCount, 1);
    assert.deepStrictEqual(closedTickets.body.tickets, [
      {
        id: "4100",
        ticketNumber: "T4100",
        summary: "Completed onboarding",
        status: "In Progress",
        lifecycle: "closed",
        client: "Example Co",
        agent: "Morgan",
      },
    ]);

    const emptyTicketSearch = await invoke(app, "GET", "/api/halo/tickets/search", {
      url: "/api/halo/tickets/search?ticketNumber=",
      cookie: ticketsComplete.headers["set-cookie"],
    });

    assert.strictEqual(emptyTicketSearch.statusCode, 400);
    assert.strictEqual(emptyTicketSearch.body.ok, false);
    assert.match(emptyTicketSearch.body.error, /search query/i);

    const invalidFilter = await invoke(app, "GET", "/api/halo/tickets", {
      url: "/api/halo/tickets?ownership=team&lifecycle=open",
      cookie: ticketsComplete.headers["set-cookie"],
    });
    assert.strictEqual(invalidFilter.statusCode, 400);
    assert.match(invalidFilter.body.error, /ownership must be one of/i);

    const invalidLifecycle = await invoke(app, "GET", "/api/halo/tickets", {
      url: "/api/halo/tickets?ownership=mine&lifecycle=archived",
      cookie: ticketsComplete.headers["set-cookie"],
    });
    assert.strictEqual(invalidLifecycle.statusCode, 400);
    assert.match(invalidLifecycle.body.error, /lifecycle must be one of/i);

    const oversizedTicketSearch = await invoke(app, "GET", "/api/halo/tickets/search", {
      url: `/api/halo/tickets/search?query=${"x".repeat(201)}`,
      cookie: ticketsComplete.headers["set-cookie"],
    });
    assert.strictEqual(oversizedTicketSearch.statusCode, 400);
    assert.match(oversizedTicketSearch.body.error, /200 characters or fewer/i);
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

  const unauthenticatedEmailMatch = await invoke(app, "POST", "/api/halo/email/match", {
    url: "/api/halo/email/match",
    headers: { authorization: "" },
    body: createEmailPayload({ bodyHtml: "", bodyText: "" }),
  });
  assert.strictEqual(unauthenticatedEmailMatch.statusCode, 401);
  assert.strictEqual(unauthenticatedEmailMatch.body.ok, false);

  const noSessionSendAutoAttach = await invoke(app, "POST", "/api/halo/email/send-auto-attach", {
    url: "/api/halo/email/send-auto-attach",
    headers: { authorization: "" },
    body: createSendPayload(),
  });
  assert.strictEqual(noSessionSendAutoAttach.statusCode, 200);
  assert.strictEqual(noSessionSendAutoAttach.body.ok, true);
  assert.strictEqual(noSessionSendAutoAttach.body.status, "no-session");

  const noSessionExplicitAttach = await invoke(
    app,
    "POST",
    "/api/halo/tickets/:ticketId/sent-email",
    {
      url: "/api/halo/tickets/1001/sent-email",
      params: { ticketId: "1001" },
      headers: { authorization: "" },
      body: createSendPayload({
        composeAttachId: "no-session-compose",
        conversationId: "",
        inReplyToMessageIds: [],
        ticketNumber: "T1001",
      }),
    }
  );
  assert.strictEqual(noSessionExplicitAttach.statusCode, 200);
  assert.strictEqual(noSessionExplicitAttach.body.ok, true);
  assert.strictEqual(noSessionExplicitAttach.body.status, "no-session");
  assert.strictEqual(noSessionExplicitAttach.body.ticketNumber, "T1001");

  let attachTokenFetchCount = 0;
  let attachActionFetchCount = 0;
  let attachImageFetchCount = 0;
  let attachFileFetchCount = 0;
  let failNextAutoAttach = true;
  let failNextSentAutoAttach = true;
  let failNextExplicitAttach = true;
  let failedStagedAttachmentName = "";
  let failNextStagedAction = false;
  const deletedStagedAttachmentIds = [];
  const attachActions = [];
  const inlineImageBytes = Buffer.concat([
    Buffer.from("89504e470d0a1a0a", "hex"),
    Buffer.from("native Halo image"),
  ]);
  const inlineImageHash = sha256Hex(inlineImageBytes);
  const inlineImageAttachmentId = "825458da-7545-42eb-bb16-a57737b8d821";
  const inlineImageToken = [
    Buffer.from(JSON.stringify({ alg: "HS256", typ: "JWT" })).toString("base64url"),
    Buffer.from(JSON.stringify({ id: inlineImageAttachmentId })).toString("base64url"),
    "test-signature",
  ].join(".");
  const ordinaryAttachmentBytes = Buffer.from("ordinary email attachment");
  const ordinaryAttachmentKey = sha256Hex("ordinary-attachment-key");
  const ordinaryAttachmentFingerprint = sha256Hex("ordinary-attachment-fingerprint");

  global.fetch = async (requestUrl, options = {}) => {
    const url = String(requestUrl);

    if (url === "https://customer.halopsa.com/auth/token") {
      attachTokenFetchCount += 1;
      return jsonResponse({
        access_token:
          attachTokenFetchCount === 1 ? "attach-access-token" : "attach-refreshed-access-token",
        expires_in: 3600,
        refresh_token: "attach-refresh-token",
      });
    }

    if (url === "https://customer.halopsa.com/api/attachment/image") {
      attachImageFetchCount += 1;
      assert.strictEqual(options.method, "POST");
      assert(options.body instanceof FormData);
      assert.strictEqual(options.body.get("ticket_id"), "1001");
      assert.strictEqual(options.body.get("image_upload_id"), "0");
      assert.strictEqual(options.body.get("image_upload_key"), "");
      assert.strictEqual(options.body.get("showforusers"), "true");
      assert.deepStrictEqual(
        Buffer.from(await options.body.get("file").arrayBuffer()),
        inlineImageBytes
      );

      if (attachImageFetchCount === 1) {
        assert.strictEqual(options.headers.Authorization, "Bearer attach-access-token");
        return jsonResponse({ message: "Access token expired" }, 401, "Unauthorized");
      }

      assert.strictEqual(options.headers.Authorization, "Bearer attach-refreshed-access-token");
      return jsonResponse({
        link: `/api/attachment/image?token=${encodeURIComponent(inlineImageToken)}`,
      });
    }

    if (url === "https://customer.halopsa.com/api/Attachment") {
      attachFileFetchCount += 1;
      assert.strictEqual(options.method, "POST");
      assert.match(
        options.headers.Authorization,
        /^Bearer attach-(?:access-token|refreshed-access-token)$/
      );
      const attachments = JSON.parse(options.body);
      assert(Array.isArray(attachments));
      assert.strictEqual(attachments.length, 1);
      const body = attachments[0];
      assert(!Object.prototype.hasOwnProperty.call(body, "ticket_id"));
      assert(!Object.prototype.hasOwnProperty.call(body, "action_id"));
      assert(!Object.prototype.hasOwnProperty.call(body, "type"));
      assert(!Object.prototype.hasOwnProperty.call(body, "desc"));
      assert(
        [
          "customer-report.pdf",
          "post-action-report.pdf",
          "partial-a.pdf",
          "partial-b.pdf",
          "action-failure.pdf",
          "automatic-compose.pdf",
        ].includes(body.filename)
      );
      assert.strictEqual(body.filesize, ordinaryAttachmentBytes.length);
      assert.strictEqual(body.showforusers, null);
      assert.strictEqual(body._uploading, true);
      assert.match(body._tempid, /^__[0-9a-f-]{36}$/);
      assert.strictEqual(body.showonchild, false);
      assert.strictEqual(body.showonrelated, false);
      assert.strictEqual(
        body.data_base64,
        `data:application/pdf;base64,${ordinaryAttachmentBytes.toString("base64")}`
      );
      if (body.filename === failedStagedAttachmentName) {
        return jsonResponse({ message: "Temporary attachment upload failure" }, 503, "Unavailable");
      }
      return jsonResponse([
        {
          id: 7000 + attachFileFetchCount,
          filename: "internal-storage-name.pdf",
          filesize: body.filesize,
          desc: body.filename,
          showforusers: false,
          type: 0,
        },
      ]);
    }

    if (/^https:\/\/customer\.halopsa\.com\/api\/Attachment\/\d+$/.test(url)) {
      assert.strictEqual(options.method, "DELETE");
      deletedStagedAttachmentIds.push(Number(url.split("/").pop()));
      return jsonResponse({}, 204, "No Content");
    }

    if (url === "https://customer.halopsa.com/api/Actions") {
      attachActionFetchCount += 1;
      assert.strictEqual(options.method, "POST");
      const actions = JSON.parse(options.body);
      assert(Array.isArray(actions));
      assert.strictEqual(actions.length, 1);
      attachActions.push(actions[0]);
      assert([1001, 2002].includes(actions[0].ticket_id));
      assert.strictEqual(actions[0].outcome, "Email");
      assert.strictEqual(actions[0].sendemail, false);
      assert(Number.isFinite(Date.parse(actions[0].datetime)));
      assert.match(actions[0].note, /Outlook email attached to ticket/);
      assert.match(actions[0].note_html, /Sender User &lt;sender@example.com&gt;/);
      if (attachActionFetchCount === 1) {
        assert.strictEqual(actions[0].attachments.length, 1);
        assert.deepStrictEqual(actions[0].attachments[0], {
          filename: "customer-report.pdf",
          filesize: ordinaryAttachmentBytes.length,
          showforusers: null,
          _uploading: false,
          _tempid: null,
          showonchild: false,
          showonrelated: false,
          data_base64: null,
          id: 7001,
          data: null,
        });
      }

      if (actions[0].email_message_id === "<failing-reply@example.com>" && failNextAutoAttach) {
        failNextAutoAttach = false;
        return jsonResponse({ message: "Temporary action failure" }, 403, "Forbidden");
      }

      if (/Fail sent reply/.test(actions[0].note_html) && failNextSentAutoAttach) {
        failNextSentAutoAttach = false;
        return jsonResponse({ message: "Temporary sent action failure" }, 403, "Forbidden");
      }

      if (/Fail composed email/.test(actions[0].note_html) && failNextExplicitAttach) {
        failNextExplicitAttach = false;
        return jsonResponse({ message: "Temporary composed action failure" }, 403, "Forbidden");
      }

      if (/Staged action failure/.test(actions[0].note_html) && failNextStagedAction) {
        failNextStagedAction = false;
        return jsonResponse({ message: "Temporary staged action failure" }, 503, "Unavailable");
      }

      if (/Concurrent composed email/.test(actions[0].note_html)) {
        await new Promise((resolve) => setTimeout(resolve, 25));
      }

      return jsonResponse({ id: 9000 + attachActionFetchCount }, 201, "Created");
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  };

  try {
    const attachCookie = await loginAndGetCookie(app);
    const stageOrdinaryAttachments = async ({ draftItemId, filenames, operationId }) => {
      const fingerprint = sha256Hex(`${operationId}-fingerprint`);
      const descriptors = filenames.map((name, index) => ({
        attachmentKey: sha256Hex(`${operationId}-${index}-${name}`),
        attachmentType: "file",
        contentSha256: sha256Hex(ordinaryAttachmentBytes),
        contentType: "application/pdf",
        name,
        reportedSize: ordinaryAttachmentBytes.length,
      }));
      const start = await invoke(app, "POST", "/api/halo/email-attachments/prefetch/start", {
        url: "/api/halo/email-attachments/prefetch/start",
        cookie: attachCookie,
        body: {
          draftItemId,
          emailAttachmentFingerprint: fingerprint,
          emailAttachments: descriptors,
          operationId,
          ticketId: 1001,
          ticketNumber: "T1001",
        },
      });
      assert.strictEqual(start.statusCode, 200);
      for (const descriptor of descriptors) {
        const staged = await invoke(
          app,
          "POST",
          "/api/halo/email-attachments/prefetch/:prefetchKey/items",
          {
            url: `/api/halo/email-attachments/prefetch/${start.body.prefetchKey}/items`,
            params: { prefetchKey: start.body.prefetchKey },
            cookie: attachCookie,
            body: {
              attachmentKey: descriptor.attachmentKey,
              contentBase64: ordinaryAttachmentBytes.toString("base64"),
              contentFormat: "base64",
              contentSha256: sha256Hex(ordinaryAttachmentBytes),
            },
          }
        );
        assert.strictEqual(staged.statusCode, 200);
        assert(["prepared", "already-prepared"].includes(staged.body.status));
      }
      return {
        descriptors,
        draftItemId,
        fingerprint,
        operationId,
        prefetchKey: start.body.prefetchKey,
      };
    };
    const stagedSendPayload = (stage, overrides = {}) =>
      createSendPayload({
        bodyHtml: "<p>Staged attachment test</p>",
        composeAttachId: `${stage.operationId}-compose`,
        emailAttachmentDraftItemId: stage.draftItemId,
        emailAttachmentFingerprint: stage.fingerprint,
        emailAttachmentOperationId: stage.operationId,
        emailAttachmentPrefetchKey: stage.prefetchKey,
        emailAttachmentStagingVersion: 2,
        emailAttachmentSummary: {
          attached: 0,
          detected: stage.descriptors.length,
          failed: 0,
          selected: stage.descriptors.length,
          skipped: 0,
          prepared: stage.descriptors.length,
          warnings: [],
        },
        includeEmailAttachments: true,
        inReplyToMessageIds: [],
        itemId: stage.draftItemId,
        ticketNumber: "T1001",
        ...overrides,
      });
    const attachmentPrefetch = await invoke(
      app,
      "POST",
      "/api/halo/email-attachments/prefetch/start",
      {
        url: "/api/halo/email-attachments/prefetch/start",
        cookie: attachCookie,
        body: {
          draftItemId: "attachment-prefetch-item-id",
          emailAttachmentFingerprint: ordinaryAttachmentFingerprint,
          emailAttachments: [
            {
              attachmentKey: ordinaryAttachmentKey,
              attachmentType: "file",
              contentSha256: sha256Hex(ordinaryAttachmentBytes),
              contentType: "application/pdf",
              name: "customer-report.pdf",
              reportedSize: ordinaryAttachmentBytes.length,
            },
          ],
          operationId: "manual-attachment-operation",
          ticketId: 1001,
          ticketNumber: "T1001",
        },
      }
    );
    assert.strictEqual(attachmentPrefetch.statusCode, 200);
    assert.strictEqual(attachmentPrefetch.body.status, "pending");
    assert.strictEqual(attachmentPrefetch.body.stagingVersion, 2);
    const attachmentUpload = await invoke(
      app,
      "POST",
      "/api/halo/email-attachments/prefetch/:prefetchKey/items",
      {
        url: `/api/halo/email-attachments/prefetch/${attachmentPrefetch.body.prefetchKey}/items`,
        params: { prefetchKey: attachmentPrefetch.body.prefetchKey },
        cookie: attachCookie,
        body: {
          attachmentKey: ordinaryAttachmentKey,
          contentBase64: ordinaryAttachmentBytes.toString("base64"),
          contentFormat: "base64",
          contentSha256: sha256Hex(ordinaryAttachmentBytes),
        },
      }
    );
    assert.strictEqual(attachmentUpload.statusCode, 200);
    assert.strictEqual(attachmentUpload.body.status, "prepared");
    assert.strictEqual(attachFileFetchCount, 0);
    const attachmentStatus = await invoke(
      app,
      "GET",
      "/api/halo/email-attachments/prefetch/:prefetchKey/status",
      {
        url: `/api/halo/email-attachments/prefetch/${attachmentPrefetch.body.prefetchKey}/status`,
        params: { prefetchKey: attachmentPrefetch.body.prefetchKey },
        cookie: attachCookie,
      }
    );
    assert.strictEqual(attachmentStatus.statusCode, 200);
    assert.deepStrictEqual(attachmentStatus.body.aggregate, {
      failed: 0,
      pending: 0,
      prepared: 1,
      selected: 1,
    });
    const invalidPrefetchOperation = await invoke(app, "POST", "/api/halo/inline-images/prefetch", {
      url: "/api/halo/inline-images/prefetch",
      cookie: attachCookie,
      body: {
        composeOperationId: "bad id",
        inlineImageFingerprint: "a".repeat(64),
        ticketId: 1001,
      },
    });
    assert.strictEqual(invalidPrefetchOperation.statusCode, 400);

    const invalidPrefetchFingerprint = await invoke(
      app,
      "POST",
      "/api/halo/inline-images/prefetch",
      {
        url: "/api/halo/inline-images/prefetch",
        cookie: attachCookie,
        body: {
          composeOperationId: "compose-operation-1",
          inlineImageFingerprint: "not-a-sha256",
          ticketId: 1001,
        },
      }
    );
    assert.strictEqual(invalidPrefetchFingerprint.statusCode, 400);

    const attach = await invoke(app, "POST", "/api/halo/tickets/:ticketId/email", {
      url: "/api/halo/tickets/1001/email",
      params: { ticketId: "1001" },
      cookie: attachCookie,
      body: createEmailPayload({
        bodyHtml:
          '<p>Hello from Outlook</p><img alt="Signature" src="cid:signature-logo">' +
          "<blockquote><p>Prior thread content should stay for first attach</p></blockquote>",
        inReplyToMessageIds: ["<prior-reply@example.com>"],
        referenceMessageIds: ["<original-message@example.com>", "<prior-reply@example.com>"],
        ticketNumber: "T1001",
        inlineImageRefs: [{ contentId: "signature-logo", sha256: inlineImageHash }],
        inlineImageUploads: [
          {
            sha256: inlineImageHash,
            name: "signature.png",
            contentType: "image/png",
            contentBase64: inlineImageBytes.toString("base64"),
          },
        ],
        inlineImageTimings: { hashingMs: 23.7, outlookReadMs: 145.2 },
        includeEmailAttachments: true,
        itemId: "attachment-prefetch-item-id",
        emailAttachmentDraftItemId: "attachment-prefetch-item-id",
        emailAttachmentFingerprint: ordinaryAttachmentFingerprint,
        emailAttachmentOperationId: "manual-attachment-operation",
        emailAttachmentPrefetchKey: attachmentPrefetch.body.prefetchKey,
        emailAttachmentStagingVersion: 2,
        emailAttachmentSummary: {
          attached: 0,
          detected: 1,
          failed: 0,
          selected: 1,
          skipped: 0,
          prepared: 1,
          warnings: [],
        },
      }),
    });

    assert.strictEqual(attach.statusCode, 200, attach.body.error);
    assert.strictEqual(attach.body.ok, true);
    assert.strictEqual(attach.body.attachMode, "full-chain");
    assert.strictEqual(attach.body.message, "Full email chain attached to Halo ticket");
    assert.strictEqual(attach.body.actionId, "9001");
    assert(attach.body.backgroundSessionId);
    assert.strictEqual(attachTokenFetchCount, 2);
    assert.strictEqual(attachImageFetchCount, 2);
    assert.strictEqual(attachActionFetchCount, 1);
    assert.strictEqual(attachFileFetchCount, 1);
    assert.strictEqual(attach.body.inlineImages.uploaded, 1);
    assert.strictEqual(attach.body.inlineImages.failed, 0);
    assert.strictEqual(attach.body.emailAttachments.attached, 1);
    assert.strictEqual(attach.body.emailAttachments.failed, 0);
    assert.strictEqual(attach.body.inlineImageTimings.hashingMs, 24);
    assert.strictEqual(attach.body.inlineImageTimings.outlookReadMs, 145);
    assert(Number.isInteger(attach.body.inlineImageTimings.actionCreationMs));
    assert.strictEqual(attachActions[0].emailsubject, "RE: Example subject");
    assert.strictEqual(attachActions[0].email_message_id, "<message@example.com>");
    assert.strictEqual(attachActions[0].actioninternetmessageid, "<message@example.com>");
    assert.notStrictEqual(attachActions[0].datetime, "2026-07-07T10:00:00.000Z");
    assert.match(attachActions[0].note_html, /<p>Hello from Outlook<\/p>/);
    assert.match(
      attachActions[0].note_html,
      new RegExp(
        `https://customer\\.halopsa\\.com/api/attachment/image\\?token=${inlineImageToken.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`
      )
    );
    assert.doesNotMatch(attachActions[0].note_html, /cid:signature-logo/i);
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

    const alreadyAttachedMatch = await invoke(app, "POST", "/api/halo/email/match", {
      url: "/api/halo/email/match",
      cookie: attachCookie,
      body: createEmailPayload({ bodyHtml: "", bodyText: "" }),
    });
    assert.strictEqual(alreadyAttachedMatch.statusCode, 200);
    assert.strictEqual(alreadyAttachedMatch.body.status, "already-attached");
    assert.strictEqual(alreadyAttachedMatch.body.ticketId, 1001);
    assert.strictEqual(attachActionFetchCount, 1);

    const replyMatch = await invoke(app, "POST", "/api/halo/email/match", {
      url: "/api/halo/email/match",
      cookie: attachCookie,
      body: createEmailPayload({
        bodyHtml: "",
        bodyText: "",
        inReplyToMessageIds: ["message@example.com"],
        internetMessageId: "<match-only-reply@example.com>",
      }),
    });
    assert.strictEqual(replyMatch.statusCode, 200);
    assert.strictEqual(replyMatch.body.status, "matched");
    assert.strictEqual(replyMatch.body.ticketNumber, "T1001");
    assert.strictEqual(attachActionFetchCount, 1);

    const automaticComposePrefetch = await invoke(
      app,
      "POST",
      "/api/halo/email-attachments/prefetch/start",
      {
        url: "/api/halo/email-attachments/prefetch/start",
        cookie: attachCookie,
        body: {
          conversationId: "compose-client-conversation-id",
          draftItemId: "compose-draft-item-id",
          emailAttachmentFingerprint: sha256Hex("automatic-compose-fingerprint"),
          emailAttachments: [
            {
              attachmentKey: sha256Hex("automatic-compose-attachment"),
              attachmentType: "file",
              contentSha256: sha256Hex(ordinaryAttachmentBytes),
              contentType: "application/pdf",
              name: "automatic-compose.pdf",
              reportedSize: ordinaryAttachmentBytes.length,
            },
          ],
          inReplyToMessageIds: ["attachment-prefetch-item-id"],
          internetMessageId: "",
          mailboxEmail: "support@example.com",
          operationId: "automatic-compose-operation",
          referenceMessageIds: [],
        },
      }
    );
    assert.strictEqual(automaticComposePrefetch.statusCode, 200);
    assert.strictEqual(automaticComposePrefetch.body.status, "pending");
    assert.strictEqual(automaticComposePrefetch.body.ticketId, "1001");
    assert.strictEqual(automaticComposePrefetch.body.ticketNumber, "T1001");
    const automaticComposeStage = await invoke(
      app,
      "POST",
      "/api/halo/email-attachments/prefetch/:prefetchKey/items",
      {
        url: `/api/halo/email-attachments/prefetch/${automaticComposePrefetch.body.prefetchKey}/items`,
        params: { prefetchKey: automaticComposePrefetch.body.prefetchKey },
        cookie: attachCookie,
        body: {
          attachmentKey: sha256Hex("automatic-compose-attachment"),
          contentBase64: ordinaryAttachmentBytes.toString("base64"),
          contentFormat: "base64",
          contentSha256: sha256Hex(ordinaryAttachmentBytes),
        },
      }
    );
    assert.strictEqual(automaticComposeStage.statusCode, 200);
    assert.strictEqual(automaticComposeStage.body.status, "prepared");

    const noMatchLookup = await invoke(app, "POST", "/api/halo/email/match", {
      url: "/api/halo/email/match",
      cookie: attachCookie,
      body: createEmailPayload({
        bodyHtml: "",
        bodyText: "",
        conversationId: "lookup-only-unrelated-conversation",
        internetMessageId: "<lookup-only-unrelated@example.com>",
      }),
    });
    assert.strictEqual(noMatchLookup.statusCode, 200);
    assert.strictEqual(noMatchLookup.body.status, "no-match");
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
        includeEmailAttachments: true,
        emailAttachmentDraftItemId: "conversation-reply-item-id",
        emailAttachmentFingerprint: "f".repeat(64),
        emailAttachmentOperationId: "missing-prefetch-operation",
        emailAttachmentPrefetchKey: "missing-attachment-prefetch",
        emailAttachmentStagingVersion: 2,
        emailAttachmentSummary: {
          attached: 0,
          detected: 1,
          failed: 0,
          selected: 1,
          skipped: 0,
          prepared: 0,
          warnings: [],
        },
      }),
    });
    assert.strictEqual(conversationAttach.statusCode, 409);
    assert.strictEqual(conversationAttach.body.ok, false);
    assert.strictEqual(conversationAttach.body.status, "attachments-not-ready");
    assert.strictEqual(attachActionFetchCount, 3);

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
    assert.strictEqual(attachActionFetchCount, 3);

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
    assert.strictEqual(attachActionFetchCount, 3);

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
    assert.strictEqual(attachActionFetchCount, 4);

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
    assert.strictEqual(attachActionFetchCount, 5);

    const mappedAttachmentDiagnosticStart = routeDiagnosticLines.length;
    const mappedAttachmentNotReady = await invoke(app, "POST", "/api/halo/email/send-auto-attach", {
      url: "/api/halo/email/send-auto-attach",
      cookie: attachCookie,
      body: createSendPayload({
        bodyHtml: "<p>Private mapped attachment diagnostic body</p>",
        emailAttachmentDraftItemId: "mapped-attachment-draft-id",
        emailAttachmentFingerprint: "a".repeat(64),
        emailAttachmentOperationId: "mapped-attachment-operation",
        emailAttachmentPrefetchKey: "",
        emailAttachmentStagingVersion: 2,
        emailAttachmentSummary: {
          attached: 0,
          detected: 1,
          failed: 0,
          prepared: 0,
          selected: 1,
          skipped: 0,
        },
        includeEmailAttachments: true,
        inReplyToMessageIds: ["<message@example.com>"],
        itemId: "mapped-attachment-draft-id",
        subject: "Private mapped attachment diagnostic subject",
      }),
    });
    assert.strictEqual(mappedAttachmentNotReady.statusCode, 409);
    assert.strictEqual(mappedAttachmentNotReady.body.ok, false);
    assert.strictEqual(mappedAttachmentNotReady.body.status, "attachments-not-ready");
    assert.strictEqual(attachActionFetchCount, 5);
    const mappedAttachmentDiagnostics = routeDiagnosticLines.slice(mappedAttachmentDiagnosticStart);
    assert(
      mappedAttachmentDiagnostics.some(
        (line) =>
          line.includes("[halo-send]") &&
          line.includes('"stage":"mapping-lookup"') &&
          line.includes('"outcome":"matched"')
      )
    );
    assert(
      mappedAttachmentDiagnostics.some(
        (line) =>
          line.includes("[email-attachments]") &&
          line.includes("commit-validation-failed") &&
          line.includes('"outcome":"stage-missing"')
      )
    );
    assert(
      mappedAttachmentDiagnostics.some(
        (line) =>
          line.includes("[halo-send]") &&
          line.includes('"stage":"response-complete"') &&
          line.includes('"outcome":"stage-missing"')
      )
    );
    assert.doesNotMatch(
      JSON.stringify(mappedAttachmentDiagnostics),
      /Private mapped attachment|mapped-attachment|message@example\.com/
    );

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
    assert.strictEqual(attachActionFetchCount, 6);
    assert.match(attachActions[5].email_message_id, /^<halo-outlook-[a-f0-9]{32}@local>$/);
    assert.match(attachActions[5].note_html, /Sent reply from Outlook/);
    assert.doesNotMatch(attachActions[5].note_html, /Old quoted content/);

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
    assert.strictEqual(attachActionFetchCount, 6);

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
    assert.strictEqual(attachActionFetchCount, 7);

    const expiredBackgroundSessionId = "expired-background-session-id";
    await store.createBackgroundSession({
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
    assert.strictEqual(attachActionFetchCount, 7);

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
    assert.strictEqual(attachActionFetchCount, 7);

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
    assert.strictEqual(attachActionFetchCount, 8);

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
    assert.strictEqual(attachActionFetchCount, 9);

    const explicitComposeAttach = await invoke(
      app,
      "POST",
      "/api/halo/tickets/:ticketId/sent-email",
      {
        url: "/api/halo/tickets/1001/sent-email",
        params: { ticketId: "1001" },
        cookie: attachCookie,
        body: createSendPayload({
          bcc: [{ displayName: "Hidden User", emailAddress: "hidden@example.com" }],
          bodyHtml: "<p>Brand new composed email</p><blockquote>Content to retain</blockquote>",
          composeAttachId: "compose-operation-new-mail",
          conversationId: "",
          inReplyToMessageIds: [],
          internetMessageId: "<untrusted-draft-message-id@example.com>",
          itemId: "new-compose-draft-id",
          subject: "Brand new composed email",
          ticketNumber: "T1001",
        }),
      }
    );
    assert.strictEqual(explicitComposeAttach.statusCode, 200);
    assert.strictEqual(explicitComposeAttach.body.ok, true);
    assert.strictEqual(explicitComposeAttach.body.status, "attached");
    assert.strictEqual(explicitComposeAttach.body.ticketNumber, "T1001");
    assert.strictEqual(attachActionFetchCount, 10);
    assert.match(attachActions[9].email_message_id, /^<halo-outlook-compose-[a-f0-9]{32}@local>$/);
    assert.notStrictEqual(
      attachActions[9].email_message_id,
      "<untrusted-draft-message-id@example.com>"
    );
    assert.strictEqual(attachActions[9].emailsubject, "Brand new composed email");
    assert.strictEqual(attachActions[9].emailtolistall, "Support User <support@example.com>");
    assert.match(attachActions[9].note_html, /Sender User &lt;sender@example\.com&gt;/);
    assert.match(attachActions[9].note_html, /Copied User &lt;cc@example\.com&gt;/);
    assert.match(attachActions[9].note_html, /Content to retain/);
    assert.doesNotMatch(attachActions[9].note_html, /hidden@example\.com/);
    assert.strictEqual(attachActions[9].bcc, undefined);

    const recoveredComposeMapping = await invoke(app, "POST", "/api/halo/email/recover-mapping", {
      url: "/api/halo/email/recover-mapping",
      cookie: attachCookie,
      body: createSendPayload({
        composeAttachIds: ["invalid-candidate", "compose-operation-new-mail"],
        conversationId: "unrelated-recovery-conversation",
        inReplyToMessageIds: [],
      }),
    });
    assert.strictEqual(recoveredComposeMapping.statusCode, 200);
    assert.strictEqual(recoveredComposeMapping.body.status, "matched");
    assert.strictEqual(recoveredComposeMapping.body.candidateIndex, 1);
    assert.strictEqual(recoveredComposeMapping.body.ticketNumber, "T1001");

    const recoveredReadModeMatch = await invoke(app, "POST", "/api/halo/email/match", {
      url: "/api/halo/email/match",
      cookie: attachCookie,
      body: createEmailPayload({
        composeAttachIds: ["compose-operation-new-mail"],
        conversationId: "recovered-read-conversation",
        inReplyToMessageIds: ["<transported-original@example.com>"],
        internetMessageId: "<unattached-incoming-reply@example.com>",
        itemId: "unattached-incoming-reply-item",
        referenceMessageIds: ["<transported-original@example.com>"],
      }),
    });
    assert.strictEqual(recoveredReadModeMatch.statusCode, 200);
    assert.strictEqual(recoveredReadModeMatch.body.status, "matched");
    assert.strictEqual(recoveredReadModeMatch.body.ticketNumber, "T1001");

    const recoveredConversationFollowup = await invoke(app, "POST", "/api/halo/email/match", {
      url: "/api/halo/email/match",
      cookie: attachCookie,
      body: createEmailPayload({
        conversationId: "recovered-read-conversation",
        internetMessageId: "<later-unattached-reply@example.com>",
        itemId: "later-unattached-reply-item",
      }),
    });
    assert.strictEqual(recoveredConversationFollowup.body.status, "matched");

    const recoveredTransportReference = await invoke(app, "POST", "/api/halo/email/match", {
      url: "/api/halo/email/match",
      cookie: attachCookie,
      body: createEmailPayload({
        conversationId: "different-client-conversation",
        internetMessageId: "<another-unattached-reply@example.com>",
        itemId: "another-unattached-reply-item",
        referenceMessageIds: ["<transported-original@example.com>"],
      }),
    });
    assert.strictEqual(recoveredTransportReference.body.status, "matched");

    const recoveredCurrentMessageWasNotMarkedAttached = await invoke(
      app,
      "POST",
      "/api/halo/email/match",
      {
        url: "/api/halo/email/match",
        cookie: attachCookie,
        body: createEmailPayload({
          conversationId: "",
          inReplyToMessageIds: [],
          internetMessageId: "<unattached-incoming-reply@example.com>",
          itemId: "different-unattached-item",
          referenceMessageIds: [],
        }),
      }
    );
    assert.strictEqual(recoveredCurrentMessageWasNotMarkedAttached.body.status, "no-match");

    const forgedComposeMapping = await invoke(app, "POST", "/api/halo/email/recover-mapping", {
      url: "/api/halo/email/recover-mapping",
      cookie: attachCookie,
      body: createSendPayload({
        composeAttachIds: ["forged-compose-id"],
        conversationId: "unrelated-recovery-conversation",
        inReplyToMessageIds: [],
      }),
    });
    assert.strictEqual(forgedComposeMapping.statusCode, 200);
    assert.strictEqual(forgedComposeMapping.body.status, "no-match");

    const mismatchedRecoveryMailbox = await invoke(app, "POST", "/api/halo/email/recover-mapping", {
      url: "/api/halo/email/recover-mapping",
      cookie: attachCookie,
      body: createSendPayload({
        composeAttachIds: ["compose-operation-new-mail"],
        conversationId: "unrelated-recovery-conversation",
        inReplyToMessageIds: [],
        mailboxEmail: "other@example.com",
      }),
    });
    assert.strictEqual(mismatchedRecoveryMailbox.statusCode, 403);
    assert.strictEqual(mismatchedRecoveryMailbox.body.status, "no-match");

    const duplicateExplicitComposeAttach = await invoke(
      app,
      "POST",
      "/api/halo/tickets/:ticketId/sent-email",
      {
        url: "/api/halo/tickets/1001/sent-email",
        params: { ticketId: "1001" },
        cookie: attachCookie,
        body: createSendPayload({
          bodyHtml: "<p>Brand new composed email</p><blockquote>Content to retain</blockquote>",
          composeAttachId: "compose-operation-new-mail",
          conversationId: "",
          inReplyToMessageIds: [],
          itemId: "new-compose-draft-id",
          subject: "Brand new composed email",
          ticketNumber: "T1001",
        }),
      }
    );
    assert.strictEqual(duplicateExplicitComposeAttach.statusCode, 200);
    assert.strictEqual(duplicateExplicitComposeAttach.body.status, "already-attached");
    assert.strictEqual(attachActionFetchCount, 10);

    const explicitOverride = await invoke(app, "POST", "/api/halo/tickets/:ticketId/sent-email", {
      url: "/api/halo/tickets/2002/sent-email",
      params: { ticketId: "2002" },
      cookie: attachCookie,
      body: createSendPayload({
        bodyHtml: "<p>Explicitly remapped reply</p><blockquote>Old content</blockquote>",
        composeAttachId: "compose-operation-override",
        conversationId: "new-compose-conversation-id",
        inReplyToMessageIds: [],
        itemId: "override-compose-draft-id",
        ticketNumber: "T2002",
      }),
    });
    assert.strictEqual(explicitOverride.statusCode, 200);
    assert.strictEqual(explicitOverride.body.status, "attached");
    assert.strictEqual(explicitOverride.body.ticketNumber, "T2002");
    assert.strictEqual(attachActionFetchCount, 11);
    assert.strictEqual(attachActions[10].ticket_id, 2002);
    assert.match(attachActions[10].note_html, /Old content/);

    const remappedReply = await invoke(app, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: attachCookie,
      body: createEmailPayload({
        bodyHtml: "<p>Reply after explicit remap</p>",
        conversationId: "new-compose-conversation-id",
        internetMessageId: "<remapped-reply@example.com>",
        itemId: "remapped-reply-id",
      }),
    });
    assert.strictEqual(remappedReply.statusCode, 200);
    assert.strictEqual(remappedReply.body.status, "attached");
    assert.strictEqual(remappedReply.body.ticketNumber, "T2002");
    assert.strictEqual(attachActionFetchCount, 12);
    assert.strictEqual(attachActions[11].ticket_id, 2002);

    const failedExplicitAttach = await invoke(
      app,
      "POST",
      "/api/halo/tickets/:ticketId/sent-email",
      {
        url: "/api/halo/tickets/1001/sent-email",
        params: { ticketId: "1001" },
        cookie: attachCookie,
        body: createSendPayload({
          bodyHtml: "<p>Fail composed email</p>",
          composeAttachId: "compose-operation-failure",
          conversationId: "failed-compose-conversation",
          inReplyToMessageIds: [],
          ticketNumber: "T1001",
        }),
      }
    );
    assert.strictEqual(failedExplicitAttach.statusCode, 502);
    assert.strictEqual(failedExplicitAttach.body.status, "failed");
    assert.match(failedExplicitAttach.body.error, /Temporary composed action failure/);
    assert.strictEqual(failedExplicitAttach.body.debug.phase, "email-explicit-send-attach");
    assert.strictEqual(failedExplicitAttach.body.ticketNumber, "T1001");
    assert.strictEqual(attachActionFetchCount, 13);

    const retryExplicitAttach = await invoke(
      app,
      "POST",
      "/api/halo/tickets/:ticketId/sent-email",
      {
        url: "/api/halo/tickets/1001/sent-email",
        params: { ticketId: "1001" },
        cookie: attachCookie,
        body: createSendPayload({
          bodyHtml: "<p>Fail composed email</p>",
          composeAttachId: "compose-operation-failure",
          conversationId: "failed-compose-conversation",
          inReplyToMessageIds: [],
          ticketNumber: "T1001",
        }),
      }
    );
    assert.strictEqual(retryExplicitAttach.statusCode, 200);
    assert.strictEqual(retryExplicitAttach.body.status, "attached");
    assert.strictEqual(attachActionFetchCount, 14);

    const serverDerivedExplicitMailboxIdentity = await invoke(
      app,
      "POST",
      "/api/halo/tickets/:ticketId/sent-email",
      {
        url: "/api/halo/tickets/1001/sent-email",
        params: { ticketId: "1001" },
        cookie: attachCookie,
        body: createSendPayload({
          composeAttachId: "missing-mailbox-identity",
          conversationId: "",
          inReplyToMessageIds: [],
          mailboxEmail: "",
          ticketNumber: "T1001",
        }),
      }
    );
    assert.strictEqual(serverDerivedExplicitMailboxIdentity.statusCode, 200);
    assert.strictEqual(serverDerivedExplicitMailboxIdentity.body.status, "attached");
    assert.strictEqual(attachActionFetchCount, 15);

    const invalidComposeAttachId = await invoke(
      app,
      "POST",
      "/api/halo/tickets/:ticketId/sent-email",
      {
        url: "/api/halo/tickets/1001/sent-email",
        params: { ticketId: "1001" },
        cookie: attachCookie,
        body: createSendPayload({
          composeAttachId: "invalid compose id",
          conversationId: "",
          inReplyToMessageIds: [],
          ticketNumber: "T1001",
        }),
      }
    );
    assert.strictEqual(invalidComposeAttachId.statusCode, 400);
    assert.match(invalidComposeAttachId.body.error, /compose attachment ID/);
    assert.strictEqual(attachActionFetchCount, 15);

    const createConcurrentExplicitRequest = () => ({
      url: "/api/halo/tickets/1001/sent-email",
      params: { ticketId: "1001" },
      cookie: attachCookie,
      body: createSendPayload({
        bodyHtml: "<p>Concurrent composed email</p>",
        composeAttachId: "compose-operation-concurrent",
        conversationId: "concurrent-compose-conversation",
        inReplyToMessageIds: [],
        itemId: "concurrent-compose-draft-id",
        ticketNumber: "T1001",
      }),
    });
    const concurrentExplicitResults = await Promise.all([
      invoke(
        app,
        "POST",
        "/api/halo/tickets/:ticketId/sent-email",
        createConcurrentExplicitRequest()
      ),
      invoke(
        app,
        "POST",
        "/api/halo/tickets/:ticketId/sent-email",
        createConcurrentExplicitRequest()
      ),
    ]);
    assert.deepStrictEqual(concurrentExplicitResults.map((result) => result.body.status).sort(), [
      "already-attached",
      "attached",
    ]);
    assert.strictEqual(attachActionFetchCount, 16);
    assert.strictEqual(
      attachActions.filter((action) => /Concurrent composed email/.test(action.note_html)).length,
      1
    );

    const postActionAttachmentFingerprint = sha256Hex("post-action-attachment-fingerprint");
    const postActionAttachmentPrefetch = await invoke(
      app,
      "POST",
      "/api/halo/email-attachments/prefetch/start",
      {
        url: "/api/halo/email-attachments/prefetch/start",
        cookie: attachCookie,
        body: {
          draftItemId: "post-action-failure-draft-id",
          emailAttachmentFingerprint: postActionAttachmentFingerprint,
          emailAttachments: [
            {
              attachmentKey: ordinaryAttachmentKey,
              attachmentType: "file",
              contentSha256: sha256Hex(ordinaryAttachmentBytes),
              contentType: "application/pdf",
              name: "post-action-report.pdf",
              reportedSize: ordinaryAttachmentBytes.length,
            },
          ],
          operationId: "post-action-prefetch-operation",
          ticketId: 1001,
          ticketNumber: "T1001",
        },
      }
    );
    assert.strictEqual(postActionAttachmentPrefetch.statusCode, 200);
    assert.strictEqual(postActionAttachmentPrefetch.body.status, "pending");
    const postActionAttachmentStage = await invoke(
      app,
      "POST",
      "/api/halo/email-attachments/prefetch/:prefetchKey/items",
      {
        url: `/api/halo/email-attachments/prefetch/${postActionAttachmentPrefetch.body.prefetchKey}/items`,
        params: { prefetchKey: postActionAttachmentPrefetch.body.prefetchKey },
        cookie: attachCookie,
        body: {
          attachmentKey: ordinaryAttachmentKey,
          contentBase64: ordinaryAttachmentBytes.toString("base64"),
          contentFormat: "base64",
          contentSha256: sha256Hex(ordinaryAttachmentBytes),
        },
      }
    );
    assert.strictEqual(postActionAttachmentStage.statusCode, 200);
    assert.strictEqual(postActionAttachmentStage.body.status, "prepared");

    const originalConsumeEmailAttachmentPrefetch = store.consumeEmailAttachmentPrefetch;
    let failNextPrefetchConsume = true;
    store.consumeEmailAttachmentPrefetch = (...args) => {
      if (failNextPrefetchConsume) {
        failNextPrefetchConsume = false;
        throw new Error("Temporary prefetch consume failure");
      }
      return originalConsumeEmailAttachmentPrefetch(...args);
    };
    try {
      const createPostActionFailureRequest = () => ({
        url: "/api/halo/tickets/1001/sent-email",
        params: { ticketId: "1001" },
        cookie: attachCookie,
        body: createSendPayload({
          bodyHtml: "<p>Composed email with post-action failure</p>",
          composeAttachId: "compose-operation-post-action-failure",
          conversationId: "post-action-failure-conversation",
          emailAttachmentFingerprint: postActionAttachmentFingerprint,
          emailAttachmentDraftItemId: "post-action-failure-draft-id",
          emailAttachmentOperationId: "post-action-prefetch-operation",
          emailAttachmentPrefetchKey: postActionAttachmentPrefetch.body.prefetchKey,
          emailAttachmentStagingVersion: 2,
          emailAttachmentSummary: {
            attached: 0,
            detected: 1,
            failed: 0,
            selected: 1,
            skipped: 0,
            prepared: 1,
            warnings: [],
          },
          includeEmailAttachments: true,
          inReplyToMessageIds: [],
          itemId: "post-action-failure-draft-id",
          ticketNumber: "T1001",
        }),
      });
      const postActionFailure = await invoke(
        app,
        "POST",
        "/api/halo/tickets/:ticketId/sent-email",
        createPostActionFailureRequest()
      );
      assert.strictEqual(postActionFailure.statusCode, 502);
      assert.match(postActionFailure.body.error, /Temporary prefetch consume failure/);
      assert.strictEqual(attachActionFetchCount, 17);

      const postActionRetry = await invoke(
        app,
        "POST",
        "/api/halo/tickets/:ticketId/sent-email",
        createPostActionFailureRequest()
      );
      assert.strictEqual(postActionRetry.statusCode, 200);
      assert.strictEqual(postActionRetry.body.status, "already-attached");
      assert.strictEqual(attachActionFetchCount, 17);
    } finally {
      store.consumeEmailAttachmentPrefetch = originalConsumeEmailAttachmentPrefetch;
    }

    const partialUploadStage = await stageOrdinaryAttachments({
      draftItemId: "partial-upload-draft-id",
      filenames: ["partial-a.pdf", "partial-b.pdf"],
      operationId: "partial-upload-operation",
    });
    const deletedBeforePartialFailure = deletedStagedAttachmentIds.length;
    failedStagedAttachmentName = "partial-b.pdf";
    const partialUploadFailure = await invoke(
      app,
      "POST",
      "/api/halo/tickets/:ticketId/sent-email",
      {
        url: "/api/halo/tickets/1001/sent-email",
        params: { ticketId: "1001" },
        cookie: attachCookie,
        body: stagedSendPayload(partialUploadStage, {
          bodyHtml: "<p>Partial staged upload failure</p>",
        }),
      }
    );
    assert.strictEqual(partialUploadFailure.statusCode, 502);
    assert.match(partialUploadFailure.body.error, /Temporary attachment upload failure/);
    assert.strictEqual(deletedStagedAttachmentIds.length, deletedBeforePartialFailure + 1);
    const partialUploadRetryState = await store.getEmailAttachmentPrefetch(
      partialUploadStage.prefetchKey,
      { haloTenant: "https://customer.halopsa.com", ticketId: 1001 }
    );
    assert.strictEqual(partialUploadRetryState.status, "active");
    assert(partialUploadRetryState.items.every((item) => item.status === "prepared"));
    assert(partialUploadRetryState.items.every((item) => Buffer.isBuffer(item.contentCiphertext)));

    failedStagedAttachmentName = "";
    const partialUploadRetry = await invoke(app, "POST", "/api/halo/tickets/:ticketId/sent-email", {
      url: "/api/halo/tickets/1001/sent-email",
      params: { ticketId: "1001" },
      cookie: attachCookie,
      body: stagedSendPayload(partialUploadStage, {
        bodyHtml: "<p>Partial staged upload failure</p>",
      }),
    });
    assert.strictEqual(partialUploadRetry.statusCode, 200);
    assert.strictEqual(partialUploadRetry.body.status, "attached");

    const actionFailureStage = await stageOrdinaryAttachments({
      draftItemId: "action-failure-draft-id",
      filenames: ["action-failure.pdf"],
      operationId: "action-failure-operation",
    });
    const deletedBeforeActionFailure = deletedStagedAttachmentIds.length;
    failNextStagedAction = true;
    const stagedActionFailure = await invoke(
      app,
      "POST",
      "/api/halo/tickets/:ticketId/sent-email",
      {
        url: "/api/halo/tickets/1001/sent-email",
        params: { ticketId: "1001" },
        cookie: attachCookie,
        body: stagedSendPayload(actionFailureStage, {
          bodyHtml: "<p>Staged action failure</p>",
        }),
      }
    );
    assert.strictEqual(stagedActionFailure.statusCode, 502);
    assert.match(stagedActionFailure.body.error, /Temporary staged action failure/);
    assert.strictEqual(deletedStagedAttachmentIds.length, deletedBeforeActionFailure + 1);
    const actionFailureRetryState = await store.getEmailAttachmentPrefetch(
      actionFailureStage.prefetchKey,
      { haloTenant: "https://customer.halopsa.com", ticketId: 1001 }
    );
    assert.strictEqual(actionFailureRetryState.status, "active");
    assert.strictEqual(actionFailureRetryState.items[0].status, "prepared");
    assert(Buffer.isBuffer(actionFailureRetryState.items[0].contentCiphertext));

    const stagedActionRetry = await invoke(app, "POST", "/api/halo/tickets/:ticketId/sent-email", {
      url: "/api/halo/tickets/1001/sent-email",
      params: { ticketId: "1001" },
      cookie: attachCookie,
      body: stagedSendPayload(actionFailureStage, {
        bodyHtml: "<p>Staged action failure</p>",
      }),
    });
    assert.strictEqual(stagedActionRetry.statusCode, 200);
    assert.strictEqual(stagedActionRetry.body.status, "attached");

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
      body: createEmailPayload({ bodyHtml: "x".repeat(10 * 1024 * 1024) }),
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

  await store.close();

  const persistentDatabase = await createTestDatabase();
  const persistentStore = await persistentDatabase.createStore();
  const persistentApp = createMockApp();
  await registerTestRoutes(persistentApp, persistentStore);

  let persistentTokenFetchCount = 0;
  let persistentActionFetchCount = 0;
  const persistentActions = [];
  const persistentFetch = async (requestUrl, options = {}) => {
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
      persistentActions.push(JSON.parse(options.body)[0]);
      return jsonResponse({ id: 9900 + persistentActionFetchCount }, 201, "Created");
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  };
  global.fetch = persistentFetch;

  let persistentCookie;
  try {
    persistentCookie = await loginAndGetCookie(persistentApp);
    const persistentAttach = await invoke(
      persistentApp,
      "POST",
      "/api/halo/tickets/:ticketId/sent-email",
      {
        url: "/api/halo/tickets/1001/sent-email",
        params: { ticketId: "1001" },
        cookie: persistentCookie,
        body: createSendPayload({
          actionMode: "private-note",
          composeAttachId: "persistent-compose-operation",
          conversationId: "persistent-compose-conversation",
          inReplyToMessageIds: [],
          ticketNumber: "T1001",
        }),
      }
    );

    assert.strictEqual(persistentAttach.statusCode, 200);
    assert.strictEqual(persistentAttach.body.ok, true);
    assert.strictEqual(persistentTokenFetchCount, 1);
    assert.strictEqual(persistentActionFetchCount, 1);
    assert.strictEqual(persistentActions[0].outcome, "Private Note");
    assert.strictEqual(persistentActions[0].hiddenfromuser, true);
    assert.strictEqual(persistentActions[0].emailsubject, undefined);
  } finally {
    global.fetch = originalFetch;
    await persistentStore.close();
  }

  const restartedStore = await persistentDatabase.createStore();
  const restartedApp = createMockApp();
  await registerTestRoutes(restartedApp, restartedStore);

  try {
    global.fetch = persistentFetch;
    const restartedStatus = await invoke(restartedApp, "GET", "/api/auth/status", {
      url: "/api/auth/status",
      cookie: persistentCookie,
      headers: { authorization: "" },
    });
    assert.strictEqual(restartedStatus.statusCode, 200);
    assert.strictEqual(restartedStatus.body.authenticated, true);
    assert.strictEqual(restartedStatus.body.haloUrl, "https://customer.halopsa.com");
    assert(restartedStatus.body.backgroundSessionId);

    const restartedMatch = await invoke(restartedApp, "POST", "/api/halo/email/match", {
      url: "/api/halo/email/match",
      cookie: persistentCookie,
      headers: { authorization: "" },
      body: createEmailPayload({
        bodyHtml: "",
        bodyText: "",
        conversationId: "persistent-compose-conversation",
        internetMessageId: "<persistent-match@example.com>",
      }),
    });
    assert.strictEqual(restartedMatch.body.status, "matched");
    assert.strictEqual(restartedMatch.body.actionMode, "private-note");

    const restartedConversationAttach = await invoke(
      restartedApp,
      "POST",
      "/api/halo/email/auto-attach",
      {
        url: "/api/halo/email/auto-attach",
        cookie: persistentCookie,
        headers: { authorization: "" },
        body: createEmailPayload({
          bodyHtml: "<p>Reply after server restart</p>",
          conversationId: "persistent-compose-conversation",
          internetMessageId: "<persistent-reply@example.com>",
          itemId: "persistent-reply-item-id",
        }),
      }
    );
    assert.strictEqual(restartedConversationAttach.statusCode, 200);
    assert.strictEqual(restartedConversationAttach.body.ok, true);
    assert.strictEqual(restartedConversationAttach.body.status, "attached");
    assert.strictEqual(restartedConversationAttach.body.ticketNumber, "T1001");
    assert.strictEqual(persistentActionFetchCount, 2);
    assert.strictEqual(persistentActions[1].outcome, "Private Note");
    assert.strictEqual(persistentActions[1].hiddenfromuser, true);

    const publicOverride = await invoke(restartedApp, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: persistentCookie,
      headers: { authorization: "" },
      body: createEmailPayload({
        actionMode: "email",
        bodyHtml: "<p>Public override</p>",
        conversationId: "persistent-compose-conversation",
        internetMessageId: "<persistent-public@example.com>",
        itemId: "persistent-public-item-id",
      }),
    });
    assert.strictEqual(publicOverride.body.status, "attached");
    assert.strictEqual(persistentActions[2].outcome, "Email");

    const inheritedPublic = await invoke(restartedApp, "POST", "/api/halo/email/auto-attach", {
      url: "/api/halo/email/auto-attach",
      cookie: persistentCookie,
      headers: { authorization: "" },
      body: createEmailPayload({
        bodyHtml: "<p>Inherited public mode</p>",
        conversationId: "persistent-compose-conversation",
        internetMessageId: "<persistent-public-followup@example.com>",
        itemId: "persistent-public-followup-item-id",
      }),
    });
    assert.strictEqual(inheritedPublic.body.status, "attached");
    assert.strictEqual(persistentActions[3].outcome, "Email");
  } finally {
    global.fetch = originalFetch;
    await restartedStore.close();
    await persistentDatabase.close();
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
