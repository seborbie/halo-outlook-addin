const assert = require("assert");
const { registerHaloAuthRoutes } = require("./haloAuth");

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

function createMockReq({ url, body, cookie, params } = {}) {
  return {
    body,
    headers: {
      host: "localhost:3000",
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

async function loginAndGetCookie(app) {
  const start = await invoke(app, "POST", "/api/auth/start", {
    url: "/api/auth/start",
    body: { haloUrl: "https://customer.halopsa.com", clientId: "test-client-id" },
  });
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
    internetMessageId: "<message@example.com>",
    itemId: "outlook-item-id",
    normalizedSubject: "Example subject",
    subject: "RE: Example subject",
    to: [{ displayName: "Support User", emailAddress: "support@example.com" }],
    ...overrides,
  };
}

async function run() {
  const app = createMockApp();
  registerHaloAuthRoutes(app);

  const missingClientId = await invoke(app, "POST", "/api/auth/start", {
    url: "/api/auth/start",
    body: { haloUrl: "https://customer.halopsa.com" },
  });
  assert.strictEqual(missingClientId.statusCode, 400);
  assert.match(missingClientId.body.error, /client ID/);

  const invalidUrl = await invoke(app, "POST", "/api/auth/start", {
    url: "/api/auth/start",
    body: { haloUrl: "http://customer.halopsa.com", clientId: "test-client-id" },
  });
  assert.strictEqual(invalidUrl.statusCode, 400);
  assert.match(invalidUrl.body.error, /https/);

  const start = await invoke(app, "POST", "/api/auth/start", {
    url: "/api/auth/start",
    body: { haloUrl: "https://customer.halopsa.com/some/path", clientId: "test-client-id" },
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
  assert.strictEqual(haloAuthUrl.searchParams.get("redirect_uri"), "https://localhost:3000/auth/callback");
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
    assert.strictEqual(ping.body.debug.endpoint, "https://customer.halopsa.com/api/Tickets?count=1");
    assert.strictEqual(ping.body.debug.bodyExcerpt, '{"message":"Tickets permission missing"}');
    assert.strictEqual(ping.body.debug.requestedScope, "all");
  } finally {
    global.fetch = originalFetch;
  }

  let ticketTokenFetchCount = 0;
  let ticketListFetchCount = 0;

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

    throw new Error(`Unexpected fetch URL: ${url}`);
  };

  try {
    const ticketsStart = await invoke(app, "POST", "/api/auth/start", {
      url: "/api/auth/start",
      body: { haloUrl: "https://customer.halopsa.com", clientId: "test-client-id" },
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
  } finally {
    global.fetch = originalFetch;
  }

  const unauthenticatedAttach = await invoke(app, "POST", "/api/halo/tickets/:ticketId/email", {
    url: "/api/halo/tickets/1001/email",
    params: { ticketId: "1001" },
    body: createEmailPayload(),
  });
  assert.strictEqual(unauthenticatedAttach.statusCode, 401);
  assert.strictEqual(unauthenticatedAttach.body.ok, false);

  let attachTokenFetchCount = 0;
  let attachActionFetchCount = 0;

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
      assert.strictEqual(actions[0].ticket_id, 1001);
      assert.strictEqual(actions[0].outcome, "Email");
      assert.strictEqual(actions[0].emailsubject, "RE: Example subject");
      assert.strictEqual(actions[0].email_message_id, "<message@example.com>");
      assert.strictEqual(actions[0].actioninternetmessageid, "<message@example.com>");
      assert.match(actions[0].note, /Outlook email attached to ticket/);
      assert.match(actions[0].note_html, /Sender User &lt;sender@example.com&gt;/);
      assert.match(actions[0].note_html, /<p>Hello from Outlook<\/p>/);
      return jsonResponse({ id: 9001 }, 201, "Created");
    }

    throw new Error(`Unexpected fetch URL: ${url}`);
  };

  try {
    const attachCookie = await loginAndGetCookie(app);
    const attach = await invoke(app, "POST", "/api/halo/tickets/:ticketId/email", {
      url: "/api/halo/tickets/1001/email",
      params: { ticketId: "1001" },
      cookie: attachCookie,
      body: createEmailPayload(),
    });

    assert.strictEqual(attach.statusCode, 200);
    assert.strictEqual(attach.body.ok, true);
    assert.strictEqual(attach.body.message, "Email attached to Halo ticket");
    assert.strictEqual(attach.body.actionId, "9001");
    assert.strictEqual(attachTokenFetchCount, 1);
    assert.strictEqual(attachActionFetchCount, 1);

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
}

run()
  .then(() => {
    console.log("Halo auth smoke tests passed");
  })
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
