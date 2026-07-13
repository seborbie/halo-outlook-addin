const assert = require("node:assert/strict");
const test = require("node:test");

const {
  BugReportError,
  createGitHubIssueClient,
  getBugReportConfig,
  hashSessionToken,
  registerBugReportRoutes,
} = require("./bugReports");
const { createHaloStore } = require("./haloStore");

function createMockApp() {
  const routes = { POST: new Map() };
  return {
    locals: {},
    routes,
    post(path, handler) {
      routes.POST.set(path, handler);
    },
  };
}

function createMockResponse() {
  return {
    body: undefined,
    statusCode: 200,
    status(value) {
      this.statusCode = value;
      return this;
    },
    json(value) {
      this.body = value;
      return this;
    },
  };
}

async function invoke(app, path, request = {}) {
  const response = createMockResponse();
  const handler = app.routes.POST.get(path);
  assert(handler, `Expected POST ${path} to be registered.`);
  await handler(
    {
      body: request.body || {},
      headers: { host: "localhost:3000", ...(request.headers || {}) },
      protocol: "https",
    },
    response
  );
  return response;
}

function createHarness(options = {}) {
  const app = createMockApp();
  const store = options.store || createHaloStore({ dbPath: ":memory:" });
  const user = store.upsertUser({
    displayName: "Support User",
    email: "support@example.com",
    objectId: "object-1",
    tenantId: "tenant-1",
  });
  const issues = [];
  const githubClient =
    options.githubClient ||
    {
      async createIssue(issue) {
        issues.push(issue);
        return { number: 42 };
      },
    };

  registerBugReportRoutes(app, {
    env: {
      BUG_REPORT_GITHUB_REPOSITORY: "example/private-bug-reports",
      BUG_REPORT_GITHUB_TOKEN: "test-token",
      PUBLIC_BASE_URL: "https://production.example.com",
      ...(options.env || {}),
    },
    githubClient,
    requireMicrosoftUser: options.requireMicrosoftUser || (async () => user),
    store,
  });

  return { app, issues, store, user };
}

function tokenFromSessionResponse(response) {
  const url = new URL(response.body.url);
  return new URLSearchParams(url.hash.slice(1)).get("token");
}

test("authenticated add-in users can submit one identity-redacted GitHub issue", async () => {
  const { app, issues, store } = createHarness();
  const sessionResponse = await invoke(app, "/api/bug-reports/session", {
    body: {
      diagnostics: {
        addInVersion: "2026.7.10-beta",
        emailSubject: "must not be collected",
        officeVersion: "16.0.1",
        outlookHost: "Outlook",
        outlookPlatform: "PC",
      },
    },
  });

  assert.equal(sessionResponse.statusCode, 201);
  assert.match(sessionResponse.body.url, /^https:\/\/localhost:3000\/bugreport#token=/);
  const token = tokenFromSessionResponse(sessionResponse);
  assert(token);

  const submitResponse = await invoke(app, "/api/bug-reports", {
    body: {
      additionalContext: "Only happens on first launch.",
      description:
        "The add-in stopped responding.\u0000\n![tracking](https://example.test/pixel)\n@octocat",
      expectedBehavior: "It should load tickets.",
      reproductionSteps: "Open Outlook, then open the add-in.",
      summary: "Add-in does not load",
    },
    headers: { "x-bug-report-session": token },
  });

  assert.equal(submitResponse.statusCode, 201);
  assert.deepEqual(submitResponse.body, { ok: true, reference: 42 });
  assert.equal(issues.length, 1);
  assert.equal(issues[0].title, "[Add-in bug] Add-in does not load");
  assert.deepEqual(issues[0].labels, ["bug"]);
  assert.doesNotMatch(issues[0].body, /Support User/);
  assert.doesNotMatch(issues[0].body, /support@example\.com/);
  assert.doesNotMatch(issues[0].body, /## Reporter/);
  assert.match(issues[0].body, /Reporter name and email.*were not included/);
  assert.match(issues[0].body, /2026\.7\.10-beta/);
  assert.doesNotMatch(issues[0].body, /must not be collected/);
  assert.doesNotMatch(issues[0].body, /\u0000/);
  assert.doesNotMatch(issues[0].body, /!\[tracking\]/);
  assert.doesNotMatch(issues[0].body, /@octocat/);

  const reusedResponse = await invoke(app, "/api/bug-reports", {
    body: { description: "Duplicate", summary: "Duplicate report" },
    headers: { "x-bug-report-session": token },
  });
  assert.equal(reusedResponse.statusCode, 401);
  assert.equal(issues.length, 1);
  store.close();
});

test("non-local report links use the configured public origin", async () => {
  const { app, store } = createHarness();
  const response = await invoke(app, "/api/bug-reports/session", {
    headers: {
      host: "internal-container:3000",
      "x-forwarded-proto": "http",
    },
  });

  assert.equal(response.statusCode, 201);
  assert.match(response.body.url, /^https:\/\/production\.example\.com\/bugreport#token=/);
  store.close();
});

test("a transient GitHub failure releases the session for retry", async () => {
  let attempts = 0;
  const { app, store } = createHarness({
    githubClient: {
      async createIssue() {
        attempts += 1;
        if (attempts === 1) {
          throw new BugReportError("Could not submit the bug report right now. Please try again.", 502);
        }
        return { number: 99 };
      },
    },
  });
  const sessionResponse = await invoke(app, "/api/bug-reports/session");
  const token = tokenFromSessionResponse(sessionResponse);
  const request = {
    body: { description: "A repeatable problem.", summary: "Retry test" },
    headers: { "x-bug-report-session": token },
  };

  const failure = await invoke(app, "/api/bug-reports", request);
  assert.equal(failure.statusCode, 502);

  const retry = await invoke(app, "/api/bug-reports", request);
  assert.equal(retry.statusCode, 201);
  assert.equal(retry.body.reference, 99);
  assert.equal(attempts, 2);
  store.close();
});

test("a release failure does not replace the original validation response", async () => {
  const realStore = createHaloStore({ dbPath: ":memory:" });
  const store = {
    ...realStore,
    releaseBugReportSession() {
      throw new Error("release failed");
    },
  };
  const { app } = createHarness({ store });
  const sessionResponse = await invoke(app, "/api/bug-reports/session");
  const token = tokenFromSessionResponse(sessionResponse);
  const originalConsoleError = console.error;
  console.error = () => {};

  try {
    const response = await invoke(app, "/api/bug-reports", {
      body: { description: "", summary: "Invalid report" },
      headers: { "x-bug-report-session": token },
    });
    assert.equal(response.statusCode, 400);
    assert.match(response.body.error, /What happened is required/);
  } finally {
    console.error = originalConsoleError;
    realStore.close();
  }
});

test("a successful GitHub issue remains successful if session consumption fails", async () => {
  const realStore = createHaloStore({ dbPath: ":memory:" });
  let releaseCalls = 0;
  const store = {
    ...realStore,
    consumeBugReportSession() {
      throw new Error("consume failed");
    },
    releaseBugReportSession(sessionHash) {
      releaseCalls += 1;
      return realStore.releaseBugReportSession(sessionHash);
    },
  };
  const { app } = createHarness({ store });
  const sessionResponse = await invoke(app, "/api/bug-reports/session");
  const token = tokenFromSessionResponse(sessionResponse);
  const originalConsoleError = console.error;
  console.error = () => {};

  try {
    const response = await invoke(app, "/api/bug-reports", {
      body: { description: "A real problem.", summary: "Consumption failure" },
      headers: { "x-bug-report-session": token },
    });
    assert.deepEqual(response.body, { ok: true, reference: 42 });
    assert.equal(response.statusCode, 201);
    assert.equal(releaseCalls, 0);

    const duplicate = await invoke(app, "/api/bug-reports", {
      body: { description: "A duplicate problem.", summary: "Duplicate attempt" },
      headers: { "x-bug-report-session": token },
    });
    assert.equal(duplicate.statusCode, 401);
    assert.equal(
      realStore.claimBugReportSession(hashSessionToken(token), Date.now() + 61 * 1000),
      null
    );
  } finally {
    console.error = originalConsoleError;
    realStore.close();
  }
});

test("unconfigured reporting returns 503 without creating a session", async () => {
  const { app, store } = createHarness({
    env: {
      BUG_REPORT_GITHUB_REPOSITORY: "",
      BUG_REPORT_GITHUB_TOKEN: "",
    },
  });
  const response = await invoke(app, "/api/bug-reports/session");
  assert.equal(response.statusCode, 503);
  assert.match(response.body.error, /temporarily unavailable/i);
  store.close();
});

test("unexpected server failures do not leak implementation details", async () => {
  const app = createMockApp();
  const originalConsoleError = console.error;
  console.error = () => {};
  registerBugReportRoutes(app, {
    env: {
      BUG_REPORT_GITHUB_REPOSITORY: "example/private-bug-reports",
      BUG_REPORT_GITHUB_TOKEN: "test-token",
    },
    githubClient: { async createIssue() {} },
    requireMicrosoftUser: async () => ({ id: 1 }),
    store: {
      createBugReportSession() {
        throw new Error("SQLITE_SECRET_FILE_PATH");
      },
    },
  });

  try {
    const response = await invoke(app, "/api/bug-reports/session");
    assert.equal(response.statusCode, 500);
    assert.match(response.body.error, /Unexpected bug report submission error/);
    assert.doesNotMatch(response.body.error, /SQLITE_SECRET_FILE_PATH/);
  } finally {
    console.error = originalConsoleError;
  }
});

test("report fields and authenticated sessions are validated", async () => {
  const { app, store } = createHarness();
  const missingToken = await invoke(app, "/api/bug-reports", {
    body: { description: "Details", summary: "Missing token" },
  });
  assert.equal(missingToken.statusCode, 401);

  const unknownToken = await invoke(app, "/api/bug-reports", {
    body: { description: "Details", summary: "Unknown token" },
    headers: { "x-bug-report-session": "unknown-token" },
  });
  assert.equal(unknownToken.statusCode, 401);

  const sessionResponse = await invoke(app, "/api/bug-reports/session");
  const token = tokenFromSessionResponse(sessionResponse);
  const missingDescription = await invoke(app, "/api/bug-reports", {
    body: { description: "", summary: "Missing description" },
    headers: { "x-bug-report-session": token },
  });
  assert.equal(missingDescription.statusCode, 400);

  const longSummary = await invoke(app, "/api/bug-reports", {
    body: { description: "Details", summary: "x".repeat(121) },
    headers: { "x-bug-report-session": token },
  });
  assert.equal(longSummary.statusCode, 400);

  const retryAfterValidation = await invoke(app, "/api/bug-reports", {
    body: { description: "Now supplied", summary: "Valid after correction" },
    headers: { "x-bug-report-session": token },
  });
  assert.equal(retryAfterValidation.statusCode, 201);
  store.close();
});

test("expired sessions and unauthenticated session requests are rejected", async () => {
  const { app, store, user } = createHarness({
    requireMicrosoftUser: async () => {
      throw new BugReportError("Microsoft add-in authentication is required.", 401);
    },
  });
  const unauthenticated = await invoke(app, "/api/bug-reports/session");
  assert.equal(unauthenticated.statusCode, 401);

  store.createBugReportSession({
    diagnostics: {},
    expiresAt: Date.now() - 1,
    sessionHash: "expired-session-hash",
    userId: user.id,
  });
  assert.equal(store.claimBugReportSession("expired-session-hash", Date.now()), null);
  store.close();
});

test("the GitHub client uses a scoped issue creation request", async () => {
  let capturedUrl = "";
  let capturedOptions = null;
  const config = getBugReportConfig({
    BUG_REPORT_GITHUB_REPOSITORY: "example/private-reports",
    BUG_REPORT_GITHUB_TOKEN: "secret-token",
  });
  const client = createGitHubIssueClient(config, {
    async fetchImpl(url, options) {
      capturedUrl = url;
      capturedOptions = options;
      return new Response(JSON.stringify({ number: 7 }), {
        headers: { "Content-Type": "application/json" },
        status: 201,
      });
    },
  });

  const issue = await client.createIssue({ body: "Details", labels: ["bug"], title: "Bug" });
  assert.equal(issue.number, 7);
  assert.equal(capturedUrl, "https://api.github.com/repos/example/private-reports/issues");
  assert.equal(capturedOptions.method, "POST");
  assert.equal(capturedOptions.headers.Authorization, "Bearer secret-token");
});

test("GitHub errors are converted to retryable, non-secret failures", async () => {
  const config = getBugReportConfig({
    BUG_REPORT_GITHUB_REPOSITORY: "example/private-reports",
    BUG_REPORT_GITHUB_TOKEN: "secret-token",
  });
  const rejectedClient = createGitHubIssueClient(config, {
    async fetchImpl() {
      return new Response(JSON.stringify({ message: "Bad credentials" }), {
        headers: { "Content-Type": "application/json" },
        status: 403,
      });
    },
  });
  const networkClient = createGitHubIssueClient(config, {
    async fetchImpl() {
      throw new Error("socket details that must stay server-side");
    },
  });

  await assert.rejects(
    () => rejectedClient.createIssue({ body: "Details", labels: [], title: "Bug" }),
    (error) =>
      error.status === 502 &&
      /Please try again/.test(error.message) &&
      /Bad credentials/.test(error.debug)
  );
  await assert.rejects(
    () => networkClient.createIssue({ body: "Details", labels: [], title: "Bug" }),
    (error) =>
      error.status === 502 &&
      /Please try again/.test(error.message) &&
      !error.message.includes("socket details")
  );
});
