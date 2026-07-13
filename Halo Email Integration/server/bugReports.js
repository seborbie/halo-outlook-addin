const crypto = require("crypto");

const DEFAULT_SESSION_TTL_MINUTES = 15;
const MAX_SESSION_TTL_MINUTES = 60;
const MAX_SUMMARY_LENGTH = 120;
const MAX_FIELD_LENGTH = 5000;
const GITHUB_API_VERSION = "2022-11-28";

function registerBugReportRoutes(app, options = {}) {
  if (app.locals && app.locals.bugReportRoutesRegistered) {
    return;
  }

  const env = options.env || process.env;
  const store = options.store;
  const requireMicrosoftUser = options.requireMicrosoftUser;
  const config = getBugReportConfig(env);
  const githubClient = options.githubClient || createGitHubIssueClient(config);

  if (!store || typeof requireMicrosoftUser !== "function") {
    throw new Error("Bug report routes require the Halo store and Microsoft authentication.");
  }

  if (app.locals) {
    app.locals.bugReportRoutesRegistered = true;
  }

  app.post("/api/bug-reports/session", async (req, res) => {
    try {
      requireConfigured(config);
      const user = await requireMicrosoftUser(req);
      const diagnostics = normalizeDiagnostics(req.body && req.body.diagnostics);
      const sessionToken = crypto.randomBytes(32).toString("base64url");
      const expiresAt = Date.now() + config.sessionTtlMinutes * 60 * 1000;

      store.createBugReportSession({
        diagnostics,
        expiresAt,
        sessionHash: hashSessionToken(sessionToken),
        userId: user.id,
      });

      sendJson(res, 201, {
        expiresAt: new Date(expiresAt).toISOString(),
        url: `${getPublicOrigin(req, env)}/bugreport#token=${encodeURIComponent(sessionToken)}`,
      });
    } catch (error) {
      sendBugReportError(res, error);
    }
  });

  app.post("/api/bug-reports", async (req, res) => {
    const sessionToken = getSessionToken(req);
    const sessionHash = sessionToken ? hashSessionToken(sessionToken) : "";
    let session = null;

    try {
      requireConfigured(config);
      if (!sessionHash) {
        throw new BugReportError("This bug report link is missing or invalid.", 401);
      }

      session = store.claimBugReportSession(sessionHash, Date.now());
      if (!session) {
        throw new BugReportError(
          "This bug report link has expired or has already been used. Open a new report from the add-in.",
          401
        );
      }

      const report = normalizeBugReport(req.body);
      const issue = await githubClient.createIssue({
        body: buildIssueBody(report, session),
        labels: config.labels,
        title: `[Add-in bug] ${preventMentions(report.summary)}`,
      });

      try {
        const consumed = store.consumeBugReportSession(sessionHash, Date.now());
        if (!consumed) {
          console.error("A submitted bug report session could not be marked as consumed.");
        }
      } catch (error) {
        console.error("A submitted bug report session could not be marked as consumed.", error);
      }
      session = null;
      sendJson(res, 201, {
        ok: true,
        reference: issue.number,
      });
    } catch (error) {
      if (session) {
        try {
          store.releaseBugReportSession(sessionHash);
        } catch (releaseError) {
          console.error("A failed bug report session could not be released.", releaseError);
        }
      }
      sendBugReportError(res, error);
    }
  });
}

function getBugReportConfig(env = process.env) {
  const repository = String(env.BUG_REPORT_GITHUB_REPOSITORY || "").trim();
  const repositoryMatch = /^([^/\s]+)\/([^/\s]+)$/.exec(repository);
  const token = String(env.BUG_REPORT_GITHUB_TOKEN || "").trim();
  const requestedTtl = Number(env.BUG_REPORT_SESSION_TTL_MINUTES || DEFAULT_SESSION_TTL_MINUTES);
  const sessionTtlMinutes =
    Number.isFinite(requestedTtl) && requestedTtl > 0
      ? Math.min(requestedTtl, MAX_SESSION_TTL_MINUTES)
      : DEFAULT_SESSION_TTL_MINUTES;
  const labels = String(env.BUG_REPORT_GITHUB_LABELS || "bug,outlook-addin")
    .split(",")
    .map((label) => label.trim())
    .filter(Boolean);

  return {
    configured: Boolean(repositoryMatch && token),
    labels,
    owner: repositoryMatch ? repositoryMatch[1] : "",
    repo: repositoryMatch ? repositoryMatch[2] : "",
    repository,
    sessionTtlMinutes,
    token,
  };
}

function createGitHubIssueClient(config, options = {}) {
  const fetchImpl = options.fetchImpl || globalThis.fetch;

  return {
    async createIssue(issue) {
      if (typeof fetchImpl !== "function") {
        throw new BugReportError("GitHub issue submission is unavailable.", 502);
      }

      const abortController = new AbortController();
      const timeout = setTimeout(() => abortController.abort(), 15 * 1000);
      timeout.unref();
      let response;
      let payload;

      try {
        response = await fetchImpl(
          `https://api.github.com/repos/${encodeURIComponent(config.owner)}/${encodeURIComponent(config.repo)}/issues`,
          {
            body: JSON.stringify(issue),
            headers: {
              Accept: "application/vnd.github+json",
              Authorization: `Bearer ${config.token}`,
              "Content-Type": "application/json",
              "User-Agent": "halo-outlook-addin-bug-reporter",
              "X-GitHub-Api-Version": GITHUB_API_VERSION,
            },
            method: "POST",
            signal: abortController.signal,
          }
        );
        payload = await response.json().catch(() => ({}));
      } catch (error) {
        throw new BugReportError(
          "Could not submit the bug report right now. Please try again.",
          502,
          `GitHub issue request failed: ${error && error.message ? error.message : "network error"}`
        );
      } finally {
        clearTimeout(timeout);
      }

      if (!response.ok || !Number.isInteger(payload.number)) {
        const detail = payload && payload.message ? `: ${String(payload.message).slice(0, 300)}` : "";
        throw new BugReportError(
          "Could not submit the bug report right now. Please try again.",
          502,
          `GitHub issue creation failed with HTTP ${response.status}${detail}`
        );
      }

      return {
        number: payload.number,
      };
    },
  };
}

function normalizeBugReport(body) {
  const value = body && typeof body === "object" ? body : {};
  return {
    additionalContext: normalizeTextField(value.additionalContext, "Additional context", false),
    description: normalizeTextField(value.description, "What happened", true),
    expectedBehavior: normalizeTextField(value.expectedBehavior, "Expected behaviour", false),
    reproductionSteps: normalizeTextField(value.reproductionSteps, "Reproduction steps", false),
    summary: normalizeSummary(value.summary),
  };
}

function normalizeSummary(value) {
  const summary = String(value || "")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, "")
    .replace(/[\r\n]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!summary) {
    throw new BugReportError("Summary is required.", 400);
  }
  if (summary.length > MAX_SUMMARY_LENGTH) {
    throw new BugReportError(`Summary must be ${MAX_SUMMARY_LENGTH} characters or fewer.`, 400);
  }
  return summary;
}

function normalizeTextField(value, label, required) {
  const text = String(value || "")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, "")
    .replace(/\r\n/g, "\n")
    .trim();
  if (required && !text) {
    throw new BugReportError(`${label} is required.`, 400);
  }
  if (text.length > MAX_FIELD_LENGTH) {
    throw new BugReportError(`${label} must be ${MAX_FIELD_LENGTH} characters or fewer.`, 400);
  }
  return text;
}

function normalizeDiagnostics(value) {
  const diagnostics = value && typeof value === "object" ? value : {};
  return {
    addInVersion: normalizeDiagnosticValue(diagnostics.addInVersion),
    officeVersion: normalizeDiagnosticValue(diagnostics.officeVersion),
    outlookHost: normalizeDiagnosticValue(diagnostics.outlookHost),
    outlookPlatform: normalizeDiagnosticValue(diagnostics.outlookPlatform),
  };
}

function normalizeDiagnosticValue(value) {
  return String(value || "")
    .replace(/[\u0000-\u001f\u007f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 100);
}

function buildIssueBody(report, session, now = new Date()) {
  const diagnostics = session.diagnostics || {};
  return [
    "## What happened",
    safeMarkdown(report.description),
    "",
    "## Reproduction steps",
    safeMarkdown(report.reproductionSteps || "Not provided."),
    "",
    "## Expected behaviour",
    safeMarkdown(report.expectedBehavior || "Not provided."),
    "",
    "## Additional context",
    safeMarkdown(report.additionalContext || "Not provided."),
    "",
    "## Reporter",
    `- Name: ${safeInlineMarkdown(session.displayName || "Not provided")}`,
    `- Email: ${inlineCode(session.email || "Not provided")}`,
    "",
    "## Safe diagnostics",
    `- Add-in version: ${inlineCode(diagnostics.addInVersion || "Unknown")}`,
    `- Outlook host: ${inlineCode(diagnostics.outlookHost || "Unknown")}`,
    `- Outlook platform: ${inlineCode(diagnostics.outlookPlatform || "Unknown")}`,
    `- Office version: ${inlineCode(diagnostics.officeVersion || "Unknown")}`,
    `- Submitted: ${now.toISOString()}`,
    "",
    "_No email subject, recipients, body, attachments, or Halo ticket data were collected._",
  ].join("\n");
}

function safeMarkdown(value) {
  return preventMentions(String(value || "")).replace(/([\\`*_{}\[\]<>#|])/g, "\\$1");
}

function safeInlineMarkdown(value) {
  return preventMentions(stripControlCharacters(value)).replace(
    /([\\`*_{}\[\]()<>#+!|])/g,
    "\\$1"
  );
}

function inlineCode(value) {
  return `\`${stripControlCharacters(value).replace(/`/g, "'")}\``;
}

function stripControlCharacters(value) {
  return String(value || "").replace(/[\u0000-\u001f\u007f]/g, " ").trim();
}

function preventMentions(value) {
  return value.replace(/@/g, "@\u200b");
}

function getSessionToken(req) {
  const headers = (req && req.headers) || {};
  const value = headers["x-bug-report-session"] || headers["X-Bug-Report-Session"] || "";
  return Array.isArray(value) ? String(value[0] || "").trim() : String(value).trim();
}

function hashSessionToken(value) {
  return crypto.createHash("sha256").update(`bug-report:${value}`).digest("base64url");
}

function getPublicOrigin(req, env) {
  const headers = (req && req.headers) || {};
  const forwardedProto = String(headers["x-forwarded-proto"] || "").split(",")[0].trim();
  const protocol = forwardedProto || req.protocol || "https";
  const host = String(headers.host || "localhost:3000");
  const requestOrigin = `${protocol}://${host}`;

  if (isLoopbackHost(host)) {
    return requestOrigin;
  }

  const configured = String(env.PUBLIC_BASE_URL || "").trim().replace(/\/+$/, "");
  return configured || requestOrigin;
}

function isLoopbackHost(host) {
  const normalized = String(host || "").trim().toLowerCase();
  const hostname = normalized.startsWith("[")
    ? normalized.slice(1, normalized.indexOf("]"))
    : normalized.split(":")[0];
  return hostname === "localhost" || hostname === "127.0.0.1" || hostname === "::1";
}

function requireConfigured(config) {
  if (!config.configured) {
    throw new BugReportError("Bug reporting is temporarily unavailable.", 503);
  }
}

function sendBugReportError(res, error) {
  const status = error && Number.isInteger(error.status) ? error.status : 500;
  const isSafeError = error instanceof BugReportError || status < 500;
  const message =
    isSafeError && error && error.message
      ? error.message
      : "Unexpected bug report submission error. Please try again.";

  if (error && error.debug) {
    console.error(error.debug);
  } else if (status >= 500 && !(error instanceof BugReportError)) {
    console.error(error);
  }

  sendJson(res, status, {
    error: message,
    ok: false,
  });
}

function sendJson(res, status, body) {
  res.status(status).json(body);
}

class BugReportError extends Error {
  constructor(message, status, debug = "") {
    super(message);
    this.status = status;
    this.debug = debug;
  }
}

module.exports = {
  BugReportError,
  buildIssueBody,
  createGitHubIssueClient,
  getBugReportConfig,
  hashSessionToken,
  normalizeBugReport,
  normalizeDiagnostics,
  registerBugReportRoutes,
};
