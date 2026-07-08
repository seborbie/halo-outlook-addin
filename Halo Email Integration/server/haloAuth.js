const crypto = require("crypto");

const SESSION_COOKIE = "halo_session";
const AUTH_PATH = "/auth/authorize";
const TOKEN_PATH = "/auth/token";
const TEST_PATH = "/api/Tickets?count=1";
const TICKETS_COUNT = 50;
const DEFAULT_SCOPE = "all";
const STATE_TTL_MS = 10 * 60 * 1000;
const HANDOFF_TTL_MS = 5 * 60 * 1000;
const SESSION_TTL_MS = 24 * 60 * 60 * 1000;
const MAX_JSON_BODY_BYTES = 32 * 1024;
const MAX_EMAIL_JSON_BODY_BYTES = 2 * 1024 * 1024;

const pendingStates = new Map();
const handoffs = new Map();
const sessions = new Map();
const backgroundSessions = new Map();
const conversationMappings = new Map();
const conversationMappingsByKey = new Map();
const messageMappingsByKey = new Map();

const encryptionKey = crypto.randomBytes(32);

function registerHaloAuthRoutes(app) {
  if (app.locals && app.locals.haloAuthRoutesRegistered) {
    return;
  }

  if (app.locals) {
    app.locals.haloAuthRoutesRegistered = true;
  }

  setInterval(cleanExpiredRecords, 60 * 1000).unref();

  app.post("/api/auth/start", async (req, res) => {
    try {
      const body = await readJsonBody(req);
      const haloUrl = normalizeHaloUrl(body.haloUrl);
      const clientId = normalizeClientId(body.clientId);
      const state = randomBase64Url(32);
      const codeVerifier = randomBase64Url(64);
      const codeChallenge = base64Url(crypto.createHash("sha256").update(codeVerifier).digest());
      const now = Date.now();

      pendingStates.set(state, {
        haloUrl,
        clientId,
        scope: DEFAULT_SCOPE,
        codeVerifier,
        codeChallenge,
        expiresAt: now + STATE_TTL_MS,
      });

      sendJson(res, 200, {
        dialogUrl: `${getBaseUrl(req)}/auth/start?state=${encodeURIComponent(state)}`,
      });
    } catch (error) {
      sendJson(res, 400, { error: publicError(error) });
    }
  });

  app.get("/auth/start", (req, res) => {
    try {
      const url = getRequestUrl(req);
      const state = url.searchParams.get("state");
      const pending = state ? pendingStates.get(state) : null;

      if (!state || !pending || pending.expiresAt <= Date.now()) {
        if (state) {
          pendingStates.delete(state);
        }
        sendAuthResultPage(res, {
          status: "failed",
          message: "Halo API Auth failed",
          error: "The Halo login request expired. Start login again from the add-in.",
        });
        return;
      }

      const authUrl = new URL(resolveHaloUrl(pending.haloUrl, AUTH_PATH));
      authUrl.searchParams.set("response_type", "code");
      authUrl.searchParams.set("client_id", pending.clientId);
      authUrl.searchParams.set("redirect_uri", `${getBaseUrl(req)}/auth/callback`);
      authUrl.searchParams.set("scope", pending.scope);
      authUrl.searchParams.set("state", state);
      authUrl.searchParams.set("code_challenge", pending.codeChallenge);
      authUrl.searchParams.set("code_challenge_method", "S256");

      res.redirect(authUrl.toString());
    } catch (error) {
      sendAuthResultPage(res, {
        status: "failed",
        message: "Halo API Auth failed",
        error: publicError(error),
      });
    }
  });

  app.get("/auth/callback", async (req, res) => {
    const url = getRequestUrl(req);
    const state = url.searchParams.get("state");
    const code = url.searchParams.get("code");
    const haloError = url.searchParams.get("error");
    const pending = state ? pendingStates.get(state) : null;

    if (state) {
      pendingStates.delete(state);
    }

    try {
      if (haloError) {
        throw new Error(url.searchParams.get("error_description") || haloError);
      }

      if (!state || !code || !pending || pending.expiresAt <= Date.now()) {
        throw new Error("The Halo login response was invalid or expired.");
      }

      const tokenPayload = await exchangeAuthorizationCode({
        haloUrl: pending.haloUrl,
        clientId: pending.clientId,
        code,
        codeVerifier: pending.codeVerifier,
        redirectUri: `${getBaseUrl(req)}/auth/callback`,
        scope: pending.scope,
      });

      const encryptedToken = encryptJson(tokenPayload);
      const handoffCode = randomBase64Url(32);
      handoffs.set(handoffCode, {
        haloUrl: pending.haloUrl,
        clientId: pending.clientId,
        scope: pending.scope,
        encryptedToken,
        expiresAt: Date.now() + HANDOFF_TTL_MS,
      });

      sendAuthResultPage(res, {
        status: "success",
        message: "Halo API Auth works",
        handoffCode,
      });
    } catch (error) {
      sendAuthResultPage(res, {
        status: "failed",
        message: "Halo API Auth failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.post("/api/auth/complete", async (req, res) => {
    try {
      const body = await readJsonBody(req);
      const handoffCode = typeof body.handoffCode === "string" ? body.handoffCode : "";
      const handoff = handoffs.get(handoffCode);

      if (!handoff || handoff.expiresAt <= Date.now()) {
        if (handoffCode) {
          handoffs.delete(handoffCode);
        }
        sendJson(res, 400, { error: "The Halo login handoff expired. Start login again." });
        return;
      }

      handoffs.delete(handoffCode);

      const sessionId = randomBase64Url(32);
      const sessionHash = hashSessionId(sessionId);
      const expiresAt = Date.now() + SESSION_TTL_MS;
      const backgroundSessionId = createBackgroundSession(sessionHash, expiresAt);

      sessions.set(sessionHash, {
        haloUrl: handoff.haloUrl,
        clientId: handoff.clientId,
        scope: handoff.scope,
        encryptedToken: handoff.encryptedToken,
        expiresAt,
      });

      res.setHeader("Set-Cookie", serializeSessionCookie(sessionId, Math.floor(SESSION_TTL_MS / 1000)));
      sendJson(res, 200, {
        authenticated: true,
        backgroundSessionId,
        expiresAt: new Date(expiresAt).toISOString(),
      });
    } catch (error) {
      sendJson(res, 400, { error: publicError(error) });
    }
  });

  app.post("/api/auth/background-session", (req, res) => {
    const sessionId = getSessionIdFromRequest(req);
    const record = getSessionRecordBySessionId(sessionId);

    if (!sessionId || !record) {
      sendJson(res, 401, {
        ok: false,
        error: "No active Halo session.",
      });
      return;
    }

    sendJson(res, 200, {
      ok: true,
      backgroundSessionId: createBackgroundSession(hashSessionId(sessionId), record.expiresAt),
      expiresAt: new Date(record.expiresAt).toISOString(),
    });
  });

  app.get("/api/auth/status", (req, res) => {
    const record = getSessionRecord(req);

    sendJson(res, 200, {
      authenticated: Boolean(record),
      haloUrl: record ? record.haloUrl : null,
      expiresAt: record ? new Date(record.expiresAt).toISOString() : null,
    });
  });

  app.get("/api/halo/ping", async (req, res) => {
    try {
      const record = getSessionRecord(req);

      if (!record) {
        sendJson(res, 401, {
          ok: false,
          message: "Halo API Auth failed",
          error: "No active Halo session.",
        });
        return;
      }

      await callHaloApiWithRefresh(record, TEST_PATH, "api-test", "Halo test query failed");

      sendJson(res, 200, {
        ok: true,
        message: "Halo API Auth works",
      });
    } catch (error) {
      sendJson(res, 502, {
        ok: false,
        message: "Halo API Auth failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.get("/api/halo/tickets", async (req, res) => {
    try {
      const record = getSessionRecord(req);

      if (!record) {
        sendJson(res, 401, {
          ok: false,
          message: "Halo ticket list failed",
          error: "No active Halo session.",
        });
        return;
      }

      const path = buildMyOpenTicketsPath();
      const payload = await callHaloApiWithRefresh(
        record,
        path,
        "tickets-list",
        "Halo ticket list failed"
      );
      const tickets = normalizeTickets(payload);

      sendJson(res, 200, {
        ok: true,
        tickets,
      });
    } catch (error) {
      sendJson(res, 502, {
        ok: false,
        message: "Halo ticket list failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.post("/api/halo/tickets/:ticketId/email", async (req, res) => {
    try {
      const record = getSessionRecord(req);

      if (!record) {
        sendJson(res, 401, {
          ok: false,
          message: "Email attach failed",
          error: "No active Halo session.",
        });
        return;
      }

      const ticketId = getTicketIdFromRequest(req);
      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      const email = normalizeEmailPayload(body);
      const ticketNumber = stringifyField(body.ticketNumber);
      const existingMapping = findConversationMappingForEmail(email);
      const isInitialChainAttach = !existingMapping;
      const actionPayload = buildEmailActionPayload(ticketId, email, {
        bodyMode: isInitialChainAttach ? "full" : "trimmed",
      });
      const payload = await callHaloApiWithRefresh(record, {
        body: [actionPayload],
        method: "POST",
        path: "/api/Actions",
        phase: "email-attach",
        messagePrefix: "Halo email attach failed",
      });
      const actionId = getCreatedActionId(payload);

      storeConversationMapping({
        email,
        includeThreadMessageIds: isInitialChainAttach,
        ticketId,
        ticketNumber,
      });

      sendJson(res, 200, {
        ok: true,
        attachMode: isInitialChainAttach ? "full-chain" : "latest-reply",
        message: isInitialChainAttach
          ? "Full email chain attached to Halo ticket"
          : "Email attached to Halo ticket",
        actionId: actionId || undefined,
        backgroundSessionId: createBackgroundSessionForRequest(req) || undefined,
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Email attach failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.post("/api/halo/email/auto-attach", async (req, res) => {
    try {
      const record = getSessionRecord(req);

      if (!record) {
        sendJson(res, 401, {
          ok: false,
          message: "Email auto-attach failed",
          error: "No active Halo session.",
        });
        return;
      }

      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      const email = normalizeEmailPayload(body);
      const match = findConversationMappingForEmail(email);

      if (!match) {
        sendJson(res, 200, {
          ok: true,
          status: "no-match",
        });
        return;
      }

      if (match.status === "already-attached") {
        sendJson(res, 200, {
          ok: true,
          status: "already-attached",
          ticketId: match.mapping.ticketId,
          ticketNumber: match.mapping.ticketNumber,
          message: `This email is already attached to ticket ${getMappingTicketLabel(match.mapping)}.`,
        });
        return;
      }

      const actionPayload = buildEmailActionPayload(match.mapping.ticketId, email);
      const payload = await callHaloApiWithRefresh(record, {
        body: [actionPayload],
        method: "POST",
        path: "/api/Actions",
        phase: "email-auto-attach",
        messagePrefix: "Halo email auto-attach failed",
      });
      const actionId = getCreatedActionId(payload);

      markEmailSynced(match.mapping, email);

      sendJson(res, 200, {
        ok: true,
        status: "attached",
        ticketId: match.mapping.ticketId,
        ticketNumber: match.mapping.ticketNumber,
        message: `Email automatically added to ticket ${getMappingTicketLabel(match.mapping)}.`,
        actionId: actionId || undefined,
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Email auto-attach failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.post("/api/halo/email/send-auto-attach", async (req, res) => {
    try {
      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      const record = getSessionRecord(req) || getBackgroundSessionRecord(body.backgroundSessionId);

      if (!record) {
        sendJson(res, 200, {
          ok: true,
          status: "no-session",
        });
        return;
      }

      const email = normalizeSendEmailPayload(body);
      const match = findConversationMappingForEmail(email);

      if (!match) {
        sendJson(res, 200, {
          ok: true,
          status: "no-match",
        });
        return;
      }

      if (match.status === "already-attached") {
        sendJson(res, 200, {
          ok: true,
          status: "already-attached",
          ticketId: match.mapping.ticketId,
          ticketNumber: match.mapping.ticketNumber,
          message: `This email is already attached to ticket ${getMappingTicketLabel(match.mapping)}.`,
        });
        return;
      }

      const actionPayload = buildEmailActionPayload(match.mapping.ticketId, email);
      const payload = await callHaloApiWithTicketContext(match.mapping, () =>
        callHaloApiWithRefresh(record, {
          body: [actionPayload],
          method: "POST",
          path: "/api/Actions",
          phase: "email-send-auto-attach",
          messagePrefix: "Halo sent email auto-attach failed",
        })
      );
      const actionId = getCreatedActionId(payload);

      markEmailSynced(match.mapping, email);

      sendJson(res, 200, {
        ok: true,
        status: "attached",
        ticketId: match.mapping.ticketId,
        ticketNumber: match.mapping.ticketNumber,
        message: `Sent email added to Halo ticket ${getMappingTicketLabel(match.mapping)}.`,
        actionId: actionId || undefined,
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        status: "failed",
        message: "Sent email auto-attach failed",
        error: publicError(error),
        debug: publicDebug(error),
        ticketNumber: getErrorTicketNumber(error),
      });
    }
  });

  app.post("/api/auth/logout", (req, res) => {
    const sessionId = getSessionIdFromRequest(req);

    if (sessionId) {
      const sessionHash = hashSessionId(sessionId);
      sessions.delete(sessionHash);
      deleteBackgroundSessionsForSessionHash(sessionHash);
    }

    res.setHeader("Set-Cookie", clearSessionCookie());
    sendJson(res, 200, { authenticated: false });
  });
}

async function exchangeAuthorizationCode({ haloUrl, clientId, code, codeVerifier, redirectUri, scope }) {
  const form = new URLSearchParams();
  form.set("grant_type", "authorization_code");
  form.set("code", code);
  form.set("redirect_uri", redirectUri);
  form.set("client_id", clientId);
  form.set("code_verifier", codeVerifier);
  form.set("scope", scope || DEFAULT_SCOPE);

  const requestUrl = resolveHaloUrl(haloUrl, TOKEN_PATH);
  const response = await fetch(requestUrl, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: form.toString(),
  });

  const responseDetails = await readResponseDetails(response, requestUrl);
  if (!response.ok) {
    throw HttpError.fromResponse("Halo token exchange failed", "token-exchange", responseDetails);
  }

  if (!responseDetails.payload.access_token) {
    throw new Error("Halo token exchange did not return an access token.");
  }

  return annotateTokenPayload(responseDetails.payload);
}

async function refreshAccessToken(record, currentTokenPayload) {
  if (!currentTokenPayload.refresh_token) {
    throw new Error("The Halo access token expired and no refresh token is available.");
  }

  const form = new URLSearchParams();
  form.set("grant_type", "refresh_token");
  form.set("refresh_token", currentTokenPayload.refresh_token);
  form.set("client_id", record.clientId);
  form.set("scope", record.scope || DEFAULT_SCOPE);

  const requestUrl = resolveHaloUrl(record.haloUrl, TOKEN_PATH);
  const response = await fetch(requestUrl, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: form.toString(),
  });

  const responseDetails = await readResponseDetails(response, requestUrl);
  if (!response.ok) {
    throw HttpError.fromResponse("Halo refresh token request failed", "token-refresh", responseDetails);
  }

  const nextTokenPayload = annotateTokenPayload({
    ...currentTokenPayload,
    ...responseDetails.payload,
    refresh_token: responseDetails.payload.refresh_token || currentTokenPayload.refresh_token,
  });

  record.encryptedToken = encryptJson(nextTokenPayload);
  return nextTokenPayload;
}

async function getValidTokenPayload(record) {
  let tokenPayload = decryptJson(record.encryptedToken);

  if (isTokenExpired(tokenPayload)) {
    tokenPayload = await refreshAccessToken(record, tokenPayload);
  }

  return tokenPayload;
}

async function callHaloApiWithRefresh(record, requestOrPath, phase, messagePrefix) {
  const request =
    typeof requestOrPath === "string"
      ? { path: requestOrPath, phase, messagePrefix }
      : requestOrPath;
  let tokenPayload = await getValidTokenPayload(record);

  try {
    return await fetchHaloJson({
      body: request.body,
      haloUrl: record.haloUrl,
      messagePrefix: request.messagePrefix,
      method: request.method || "GET",
      path: request.path,
      phase: request.phase,
      scope: record.scope,
      tokenPayload,
    });
  } catch (error) {
    if (!isUnauthorizedError(error) || !tokenPayload.refresh_token) {
      throw error;
    }

    tokenPayload = await refreshAccessToken(record, tokenPayload);
    return fetchHaloJson({
      body: request.body,
      haloUrl: record.haloUrl,
      messagePrefix: request.messagePrefix,
      method: request.method || "GET",
      path: request.path,
      phase: request.phase,
      scope: record.scope,
      tokenPayload,
    });
  }
}

async function fetchHaloJson({ body, haloUrl, messagePrefix, method, path, phase, scope, tokenPayload }) {
  const requestUrl = resolveHaloUrl(haloUrl, path);
  const headers = {
    Accept: "application/json",
    Authorization: `Bearer ${tokenPayload.access_token}`,
  };
  const options = {
    method,
    headers,
  };

  if (body !== undefined) {
    headers["Content-Type"] = "application/json";
    options.body = JSON.stringify(body);
  }

  const response = await fetch(requestUrl, {
    ...options,
  });

  if (!response.ok) {
    const responseDetails = await readResponseDetails(response, requestUrl);
    const error = HttpError.fromResponse(messagePrefix, phase, responseDetails);
    error.debug.method = method;
    error.debug.requestedScope = scope || DEFAULT_SCOPE;
    throw error;
  }

  return readResponseJson(response, requestUrl);
}

async function readResponseJson(response, requestUrl) {
  const responseDetails = await readResponseDetails(response, requestUrl);
  return responseDetails.payload;
}

function buildMyOpenTicketsPath() {
  const params = new URLSearchParams();
  params.set("count", String(TICKETS_COUNT));
  params.set("open_only", "true");
  params.set("mine", "true");
  params.set("includeagent", "true");
  params.set("includestatus", "true");

  return `/api/Tickets?${params.toString()}`;
}

function getTicketIdFromRequest(req) {
  const candidate =
    (req.params && req.params.ticketId) ||
    (getRequestUrl(req).pathname.match(/\/api\/halo\/tickets\/([^/]+)\/email$/) || [])[1];
  const ticketId = Number.parseInt(String(candidate || ""), 10);

  if (!Number.isInteger(ticketId) || ticketId <= 0) {
    throw new RequestError("A valid Halo ticket ID is required.", 400);
  }

  return ticketId;
}

function normalizeEmailPayload(value) {
  if (!value || typeof value !== "object") {
    throw new RequestError("Email payload is required.", 400);
  }

  const email = {
    bodyHtml: stringifyField(value.bodyHtml),
    bodyText: stringifyField(value.bodyText),
    cc: normalizeEmailAddressList(value.cc),
    conversationId: stringifyField(value.conversationId),
    dateTimeCreated: normalizeIsoDate(value.dateTimeCreated),
    from: normalizeEmailAddress(value.from),
    inReplyToMessageIds: normalizeMessageIdList(value.inReplyToMessageIds),
    internetHeaders: stringifyField(value.internetHeaders),
    internetMessageId: stringifyField(value.internetMessageId),
    itemId: stringifyField(value.itemId),
    mailboxEmail: normalizeMailboxEmail(value.mailboxEmail),
    normalizedSubject: stringifyField(value.normalizedSubject),
    referenceMessageIds: normalizeMessageIdList(value.referenceMessageIds),
    subject: stringifyField(value.subject),
    timeZone: normalizeTimeZone(value.timeZone),
    to: normalizeEmailAddressList(value.to),
  };

  if (!email.internetMessageId) {
    throw new RequestError("Open an existing received email, then choose a Halo ticket.", 400);
  }

  if (!email.bodyHtml && !email.bodyText) {
    throw new RequestError("Could not read an email body to attach.", 400);
  }

  return email;
}

function normalizeSendEmailPayload(value) {
  if (!value || typeof value !== "object") {
    throw new RequestError("Email payload is required.", 400);
  }

  const email = {
    bodyHtml: stringifyField(value.bodyHtml),
    bodyText: stringifyField(value.bodyText),
    cc: normalizeEmailAddressList(value.cc),
    conversationId: stringifyField(value.conversationId),
    dateTimeCreated: normalizeIsoDate(value.dateTimeCreated),
    from: normalizeEmailAddress(value.from),
    inReplyToMessageIds: normalizeMessageIdList(value.inReplyToMessageIds),
    internetHeaders: stringifyField(value.internetHeaders),
    internetMessageId: stringifyField(value.internetMessageId),
    itemId: stringifyField(value.itemId),
    mailboxEmail: normalizeMailboxEmail(value.mailboxEmail),
    normalizedSubject: stringifyField(value.normalizedSubject),
    referenceMessageIds: normalizeMessageIdList(value.referenceMessageIds),
    subject: stringifyField(value.subject),
    timeZone: normalizeTimeZone(value.timeZone),
    to: normalizeEmailAddressList(value.to),
  };

  if (!email.bodyHtml && !email.bodyText) {
    throw new RequestError("Could not read an email body to attach.", 400);
  }

  if (!email.inReplyToMessageIds.length && !email.conversationId) {
    throw new RequestError("No mapped reply identifiers were available.", 400);
  }

  email.internetMessageId = email.internetMessageId || buildSyntheticMessageId(email);
  return email;
}

function buildSyntheticMessageId(email) {
  const stableKey = email.itemId ? hashStableValue(email.itemId) : buildOutgoingBodyHash(email);
  return `<halo-outlook-${stableKey}@local>`;
}

function hashStableValue(value) {
  return crypto.createHash("sha256").update(String(value)).digest("hex").slice(0, 32);
}

function buildOutgoingBodyHash(email) {
  return crypto
    .createHash("sha256")
    .update(
      JSON.stringify({
        mailboxEmail: email.mailboxEmail,
        conversationId: email.conversationId,
        inReplyToMessageIds: email.inReplyToMessageIds,
        subject: email.subject,
        bodyHtml: email.bodyHtml,
        bodyText: email.bodyText,
      })
    )
    .digest("hex")
    .slice(0, 32);
}

function normalizeMailboxEmail(value) {
  return stringifyField(value).toLowerCase();
}

function normalizeTimeZone(value) {
  const timeZone = stringifyField(value);

  if (!timeZone) {
    return "";
  }

  try {
    Intl.DateTimeFormat("en-GB", { timeZone }).format(new Date());
    return timeZone;
  } catch {
    return "";
  }
}

function normalizeMessageIdList(value) {
  if (!Array.isArray(value)) {
    const singleValue = normalizeMessageId(value);
    return singleValue ? [singleValue] : [];
  }

  const seen = new Set();
  const messageIds = [];

  value.forEach((entry) => {
    const messageId = normalizeMessageId(entry);
    const key = normalizeMessageIdKey(messageId);

    if (messageId && key && !seen.has(key)) {
      seen.add(key);
      messageIds.push(messageId);
    }
  });

  return messageIds;
}

function normalizeMessageId(value) {
  return stringifyField(value);
}

function normalizeMessageIdKey(value) {
  return normalizeMessageId(value).toLowerCase();
}

function storeConversationMapping({ email, includeThreadMessageIds = false, ticketId, ticketNumber }) {
  const mailboxEmail = normalizeMailboxEmail(email.mailboxEmail);

  if (!mailboxEmail || !email.internetMessageId) {
    return null;
  }

  let mapping =
    getMappingByMessageId(mailboxEmail, email.internetMessageId) ||
    getMappingByConversationId(mailboxEmail, email.conversationId);
  const now = Date.now();

  if (!mapping) {
    mapping = {
      id: randomBase64Url(16),
      mailboxEmail,
      ticketId,
      ticketNumber: ticketNumber || String(ticketId),
      conversationId: email.conversationId || "",
      normalizedSubject: email.normalizedSubject || "",
      syncedMessageIds: new Set(),
      createdAt: now,
      updatedAt: now,
    };
    conversationMappings.set(mapping.id, mapping);
  }

  mapping.mailboxEmail = mailboxEmail;
  mapping.ticketId = ticketId;
  mapping.ticketNumber = ticketNumber || mapping.ticketNumber || String(ticketId);
  mapping.conversationId = email.conversationId || mapping.conversationId || "";
  mapping.normalizedSubject = email.normalizedSubject || mapping.normalizedSubject || "";
  mapping.updatedAt = now;
  markEmailSynced(mapping, email, { includeThreadMessageIds });

  return mapping;
}

function markEmailSynced(mapping, email, options = {}) {
  const mailboxEmail = normalizeMailboxEmail(mapping.mailboxEmail);
  const messageIds = [email.internetMessageId];

  if (options.includeThreadMessageIds) {
    messageIds.push(...email.inReplyToMessageIds, ...email.referenceMessageIds);
  }

  messageIds.forEach((messageId) => {
    const messageIdKey = normalizeMessageIdKey(messageId);
    if (!messageIdKey) {
      return;
    }

    mapping.syncedMessageIds.add(messageIdKey);
    messageMappingsByKey.set(getMessageMappingKey(mailboxEmail, messageIdKey), mapping.id);
  });

  if (email.conversationId) {
    mapping.conversationId = email.conversationId;
    conversationMappingsByKey.set(
      getConversationMappingKey(mailboxEmail, email.conversationId),
      mapping.id
    );
  }

  mapping.updatedAt = Date.now();
}

function findConversationMappingForEmail(email) {
  const mailboxEmail = normalizeMailboxEmail(email.mailboxEmail);

  if (!mailboxEmail) {
    return null;
  }

  const existingMessageMapping = getMappingByMessageId(mailboxEmail, email.internetMessageId);
  if (existingMessageMapping) {
    return {
      mapping: existingMessageMapping,
      status: "already-attached",
    };
  }

  const threadMessageIds = email.inReplyToMessageIds.concat(email.referenceMessageIds);
  for (const messageId of threadMessageIds) {
    const mapping = getMappingByMessageId(mailboxEmail, messageId);
    if (mapping) {
      return {
        mapping,
        status: "match",
      };
    }
  }

  const conversationMapping = getMappingByConversationId(mailboxEmail, email.conversationId);
  if (conversationMapping) {
    return {
      mapping: conversationMapping,
      status: "match",
    };
  }

  return null;
}

function getMappingByMessageId(mailboxEmail, messageId) {
  const messageIdKey = normalizeMessageIdKey(messageId);
  if (!mailboxEmail || !messageIdKey) {
    return null;
  }

  return conversationMappings.get(getMessageMappingId(mailboxEmail, messageIdKey)) || null;
}

function getMappingByConversationId(mailboxEmail, conversationId) {
  if (!mailboxEmail || !conversationId) {
    return null;
  }

  return conversationMappings.get(
    conversationMappingsByKey.get(getConversationMappingKey(mailboxEmail, conversationId))
  ) || null;
}

function getMessageMappingId(mailboxEmail, messageIdKey) {
  return messageMappingsByKey.get(getMessageMappingKey(mailboxEmail, messageIdKey));
}

function getMessageMappingKey(mailboxEmail, messageIdKey) {
  return `${mailboxEmail}|${messageIdKey}`;
}

function getConversationMappingKey(mailboxEmail, conversationId) {
  return `${mailboxEmail}|${conversationId}`;
}

function getMappingTicketLabel(mapping) {
  return mapping.ticketNumber || String(mapping.ticketId);
}

async function callHaloApiWithTicketContext(mapping, callback) {
  try {
    return await callback();
  } catch (error) {
    error.ticketNumber = getMappingTicketLabel(mapping);
    throw error;
  }
}

function normalizeEmailAddressList(value) {
  if (Array.isArray(value)) {
    return value.map((entry) => normalizeEmailAddress(entry)).filter(Boolean);
  }

  const singleValue = normalizeEmailAddress(value);
  return singleValue ? [singleValue] : [];
}

function normalizeEmailAddress(value) {
  if (!value) {
    return "";
  }

  if (typeof value === "string") {
    return value.trim();
  }

  if (typeof value === "object") {
    const displayName = stringifyField(value.displayName || value.name);
    const emailAddress = stringifyField(value.emailAddress || value.address);

    if (displayName && emailAddress) {
      return `${displayName} <${emailAddress}>`;
    }

    return displayName || emailAddress;
  }

  return "";
}

function normalizeIsoDate(value) {
  if (!value) {
    return new Date().toISOString();
  }

  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? new Date().toISOString() : date.toISOString();
}

function buildEmailActionPayload(ticketId, email, options = {}) {
  const body = getEmailBodyForAction(email, options.bodyMode || "trimmed");
  const htmlBody = body.bodyHtml || textToHtml(body.bodyText);
  const noteHtml = buildEmailNoteHtml(email, htmlBody);
  const note = buildEmailNoteText(email);
  const subject = email.subject || email.normalizedSubject || "(no subject)";
  const actionDatetime = getCurrentActionDateTime();

  return {
    ticket_id: ticketId,
    outcome: "Email",
    note,
    note_html: noteHtml,
    emailbody_html: noteHtml,
    emailsubject: subject,
    email_message_id: email.internetMessageId,
    actioninternetmessageid: email.internetMessageId,
    emailtolistall: email.to.join("; "),
    whowith: email.from,
    datetime: actionDatetime,
  };
}

function getEmailBodyForAction(email, bodyMode) {
  if (bodyMode === "full") {
    return {
      bodyHtml: email.bodyHtml || "",
      bodyText: email.bodyText || "",
    };
  }

  return trimEmailBody(email);
}

function getCurrentActionDateTime() {
  return new Date().toISOString();
}

function buildEmailNoteText(email) {
  const subject = email.subject || email.normalizedSubject || "(no subject)";
  const emailDate = formatEmailDate(email.dateTimeCreated, email.timeZone);
  return [
    "Outlook email attached to ticket.",
    `From: ${email.from || "(unknown)"}`,
    `To: ${email.to.join("; ") || "(none)"}`,
    email.cc.length ? `Cc: ${email.cc.join("; ")}` : "",
    `Subject: ${subject}`,
    `Email date: ${emailDate}`,
  ]
    .filter(Boolean)
    .join("\n");
}

function buildEmailNoteHtml(email, htmlBody) {
  const emailDate = formatEmailDate(email.dateTimeCreated, email.timeZone);
  const rows = [
    ["From", email.from || "(unknown)"],
    ["To", email.to.join("; ") || "(none)"],
    ["Cc", email.cc.join("; ")],
    ["Subject", email.subject || email.normalizedSubject || "(no subject)"],
    ["Email date", emailDate],
  ].filter((row) => row[1]);

  const metadataRows = rows
    .map(
      ([label, value]) =>
        `<tr><th style="text-align:left;padding:2px 12px 2px 0;">${escapeHtml(
          label
        )}</th><td style="padding:2px 0;">${escapeHtml(value)}</td></tr>`
    )
    .join("");

  return `<div><p><strong>Outlook email attached to ticket.</strong></p><table>${metadataRows}</table><hr>${htmlBody}</div>`;
}

function formatEmailDate(value, timeZone) {
  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return stringifyField(value);
  }

  const formatOptions = {
    day: "2-digit",
    hour: "2-digit",
    hour12: false,
    minute: "2-digit",
    month: "2-digit",
    second: "2-digit",
    timeZoneName: "short",
    year: "numeric",
  };

  try {
    const formatter = timeZone
      ? new Intl.DateTimeFormat("en-GB", { ...formatOptions, timeZone })
      : new Intl.DateTimeFormat("en-GB", formatOptions);
    return formatter.format(date);
  } catch {
    return date.toISOString();
  }
}

function trimEmailBody(email) {
  const bodyHtml = email.bodyHtml ? trimQuotedHtml(email.bodyHtml) || email.bodyHtml : "";
  const bodyText = email.bodyText ? trimQuotedText(email.bodyText) || email.bodyText : "";

  return {
    bodyHtml,
    bodyText,
  };
}

function trimQuotedHtml(value) {
  const trimIndex = getFirstUsableIndex([
    getPatternIndex(value, /<blockquote\b/i),
    getPatternIndex(value, /<[^>]+\bclass=["'][^"']*(?:gmail_quote|moz-cite-prefix|yahoo_quoted)[^"']*["'][^>]*>/i),
    getPatternIndex(value, /\bOn\s+[\s\S]{1,500}?\s+wrote:/i),
    getOutlookHeaderIndex(value),
  ]);

  const trimmed = trimIndex >= 0 ? value.slice(0, trimIndex) : value;
  return hasMeaningfulHtml(trimmed) ? trimmed.trim() : "";
}

function trimQuotedText(value) {
  const trimIndex = getFirstUsableIndex([
    getPatternIndex(value, /^On .+ wrote:$/im),
    getPatternIndex(value, /^-{2,}\s*Original Message\s*-{2,}$/im),
    getOutlookHeaderIndex(value),
  ]);

  const trimmed = trimIndex >= 0 ? value.slice(0, trimIndex) : value;
  return trimmed.trim() ? trimmed.trim() : "";
}

function getFirstUsableIndex(indexes) {
  return indexes.filter((index) => index >= 0).sort((left, right) => left - right)[0] ?? -1;
}

function getPatternIndex(value, pattern) {
  const match = pattern.exec(value);
  return match ? match.index : -1;
}

function getOutlookHeaderIndex(value) {
  const match = /(?:^|[\r\n]|<[^>]+>)\s*(?:<b>|<strong>)?From:(?:<\/b>|<\/strong>)?/i.exec(
    value
  );

  if (!match) {
    return -1;
  }

  const headerBlock = value.slice(match.index, match.index + 1500);
  if (
    /(?:^|[\r\n]|<[^>]+>)\s*(?:<b>|<strong>)?Sent:/i.test(headerBlock) &&
    /(?:^|[\r\n]|<[^>]+>)\s*(?:<b>|<strong>)?To:/i.test(headerBlock) &&
    /(?:^|[\r\n]|<[^>]+>)\s*(?:<b>|<strong>)?Subject:/i.test(headerBlock)
  ) {
    return match.index;
  }

  return -1;
}

function hasMeaningfulHtml(value) {
  return stripHtml(value).trim().length > 0;
}

function stripHtml(value) {
  return String(value)
    .replace(/<style\b[\s\S]*?<\/style>/gi, " ")
    .replace(/<script\b[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/\s+/g, " ");
}

function textToHtml(value) {
  return `<div>${escapeHtml(value).replace(/\r?\n/g, "<br>")}</div>`;
}

function getCreatedActionId(payload) {
  const action = Array.isArray(payload) ? payload[0] : payload;
  return stringifyField(getFirstField(action || {}, ["id", "action_id", "actionid"]));
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normalizeTickets(payload, currentAgentId) {
  let tickets = getTicketArray(payload).filter((ticket) => isOpenTicket(ticket));

  if (currentAgentId) {
    tickets = tickets.filter((ticket) => isAssignedToAgent(ticket, currentAgentId));
  }

  return tickets
    .map((ticket) => toTicketSummary(ticket))
    .filter((ticket) => ticket.id || ticket.ticketNumber || ticket.summary);
}

function getTicketArray(payload) {
  if (Array.isArray(payload)) {
    return payload;
  }

  if (!payload || typeof payload !== "object") {
    return [];
  }

  const possibleKeys = ["tickets", "Tickets", "items", "Items", "results", "Results", "data", "Data"];
  for (const key of possibleKeys) {
    if (Array.isArray(payload[key])) {
      return payload[key];
    }
  }

  return [];
}

function isOpenTicket(ticket) {
  const closedValues = [
    ticket.closed,
    ticket.is_closed,
    ticket.isclosed,
    ticket.isClosed,
    ticket.status && ticket.status.closed,
    ticket.status && ticket.status.is_closed,
    ticket.status && ticket.status.isClosed,
  ];

  if (closedValues.some((value) => value === true || value === 1 || value === "true")) {
    return false;
  }

  const statusText = getTicketStatus(ticket).toLowerCase();
  if (!statusText) {
    return true;
  }

  return !["closed", "resolved", "complete", "completed", "cancelled", "canceled"].some((word) =>
    statusText.includes(word)
  );
}

function isAssignedToAgent(ticket, currentAgentId) {
  const current = String(currentAgentId);
  const candidates = [
    ticket.agent_id,
    ticket.agentid,
    ticket.agentId,
    ticket.assigned_agent_id,
    ticket.assignedagentid,
    ticket.assignedAgentId,
    ticket.owner_id,
    ticket.ownerid,
    ticket.ownerId,
    ticket.agent && ticket.agent.id,
    ticket.assigned_agent && ticket.assigned_agent.id,
    ticket.assignedAgent && ticket.assignedAgent.id,
  ]
    .filter((value) => value !== undefined && value !== null && value !== "")
    .map((value) => String(value));

  if (!candidates.length) {
    return true;
  }

  return candidates.includes(current);
}

function toTicketSummary(ticket) {
  return {
    id: stringifyField(getFirstField(ticket, ["id", "ticket_id", "ticketid", "ticketId"])),
    ticketNumber: stringifyField(
      getFirstField(ticket, ["ticketnumber", "ticket_number", "number", "ref", "reference"])
    ),
    summary: stringifyField(
      getFirstField(ticket, ["summary", "title", "subject", "details", "description"])
    ),
    status: getTicketStatus(ticket),
    client: stringifyField(
      getFirstField(ticket, ["client_name", "clientname", "clientName", "customer_name"])
    ),
    agent: stringifyField(getFirstField(ticket, ["agent_name", "agentname", "agentName"])),
  };
}

function getTicketStatus(ticket) {
  const status = getFirstField(ticket, ["status", "status_name", "statusname", "statusName"]);

  if (status && typeof status === "object") {
    return stringifyField(
      getFirstField(status, ["name", "label", "description", "status_name", "statusName"])
    );
  }

  return stringifyField(status);
}

function getFirstField(source, keys) {
  if (!source || typeof source !== "object") {
    return "";
  }

  for (const key of keys) {
    const value = source[key];
    if (value !== undefined && value !== null && value !== "") {
      return value;
    }
  }

  return "";
}

function stringifyField(value) {
  if (value === undefined || value === null) {
    return "";
  }

  return String(value).trim();
}

function readJsonBody(req, maxBytes = MAX_JSON_BODY_BYTES) {
  if (req.body && typeof req.body === "object") {
    const bodyBytes = Buffer.byteLength(JSON.stringify(req.body), "utf8");
    if (bodyBytes > maxBytes) {
      return Promise.reject(new RequestError("Request body is too large.", 413));
    }

    return Promise.resolve(req.body);
  }

  return new Promise((resolve, reject) => {
    let body = "";
    let rejected = false;

    req.on("data", (chunk) => {
      body += chunk;
      if (!rejected && Buffer.byteLength(body, "utf8") > maxBytes) {
        rejected = true;
        reject(new RequestError("Request body is too large.", 413));
        req.destroy();
      }
    });

    req.on("end", () => {
      if (!body) {
        resolve({});
        return;
      }

      try {
        resolve(JSON.parse(body));
      } catch (error) {
        reject(new Error("Request body must be valid JSON."));
      }
    });

    req.on("error", (error) => {
      if (!rejected) {
        reject(error);
      }
    });
  });
}

async function readResponseDetails(response, requestUrl) {
  const text = await response.text();
  const contentType = response.headers.get("content-type") || "";
  let payload = {};

  if (text) {
    try {
      payload = JSON.parse(text);
    } catch (error) {
      payload = {};
    }
  }

  return {
    bodyText: text,
    contentType,
    payload,
    requestUrl,
    status: response.status,
    statusText: response.statusText || "",
  };
}

function sendAuthResultPage(res, payload) {
  const safePayload = JSON.stringify({
    type: "halo-auth",
    ...payload,
  }).replace(/</g, "\\u003c");

  sendHtml(
    res,
    200,
    `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=Edge">
  <title>Halo login</title>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head>
<body>
  <p>Completing Halo login...</p>
  <script>
    (function () {
      var payload = ${safePayload};
      function writeFallback() {
        document.body.textContent = "";
        var message = document.createElement("p");
        message.textContent = payload.message || "";
        document.body.appendChild(message);
        if (payload.error) {
          var error = document.createElement("p");
          error.textContent = payload.error;
          document.body.appendChild(error);
        }
      }
      function send() {
        if (window.Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
          Office.context.ui.messageParent(JSON.stringify(payload));
          return;
        }
        writeFallback();
      }
      if (window.Office && Office.onReady) {
        Office.onReady(send);
      } else {
        window.setTimeout(send, 250);
      }
    })();
  </script>
</body>
</html>`
  );
}

function normalizeHaloUrl(value) {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error("Enter your Halo URL.");
  }

  let url;
  try {
    url = new URL(value.trim());
  } catch (error) {
    throw new Error("Enter a valid Halo URL, including https://.");
  }

  if (url.protocol !== "https:") {
    throw new Error("Halo URL must use https://.");
  }

  return url.origin;
}

function normalizeClientId(value) {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error("Enter your Halo API application client ID.");
  }

  return value.trim();
}

function resolveHaloUrl(haloUrl, path) {
  if (/^https:\/\//i.test(path)) {
    return path;
  }

  return new URL(path.replace(/^\//, ""), `${haloUrl}/`).toString();
}

function getRequestUrl(req) {
  return new URL(req.originalUrl || req.url, getBaseUrl(req));
}

function getBaseUrl(req) {
  const forwardedProto = req.headers["x-forwarded-proto"];
  const proto = Array.isArray(forwardedProto) ? forwardedProto[0] : forwardedProto;
  return `${proto || "https"}://${req.headers.host || "localhost:3000"}`;
}

function annotateTokenPayload(payload) {
  const now = Date.now();
  const expiresIn = Number(payload.expires_in);

  return {
    ...payload,
    obtained_at: now,
    expires_at: Number.isFinite(expiresIn) && expiresIn > 0 ? now + expiresIn * 1000 : null,
  };
}

function isTokenExpired(tokenPayload) {
  return Boolean(tokenPayload.expires_at && Date.now() > tokenPayload.expires_at - 30 * 1000);
}

function encryptJson(value) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", encryptionKey, iv);
  const ciphertext = Buffer.concat([cipher.update(JSON.stringify(value), "utf8"), cipher.final()]);

  return {
    iv: iv.toString("base64"),
    ciphertext: ciphertext.toString("base64"),
    tag: cipher.getAuthTag().toString("base64"),
  };
}

function decryptJson(value) {
  const decipher = crypto.createDecipheriv(
    "aes-256-gcm",
    encryptionKey,
    Buffer.from(value.iv, "base64")
  );
  decipher.setAuthTag(Buffer.from(value.tag, "base64"));

  const plaintext = Buffer.concat([
    decipher.update(Buffer.from(value.ciphertext, "base64")),
    decipher.final(),
  ]).toString("utf8");

  return JSON.parse(plaintext);
}

function getSessionRecord(req) {
  return getSessionRecordBySessionId(getSessionIdFromRequest(req));
}

function getSessionRecordBySessionId(sessionId) {
  if (!sessionId) {
    return null;
  }

  const sessionHash = hashSessionId(sessionId);
  const record = sessions.get(sessionHash);

  if (!record) {
    return null;
  }

  if (record.expiresAt <= Date.now()) {
    sessions.delete(sessionHash);
    return null;
  }

  return record;
}

function createBackgroundSession(sessionHash, expiresAt) {
  const backgroundSessionId = randomBase64Url(32);
  backgroundSessions.set(hashBackgroundSessionId(backgroundSessionId), {
    sessionHash,
    expiresAt,
  });

  return backgroundSessionId;
}

function createBackgroundSessionForRequest(req) {
  const sessionId = getSessionIdFromRequest(req);
  const record = getSessionRecordBySessionId(sessionId);

  if (!sessionId || !record) {
    return "";
  }

  return createBackgroundSession(hashSessionId(sessionId), record.expiresAt);
}

function getBackgroundSessionRecord(backgroundSessionId) {
  const backgroundSession = backgroundSessions.get(hashBackgroundSessionId(backgroundSessionId));

  if (!backgroundSession) {
    return null;
  }

  if (backgroundSession.expiresAt <= Date.now()) {
    backgroundSessions.delete(hashBackgroundSessionId(backgroundSessionId));
    return null;
  }

  const record = sessions.get(backgroundSession.sessionHash);
  if (!record || record.expiresAt <= Date.now()) {
    backgroundSessions.delete(hashBackgroundSessionId(backgroundSessionId));
    if (record) {
      sessions.delete(backgroundSession.sessionHash);
    }
    return null;
  }

  return record;
}

function deleteBackgroundSessionsForSessionHash(sessionHash) {
  for (const [key, record] of backgroundSessions.entries()) {
    if (record.sessionHash === sessionHash) {
      backgroundSessions.delete(key);
    }
  }
}

function getSessionIdFromRequest(req) {
  return parseCookies(req.headers.cookie || "")[SESSION_COOKIE] || "";
}

function parseCookies(cookieHeader) {
  return cookieHeader.split(";").reduce((cookies, part) => {
    const index = part.indexOf("=");
    if (index < 0) {
      return cookies;
    }

    const key = part.slice(0, index).trim();
    const value = part.slice(index + 1).trim();
    if (key) {
      cookies[key] = decodeURIComponent(value);
    }

    return cookies;
  }, {});
}

function serializeSessionCookie(sessionId, maxAgeSeconds) {
  return `${SESSION_COOKIE}=${encodeURIComponent(sessionId)}; Path=/; HttpOnly; Secure; SameSite=None; Max-Age=${maxAgeSeconds}`;
}

function clearSessionCookie() {
  return `${SESSION_COOKIE}=; Path=/; HttpOnly; Secure; SameSite=None; Max-Age=0`;
}

function hashSessionId(sessionId) {
  return crypto.createHash("sha256").update(sessionId).digest("hex");
}

function hashBackgroundSessionId(backgroundSessionId) {
  return crypto.createHash("sha256").update(backgroundSessionId || "").digest("hex");
}

function randomBase64Url(byteLength) {
  return base64Url(crypto.randomBytes(byteLength));
}

function base64Url(buffer) {
  return buffer
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function cleanExpiredRecords() {
  const now = Date.now();
  deleteExpired(pendingStates, now);
  deleteExpired(handoffs, now);
  deleteExpired(sessions, now);
  deleteExpired(backgroundSessions, now);
}

function deleteExpired(map, now) {
  for (const [key, record] of map.entries()) {
    if (record.expiresAt <= now) {
      map.delete(key);
    }
  }
}

function sendJson(res, status, body) {
  res.status(status).json(body);
}

function sendHtml(res, status, html) {
  res.status(status).set("Content-Type", "text/html; charset=utf-8").send(html);
}

function publicError(error) {
  return error && error.message ? error.message : "Unexpected Halo auth error.";
}

function publicDebug(error) {
  if (error instanceof HttpError) {
    return error.debug;
  }

  if (error instanceof RequestError) {
    return error.debug;
  }

  return null;
}

function getErrorStatus(error, fallbackStatus) {
  return error instanceof RequestError ? error.status : fallbackStatus;
}

function getErrorTicketNumber(error) {
  return error && error.ticketNumber ? error.ticketNumber : "";
}

function safeResponseError(responseDetails) {
  const payload = responseDetails.payload || {};
  const payloadError = payload.error_description || payload.error || payload.message;
  if (payloadError) {
    return String(payloadError).slice(0, 500);
  }

  const bodyExcerpt = getBodyExcerpt(responseDetails.bodyText);
  if (bodyExcerpt) {
    return bodyExcerpt;
  }

  return "empty response body";
}

function getBodyExcerpt(bodyText) {
  return bodyText ? bodyText.replace(/\s+/g, " ").trim().slice(0, 500) : "";
}

function isUnauthorizedError(error) {
  return error instanceof HttpError && error.status === 401;
}

class HttpError extends Error {
  constructor(message, status, debug) {
    super(message);
    this.status = status;
    this.debug = debug;
  }

  static fromResponse(messagePrefix, phase, responseDetails) {
    const statusLabel = `${responseDetails.status}${responseDetails.statusText ? ` ${responseDetails.statusText}` : ""}`;
    const responseError = safeResponseError(responseDetails);

    return new HttpError(`${messagePrefix}: HTTP ${statusLabel} - ${responseError}`, responseDetails.status, {
      bodyExcerpt: getBodyExcerpt(responseDetails.bodyText),
      contentType: responseDetails.contentType || "(none)",
      endpoint: responseDetails.requestUrl,
      phase,
      status: responseDetails.status,
      statusText: responseDetails.statusText || "",
    });
  }
}

class RequestError extends Error {
  constructor(message, status, debug) {
    super(message);
    this.status = status;
    this.debug = debug || null;
  }
}

module.exports = {
  registerHaloAuthRoutes,
};
