const crypto = require("crypto");
const { registerBugReportRoutes } = require("./bugReports");
const {
  decryptStagedAttachment,
  getEmailAttachmentPreparationStatus,
  resolveEmailAttachments,
  stageEmailAttachmentItem,
  startEmailAttachmentPrefetch,
} = require("./emailAttachments");
const { lookupInlineImages, prefetchInlineImages, resolveInlineImages } = require("./inlineImages");
const { createMicrosoftAuthVerifier, getMicrosoftAuthConfig } = require("./microsoftAuth");
const { createTokenCrypto } = require("./tokenCrypto");
const {
  buildTicketPayload,
  getCreatedTicket,
  hydrateTicketCreationFieldOptions,
  normalizeLookupResults,
  normalizeRequesters,
  normalizeTicketTypeSchema,
  normalizeTicketTypes,
  validateCreationInput,
} = require("./ticketCreation");

const SESSION_COOKIE = "halo_session";
const AUTH_PATH = "/auth/authorize";
const TOKEN_PATH = "/auth/token";
const TEST_PATH = "/api/Tickets?count=1";
const TICKETS_COUNT = 50;
const TICKET_SEARCH_MAX_LENGTH = 200;
const DEFAULT_SCOPE = "all";
const STATE_TTL_MS = 10 * 60 * 1000;
const HANDOFF_TTL_MS = 5 * 60 * 1000;
const SESSION_TTL_MS = 24 * 60 * 60 * 1000;
const MAX_JSON_BODY_BYTES = 32 * 1024;
const MAX_TICKET_CREATION_INTENT_JSON_BODY_BYTES = 1024 * 1024;
const MAX_EMAIL_JSON_BODY_BYTES = 10 * 1024 * 1024;
const MAX_ATTACHMENT_UPLOAD_JSON_BODY_BYTES = 36 * 1024 * 1024;
const TICKET_CREATION_METADATA_TTL_MS = 30 * 60 * 1000;
const TICKET_CREATION_METADATA_STALE_MS = 24 * 60 * 60 * 1000;
const TICKET_CREATION_INTENT_TTL_MS = 30 * 24 * 60 * 60 * 1000;

const pendingStates = new Map();
const handoffs = new Map();
const explicitSendOperations = new Map();
const ticketCreationOperations = new Map();

let authStore = null;
let tokenCrypto = null;
let microsoftAuthVerifier = null;
let microsoftAuthConfig = null;
let cleaningEmailAttachmentPrefetches = false;
let cleaningExpiredRecords = false;

function registerHaloAuthRoutes(app, options = {}) {
  if (app.locals && app.locals.haloAuthRoutesRegistered) {
    return;
  }

  const env = options.env || process.env;
  const haloUrl = normalizeHaloUrl(env.HALO_URL, "HALO_URL");
  const clientId = normalizeClientId(env.HALO_CLIENT_ID, "HALO_CLIENT_ID");

  if (!options.store) {
    throw new Error("Halo auth routes require an initialized PostgreSQL store.");
  }
  authStore = options.store;
  tokenCrypto = options.tokenCrypto || tokenCrypto || createTokenCrypto(env);
  microsoftAuthVerifier =
    options.microsoftAuthVerifier ||
    microsoftAuthVerifier ||
    createMicrosoftAuthVerifier(options.microsoftAuth || {});
  microsoftAuthConfig = getMicrosoftAuthConfig(options.microsoftAuth || {});
  const logInlineImageDiagnostic = createInlineImageDiagnosticLogger(options.logger, env);
  const logEmailAttachmentDiagnostic = createEmailAttachmentDiagnosticLogger(options.logger, env);
  const logSendDiagnostic = createSendDiagnosticLogger(options.logger, env);

  if (app.locals) {
    app.locals.haloAuthRoutesRegistered = true;
  }

  setInterval(() => {
    void cleanExpiredRecords().catch((error) => {
      console.error("Scheduled PostgreSQL cleanup failed.", error);
    });
  }, 60 * 1000).unref();
  logSendDiagnostic("server", {
    elapsedMs: 0,
    outcome: "ok",
    stage: "diagnostics-ready",
  });

  app.post("/api/diagnostics/send-event", async (req, res) => {
    try {
      const body = await readJsonBody(req, 4096);
      logSendDiagnostic("client", body);
      sendJson(res, 200, { ok: true });
    } catch {
      sendJson(res, 200, { ok: false });
    }
  });

  app.get("/api/auth/config", (req, res) => {
    sendJson(res, 200, microsoftAuthConfig);
  });

  app.post("/api/auth/start", async (req, res) => {
    try {
      const user = await requireMicrosoftUser(req);
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
        userId: user.id,
        expiresAt: now + STATE_TTL_MS,
      });

      sendJson(res, 200, {
        dialogUrl: `${getBaseUrl(req)}/auth/start?state=${encodeURIComponent(state)}`,
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 400), { error: publicError(error) });
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
        userId: pending.userId,
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
      const user = await requireMicrosoftUser(req);

      if (!handoff || handoff.expiresAt <= Date.now()) {
        if (handoffCode) {
          handoffs.delete(handoffCode);
        }
        sendJson(res, 400, { error: "The Halo login handoff expired. Start login again." });
        return;
      }

      handoffs.delete(handoffCode);

      if (handoff.userId !== user.id) {
        sendJson(res, 403, {
          error: "The Halo login handoff belongs to a different Microsoft user.",
        });
        return;
      }

      const grant = await authStore.saveHaloGrant({
        userId: user.id,
        haloUrl: handoff.haloUrl,
        clientId: handoff.clientId,
        scope: handoff.scope,
        encryptedToken: handoff.encryptedToken,
      });
      const { backgroundSessionId, expiresAt } = await createSessionForGrant(res, user.id, grant);

      sendJson(res, 200, {
        authenticated: true,
        backgroundSessionId,
        expiresAt: new Date(expiresAt).toISOString(),
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 400), { error: publicError(error) });
    }
  });

  app.post("/api/auth/background-session", async (req, res) => {
    let record;
    try {
      record = await getOrCreateSessionRecord(req, res);
    } catch (error) {
      sendJson(res, getErrorStatus(error, 401), {
        ok: false,
        error: publicError(error),
      });
      return;
    }

    if (!record) {
      sendJson(res, 401, {
        ok: false,
        error: "No active Halo session.",
      });
      return;
    }

    sendJson(res, 200, {
      ok: true,
      backgroundSessionId: await createBackgroundSession(record.sessionHash, record.expiresAt),
      expiresAt: new Date(record.expiresAt).toISOString(),
    });
  });

  app.get("/api/auth/status", async (req, res) => {
    try {
      const record = await getOrCreateSessionRecord(req, res);
      const body = {
        authenticated: Boolean(record),
        haloUrl: record ? record.haloUrl : null,
        expiresAt: record ? new Date(record.expiresAt).toISOString() : null,
      };

      if (record) {
        body.backgroundSessionId = await createBackgroundSession(
          record.sessionHash,
          record.expiresAt
        );
      }

      sendJson(res, 200, body);
    } catch (error) {
      sendJson(res, 401, { authenticated: false, error: publicError(error) });
    }
  });

  app.get("/api/halo/ping", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);

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
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Halo API Auth failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.get("/api/halo/tickets", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);

      if (!record) {
        sendJson(res, 401, {
          ok: false,
          message: "Halo ticket list failed",
          error: "No active Halo session.",
        });
        return;
      }

      const ownership = getTicketOwnership(req);
      const lifecycle = getTicketLifecycle(req);
      const path = buildTicketsPath({ ownership, lifecycle });
      const payload = await callHaloApiWithRefresh(
        record,
        path,
        "tickets-list",
        "Halo ticket list failed"
      );
      const tickets = normalizeTicketsForLifecycle(payload, lifecycle);

      sendJson(res, 200, {
        ok: true,
        tickets,
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Halo ticket list failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.get("/api/halo/tickets/search", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);

      if (!record) {
        sendJson(res, 401, {
          ok: false,
          message: "Halo ticket search failed",
          error: "No active Halo session.",
        });
        return;
      }

      const query = getTicketSearchQuery(req);
      const ownership = getTicketOwnership(req);
      const lifecycle = getTicketLifecycle(req);
      const payload = await callHaloApiWithRefresh(
        record,
        buildTicketsPath({ query, ownership, lifecycle }),
        "tickets-search",
        "Halo ticket search failed"
      );
      const tickets = promoteExactTicketMatches(
        normalizeTicketsForLifecycle(payload, lifecycle),
        query
      );

      sendJson(res, 200, {
        ok: true,
        tickets,
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Halo ticket search failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.get("/api/halo/ticket-creation/types", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const refresh = getBooleanQueryParameter(req, "refresh");
      const result = await loadTicketCreationTypes(record, { refresh });
      sendJson(res, 200, { ok: true, ...result });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Halo ticket type discovery failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.get("/api/halo/ticket-creation/types/:typeId/schema", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const typeId = normalizePositiveInteger(
        req.params && req.params.typeId,
        "A valid Halo ticket type ID is required."
      );
      const refresh = getBooleanQueryParameter(req, "refresh");
      const typesResult = await loadTicketCreationTypes(record, { refresh: false });
      const type = typesResult.types.find((candidate) => Number(candidate.id) === typeId);
      if (!type) {
        throw new RequestError("This Halo ticket type is unavailable to the current agent.", 404);
      }
      const result = await loadTicketCreationSchema(record, type, { refresh });
      sendJson(res, 200, { ok: true, ...result });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Halo ticket field discovery failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.get("/api/halo/ticket-creation/requesters", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const url = getRequestUrl(req);
      const query = stringifyField(url.searchParams.get("query")).slice(0, 200);
      if (!query) {
        throw new RequestError("Enter an email address or requester name.", 400);
      }
      const requesters = await searchHaloRequesters(record, query);
      sendJson(res, 200, { ok: true, requesters });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Halo requester search failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.get("/api/halo/ticket-creation/lookups/:entity", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const entity = stringifyField(req.params && req.params.entity).toLowerCase();
      const endpointByEntity = {
        agent: "Agent",
        asset: "Asset",
        client: "Client",
        site: "Site",
        team: "Team",
        user: "Users",
      };
      if (!endpointByEntity[entity]) {
        throw new RequestError("This Halo lookup type is unsupported.", 400);
      }
      const query = stringifyField(getRequestUrl(req).searchParams.get("query")).slice(0, 200);
      if (!query) {
        throw new RequestError("Enter a Halo lookup search value.", 400);
      }
      const payload = await callHaloApiWithRefresh(
        record,
        `/api/${endpointByEntity[entity]}?search=${encodeURIComponent(
          query
        )}&count=20&includeinactive=false`,
        "ticket-creation-lookup",
        "Halo lookup failed"
      );
      sendJson(res, 200, {
        ok: true,
        results: normalizeLookupResults(payload, entity),
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Halo lookup failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.post("/api/halo/ticket-creation/from-email", async (req, res) => {
    let operationId = "";
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      operationId = normalizeOpaqueIdentifier(
        body.operationId,
        "A valid ticket creation operation ID is required."
      );
      const email = normalizeEmailPayload(body);
      const stored = await getStoredTicketCreationIntent(record, operationId);
      const intent =
        stored && stored.record.ticketId
          ? stored.intent
          : await normalizeTicketCreationIntent(record, body.creation || body);
      if (!stored || !stored.record.ticketId) {
        await saveTicketCreationIntent(record, operationId, intent);
      }
      const result = await runExclusiveTicketCreation(record, operationId, () =>
        createTicketAndEmailAction(record, operationId, intent, email, body, {
          bodyMode: "full",
          logEmailAttachmentDiagnostic,
          logInlineImageDiagnostic,
        })
      );
      sendJson(res, 201, {
        ok: true,
        backgroundSessionId: (await createBackgroundSessionForRequest(req)) || undefined,
        ...result,
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        status: getAttachmentFailureStatus(error),
        message: "Halo ticket creation failed",
        error: publicError(error),
        debug: publicDebug(error),
        operationId: operationId || undefined,
        ticketNumber: getErrorTicketNumber(error) || undefined,
      });
    }
  });

  app.post("/api/halo/ticket-creation/intents", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const body = await readJsonBody(req, MAX_TICKET_CREATION_INTENT_JSON_BODY_BYTES);
      const operationId = normalizeOpaqueIdentifier(
        body.operationId,
        "A valid ticket creation operation ID is required."
      );
      const existing = await getStoredTicketCreationIntent(record, operationId);
      if (existing && existing.record.ticketId) {
        sendJson(res, 200, {
          ok: true,
          operationId,
          status: existing.record.status,
          ticketNumber: existing.record.ticketNumber || undefined,
          expiresAt: new Date(existing.record.expiresAt).toISOString(),
        });
        return;
      }
      const intent = await normalizeTicketCreationIntent(record, body);
      const saved = await saveTicketCreationIntent(record, operationId, intent);
      sendJson(res, 201, {
        ok: true,
        operationId,
        expiresAt: new Date(saved.expiresAt).toISOString(),
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Ticket creation intent could not be saved",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  const updateTicketCreationIntentRoute = async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const operationId = normalizeOpaqueIdentifier(
        req.params && req.params.operationId,
        "A valid ticket creation operation ID is required."
      );
      const existing = await getStoredTicketCreationIntent(record, operationId);
      if (!existing || existing.record.status !== "pending") {
        throw new RequestError("This ticket creation intent can no longer be edited.", 409);
      }
      const body = await readJsonBody(req, MAX_TICKET_CREATION_INTENT_JSON_BODY_BYTES);
      const nextIntent = await normalizeTicketCreationIntent(record, {
        ...existing.intent,
        ...body,
      });
      await saveTicketCreationIntent(record, operationId, nextIntent);
      sendJson(res, 200, { ok: true, operationId });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Ticket creation intent could not be updated",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  };
  if (typeof app.patch === "function") {
    app.patch("/api/halo/ticket-creation/intents/:operationId", updateTicketCreationIntentRoute);
  }
  app.post("/api/halo/ticket-creation/intents/:operationId", updateTicketCreationIntentRoute);

  app.delete("/api/halo/ticket-creation/intents/:operationId", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const operationId = normalizeOpaqueIdentifier(
        req.params && req.params.operationId,
        "A valid ticket creation operation ID is required."
      );
      const existing = await getStoredTicketCreationIntent(record, operationId);
      if (existing && existing.intent.emailAttachmentPrefetchKey) {
        const staged = await authStore.getEmailAttachmentPrefetch(
          existing.intent.emailAttachmentPrefetchKey,
          { haloTenant: record.haloUrl, userId: record.userId }
        );
        if (staged && staged.stagingVersion === 2) {
          await authStore.deleteEmailAttachmentPrefetch(staged.prefetchKey);
        } else {
          await authStore.markEmailAttachmentPrefetchForCleanup(
            existing.intent.emailAttachmentPrefetchKey
          );
        }
      }
      const deleted = await authStore.deleteTicketCreationIntent(operationId, {
        haloTenant: record.haloUrl,
        userId: record.userId,
      });
      sendJson(res, 200, { ok: true, status: deleted ? "removed" : "not-active" });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Ticket creation intent could not be removed",
        error: publicError(error),
      });
    }
  });

  app.post("/api/halo/ticket-creation/intents/:operationId/send", async (req, res) => {
    let ticketNumber = "";
    try {
      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      const record =
        (await getSessionRecord(req)) ||
        (await getBackgroundSessionRecord(body.backgroundSessionId)) ||
        (await getBearerGrantRecord(req));
      if (!record) {
        sendJson(res, 200, { ok: true, status: "no-session" });
        return;
      }
      const operationId = normalizeOpaqueIdentifier(
        req.params && req.params.operationId,
        "A valid ticket creation operation ID is required."
      );
      const stored = await getStoredTicketCreationIntent(record, operationId);
      if (!stored) {
        throw new RequestError(
          "The saved ticket creation details expired. Reopen the Halo add-in and review them.",
          409
        );
      }
      ticketNumber = stored.record.ticketNumber || "";
      const email = normalizeExplicitSendEmailPayload(
        {
          ...body,
          composeAttachId: operationId,
        },
        record.userEmail
      );
      const finalIntent = stored.record.ticketId
        ? stored.intent
        : await prepareFinalComposeIntent(record, stored.intent, email);
      const result = await runExclusiveTicketCreation(record, operationId, () =>
        createTicketAndEmailAction(record, operationId, finalIntent, email, body, {
          addWarningFooter: true,
          bodyMode: "full",
          logEmailAttachmentDiagnostic,
          logInlineImageDiagnostic,
        })
      );
      sendJson(res, 200, { ok: true, ...result });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        status: getAttachmentFailureStatus(error),
        message: "Composed email ticket creation failed",
        error: publicError(error),
        debug: publicDebug(error),
        ticketNumber: getErrorTicketNumber(error) || ticketNumber || undefined,
      });
    }
  });

  app.post("/api/halo/inline-images/lookup", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }

      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      const result = await lookupInlineImages(authStore, record.haloUrl, body.images);
      sendJson(res, 200, { ok: true, ...result });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Inline image lookup failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.post("/api/halo/inline-images/prefetch", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }

      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      const ticketId = normalizePositiveInteger(
        body.ticketId,
        "A valid Halo ticket ID is required."
      );
      const composeOperationId = normalizeOpaqueIdentifier(
        body.composeOperationId,
        "A valid compose operation ID is required."
      );
      if (!/^[a-f0-9]{64}$/i.test(String(body.inlineImageFingerprint || ""))) {
        throw new RequestError("A valid inline image fingerprint is required.", 400);
      }
      const result = await prefetchInlineImages({
        composeOperationId,
        haloTenant: record.haloUrl,
        input: body,
        showForUsers: normalizeActionMode(body.actionMode) !== "private-note",
        store: authStore,
        ticketId,
        uploadImage: (image) =>
          uploadHaloInlineImage(record, image, normalizeActionMode(body.actionMode)),
        onDiagnostic: logInlineImageDiagnostic,
      });
      sendJson(res, 200, { ok: true, ...result });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Inline image preparation failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.post("/api/halo/email/recover-mapping", async (req, res) => {
    const startedAt = Date.now();
    try {
      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      const record = await getEmailAttachmentRequestRecord(req, body);
      if (!record) {
        sendJson(res, 200, { ok: true, status: "no-session" });
        return;
      }
      const email = normalizeRecoveryEmailPayload(body, record.userEmail);
      const candidates = normalizeRecoveryComposeIds(body.composeAttachIds);
      const result = await findRecoveryMapping(record, email, candidates);
      logSendDiagnostic("server", {
        candidateCount: candidates.length,
        elapsedMs: Date.now() - startedAt,
        outcome: result ? "matched" : "no-match",
        stage: "recovery-complete",
      });
      if (!result) {
        sendJson(res, 200, { ok: true, status: "no-match" });
        return;
      }
      sendJson(res, 200, {
        ok: true,
        status: "matched",
        actionMode: normalizeActionMode(result.match.mapping.actionMode),
        candidateIndex: result.candidateIndex,
        ticketId: result.match.mapping.ticketId,
        ticketNumber: result.match.mapping.ticketNumber,
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 400), {
        ok: false,
        status: "no-match",
        error: publicError(error),
      });
    }
  });

  app.post("/api/halo/email-attachments/prefetch/start", async (req, res) => {
    const preparationStartedAt = Date.now();
    try {
      logEmailAttachmentDiagnostic("prefetch-started", {
        durationMs: 0,
        outcome: "started",
        stage: "prefetch-start",
      });
      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      const record = await getEmailAttachmentRequestRecord(req, body);
      if (!record) {
        logEmailAttachmentDiagnostic("prefetch-failed", {
          durationMs: Date.now() - preparationStartedAt,
          outcome: "no-session",
          stage: "authentication",
        });
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }

      let ticketId = body.ticketId
        ? normalizePositiveInteger(body.ticketId, "A valid Halo ticket ID is required.")
        : 0;
      let ticketNumber = stringifyField(body.ticketNumber);
      let resolvedMapping = null;
      const creationOperationId = body.creationOperationId
        ? normalizeOpaqueIdentifier(
            body.creationOperationId,
            "A valid ticket creation operation ID is required."
          )
        : "";
      if (creationOperationId) {
        const creationIntent = await authStore.getTicketCreationIntent(creationOperationId, {
          haloTenant: record.haloUrl,
          userId: record.userId,
        });
        if (
          !creationIntent ||
          !["pending", "ticket-created", "partial-failure"].includes(creationIntent.status)
        ) {
          throw new RequestError("The ticket creation intent is unavailable.", 409);
        }
        ticketId = 0;
        ticketNumber = "";
      } else if (!ticketId) {
        const email = normalizeEmailMatchPayload(body);
        const match = await findConversationOrRecoveredMapping(record, email, body);
        logEmailAttachmentDiagnostic("mapping-lookup", {
          conversationIdentifierCount: email.conversationId ? 1 : 0,
          inReplyToCount: email.inReplyToMessageIds.length,
          outcome: match ? "matched" : "no-match",
          stage: "mapping",
        });
        if (!match || match.status === "already-attached") {
          logEmailAttachmentDiagnostic("prefetch-completed", {
            durationMs: Date.now() - preparationStartedAt,
            outcome: match && match.status === "already-attached" ? "already-attached" : "no-match",
            stage: "mapping",
          });
          sendJson(res, 200, {
            ok: true,
            status: match && match.status === "already-attached" ? "already-attached" : "no-match",
          });
          return;
        }
        ticketId = match.mapping.ticketId;
        ticketNumber = match.mapping.ticketNumber;
        resolvedMapping = match.mapping;
      }

      const result = await startEmailAttachmentPrefetch({
        attachmentFingerprint: body.emailAttachmentFingerprint,
        descriptors: body.emailAttachments,
        draftItemId: body.draftItemId,
        haloTenant: record.haloUrl,
        operationId: creationOperationId || body.operationId,
        store: authStore,
        ticketId,
        userId: record.userId,
      });
      await cleanupStoredHaloAttachments(record, result.prefetchKey, result.removedUploads);
      logEmailAttachmentDiagnostic("prefetch-completed", {
        count: result.aggregate.selected,
        durationMs: Date.now() - preparationStartedAt,
        failed: result.aggregate.failed,
        outcome: result.pendingAttachmentKeys.length ? "pending" : "ready",
        pending: result.aggregate.pending,
        prepared: result.aggregate.prepared,
        selected: result.aggregate.selected,
        stage: "prefetch-start",
      });
      sendJson(res, 200, {
        aggregate: result.aggregate,
        ok: true,
        pendingAttachmentKeys: result.pendingAttachmentKeys,
        prefetchKey: result.prefetchKey,
        stagingVersion: result.stagingVersion,
        status: result.pendingAttachmentKeys.length ? "pending" : "ready",
        actionMode: normalizeActionMode(
          resolvedMapping ? resolvedMapping.actionMode : body.actionMode
        ),
        ticketId: String(ticketId),
        ticketNumber: ticketNumber || String(ticketId),
      });
    } catch (error) {
      logEmailAttachmentDiagnostic("prefetch-failed", {
        durationMs: Date.now() - preparationStartedAt,
        failed: 1,
        outcome: getEmailAttachmentDiagnosticOutcome(error),
        stage: "prefetch-start",
      });
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Email attachment preparation failed",
        error: publicError(error),
      });
    }
  });

  app.post("/api/halo/email-attachments/prefetch/:prefetchKey/items", async (req, res) => {
    const uploadStartedAt = Date.now();
    try {
      const body = await readJsonBody(req, MAX_ATTACHMENT_UPLOAD_JSON_BODY_BYTES);
      const record = await getEmailAttachmentRequestRecord(req, body);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const prefetchKey = normalizeOpaqueIdentifier(
        req.params && req.params.prefetchKey,
        "A valid email attachment prefetch key is required."
      );
      const prepared = await authStore.getEmailAttachmentPrefetch(prefetchKey, {
        haloTenant: record.haloUrl,
        userId: record.userId,
      });
      if (!prepared) {
        throw new RequestError("The email attachment preparation was not found.", 404);
      }
      const result = await stageEmailAttachmentItem({
        attachmentKey: body.attachmentKey,
        contentBase64: body.contentBase64,
        contentFormat: body.contentFormat,
        contentSha256: body.contentSha256,
        haloTenant: record.haloUrl,
        prefetchKey,
        store: authStore,
        ticketId: prepared.ticketId,
        tokenCrypto,
        userId: record.userId,
      });
      logEmailAttachmentDiagnostic("preparation-completed", {
        durationMs: Date.now() - uploadStartedAt,
        outcome: result.status,
        prepared: result.status === "prepared" ? 1 : 0,
        stage: "prefetch-item",
      });
      sendJson(res, 200, { ok: true, status: result.status });
    } catch (error) {
      logEmailAttachmentDiagnostic("preparation-failed", {
        durationMs: Date.now() - uploadStartedAt,
        failed: 1,
        outcome: getEmailAttachmentDiagnosticOutcome(error),
        stage: "prefetch-item",
      });
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Email attachment preparation failed",
        error: publicError(error),
      });
    }
  });

  app.get("/api/halo/email-attachments/prefetch/:prefetchKey/status", async (req, res) => {
    try {
      const record = await getEmailAttachmentRequestRecord(req, req.query || {});
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const prefetchKey = normalizeOpaqueIdentifier(
        req.params && req.params.prefetchKey,
        "A valid email attachment prefetch key is required."
      );
      const result = await getEmailAttachmentPreparationStatus({
        haloTenant: record.haloUrl,
        prefetchKey,
        store: authStore,
        userId: record.userId,
      });
      sendJson(res, 200, { ok: true, ...result });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        status: error && error.code === "attachments-not-ready" ? error.code : "failed",
        error: publicError(error),
      });
    }
  });

  app.delete("/api/halo/email-attachments/prefetch/:prefetchKey", async (req, res) => {
    try {
      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      const record = await getEmailAttachmentRequestRecord(req, body);
      if (!record) {
        sendJson(res, 401, { ok: false, error: "No active Halo session." });
        return;
      }
      const prefetchKey = normalizeOpaqueIdentifier(
        req.params && req.params.prefetchKey,
        "A valid email attachment prefetch key is required."
      );
      const prepared = await authStore.getEmailAttachmentPrefetch(prefetchKey, {
        haloTenant: record.haloUrl,
        userId: record.userId,
      });
      if (!prepared || prepared.status === "consumed") {
        sendJson(res, 200, { ok: true, status: "not-active" });
        return;
      }
      if (prepared.stagingVersion === 2) {
        await authStore.deleteEmailAttachmentPrefetch(prefetchKey);
        sendJson(res, 200, { ok: true, status: "cancelled" });
        return;
      }
      await authStore.markEmailAttachmentPrefetchForCleanup(prefetchKey);
      const cleaned = await cleanupStoredHaloAttachments(
        record,
        prefetchKey,
        prepared.items
          .filter((item) => item.haloAttachmentId)
          .map((item) => ({ attachmentKey: item.attachmentKey, id: item.haloAttachmentId }))
      );
      if (cleaned) {
        await authStore.deleteEmailAttachmentPrefetch(prefetchKey);
      }
      sendJson(res, 200, { ok: true, status: cleaned ? "cancelled" : "cleanup-pending" });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Email attachment cleanup failed",
        error: publicError(error),
      });
    }
  });

  app.post("/api/halo/email/match", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);
      if (!record) {
        sendJson(res, 401, {
          ok: false,
          message: "Email mapping lookup failed",
          error: "No active Halo session.",
        });
        return;
      }

      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      const email = normalizeEmailMatchPayload(body);
      const recovery = await findConversationOrRecoveredMappingResult(record, email, body);
      const match = recovery && recovery.match;

      if (!match) {
        sendJson(res, 200, { ok: true, status: "no-match" });
        return;
      }

      if (recovery.candidateIndex >= 0) {
        await backfillRecoveredConversationMapping(match.mapping, email);
      }

      sendJson(res, 200, {
        ok: true,
        status: match.status === "already-attached" ? "already-attached" : "matched",
        ticketId: match.mapping.ticketId,
        ticketNumber: match.mapping.ticketNumber,
        actionMode: normalizeActionMode(match.mapping.actionMode),
        message:
          match.status === "already-attached"
            ? `This email is already attached to ticket ${getMappingTicketLabel(match.mapping)}.`
            : undefined,
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        message: "Email mapping lookup failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.post("/api/halo/tickets/:ticketId/email", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);

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
      const existingMapping = await findConversationMappingForEmail(email);
      const isInitialChainAttach = !existingMapping;
      const preparedAction = await buildEmailActionWithInlineImages(record, ticketId, email, body, {
        bodyMode: isInitialChainAttach ? "full" : "trimmed",
        onDiagnostic: logInlineImageDiagnostic,
        onEmailAttachmentDiagnostic: logEmailAttachmentDiagnostic,
      });
      const actionCreationStart = Date.now();
      const payload = await createPreparedHaloAction(record, preparedAction, () =>
        callHaloApiWithRefresh(record, {
          body: [preparedAction.payload],
          method: "POST",
          path: "/api/Actions",
          phase: "email-attach",
          messagePrefix: "Halo email attach failed",
        })
      );
      preparedAction.timings.actionCreationMs = Date.now() - actionCreationStart;
      const actionId = getCreatedActionId(payload);
      await consumeEmailAttachmentPrefetch(preparedAction);

      await storeConversationMapping({
        email,
        includeThreadMessageIds: isInitialChainAttach,
        ticketId,
        ticketNumber,
      });

      sendJson(res, 200, {
        ok: true,
        status: getAttachedStatus(preparedAction),
        attachMode: isInitialChainAttach ? "full-chain" : "latest-reply",
        message:
          email.actionMode === "private-note"
            ? isInitialChainAttach
              ? "Full email chain attached as a private note"
              : "Email attached as a private note"
            : isInitialChainAttach
              ? "Full email chain attached to Halo ticket"
              : "Email attached to Halo ticket",
        actionId: actionId || undefined,
        backgroundSessionId: (await createBackgroundSessionForRequest(req)) || undefined,
        emailAttachments: preparedAction.emailAttachments,
        inlineImages: preparedAction.inlineImages,
        inlineImageTimings: preparedAction.timings,
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

  app.post("/api/halo/tickets/:ticketId/sent-email", async (req, res) => {
    const sendStartedAt = Date.now();
    let ticketNumber = "";

    try {
      logSendDiagnostic("server", {
        elapsedMs: 0,
        outcome: "started",
        stage: "request-received",
      });
      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      logSendDiagnostic("server", {
        attachmentCount: getSendDiagnosticAttachmentCount(body),
        elapsedMs: Date.now() - sendStartedAt,
        includeAttachments: Boolean(body.includeEmailAttachments),
        inlineImageCount: Array.isArray(body.inlineImageRefs) ? body.inlineImageRefs.length : 0,
        outcome: "ok",
        stage: "payload-read",
      });
      const ticketId = getTicketIdFromRequest(req);
      ticketNumber = stringifyField(body.ticketNumber) || String(ticketId);
      let authSource = "cookie";
      let record = await getSessionRecord(req);
      if (!record) {
        authSource = "background-session";
        record = await getBackgroundSessionRecord(body.backgroundSessionId);
      }
      if (!record) {
        authSource = "microsoft-bearer";
        record = await getBearerGrantRecord(req);
      }

      if (!record) {
        logSendDiagnostic("server", {
          authSource: "none",
          elapsedMs: Date.now() - sendStartedAt,
          outcome: "no-session",
          stage: "authentication-complete",
        });
        sendJson(res, 200, {
          ok: true,
          status: "no-session",
          ticketId: String(ticketId),
          ticketNumber,
        });
        return;
      }

      logSendDiagnostic("server", {
        authSource,
        elapsedMs: Date.now() - sendStartedAt,
        outcome: "ok",
        stage: "authentication-complete",
      });
      const email = normalizeExplicitSendEmailPayload(body, record.userEmail);
      logSendDiagnostic("server", {
        elapsedMs: Date.now() - sendStartedAt,
        outcome: "ok",
        stage: "payload-validated",
      });
      const result = await runExclusiveExplicitSend(
        buildExplicitSendOperationKey(record, email),
        async () => {
          const existingMatch = await findConversationMappingForEmail(email);

          if (existingMatch && existingMatch.status === "already-attached") {
            logSendDiagnostic("server", {
              elapsedMs: Date.now() - sendStartedAt,
              outcome: "already-attached",
              stage: "idempotency-check",
            });
            await consumeEmailAttachmentPrefetchKey(body.emailAttachmentPrefetchKey);
            return {
              ok: true,
              status: "already-attached",
              ticketId: existingMatch.mapping.ticketId,
              ticketNumber: existingMatch.mapping.ticketNumber,
              message: `This email is already attached to ticket ${getMappingTicketLabel(
                existingMatch.mapping
              )}.`,
            };
          }

          if (body.emailAttachmentPrefetchKey) {
            const committedAttachmentOperation = await authStore.getEmailAttachmentPrefetch(
              body.emailAttachmentPrefetchKey,
              { haloTenant: record.haloUrl, ticketId, userId: record.userId }
            );
            if (
              committedAttachmentOperation &&
              (committedAttachmentOperation.status === "consumed" ||
                committedAttachmentOperation.haloActionId)
            ) {
              await consumeEmailAttachmentPrefetchKey(body.emailAttachmentPrefetchKey);
              await storeConversationMapping({
                email,
                includeThreadMessageIds: !existingMatch,
                ticketId,
                ticketNumber,
              });
              return {
                ok: true,
                status: "already-attached",
                ticketId: String(ticketId),
                ticketNumber,
                message: `This email is already attached to ticket ${ticketNumber}.`,
              };
            }
          }

          const isInitialChainAttach = !existingMatch;
          const selectedMapping = { ticketId, ticketNumber };
          logSendDiagnostic("server", {
            elapsedMs: Date.now() - sendStartedAt,
            outcome: "started",
            stage: "assets-prepare-start",
          });
          const preparedAction = await buildEmailActionWithInlineImages(
            record,
            ticketId,
            email,
            body,
            {
              addWarningFooter: true,
              bodyMode: "full",
              onDiagnostic: logInlineImageDiagnostic,
              onEmailAttachmentDiagnostic: logEmailAttachmentDiagnostic,
            }
          );
          logSendDiagnostic("server", {
            attachmentCount: preparedAction.emailAttachments.selected,
            elapsedMs: Date.now() - sendStartedAt,
            inlineImageCount: preparedAction.inlineImages.referenced,
            outcome: getAttachedStatus(preparedAction),
            stage: "assets-prepare-complete",
          });
          const actionCreationStart = Date.now();
          logSendDiagnostic("server", {
            elapsedMs: Date.now() - sendStartedAt,
            outcome: "started",
            stage: "halo-action-start",
          });
          const payload = await createPreparedHaloAction(record, preparedAction, () =>
            callHaloApiWithTicketContext(selectedMapping, () =>
              callHaloApiWithRefresh(record, {
                body: [preparedAction.payload],
                method: "POST",
                path: "/api/Actions",
                phase: "email-explicit-send-attach",
                messagePrefix: "Halo composed email attach failed",
              })
            )
          );
          preparedAction.timings.actionCreationMs = Date.now() - actionCreationStart;
          logSendDiagnostic("server", {
            elapsedMs: Date.now() - sendStartedAt,
            outcome: "ok",
            stage: "halo-action-complete",
          });
          const actionId = getCreatedActionId(payload);

          await consumeEmailAttachmentPrefetch(preparedAction);

          await storeConversationMapping({
            email,
            includeThreadMessageIds: isInitialChainAttach,
            ticketId,
            ticketNumber,
          });
          logSendDiagnostic("server", {
            elapsedMs: Date.now() - sendStartedAt,
            outcome: "ok",
            stage: "mapping-complete",
          });
          return {
            ok: true,
            status: getAttachedStatus(preparedAction),
            ticketId: String(ticketId),
            ticketNumber,
            message:
              email.actionMode === "private-note"
                ? `Sent email attached as a private note to ticket ${ticketNumber}.`
                : `Sent email added to Halo ticket ${ticketNumber}.`,
            actionId: actionId || undefined,
            emailAttachments: preparedAction.emailAttachments,
            inlineImages: preparedAction.inlineImages,
            inlineImageTimings: preparedAction.timings,
          };
        }
      );

      logSendDiagnostic("server", {
        elapsedMs: Date.now() - sendStartedAt,
        outcome: result.status || "ok",
        stage: "response-complete",
      });
      sendJson(res, 200, result);
    } catch (error) {
      logSendDiagnostic("server", {
        elapsedMs: Date.now() - sendStartedAt,
        errorKind: getSendDiagnosticErrorKind(error),
        outcome: "failed",
        stage: "response-complete",
        statusCode: getErrorStatus(error, 502),
      });
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        status: getAttachmentFailureStatus(error),
        message: "Composed email attach failed",
        error: publicError(error),
        debug: publicDebug(error),
        ticketNumber: getErrorTicketNumber(error) || ticketNumber,
      });
    }
  });

  app.post("/api/halo/email/auto-attach", async (req, res) => {
    try {
      const record = await getSessionOrBearerGrant(req);

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
      const match = await findConversationOrRecoveredMapping(record, email, body);

      if (!match) {
        sendJson(res, 200, {
          ok: true,
          status: "no-match",
        });
        return;
      }

      if (match.status === "already-attached") {
        await consumeEmailAttachmentPrefetchKey(body.emailAttachmentPrefetchKey);
        sendJson(res, 200, {
          ok: true,
          status: "already-attached",
          ticketId: match.mapping.ticketId,
          ticketNumber: match.mapping.ticketNumber,
          message: `This email is already attached to ticket ${getMappingTicketLabel(match.mapping)}.`,
        });
        return;
      }

      const actionMode = Object.prototype.hasOwnProperty.call(body, "actionMode")
        ? normalizeActionMode(body.actionMode)
        : normalizeActionMode(match.mapping.actionMode);
      body.actionMode = actionMode;
      email.actionMode = actionMode;

      if (body.emailAttachmentPrefetchKey) {
        const committedAttachmentOperation = await authStore.getEmailAttachmentPrefetch(
          body.emailAttachmentPrefetchKey,
          {
            haloTenant: record.haloUrl,
            ticketId: match.mapping.ticketId,
            userId: record.userId,
          }
        );
        if (
          committedAttachmentOperation &&
          (committedAttachmentOperation.status === "consumed" ||
            committedAttachmentOperation.haloActionId)
        ) {
          await consumeEmailAttachmentPrefetchKey(body.emailAttachmentPrefetchKey);
          match.mapping.actionMode = actionMode;
          await markEmailSynced(match.mapping, email);
          sendJson(res, 200, {
            ok: true,
            status: "already-attached",
            ticketId: match.mapping.ticketId,
            ticketNumber: match.mapping.ticketNumber,
          });
          return;
        }
      }

      const preparedAction = await buildEmailActionWithInlineImages(
        record,
        match.mapping.ticketId,
        email,
        body,
        {
          onDiagnostic: logInlineImageDiagnostic,
          onEmailAttachmentDiagnostic: logEmailAttachmentDiagnostic,
        }
      );
      const actionCreationStart = Date.now();
      const payload = await createPreparedHaloAction(record, preparedAction, () =>
        callHaloApiWithRefresh(record, {
          body: [preparedAction.payload],
          method: "POST",
          path: "/api/Actions",
          phase: "email-auto-attach",
          messagePrefix: "Halo email auto-attach failed",
        })
      );
      preparedAction.timings.actionCreationMs = Date.now() - actionCreationStart;
      const actionId = getCreatedActionId(payload);
      await consumeEmailAttachmentPrefetch(preparedAction);

      match.mapping.actionMode = actionMode;
      await markEmailSynced(match.mapping, email);

      sendJson(res, 200, {
        ok: true,
        status: getAttachedStatus(preparedAction),
        ticketId: match.mapping.ticketId,
        ticketNumber: match.mapping.ticketNumber,
        message:
          actionMode === "private-note"
            ? `Email attached as a private note to ticket ${getMappingTicketLabel(match.mapping)}.`
            : `Email automatically added to ticket ${getMappingTicketLabel(match.mapping)}.`,
        actionId: actionId || undefined,
        emailAttachments: preparedAction.emailAttachments,
        inlineImages: preparedAction.inlineImages,
        inlineImageTimings: preparedAction.timings,
      });
    } catch (error) {
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        status: getAttachmentFailureStatus(error),
        message: "Email auto-attach failed",
        error: publicError(error),
        debug: publicDebug(error),
      });
    }
  });

  app.post("/api/halo/email/send-auto-attach", async (req, res) => {
    const sendStartedAt = Date.now();
    let requestedAttachmentCount = 0;
    try {
      logSendDiagnostic("server", {
        elapsedMs: 0,
        outcome: "started",
        stage: "request-received",
      });
      const body = await readJsonBody(req, MAX_EMAIL_JSON_BODY_BYTES);
      requestedAttachmentCount = getSendDiagnosticAttachmentCount(body);
      logSendDiagnostic("server", {
        attachmentCount: requestedAttachmentCount,
        elapsedMs: Date.now() - sendStartedAt,
        includeAttachments: Boolean(body.includeEmailAttachments),
        inlineImageCount: Array.isArray(body.inlineImageRefs) ? body.inlineImageRefs.length : 0,
        outcome: "ok",
        stage: "payload-read",
      });
      let authSource = "cookie";
      let record = await getSessionRecord(req);
      if (!record) {
        authSource = "background-session";
        record = await getBackgroundSessionRecord(body.backgroundSessionId);
      }
      if (!record) {
        authSource = "microsoft-bearer";
        record = await getBearerGrantRecord(req);
      }

      if (!record) {
        logSendDiagnostic("server", {
          authSource: "none",
          elapsedMs: Date.now() - sendStartedAt,
          outcome: "no-session",
          stage: "authentication-complete",
        });
        logSendDiagnostic("server", {
          elapsedMs: Date.now() - sendStartedAt,
          outcome: "no-session",
          stage: "response-complete",
        });
        sendJson(res, 200, {
          ok: true,
          status: "no-session",
        });
        return;
      }

      logSendDiagnostic("server", {
        authSource,
        elapsedMs: Date.now() - sendStartedAt,
        outcome: "ok",
        stage: "authentication-complete",
      });
      const email = normalizeSendEmailPayload(body);
      logSendDiagnostic("server", {
        elapsedMs: Date.now() - sendStartedAt,
        outcome: "ok",
        stage: "payload-validated",
      });
      const match = await findConversationOrRecoveredMapping(record, email, body);

      if (!match) {
        logSendDiagnostic("server", {
          conversationIdentifierCount: email.conversationId ? 1 : 0,
          elapsedMs: Date.now() - sendStartedAt,
          inReplyToCount: email.inReplyToMessageIds.length,
          outcome: "no-match",
          stage: "mapping-lookup",
        });
        logSendDiagnostic("server", {
          elapsedMs: Date.now() - sendStartedAt,
          outcome: "no-match",
          stage: "response-complete",
        });
        sendJson(res, 200, {
          ok: true,
          status: "no-match",
        });
        return;
      }

      logSendDiagnostic("server", {
        conversationIdentifierCount: email.conversationId ? 1 : 0,
        elapsedMs: Date.now() - sendStartedAt,
        inReplyToCount: email.inReplyToMessageIds.length,
        outcome: "matched",
        stage: "mapping-lookup",
      });
      if (match.status === "already-attached") {
        logSendDiagnostic("server", {
          elapsedMs: Date.now() - sendStartedAt,
          outcome: "already-attached",
          stage: "idempotency-check",
        });
        await consumeEmailAttachmentPrefetchKey(body.emailAttachmentPrefetchKey);
        logSendDiagnostic("server", {
          elapsedMs: Date.now() - sendStartedAt,
          outcome: "already-attached",
          stage: "response-complete",
        });
        sendJson(res, 200, {
          ok: true,
          status: "already-attached",
          ticketId: match.mapping.ticketId,
          ticketNumber: match.mapping.ticketNumber,
          message: `This email is already attached to ticket ${getMappingTicketLabel(match.mapping)}.`,
        });
        return;
      }

      const actionMode = Object.prototype.hasOwnProperty.call(body, "actionMode")
        ? normalizeActionMode(body.actionMode)
        : normalizeActionMode(match.mapping.actionMode);
      body.actionMode = actionMode;
      email.actionMode = actionMode;

      if (body.emailAttachmentPrefetchKey) {
        const committedAttachmentOperation = await authStore.getEmailAttachmentPrefetch(
          body.emailAttachmentPrefetchKey,
          {
            haloTenant: record.haloUrl,
            ticketId: match.mapping.ticketId,
            userId: record.userId,
          }
        );
        if (
          committedAttachmentOperation &&
          (committedAttachmentOperation.status === "consumed" ||
            committedAttachmentOperation.haloActionId)
        ) {
          logSendDiagnostic("server", {
            elapsedMs: Date.now() - sendStartedAt,
            outcome: "already-attached",
            stage: "idempotency-check",
          });
          await consumeEmailAttachmentPrefetchKey(body.emailAttachmentPrefetchKey);
          match.mapping.actionMode = actionMode;
          await markEmailSynced(match.mapping, email);
          logSendDiagnostic("server", {
            elapsedMs: Date.now() - sendStartedAt,
            outcome: "already-attached",
            stage: "response-complete",
          });
          sendJson(res, 200, {
            ok: true,
            status: "already-attached",
            ticketId: match.mapping.ticketId,
            ticketNumber: match.mapping.ticketNumber,
          });
          return;
        }
      }

      logSendDiagnostic("server", {
        elapsedMs: Date.now() - sendStartedAt,
        outcome: "started",
        stage: "assets-prepare-start",
      });
      const preparedAction = await buildEmailActionWithInlineImages(
        record,
        match.mapping.ticketId,
        email,
        body,
        {
          addWarningFooter: true,
          onDiagnostic: logInlineImageDiagnostic,
          onEmailAttachmentDiagnostic: logEmailAttachmentDiagnostic,
        }
      );
      logSendDiagnostic("server", {
        attachmentCount: preparedAction.emailAttachments.selected,
        elapsedMs: Date.now() - sendStartedAt,
        inlineImageCount: preparedAction.inlineImages.referenced,
        outcome: getAttachedStatus(preparedAction),
        stage: "assets-prepare-complete",
      });
      const actionCreationStart = Date.now();
      logSendDiagnostic("server", {
        elapsedMs: Date.now() - sendStartedAt,
        outcome: "started",
        stage: "halo-action-start",
      });
      const payload = await createPreparedHaloAction(record, preparedAction, () =>
        callHaloApiWithTicketContext(match.mapping, () =>
          callHaloApiWithRefresh(record, {
            body: [preparedAction.payload],
            method: "POST",
            path: "/api/Actions",
            phase: "email-send-auto-attach",
            messagePrefix: "Halo sent email auto-attach failed",
          })
        )
      );
      preparedAction.timings.actionCreationMs = Date.now() - actionCreationStart;
      logSendDiagnostic("server", {
        elapsedMs: Date.now() - sendStartedAt,
        outcome: "ok",
        stage: "halo-action-complete",
      });
      const actionId = getCreatedActionId(payload);
      await consumeEmailAttachmentPrefetch(preparedAction);

      match.mapping.actionMode = actionMode;
      await markEmailSynced(match.mapping, email);
      logSendDiagnostic("server", {
        elapsedMs: Date.now() - sendStartedAt,
        outcome: "ok",
        stage: "mapping-complete",
      });

      const responseStatus = getAttachedStatus(preparedAction);
      logSendDiagnostic("server", {
        elapsedMs: Date.now() - sendStartedAt,
        outcome: responseStatus,
        stage: "response-complete",
      });

      sendJson(res, 200, {
        ok: true,
        status: responseStatus,
        ticketId: match.mapping.ticketId,
        ticketNumber: match.mapping.ticketNumber,
        message:
          actionMode === "private-note"
            ? `Sent email attached as a private note to ticket ${getMappingTicketLabel(match.mapping)}.`
            : `Sent email added to Halo ticket ${getMappingTicketLabel(match.mapping)}.`,
        actionId: actionId || undefined,
        emailAttachments: preparedAction.emailAttachments,
        inlineImages: preparedAction.inlineImages,
        inlineImageTimings: preparedAction.timings,
      });
    } catch (error) {
      logSendDiagnostic("server", {
        attachmentCount: requestedAttachmentCount,
        elapsedMs: Date.now() - sendStartedAt,
        errorKind: getSendDiagnosticErrorKind(error),
        outcome:
          error && error.code === "attachments-not-ready"
            ? getEmailAttachmentDiagnosticOutcome(error, "attachments-not-ready")
            : "failed",
        stage: "response-complete",
        statusCode: getErrorStatus(error, 502),
      });
      sendJson(res, getErrorStatus(error, 502), {
        ok: false,
        status: getAttachmentFailureStatus(error),
        message: "Sent email auto-attach failed",
        error: publicError(error),
        debug: publicDebug(error),
        ticketNumber: getErrorTicketNumber(error),
      });
    }
  });

  app.post("/api/auth/logout", async (req, res) => {
    const sessionId = getSessionIdFromRequest(req);
    let userId = null;

    if (sessionId) {
      const sessionHash = hashSessionId(sessionId);
      const record = await getSessionRecordBySessionId(sessionId);
      if (record) {
        userId = record.userId;
      } else {
        await authStore.deleteBackgroundSessionsForSessionHash(sessionHash);
        await authStore.deleteSession(sessionHash);
      }
    }

    if (!userId) {
      try {
        const user = await getMicrosoftUserFromRequest(req);
        userId = user ? user.id : null;
      } catch {
        userId = null;
      }
    }

    if (userId) {
      await authStore.deleteSessionsForUser(userId);
    }

    res.setHeader("Set-Cookie", clearSessionCookie());
    sendJson(res, 200, { authenticated: false });
  });

  registerBugReportRoutes(app, {
    env,
    githubClient: options.githubClient,
    requireMicrosoftUser,
    store: authStore,
  });
}

async function loadTicketCreationTypes(record, { refresh = false } = {}) {
  const cacheKey = "ticket-types:v1";
  const cached = await authStore.getTicketCreationMetadata(cacheKey, {
    haloTenant: record.haloUrl,
    userId: record.userId,
  });
  const now = Date.now();
  if (!refresh && cached && cached.expiresAt > now) {
    return { ...cached.payload, cached: true, stale: false };
  }
  try {
    const payload = await callHaloApiWithRefresh(
      record,
      "/api/TicketType?count=1000&can_create_only=true&canagentsselect=true&showinactive=false&include_mandatory_field_check=true",
      "ticket-creation-types",
      "Halo ticket type discovery failed"
    );
    const result = {
      fetchedAt: new Date(now).toISOString(),
      types: normalizeTicketTypes(payload),
    };
    await authStore.saveTicketCreationMetadata({
      cacheKey,
      expiresAt: now + TICKET_CREATION_METADATA_TTL_MS,
      haloTenant: record.haloUrl,
      payload: result,
      userId: record.userId,
    });
    return { ...result, cached: false, stale: false };
  } catch (error) {
    if (cached && cached.fetchedAt + TICKET_CREATION_METADATA_STALE_MS > now) {
      return { ...cached.payload, cached: true, stale: true };
    }
    throw error;
  }
}

async function loadTicketCreationSchema(record, type, { refresh = false } = {}) {
  // v4 adds per-ticket-type Agent New visibility and mandatory detection.
  const cacheKey = `ticket-type-schema:v4:${type.id}`;
  const cached = await authStore.getTicketCreationMetadata(cacheKey, {
    haloTenant: record.haloUrl,
    userId: record.userId,
  });
  const now = Date.now();
  if (!refresh && cached && cached.expiresAt > now) {
    return { ...cached.payload, cached: true, stale: false };
  }
  try {
    const detailPath = `/api/TicketType/${encodeURIComponent(
      type.id
    )}?includeconfig=true&includedetails=true&isnewticket=true&include_mandatory_field_check=true`;
    const [detailResult, fieldResult] = await Promise.allSettled([
      callHaloApiWithRefresh(
        record,
        detailPath,
        "ticket-creation-schema",
        "Halo ticket field discovery failed"
      ),
      callHaloApiWithRefresh(
        record,
        `/api/TicketTypeField?count=1000&isrtconfig=true&tickettype_id=${encodeURIComponent(type.id)}`,
        "ticket-creation-fields",
        "Halo ticket field discovery failed"
      ),
    ]);
    if (detailResult.status === "rejected") {
      throw detailResult.reason;
    }
    if (
      fieldResult.status === "rejected" &&
      !isOptionalHaloMetadataEndpointUnavailable(fieldResult.reason)
    ) {
      throw fieldResult.reason;
    }
    const detailPayload = detailResult.value;
    const fieldPayload = fieldResult.status === "fulfilled" ? fieldResult.value : [];
    let schema = normalizeTicketTypeSchema(type.id, type.name, detailPayload, fieldPayload);
    const optionPayloads = {};
    const optionRequestsByPath = new Map();
    for (const field of schema.fields) {
      if (!field.optionSource || field.options.length) {
        continue;
      }
      const path =
        field.optionSource === "severity"
          ? "/api/Lookup?lookupid=27"
          : `/api/FieldInfo/${encodeURIComponent(
              field.id
            )}?includedetails=true&getlookupvalues=true&entityid=0`;
      if (!optionRequestsByPath.has(path)) {
        optionRequestsByPath.set(
          path,
          callHaloApiWithRefresh(
            record,
            path,
            "ticket-creation-field-options",
            "Halo ticket field option discovery failed"
          )
        );
      }
      optionPayloads[field.key] = { path };
    }
    const optionResults = new Map();
    await Promise.all(
      Array.from(optionRequestsByPath.entries()).map(async ([path, request]) => {
        try {
          optionResults.set(path, await request);
        } catch (error) {
          if (!isOptionalHaloMetadataEndpointUnavailable(error)) {
            throw error;
          }
          optionResults.set(path, []);
        }
      })
    );
    Object.entries(optionPayloads).forEach(([fieldKey, value]) => {
      optionPayloads[fieldKey] = optionResults.get(value.path) || [];
    });
    schema = hydrateTicketCreationFieldOptions(schema, optionPayloads);
    const result = { fetchedAt: new Date(now).toISOString(), schema };
    await authStore.saveTicketCreationMetadata({
      cacheKey,
      expiresAt: now + TICKET_CREATION_METADATA_TTL_MS,
      haloTenant: record.haloUrl,
      payload: result,
      userId: record.userId,
    });
    return { ...result, cached: false, stale: false };
  } catch (error) {
    if (cached && cached.fetchedAt + TICKET_CREATION_METADATA_STALE_MS > now) {
      return { ...cached.payload, cached: true, stale: true };
    }
    throw error;
  }
}

async function searchHaloRequesters(record, query) {
  const payload = await callHaloApiWithRefresh(
    record,
    `/api/Users?search=${encodeURIComponent(
      query
    )}&count=20&includeinactive=false&includeclient=true&includesite=true`,
    "ticket-creation-requesters",
    "Halo requester search failed"
  );
  return normalizeRequesters(payload).slice(0, 20);
}

async function normalizeTicketCreationIntent(record, input) {
  const source = input && typeof input === "object" ? input : {};
  const typeId = normalizePositiveInteger(
    source.typeId,
    "A valid Halo ticket type ID is required."
  );
  const typesResult = await loadTicketCreationTypes(record, { refresh: false });
  const type = typesResult.types.find((candidate) => Number(candidate.id) === typeId);
  if (!type) {
    throw new RequestError("This Halo ticket type is unavailable to the current agent.", 404);
  }
  const schemaResult = await loadTicketCreationSchema(record, type, { refresh: false });
  const requesterMode = source.requesterMode === "auto" ? "auto" : "explicit";
  const requester = normalizeCreationRequester(source.requester);
  const submittedValues = {
    ...(source.values && typeof source.values === "object" ? source.values : {}),
  };
  if (requester.id) {
    submittedValues["core:user_id"] = requester.id;
  }
  if (requester.clientId) {
    submittedValues["core:client_id"] = requester.clientId;
  }
  if (requester.siteId) {
    submittedValues["core:site_id"] = requester.siteId;
  }
  const validated = validateCreationInput(schemaResult.schema, {
    schemaRevision: source.schemaRevision,
    summary: source.summary,
    values: submittedValues,
  });
  if (requesterMode === "explicit" && !requester.id) {
    throw new RequestError("Choose a Halo requester before creating the ticket.", 400);
  }
  return {
    typeId: String(typeId),
    typeName: type.name,
    schemaRevision: schemaResult.schema.revision,
    summary: validated.summary,
    summaryMode: source.summaryMode === "auto" ? "auto" : "fixed",
    values: validated.values,
    requesterMode,
    requester,
    draftItemId: stringifyField(source.draftItemId),
    emailAttachmentDecision:
      source.emailAttachmentDecision === "include" || source.emailAttachmentDecision === "exclude"
        ? source.emailAttachmentDecision
        : undefined,
    emailAttachmentFingerprint: stringifyField(source.emailAttachmentFingerprint),
    emailAttachmentPrefetchKey: stringifyField(source.emailAttachmentPrefetchKey),
    emailAttachmentStagingVersion:
      Number(source.emailAttachmentStagingVersion) === 2 ? 2 : undefined,
    emailAttachmentSummary:
      source.emailAttachmentSummary && typeof source.emailAttachmentSummary === "object"
        ? source.emailAttachmentSummary
        : undefined,
    actionMode: normalizeActionMode(source.actionMode),
  };
}

function normalizeCreationRequester(value) {
  const source = value && typeof value === "object" ? value : {};
  const normalizeId = (candidate) => {
    const id = Number.parseInt(String(candidate || ""), 10);
    return Number.isSafeInteger(id) && id > 0 ? String(id) : "";
  };
  return {
    id: normalizeId(source.id || source.userId),
    name: stringifyField(source.name).slice(0, 300),
    emailAddress: normalizeMailboxEmail(source.emailAddress || source.email),
    clientId: normalizeId(source.clientId),
    clientName: stringifyField(source.clientName).slice(0, 300),
    siteId: normalizeId(source.siteId),
    siteName: stringifyField(source.siteName).slice(0, 300),
  };
}

async function saveTicketCreationIntent(record, operationId, intent) {
  return authStore.upsertTicketCreationIntent({
    encryptedIntent: encryptJson(intent),
    expiresAt: Date.now() + TICKET_CREATION_INTENT_TTL_MS,
    haloTenant: record.haloUrl,
    operationId,
    userId: record.userId,
  });
}

async function getStoredTicketCreationIntent(record, operationId) {
  const stored = await authStore.getTicketCreationIntent(operationId, {
    haloTenant: record.haloUrl,
    userId: record.userId,
  });
  if (!stored) {
    return null;
  }
  try {
    return { record: stored, intent: decryptJson(stored.encryptedIntent) };
  } catch {
    return null;
  }
}

async function prepareFinalComposeIntent(record, intent, email) {
  const result = { ...intent };
  if (result.summaryMode === "auto") {
    result.summary = email.subject || email.normalizedSubject || "(no subject)";
  }
  if (result.requesterMode === "auto") {
    result.requester = await resolveComposeRequester(record, email);
  }
  return normalizeTicketCreationIntent(record, result);
}

async function resolveComposeRequester(record, email) {
  const mailboxEmail = normalizeMailboxEmail(email.mailboxEmail);
  const recipient = (email.to || [])
    .map((value) => extractEmailAddress(value))
    .find((value) => value && value !== mailboxEmail);
  if (!recipient) {
    throw new RequestError(
      "No external recipient could be resolved as the Halo requester. Reopen the add-in and choose one.",
      409
    );
  }
  const matches = await searchHaloRequesters(record, recipient);
  const exact = matches.filter(
    (candidate) => normalizeMailboxEmail(candidate.emailAddress) === recipient
  );
  if (exact.length !== 1) {
    throw new RequestError(
      `Halo could not uniquely match ${recipient}. Reopen the add-in and choose a requester.`,
      409
    );
  }
  return exact[0];
}

function extractEmailAddress(value) {
  const rendered = stringifyField(value).toLowerCase();
  const bracketed = /<([^>]+)>/.exec(rendered);
  const candidate = bracketed ? bracketed[1] : rendered;
  const match = /[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9.-]+/.exec(candidate);
  return match ? match[0] : "";
}

async function createTicketAndEmailAction(
  record,
  operationId,
  intent,
  email,
  input,
  {
    addWarningFooter = false,
    bodyMode = "full",
    logEmailAttachmentDiagnostic,
    logInlineImageDiagnostic,
  }
) {
  let stored = await getStoredTicketCreationIntent(record, operationId);
  if (!stored) {
    await saveTicketCreationIntent(record, operationId, intent);
    stored = await getStoredTicketCreationIntent(record, operationId);
  }
  if (stored.record.status === "complete") {
    return {
      status: "attached",
      ticketId: String(stored.record.ticketId),
      ticketNumber: stored.record.ticketNumber,
      actionId: stored.record.actionId || undefined,
      message:
        intent.actionMode === "private-note"
          ? `Email attached as a private note to newly created Halo ticket ${stored.record.ticketNumber}.`
          : `Email added to newly created Halo ticket ${stored.record.ticketNumber}.`,
    };
  }

  const actionInput = { ...input };
  if (!Object.prototype.hasOwnProperty.call(actionInput, "actionMode")) {
    actionInput.actionMode = intent.actionMode;
  }
  email.actionMode = normalizeActionMode(actionInput.actionMode);
  if (
    !Object.prototype.hasOwnProperty.call(actionInput, "includeEmailAttachments") &&
    !actionInput.emailAttachmentPrefetchKey &&
    intent.emailAttachmentDecision === "include" &&
    intent.emailAttachmentPrefetchKey
  ) {
    actionInput.includeEmailAttachments = true;
    actionInput.emailAttachmentFingerprint = intent.emailAttachmentFingerprint;
    actionInput.emailAttachmentPrefetchKey = intent.emailAttachmentPrefetchKey;
    actionInput.emailAttachmentStagingVersion = intent.emailAttachmentStagingVersion;
    actionInput.emailAttachmentDraftItemId = intent.draftItemId;
    actionInput.emailAttachmentOperationId = operationId;
    actionInput.emailAttachmentSummary = intent.emailAttachmentSummary;
  }

  let ticketId = Number(stored.record.ticketId) || 0;
  let ticketNumber = stored.record.ticketNumber || "";
  if (!ticketId) {
    const typesResult = await loadTicketCreationTypes(record, { refresh: true });
    if (typesResult.stale) {
      throw new RequestError(
        "Halo ticket types could not be revalidated. Reopen the creation form and review it.",
        409
      );
    }
    const type = typesResult.types.find((candidate) => candidate.id === intent.typeId);
    if (!type) {
      throw new RequestError("This Halo ticket type is no longer available.", 409);
    }
    const schemaResult = await loadTicketCreationSchema(record, type, { refresh: true });
    if (schemaResult.stale) {
      throw new RequestError(
        "Halo ticket fields could not be revalidated. Reopen the creation form and review them.",
        409
      );
    }
    const validated = validateCreationInput(schemaResult.schema, intent);
    const ticketPayload = buildTicketPayload({
      schema: schemaResult.schema,
      summary: validated.summary,
      values: validated.values,
      requester: intent.requester,
    });
    const creationPayload = await callHaloApiWithRefresh(record, {
      body: [ticketPayload],
      method: "POST",
      path: "/api/Tickets",
      phase: "ticket-create-from-email",
      messagePrefix: "Halo ticket creation failed",
      retryTransient: true,
    });
    const created = getCreatedTicket(creationPayload);
    if (!created) {
      throw new RequestError("Halo did not return the created ticket ID.", 502);
    }
    ticketId = Number(created.id);
    ticketNumber = created.ticketNumber;
    await authStore.updateTicketCreationIntent(
      operationId,
      {
        encryptedIntent: encryptJson(intent),
        lastError: "",
        status: "ticket-created",
        ticketId,
        ticketNumber,
      },
      { haloTenant: record.haloUrl, userId: record.userId }
    );
  }

  if (actionInput.emailAttachmentPrefetchKey) {
    await authStore.rebindEmailAttachmentPrefetch(actionInput.emailAttachmentPrefetchKey, {
      haloTenant: record.haloUrl,
      ticketId,
      userId: record.userId,
    });
    const committedAttachmentOperation = await authStore.getEmailAttachmentPrefetch(
      actionInput.emailAttachmentPrefetchKey,
      { haloTenant: record.haloUrl, ticketId, userId: record.userId }
    );
    if (
      committedAttachmentOperation &&
      (committedAttachmentOperation.status === "consumed" ||
        committedAttachmentOperation.haloActionId)
    ) {
      const receiptActionId = committedAttachmentOperation.haloActionId || "created";
      await authStore.updateTicketCreationIntent(
        operationId,
        {
          actionId: receiptActionId,
          lastError: "",
          status: "ticket-created",
          ticketId,
          ticketNumber,
        },
        { haloTenant: record.haloUrl, userId: record.userId }
      );
      await authStore.consumeEmailAttachmentPrefetch(actionInput.emailAttachmentPrefetchKey);
      await storeConversationMapping({
        email,
        includeThreadMessageIds: true,
        ticketId,
        ticketNumber,
      });
      await authStore.updateTicketCreationIntent(
        operationId,
        { actionId: receiptActionId, lastError: "", status: "complete", ticketId, ticketNumber },
        { haloTenant: record.haloUrl, userId: record.userId }
      );
      return {
        status: "attached",
        ticketId: String(ticketId),
        ticketNumber,
        actionId: receiptActionId === "created" ? undefined : receiptActionId,
        message:
          intent.actionMode === "private-note"
            ? `Email attached as a private note to newly created Halo ticket ${ticketNumber}.`
            : `Email added to newly created Halo ticket ${ticketNumber}.`,
      };
    }
  }

  if (stored.record.actionId) {
    if (actionInput.emailAttachmentPrefetchKey) {
      await authStore.consumeEmailAttachmentPrefetch(actionInput.emailAttachmentPrefetchKey);
    }
    await storeConversationMapping({
      email,
      includeThreadMessageIds: true,
      ticketId,
      ticketNumber,
    });
    await authStore.updateTicketCreationIntent(
      operationId,
      { lastError: "", status: "complete", ticketId, ticketNumber },
      { haloTenant: record.haloUrl, userId: record.userId }
    );
    return {
      status: "attached",
      ticketId: String(ticketId),
      ticketNumber,
      actionId: stored.record.actionId === "created" ? undefined : stored.record.actionId,
      message:
        intent.actionMode === "private-note"
          ? `Email attached as a private note to newly created Halo ticket ${ticketNumber}.`
          : `Email added to newly created Halo ticket ${ticketNumber}.`,
    };
  }

  try {
    const preparedAction = await buildEmailActionWithInlineImages(
      record,
      ticketId,
      email,
      actionInput,
      {
        addWarningFooter,
        bodyMode,
        onDiagnostic: logInlineImageDiagnostic,
        onEmailAttachmentDiagnostic: logEmailAttachmentDiagnostic,
      }
    );
    const payload = await createPreparedHaloAction(record, preparedAction, () =>
      callHaloApiWithTicketContext({ ticketId, ticketNumber }, () =>
        callHaloApiWithRefresh(record, {
          body: [preparedAction.payload],
          method: "POST",
          path: "/api/Actions",
          phase: "ticket-create-email-action",
          messagePrefix: "Halo created-ticket email attach failed",
          retryTransient: true,
        })
      )
    );
    const actionId = getCreatedActionId(payload);
    await authStore.updateTicketCreationIntent(
      operationId,
      {
        actionId: actionId || "created",
        lastError: "",
        status: "ticket-created",
        ticketId,
        ticketNumber,
      },
      { haloTenant: record.haloUrl, userId: record.userId }
    );
    await consumeEmailAttachmentPrefetch(preparedAction);
    await storeConversationMapping({
      email,
      includeThreadMessageIds: true,
      ticketId,
      ticketNumber,
    });
    await authStore.updateTicketCreationIntent(
      operationId,
      { actionId, lastError: "", status: "complete", ticketId, ticketNumber },
      { haloTenant: record.haloUrl, userId: record.userId }
    );
    return {
      status: getAttachedStatus(preparedAction),
      ticketId: String(ticketId),
      ticketNumber,
      actionId: actionId || undefined,
      message:
        intent.actionMode === "private-note"
          ? `Email attached as a private note to newly created Halo ticket ${ticketNumber}.`
          : `Email added to newly created Halo ticket ${ticketNumber}.`,
      emailAttachments: preparedAction.emailAttachments,
      inlineImages: preparedAction.inlineImages,
      inlineImageTimings: preparedAction.timings,
    };
  } catch (error) {
    await authStore.updateTicketCreationIntent(
      operationId,
      {
        lastError: publicError(error),
        status: "partial-failure",
        ticketId,
        ticketNumber,
      },
      { haloTenant: record.haloUrl, userId: record.userId }
    );
    error.ticketNumber = ticketNumber;
    throw error;
  }
}

async function runExclusiveTicketCreation(record, operationId, operation) {
  const key = `${record.userId}|${record.haloUrl}|${operationId}`;
  const previous = ticketCreationOperations.get(key) || Promise.resolve();
  let release;
  const current = new Promise((resolve) => {
    release = resolve;
  });
  ticketCreationOperations.set(key, current);
  await previous.catch(() => undefined);
  try {
    return await operation();
  } finally {
    release();
    if (ticketCreationOperations.get(key) === current) {
      ticketCreationOperations.delete(key);
    }
  }
}

function getBooleanQueryParameter(req, name) {
  const value = getRequestUrl(req).searchParams.get(name);
  return value === "1" || value === "true";
}

async function exchangeAuthorizationCode({
  haloUrl,
  clientId,
  code,
  codeVerifier,
  redirectUri,
  scope,
}) {
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
    if (record.grantId) {
      await authStore.invalidateGrantById(record.grantId);
    }
    throw HttpError.fromResponse(
      "Halo refresh token request failed",
      "token-refresh",
      responseDetails
    );
  }

  const nextTokenPayload = annotateTokenPayload({
    ...currentTokenPayload,
    ...responseDetails.payload,
    refresh_token: responseDetails.payload.refresh_token || currentTokenPayload.refresh_token,
  });

  record.encryptedToken = encryptJson(nextTokenPayload);
  if (record.grantId) {
    await authStore.updateGrantToken(record.grantId, record.encryptedToken);
  }
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
  const method = request.method || "GET";

  const execute = () =>
    fetchHaloJson({
      body: request.body,
      haloUrl: record.haloUrl,
      messagePrefix: request.messagePrefix,
      method,
      path: request.path,
      phase: request.phase,
      scope: record.scope,
      tokenPayload,
    });

  try {
    return await execute();
  } catch (error) {
    if (isUnauthorizedError(error) && tokenPayload.refresh_token) {
      tokenPayload = await refreshAccessToken(record, tokenPayload);
      return execute();
    }

    const mayRetryNetworkFailure = method === "GET" && !(error instanceof HttpError);
    if (
      ((method === "GET" || request.retryTransient) && isTransientHaloResponseError(error)) ||
      mayRetryNetworkFailure
    ) {
      return execute();
    }
    throw error;
  }
}

function isTransientHaloResponseError(error) {
  return error instanceof HttpError && [408, 425, 429, 500, 502, 503, 504].includes(error.status);
}

function isOptionalHaloMetadataEndpointUnavailable(error) {
  return error instanceof HttpError && (error.status === 404 || error.status === 405);
}

async function fetchHaloJson({
  body,
  haloUrl,
  messagePrefix,
  method,
  path,
  phase,
  scope,
  tokenPayload,
}) {
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

async function uploadHaloInlineImage(record, image, actionMode = "email") {
  let tokenPayload = await getValidTokenPayload(record);
  let uploadPayload;

  try {
    uploadPayload = await fetchHaloAttachment(record, tokenPayload, image, actionMode);
  } catch (error) {
    if (!isUnauthorizedError(error) || !tokenPayload.refresh_token) {
      throw error;
    }
    tokenPayload = await refreshAccessToken(record, tokenPayload);
    uploadPayload = await fetchHaloAttachment(record, tokenPayload, image, actionMode);
  }

  const renderableUrl = extractHaloAttachmentUrl(uploadPayload, record.haloUrl);
  if (!renderableUrl) {
    throw new RequestError("Halo did not return a native inline-image URL.", 502);
  }

  const attachmentId = extractHaloAttachmentId(uploadPayload, renderableUrl);
  if (!attachmentId) {
    throw new RequestError("Halo did not return an identifiable inline-image attachment.", 502);
  }

  return { attachmentId: String(attachmentId), renderableUrl };
}

async function uploadHaloEmailAttachment(record, attachment, actionMode = "email") {
  const temporaryId = `__${crypto.randomUUID()}`;
  const payload = await callHaloApiWithRefresh(record, {
    body: [
      {
        data_base64: `data:${attachment.contentType};base64,${attachment.contentBase64}`,
        filename: attachment.name,
        filesize: attachment.decodedSize,
        showforusers: actionMode === "private-note" ? false : null,
        _uploading: true,
        _tempid: temporaryId,
        showonchild: false,
        showonrelated: false,
      },
    ],
    method: "POST",
    path: "/api/Attachment",
    phase: "email-attachment-upload",
    messagePrefix: "Halo email attachment upload failed",
  });
  const uploaded = normalizeHaloEmailAttachment(payload, attachment, actionMode);
  if (!uploaded.id) {
    throw new RequestError("Halo did not return an attachment ID.", 502);
  }
  return uploaded;
}

function normalizeHaloEmailAttachment(payload, source, actionMode = "email") {
  const value = Array.isArray(payload) ? payload[0] || {} : payload || {};
  const id = value.id || value.attachment_id || value.attachmentid || "";
  return {
    decodedSize: source.decodedSize,
    // Halo returns its internal storage filename here. The original sanitized
    // name is the display filename that its own action editor sends onward.
    filename: stringifyField(source.name),
    filesize: Number(value.filesize || value.file_size || source.decodedSize) || source.decodedSize,
    id: isHaloAttachmentIdentifier(id) ? String(id) : "",
    showforusers: actionMode === "private-note" ? false : null,
    type: Number(value.type || 0),
  };
}

async function deleteHaloEmailAttachment(record, attachmentId) {
  if (!isHaloAttachmentIdentifier(attachmentId)) {
    return true;
  }
  try {
    await callHaloApiWithRefresh(record, {
      method: "DELETE",
      path: `/api/Attachment/${encodeURIComponent(String(attachmentId))}`,
      phase: "email-attachment-cleanup",
      messagePrefix: "Halo email attachment cleanup failed",
    });
    return true;
  } catch {
    return false;
  }
}

async function cleanupHaloAttachments(record, attachments) {
  const values = Array.isArray(attachments) ? attachments : [];
  const results = await Promise.all(
    values.map((attachment) =>
      deleteHaloEmailAttachment(
        record,
        attachment && (attachment.id || attachment.haloAttachmentId)
      )
    )
  );
  return results.every(Boolean);
}

async function cleanupStoredHaloAttachments(record, prefetchKey, attachments) {
  const values = Array.isArray(attachments) ? attachments : [];
  const results = await Promise.all(
    values.map(async (attachment) => {
      const cleaned = await deleteHaloEmailAttachment(
        record,
        attachment && (attachment.id || attachment.haloAttachmentId)
      );
      if (cleaned && attachment && attachment.attachmentKey) {
        await authStore.markEmailAttachmentPrefetchItemCleaned(
          prefetchKey,
          attachment.attachmentKey
        );
      }
      return cleaned;
    })
  );
  return results.every(Boolean);
}

async function getEmailAttachmentRequestRecord(req, body) {
  return (
    (await getSessionRecord(req)) ||
    (await getBackgroundSessionRecord(body && body.backgroundSessionId)) ||
    (await getBearerGrantRecord(req))
  );
}

function createSendDiagnosticLogger(logger, env) {
  if ((env && env.SEND_EVENT_DIAGNOSTICS === "0") || (env && env.NODE_ENV === "test" && !logger)) {
    return () => {};
  }

  const sink = logger || console;
  const writer =
    typeof sink === "function"
      ? sink
      : typeof sink.info === "function"
        ? sink.info.bind(sink)
        : sink.log.bind(sink);
  const booleanKeys = new Set(["hasBackgroundSession", "hasMailboxIdentity", "includeAttachments"]);
  const numberKeys = new Set([
    "attachmentCount",
    "attemptCount",
    "candidateCount",
    "conversationIdentifierCount",
    "elapsedMs",
    "failedCount",
    "inlineImageCount",
    "inReplyToCount",
    "skippedCount",
    "statusCode",
    "uploadedCount",
  ]);
  const tokenKeys = new Set(["attachmentError", "authSource", "errorKind", "outcome", "stage"]);
  const allowedTokens = {
    attachmentError: new Set([
      "invalid-attachment-id",
      "invalid-content",
      "none",
      "not-supported",
      "read-failed",
      "timeout",
    ]),
    authSource: new Set(["background-session", "cookie", "microsoft-bearer", "none"]),
    errorKind: new Set(["authentication", "halo-http", "internal", "validation"]),
    outcome: new Set([
      "allowed",
      "allowed-auto-request-failed",
      "allowed-auto-result",
      "allowed-exception",
      "allowed-no-email",
      "already-attached",
      "attached",
      "attached-with-attachment-warnings",
      "attached-with-image-warnings",
      "attached-with-warnings",
      "automatic",
      "attachments-not-ready",
      "blocked",
      "blocked-auto-request-failed",
      "commit-unavailable",
      "content-changed",
      "create-ticket",
      "excluded",
      "explicit-ticket",
      "failed",
      "integrity-failed",
      "inventory-mismatch",
      "inventory-read-failed",
      "legacy-state",
      "matched",
      "matched-staged-operation",
      "mismatch",
      "missing-response",
      "network-error",
      "no-attachments",
      "no-match",
      "no-session",
      "not-ready",
      "ok",
      "operation-finalized",
      "pending",
      "preparation-incomplete",
      "prepared",
      "ready",
      "reservation-failed",
      "stage-missing",
      "stage-unavailable",
      "state-mismatch",
      "state-missing",
      "started",
      "timeout",
    ]),
    stage: new Set([
      "attachment-change-event",
      "attachment-inventory",
      "attachment-prefetch-complete",
      "attachment-read-complete",
      "attachment-read-retry",
      "attachment-staging",
      "attachment-state",
      "assets-complete",
      "assets-prepare-complete",
      "assets-prepare-start",
      "assets-start",
      "authentication-complete",
      "compose-read-complete",
      "compose-read-start",
      "draft-save-complete",
      "draft-save-start",
      "diagnostics-ready",
      "event-complete",
      "event-start",
      "halo-action-complete",
      "halo-action-start",
      "idempotency-check",
      "mapping-complete",
      "mapping-lookup",
      "marker-read",
      "payload-read",
      "payload-validated",
      "request-complete",
      "request-received",
      "request-start",
      "recovery-branch",
      "recovery-complete",
      "recovery-start",
      "response-complete",
      "runtime-loaded",
      "watchdog-fired",
    ]),
  };

  return (source, details = {}) => {
    const safeDetails = { source: normalizeSendDiagnosticToken(source, "unknown") };
    for (const [key, value] of Object.entries(details)) {
      if (booleanKeys.has(key) && typeof value === "boolean") {
        safeDetails[key] = value;
      } else if (numberKeys.has(key)) {
        const number = Number(value);
        if (Number.isFinite(number) && number >= 0) {
          safeDetails[key] = Math.min(60 * 60 * 1000, Math.round(number));
        }
      } else if (tokenKeys.has(key)) {
        const token = normalizeSendDiagnosticToken(value, "invalid");
        safeDetails[key] = allowedTokens[key].has(token) ? token : "invalid";
      }
    }
    writer(`[halo-send] ${JSON.stringify(safeDetails)}`);
  };
}

function normalizeSendDiagnosticToken(value, fallback) {
  const token = String(value || "")
    .trim()
    .toLowerCase();
  return /^[a-z0-9-]{1,48}$/.test(token) ? token : fallback;
}

function getSendDiagnosticAttachmentCount(body) {
  const summary = body && body.emailAttachmentSummary;
  const selected = Number(summary && summary.selected);
  return Number.isFinite(selected) && selected > 0 ? Math.round(selected) : 0;
}

function getSendDiagnosticErrorKind(error) {
  if (error instanceof RequestError || (error && error.isRequestError)) {
    return "validation";
  }
  if (error instanceof HttpError) {
    return "halo-http";
  }
  const status = Number(error && error.status);
  if (status === 401 || status === 403) {
    return "authentication";
  }
  return "internal";
}

function getEmailAttachmentDiagnosticOutcome(error, fallback = "failed") {
  const reason = normalizeSendDiagnosticToken(error && error.attachmentDiagnosticReason, fallback);
  const allowedReasons = new Set([
    "commit-unavailable",
    "content-changed",
    "integrity-failed",
    "inventory-mismatch",
    "legacy-state",
    "not-ready",
    "operation-finalized",
    "preparation-incomplete",
    "reservation-failed",
    "stage-missing",
    "stage-unavailable",
  ]);
  return allowedReasons.has(reason) ? reason : fallback;
}

function createInlineImageDiagnosticLogger(logger, env) {
  if (
    (env && env.INLINE_IMAGE_DIAGNOSTICS === "0") ||
    (env && env.NODE_ENV === "test" && !logger)
  ) {
    return () => {};
  }

  const sink = logger || console;
  const writer =
    typeof sink === "function"
      ? sink
      : typeof sink.info === "function"
        ? sink.info.bind(sink)
        : sink.log.bind(sink);
  const allowedKeys = new Set([
    "cacheHits",
    "durationMs",
    "failed",
    "referenced",
    "totalMs",
    "uploaded",
  ]);

  return (event, details = {}) => {
    const safeDetails = {};
    for (const [key, value] of Object.entries(details)) {
      if (allowedKeys.has(key) && ["number", "string", "boolean"].includes(typeof value)) {
        safeDetails[key] = value;
      }
    }
    writer(`[inline-images] ${JSON.stringify({ event, ...safeDetails })}`);
  };
}

function createEmailAttachmentDiagnosticLogger(logger, env) {
  if (
    (env && env.EMAIL_ATTACHMENT_DIAGNOSTICS === "0") ||
    (env && env.NODE_ENV === "test" && !logger)
  ) {
    return () => {};
  }

  const sink = logger || console;
  const infoWriter =
    typeof sink === "function"
      ? sink
      : typeof sink.info === "function"
        ? sink.info.bind(sink)
        : sink.log.bind(sink);
  const errorWriter =
    typeof sink === "function"
      ? sink
      : typeof sink.error === "function"
        ? sink.error.bind(sink)
        : infoWriter;
  const numberKeys = new Set([
    "conversationIdentifierCount",
    "count",
    "detected",
    "durationMs",
    "failed",
    "inReplyToCount",
    "pending",
    "prepared",
    "selected",
    "skipped",
    "uploaded",
  ]);
  const tokenKeys = new Set(["outcome", "stage"]);

  return (event, details = {}) => {
    const safeDetails = {};
    for (const [key, value] of Object.entries(details)) {
      if (numberKeys.has(key)) {
        const number = Number(value);
        if (Number.isFinite(number) && number >= 0) {
          safeDetails[key] = Math.min(60 * 60 * 1000, Math.round(number));
        }
      } else if (tokenKeys.has(key)) {
        safeDetails[key] = normalizeSendDiagnosticToken(value, "invalid");
      }
    }
    const safeEvent = normalizeSendDiagnosticToken(event, "invalid");
    const writer = safeEvent.includes("failed") ? errorWriter : infoWriter;
    writer(`[email-attachments] ${JSON.stringify({ event: safeEvent, ...safeDetails })}`);
  };
}

function sanitizeEmailAttachmentDiagnosticMessage(value, sensitiveValues = [], maxLength = 500) {
  let message = String(value || "").slice(0, 4000);
  for (const sensitiveValue of sensitiveValues) {
    const normalized = String(sensitiveValue || "");
    if (normalized) {
      message = message.replace(new RegExp(escapeRegExp(normalized), "gi"), "[redacted]");
    }
  }

  return message
    .replace(/(["']?data_base64["']?\s*[:=]\s*["']?)[a-z0-9+/_=-]+/gi, "$1[redacted]")
    .replace(/\bhttps?:\/\/\S+/gi, "[redacted-url]")
    .replace(/\b[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}\b/gi, "[redacted-email]")
    .replace(/\b[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}\b/gi, "[redacted-id]")
    .replace(/[a-z0-9+/_=-]{48,}/gi, "[redacted]")
    .replace(/\b\d+\b/g, "[number]")
    .replace(/[\u0000-\u001f\u007f]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLength);
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function fetchHaloAttachment(record, tokenPayload, image, actionMode = "email") {
  const requestUrl = resolveHaloUrl(record.haloUrl, "/api/attachment/image");
  const form = new FormData();
  form.append("file", new Blob([image.bytes], { type: image.mediaType }), image.name);
  form.append("ticket_id", String(image.ticketId));
  form.append("image_upload_id", "0");
  form.append("image_upload_key", "");
  form.append("showforusers", actionMode === "private-note" ? "false" : "true");

  const response = await fetch(requestUrl, {
    method: "POST",
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${tokenPayload.access_token}`,
    },
    body: form,
  });
  if (!response.ok) {
    const responseDetails = await readResponseDetails(response, requestUrl);
    throw HttpError.fromResponse(
      "Halo inline image upload failed",
      "inline-image-upload",
      responseDetails
    );
  }
  const payload = await readResponseJson(response, requestUrl);
  const location = response.headers.get("location") || "";
  if (payload && typeof payload === "object" && !Array.isArray(payload)) {
    return { ...payload, location };
  }
  return { payload, location };
}

function extractHaloAttachmentId(payload, renderableUrl = "") {
  const tokenAttachmentId = extractAttachmentIdFromInlineImageUrl(renderableUrl);
  if (tokenAttachmentId) {
    return tokenAttachmentId;
  }

  const attachmentId = findPayloadValue(payload, (key, value) => {
    const normalizedKey = key.replace(/[^a-z0-9]/gi, "").toLowerCase();
    return normalizedKey === "attachmentid" && isHaloAttachmentIdentifier(value);
  });
  if (attachmentId) {
    return attachmentId;
  }

  if (
    payload &&
    typeof payload === "object" &&
    !Array.isArray(payload) &&
    isHaloAttachmentIdentifier(payload.id)
  ) {
    return payload.id;
  }

  return "";
}

function isHaloAttachmentIdentifier(value) {
  return /^[A-Za-z0-9_-]+$/.test(String(value || ""));
}

function extractAttachmentIdFromInlineImageUrl(value) {
  if (!value) {
    return "";
  }
  try {
    const url = new URL(String(value));
    const token = url.searchParams.get("token") || "";
    const parts = token.split(".");
    if (parts.length !== 3) {
      return "";
    }

    const tokenPayload = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8"));
    const attachmentId = tokenPayload && (tokenPayload.id || tokenPayload.attachment_id);
    return isHaloAttachmentIdentifier(attachmentId) ? String(attachmentId) : "";
  } catch {
    return "";
  }
}

function extractHaloAttachmentUrl(payload, haloUrl) {
  const token = findPayloadValue(payload, (key, value) => {
    const normalizedKey = key.replace(/[^a-z0-9]/gi, "").toLowerCase();
    return normalizedKey === "token" && typeof value === "string" && value.length > 5;
  });
  const value = findPayloadValue(payload, (key, candidate) => {
    if (typeof candidate !== "string") {
      return false;
    }
    const normalizedKey = key.replace(/[^a-z0-9]/gi, "").toLowerCase();
    return (
      ["url", "link", "path", "attachmenturl", "imageurl"].includes(normalizedKey) &&
      /\/api\/attachment\/image(?:\?|$)/i.test(candidate)
    );
  });

  if (value) {
    return validateHaloInlineImageUrl(value, haloUrl);
  }
  if (token) {
    return validateHaloInlineImageUrl(
      `/api/attachment/image?token=${encodeURIComponent(String(token))}`,
      haloUrl
    );
  }
  return "";
}

function validateHaloInlineImageUrl(value, haloUrl) {
  try {
    const tenant = new URL(String(haloUrl));
    const url = new URL(String(value), tenant.origin);
    const normalizedPath = url.pathname.replace(/\/+$/, "").toLowerCase();
    if (
      tenant.protocol !== "https:" ||
      url.protocol !== "https:" ||
      url.origin.toLowerCase() !== tenant.origin.toLowerCase() ||
      normalizedPath !== "/api/attachment/image" ||
      !url.searchParams.get("token") ||
      url.username ||
      url.password ||
      url.hash
    ) {
      return "";
    }
    return url.toString();
  } catch {
    return "";
  }
}

function findPayloadValue(payload, predicate, key = "") {
  if (predicate(key, payload)) {
    return payload;
  }
  if (!payload || typeof payload !== "object") {
    return "";
  }
  for (const [childKey, value] of Object.entries(payload)) {
    const match = findPayloadValue(value, predicate, childKey);
    if (match !== "") {
      return match;
    }
  }
  return "";
}

async function readResponseJson(response, requestUrl) {
  const responseDetails = await readResponseDetails(response, requestUrl);
  return responseDetails.payload;
}

function buildTicketsPath({ query = "", ownership = "mine", lifecycle = "open" } = {}) {
  const params = new URLSearchParams();
  params.set("count", String(TICKETS_COUNT));
  if (lifecycle === "open") {
    params.set("open_only", "true");
  }
  if (ownership === "mine") {
    params.set("mine", "true");
  }
  if (query) {
    params.set("search", query);
  }
  params.set("includeagent", "true");
  params.set("includestatus", "true");

  return `/api/Tickets?${params.toString()}`;
}

function getTicketOwnership(req) {
  return getTicketSearchOption(req, "ownership", ["mine", "all"], "mine");
}

function getTicketLifecycle(req) {
  return getTicketSearchOption(req, "lifecycle", ["open", "closed", "all"], "open");
}

function getTicketSearchOption(req, name, allowedValues, fallback) {
  const value = stringifyField(getRequestUrl(req).searchParams.get(name)).toLowerCase();
  if (!value) {
    return fallback;
  }
  if (!allowedValues.includes(value)) {
    throw new RequestError(`${name} must be one of: ${allowedValues.join(", ")}.`, 400);
  }
  return value;
}

function getTicketSearchQuery(req) {
  const params = getRequestUrl(req).searchParams;
  const value = params.get("query") || params.get("ticketNumber");
  const rawQuery = stringifyField(value);
  const legacyIdMatch = rawQuery.match(/^\[?\s*id\s*:\s*(.*?)\s*\]?$/i);
  const query = (legacyIdMatch ? legacyIdMatch[1] : rawQuery).trim();

  if (!query) {
    throw new RequestError("Enter a ticket search query.", 400);
  }

  if (query.length > TICKET_SEARCH_MAX_LENGTH) {
    throw new RequestError(
      `Ticket searches must be ${TICKET_SEARCH_MAX_LENGTH} characters or fewer.`,
      400
    );
  }

  return query;
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

function normalizePositiveInteger(value, message) {
  const number = Number.parseInt(String(value || ""), 10);
  if (!Number.isInteger(number) || number <= 0) {
    throw new RequestError(message, 400);
  }
  return number;
}

function normalizeOpaqueIdentifier(value, message) {
  const identifier = String(value || "").trim();
  if (!/^[A-Za-z0-9_-]{8,200}$/.test(identifier)) {
    throw new RequestError(message, 400);
  }
  return identifier;
}

function normalizeEmailMatchPayload(value) {
  if (!value || typeof value !== "object") {
    throw new RequestError("Email payload is required.", 400);
  }

  const email = {
    conversationId: stringifyField(value.conversationId),
    inReplyToMessageIds: normalizeMessageIdList(value.inReplyToMessageIds),
    internetMessageId: stringifyField(value.internetMessageId),
    itemId: stringifyField(value.itemId),
    mailboxEmail: normalizeMailboxEmail(value.mailboxEmail),
    referenceMessageIds: normalizeMessageIdList(value.referenceMessageIds),
  };

  if (
    !email.internetMessageId &&
    !email.itemId &&
    !email.inReplyToMessageIds.length &&
    !email.referenceMessageIds.length &&
    !email.conversationId
  ) {
    throw new RequestError("No mapped email identifiers were available.", 400);
  }

  return email;
}

function normalizeEmailPayload(value) {
  if (!value || typeof value !== "object") {
    throw new RequestError("Email payload is required.", 400);
  }

  const email = {
    actionMode: normalizeActionMode(value.actionMode),
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
    actionMode: normalizeActionMode(value.actionMode),
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

function normalizeRecoveryEmailPayload(value, fallbackMailboxEmail) {
  const source = value && typeof value === "object" ? value : {};
  const mailboxEmail =
    normalizeMailboxEmail(source.mailboxEmail) || normalizeMailboxEmail(fallbackMailboxEmail);
  const authenticatedMailbox = normalizeMailboxEmail(fallbackMailboxEmail);
  if (!mailboxEmail || (authenticatedMailbox && mailboxEmail !== authenticatedMailbox)) {
    throw new RequestError("The mailbox identity could not be validated.", 403);
  }
  return {
    conversationId: stringifyField(source.conversationId),
    inReplyToMessageIds: normalizeMessageIdList(source.inReplyToMessageIds),
    internetMessageId: stringifyField(source.internetMessageId),
    itemId: stringifyField(source.itemId),
    mailboxEmail,
    referenceMessageIds: normalizeMessageIdList(source.referenceMessageIds),
  };
}

function normalizeRecoveryComposeIds(value) {
  if (!Array.isArray(value)) {
    value = value ? [value] : [];
  }
  const seen = new Set();
  const result = [];
  value.slice(0, 12).forEach((entry) => {
    const id = stringifyField(entry);
    if (id && id.length <= 200 && /^[A-Za-z0-9_-]+$/.test(id) && !seen.has(id)) {
      seen.add(id);
      result.push(id);
    }
  });
  return result;
}

async function findRecoveryMapping(record, email, composeAttachIds) {
  const normal = await findConversationMappingForEmail(email);
  if (normal) {
    return { match: normal, candidateIndex: -1 };
  }
  const mailboxEmail = normalizeMailboxEmail(record.userEmail || email.mailboxEmail);
  for (let index = 0; index < composeAttachIds.length; index += 1) {
    const syntheticId = buildComposeAttachMessageId({ mailboxEmail }, composeAttachIds[index]);
    const mapping = await getMappingByMessageId(mailboxEmail, syntheticId);
    if (mapping) {
      return { match: { mapping, status: "match" }, candidateIndex: index };
    }
  }
  return null;
}

async function findConversationOrRecoveredMapping(record, email, body) {
  const result = await findConversationOrRecoveredMappingResult(record, email, body);
  return result ? result.match : null;
}

async function findConversationOrRecoveredMappingResult(record, email, body) {
  const authenticatedMailbox = normalizeMailboxEmail(record && record.userEmail);
  if (authenticatedMailbox && authenticatedMailbox !== normalizeMailboxEmail(email.mailboxEmail)) {
    return null;
  }
  const candidates = normalizeRecoveryComposeIds(
    body && (body.recoveredComposeAttachId || body.composeAttachIds)
  );
  const result = await findRecoveryMapping(record, email, candidates);
  return result;
}

function normalizeExplicitSendEmailPayload(value, fallbackMailboxEmail) {
  if (!value || typeof value !== "object") {
    throw new RequestError("Email payload is required.", 400);
  }

  const composeAttachId = stringifyField(value.composeAttachId);
  if (
    !composeAttachId ||
    composeAttachId.length > 200 ||
    !/^[A-Za-z0-9_-]+$/.test(composeAttachId)
  ) {
    throw new RequestError("A valid compose attachment ID is required.", 400);
  }

  const email = {
    actionMode: normalizeActionMode(value.actionMode),
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
    mailboxEmail:
      normalizeMailboxEmail(value.mailboxEmail) || normalizeMailboxEmail(fallbackMailboxEmail),
    normalizedSubject: stringifyField(value.normalizedSubject),
    referenceMessageIds: normalizeMessageIdList(value.referenceMessageIds),
    subject: stringifyField(value.subject),
    timeZone: normalizeTimeZone(value.timeZone),
    to: normalizeEmailAddressList(value.to),
  };

  if (!email.mailboxEmail) {
    throw new RequestError("Mailbox identity is required to attach a composed email.", 400);
  }

  if (!email.bodyHtml && !email.bodyText) {
    throw new RequestError("Could not read an email body to attach.", 400);
  }

  // A compose item has no final transport message ID yet. Always derive the
  // idempotency key from the server-validated mailbox and compose operation,
  // rather than trusting a client/draft value that Outlook may later replace.
  email.internetMessageId = buildComposeAttachMessageId(email, composeAttachId);
  return email;
}

function buildSyntheticMessageId(email) {
  const stableKey = email.itemId ? hashStableValue(email.itemId) : buildOutgoingBodyHash(email);
  return `<halo-outlook-${stableKey}@local>`;
}

function buildComposeAttachMessageId(email, composeAttachId) {
  const stableKey = hashStableValue(
    JSON.stringify({
      composeAttachId,
      mailboxEmail: normalizeMailboxEmail(email.mailboxEmail),
    })
  );
  return `<halo-outlook-compose-${stableKey}@local>`;
}

function buildExplicitSendOperationKey(record, email) {
  return [
    stringifyField(record && record.userId),
    normalizeMailboxEmail(email.mailboxEmail),
    normalizeMessageIdKey(email.internetMessageId),
  ].join("|");
}

async function runExclusiveExplicitSend(operationKey, operation) {
  const previousOperation = explicitSendOperations.get(operationKey) || Promise.resolve();
  let releaseOperation;
  const currentOperation = new Promise((resolve) => {
    releaseOperation = resolve;
  });
  explicitSendOperations.set(operationKey, currentOperation);

  await previousOperation.catch(() => undefined);
  try {
    return await operation();
  } finally {
    releaseOperation();
    if (explicitSendOperations.get(operationKey) === currentOperation) {
      explicitSendOperations.delete(operationKey);
    }
  }
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

function normalizeMessageIdKeys(value) {
  const raw = normalizeMessageIdKey(value);
  if (!raw) {
    return [];
  }
  const unwrapped = raw.startsWith("<") && raw.endsWith(">") ? raw.slice(1, -1).trim() : raw;
  return unwrapped && unwrapped !== raw ? [raw, unwrapped] : [raw];
}

async function storeConversationMapping({
  actionMode,
  email,
  includeThreadMessageIds = false,
  ticketId,
  ticketNumber,
}) {
  const mailboxEmail = normalizeMailboxEmail(email.mailboxEmail);

  if (!mailboxEmail || !email.internetMessageId) {
    return null;
  }

  let mapping =
    (await getMappingByMessageId(mailboxEmail, email.internetMessageId)) ||
    (await getMappingByConversationId(mailboxEmail, email.conversationId));
  const now = Date.now();

  if (!mapping) {
    mapping = {
      id: randomBase64Url(16),
      mailboxEmail,
      ticketId,
      ticketNumber: ticketNumber || String(ticketId),
      conversationId: email.conversationId || "",
      normalizedSubject: email.normalizedSubject || "",
      actionMode: normalizeActionMode(actionMode || email.actionMode),
      syncedMessageIds: new Set(),
      createdAt: now,
      updatedAt: now,
    };
  }

  mapping.mailboxEmail = mailboxEmail;
  mapping.ticketId = ticketId;
  mapping.ticketNumber = ticketNumber || mapping.ticketNumber || String(ticketId);
  mapping.conversationId = email.conversationId || mapping.conversationId || "";
  mapping.normalizedSubject = email.normalizedSubject || mapping.normalizedSubject || "";
  mapping.actionMode = normalizeActionMode(actionMode || email.actionMode);
  mapping.updatedAt = now;
  await authStore.saveConversationMapping(mapping);
  await markEmailSynced(mapping, email, { includeThreadMessageIds });

  return mapping;
}

async function markEmailSynced(mapping, email, options = {}) {
  const mailboxEmail = normalizeMailboxEmail(mapping.mailboxEmail);
  const messageIds = [email.internetMessageId, email.itemId];

  if (options.includeThreadMessageIds) {
    messageIds.push(...email.inReplyToMessageIds, ...email.referenceMessageIds);
  }

  for (const messageId of messageIds) {
    for (const messageIdKey of normalizeMessageIdKeys(messageId)) {
      mapping.syncedMessageIds.add(messageIdKey);
      await authStore.saveMessageMapping({
        mailboxEmail,
        mappingId: mapping.id,
        messageIdKey,
      });
    }
  }

  if (email.conversationId) {
    mapping.conversationId = email.conversationId;
    await authStore.saveConversationMapping(mapping);
  }

  mapping.updatedAt = Date.now();
  await authStore.saveConversationMapping(mapping);
}

async function backfillRecoveredConversationMapping(mapping, email) {
  const mailboxEmail = normalizeMailboxEmail(mapping.mailboxEmail);
  const threadMessageIds = normalizeMessageIdList([
    ...email.inReplyToMessageIds,
    ...email.referenceMessageIds,
  ]);

  for (const messageId of threadMessageIds) {
    for (const messageIdKey of normalizeMessageIdKeys(messageId)) {
      mapping.syncedMessageIds.add(messageIdKey);
      await authStore.saveMessageMapping({
        mailboxEmail,
        mappingId: mapping.id,
        messageIdKey,
      });
    }
  }

  if (email.conversationId) {
    mapping.conversationId = email.conversationId;
  }
  mapping.updatedAt = Date.now();
  await authStore.saveConversationMapping(mapping);
}

async function findConversationMappingForEmail(email) {
  const mailboxEmail = normalizeMailboxEmail(email.mailboxEmail);

  if (!mailboxEmail) {
    return null;
  }

  const existingMessageMapping = await getMappingByMessageId(mailboxEmail, email.internetMessageId);
  if (existingMessageMapping) {
    return {
      mapping: existingMessageMapping,
      status: "already-attached",
    };
  }

  const threadMessageIds = email.inReplyToMessageIds.concat(email.referenceMessageIds);
  for (const messageId of threadMessageIds) {
    const mapping = await getMappingByMessageId(mailboxEmail, messageId);
    if (mapping) {
      return {
        mapping,
        status: "match",
      };
    }
  }

  const conversationMapping = await getMappingByConversationId(mailboxEmail, email.conversationId);
  if (conversationMapping) {
    return {
      mapping: conversationMapping,
      status: "match",
    };
  }

  return null;
}

async function getMappingByMessageId(mailboxEmail, messageId) {
  const messageIdKeys = normalizeMessageIdKeys(messageId);
  if (!mailboxEmail || !messageIdKeys.length) {
    return null;
  }
  for (const messageIdKey of messageIdKeys) {
    const mapping = await authStore.getMappingByMessageId(mailboxEmail, messageIdKey);
    if (mapping) {
      return mapping;
    }
  }
  return null;
}

function getMappingByConversationId(mailboxEmail, conversationId) {
  if (!mailboxEmail || !conversationId) {
    return null;
  }

  return authStore.getMappingByConversationId(mailboxEmail, conversationId);
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

async function buildEmailActionWithInlineImages(record, ticketId, email, input, options = {}) {
  const actionMode = normalizeActionMode(input && input.actionMode);
  const selectedBody = getEmailBodyForAction(email, options.bodyMode || "trimmed");
  const attachmentStartedAt = Date.now();
  const attachmentCount = getSendDiagnosticAttachmentCount(input);
  const onEmailAttachmentDiagnostic =
    typeof options.onEmailAttachmentDiagnostic === "function"
      ? options.onEmailAttachmentDiagnostic
      : () => {};
  let emailAttachmentResult;
  onEmailAttachmentDiagnostic("commit-validation-started", {
    count: attachmentCount,
    outcome: "started",
    selected: attachmentCount,
    stage: "commit-validation",
  });
  try {
    emailAttachmentResult = await resolveEmailAttachments({
      haloTenant: record.haloUrl,
      input,
      store: authStore,
      ticketId,
      userId: record.userId,
    });
  } catch (error) {
    onEmailAttachmentDiagnostic("commit-validation-failed", {
      count: attachmentCount,
      durationMs: Date.now() - attachmentStartedAt,
      failed: attachmentCount || 1,
      outcome: getEmailAttachmentDiagnosticOutcome(error),
      selected: attachmentCount,
      stage: "commit-validation",
    });
    throw error;
  }
  onEmailAttachmentDiagnostic("commit-validation-completed", {
    count: emailAttachmentResult.summary.selected,
    durationMs: Date.now() - attachmentStartedAt,
    outcome: "ready",
    prepared: emailAttachmentResult.summary.prepared,
    selected: emailAttachmentResult.summary.selected,
    skipped: emailAttachmentResult.summary.skipped,
    stage: "commit-validation",
  });
  let temporaryHaloAttachments = [];
  let inlineResult;
  try {
    temporaryHaloAttachments = await uploadStagedEmailAttachments(
      record,
      emailAttachmentResult,
      onEmailAttachmentDiagnostic,
      actionMode
    );
    inlineResult = await resolveInlineImages({
      addWarningFooter: Boolean(options.addWarningFooter),
      bodyHtml: selectedBody.bodyHtml || textToHtml(selectedBody.bodyText),
      haloTenant: record.haloUrl,
      input,
      onDiagnostic: options.onDiagnostic,
      store: authStore,
      ticketId,
      showForUsers: actionMode !== "private-note",
      uploadImage: (image) => uploadHaloInlineImage(record, image, actionMode),
    });
  } catch (error) {
    await cleanupHaloAttachments(record, temporaryHaloAttachments);
    if (emailAttachmentResult.prefetchKey) {
      await authStore.releaseEmailAttachmentPrefetchCommit(emailAttachmentResult.prefetchKey);
    }
    throw error;
  }
  emailAttachmentResult.summary.attached = temporaryHaloAttachments.length;
  const hasAttachmentWarnings = emailAttachmentResult.summary.skipped > 0;
  const resolvedBodyHtml =
    options.addWarningFooter && hasAttachmentWarnings
      ? appendEmailAttachmentWarningFooter(inlineResult.bodyHtml)
      : inlineResult.bodyHtml;
  const actionStart = Date.now();
  const payload = buildEmailActionPayload(ticketId, email, {
    actionMode,
    attachments: temporaryHaloAttachments,
    bodyHtml: resolvedBodyHtml,
    bodyMode: "full",
    bodyText: selectedBody.bodyText,
  });

  return {
    emailAttachmentPrefetchKey: emailAttachmentResult.prefetchKey,
    emailAttachments: emailAttachmentResult.summary,
    inlineImages: inlineResult.summary,
    payload,
    temporaryHaloAttachments,
    timings: {
      ...inlineResult.timings,
      outlookReadMs: normalizeDiagnosticTiming(
        input && input.inlineImageTimings && input.inlineImageTimings.outlookReadMs
      ),
      hashingMs: normalizeDiagnosticTiming(
        input && input.inlineImageTimings && input.inlineImageTimings.hashingMs
      ),
      actionPayloadMs: Date.now() - actionStart,
    },
  };
}

async function uploadStagedEmailAttachments(
  record,
  attachmentResult,
  onDiagnostic = () => {},
  actionMode = "email"
) {
  const items = attachmentResult.stagedItems || [];
  if (!items.length) {
    onDiagnostic("commit-upload-completed", {
      count: 0,
      durationMs: 0,
      outcome: "ready",
      stage: "commit-upload",
      uploaded: 0,
    });
    return [];
  }
  const uploadStartedAt = Date.now();
  const uploaded = [];
  let nextIndex = 0;
  let firstError = null;
  const worker = async () => {
    while (!firstError && nextIndex < items.length) {
      const item = items[nextIndex++];
      let bytes;
      try {
        bytes = decryptStagedAttachment(tokenCrypto, attachmentResult.record, item);
        const haloAttachment = await uploadHaloEmailAttachment(
          record,
          {
            contentBase64: bytes.toString("base64"),
            contentType: item.contentType,
            decodedSize: bytes.length,
            name: item.filename,
          },
          actionMode
        );
        uploaded.push(toHaloActionAttachment(haloAttachment, actionMode));
      } catch (error) {
        firstError = firstError || error;
      } finally {
        if (bytes) {
          bytes.fill(0);
        }
      }
    }
  };
  onDiagnostic("commit-upload-started", {
    count: items.length,
    outcome: "started",
    stage: "commit-upload",
  });
  await Promise.all(Array.from({ length: Math.min(2, items.length) }, () => worker()));
  if (firstError) {
    const cleaned = await cleanupHaloAttachments(record, uploaded);
    onDiagnostic("commit-upload-failed", {
      count: items.length,
      durationMs: Date.now() - uploadStartedAt,
      failed: 1,
      outcome: getEmailAttachmentDiagnosticOutcome(firstError, "halo-upload-failed"),
      stage: cleaned ? "commit-upload" : "commit-cleanup",
      uploaded: uploaded.length,
    });
    uploaded.length = 0;
    throw firstError;
  }
  onDiagnostic("commit-upload-completed", {
    count: items.length,
    durationMs: Date.now() - uploadStartedAt,
    outcome: "ok",
    stage: "commit-upload",
    uploaded: uploaded.length,
  });
  return uploaded;
}

function toHaloActionAttachment(upload, actionMode = "email") {
  const rawId = String(upload.id || "");
  const id = /^\d+$/.test(rawId) && Number.isSafeInteger(Number(rawId)) ? Number(rawId) : rawId;
  return {
    filename: upload.filename,
    filesize: Number(upload.filesize || upload.decodedSize || 0),
    showforusers: actionMode === "private-note" ? false : null,
    _uploading: false,
    _tempid: null,
    showonchild: false,
    showonrelated: false,
    data_base64: null,
    id,
    data: null,
  };
}

async function createPreparedHaloAction(record, preparedAction, createAction) {
  let actionCreated = false;
  try {
    const payload = await createAction();
    actionCreated = true;
    if (preparedAction.emailAttachmentPrefetchKey) {
      await authStore.markEmailAttachmentPrefetchActionCreated(
        preparedAction.emailAttachmentPrefetchKey,
        getCreatedActionId(payload) || "created"
      );
    }
    return payload;
  } catch (error) {
    if (!actionCreated) {
      await cleanupHaloAttachments(record, preparedAction.temporaryHaloAttachments);
      if (preparedAction.emailAttachmentPrefetchKey) {
        await authStore.releaseEmailAttachmentPrefetchCommit(
          preparedAction.emailAttachmentPrefetchKey
        );
      }
    }
    throw error;
  }
}

function normalizeDiagnosticTiming(value) {
  const duration = Number(value);
  if (!Number.isFinite(duration) || duration < 0) {
    return 0;
  }
  return Math.min(Math.round(duration), 60_000);
}

function buildEmailActionPayload(ticketId, email, options = {}) {
  const body = getEmailBodyForAction(email, options.bodyMode || "trimmed");
  const htmlBody =
    options.bodyHtml || body.bodyHtml || textToHtml(options.bodyText || body.bodyText);
  const noteHtml = buildEmailNoteHtml(email, htmlBody);
  const note = buildEmailNoteText(email);
  const subject = email.subject || email.normalizedSubject || "(no subject)";
  const actionDatetime = getCurrentActionDateTime();

  const actionMode = normalizeActionMode(options.actionMode);
  const payload = {
    ticket_id: ticketId,
    outcome: actionMode === "private-note" ? "Private Note" : "Email",
    sendemail: false,
    note,
    note_html: noteHtml,
    whowith: email.from,
    datetime: actionDatetime,
  };
  if (actionMode === "private-note") {
    payload.hiddenfromuser = true;
  } else {
    payload.emailbody_html = noteHtml;
    payload.emailsubject = subject;
    payload.email_message_id = email.internetMessageId;
    payload.actioninternetmessageid = email.internetMessageId;
    payload.emailtolistall = email.to.join("; ");
  }
  if (Array.isArray(options.attachments) && options.attachments.length) {
    payload.attachments = options.attachments;
  }
  return payload;
}

function appendEmailAttachmentWarningFooter(bodyHtml) {
  return `${bodyHtml}<p style="color:#8a5a00;font-size:12px;"><em>One or more email attachments could not be added to Halo.</em></p>`;
}

async function consumeEmailAttachmentPrefetch(preparedAction) {
  await consumeEmailAttachmentPrefetchKey(
    preparedAction && preparedAction.emailAttachmentPrefetchKey
  );
}

async function consumeEmailAttachmentPrefetchKey(prefetchKey) {
  const normalizedKey = stringifyField(prefetchKey);
  if (normalizedKey) {
    await authStore.consumeEmailAttachmentPrefetch(normalizedKey);
  }
}

function getAttachedStatus(preparedAction) {
  const imageWarnings = preparedAction.inlineImages.failed > 0;
  const attachmentWarnings =
    preparedAction.emailAttachments.failed > 0 || preparedAction.emailAttachments.skipped > 0;
  if (imageWarnings && attachmentWarnings) {
    return "attached-with-warnings";
  }
  if (attachmentWarnings) {
    return "attached-with-attachment-warnings";
  }
  if (imageWarnings) {
    return "attached-with-image-warnings";
  }
  return "attached";
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
    getPatternIndex(
      value,
      /<[^>]+\bclass=["'][^"']*(?:gmail_quote|moz-cite-prefix|yahoo_quoted)[^"']*["'][^>]*>/i
    ),
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
  const match = /(?:^|[\r\n]|<[^>]+>)\s*(?:<b>|<strong>)?From:(?:<\/b>|<\/strong>)?/i.exec(value);

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
  const direct = stringifyField(getFirstField(action || {}, ["id", "action_id", "actionid"]));
  if (direct) {
    return direct;
  }
  return stringifyField(
    findPayloadValue(
      payload,
      (key, value) =>
        ["id", "action_id", "actionid"].includes(String(key || "").toLowerCase()) &&
        ["number", "string"].includes(typeof value)
    )
  );
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normalizeTickets(payload, currentAgentId, options = {}) {
  let tickets = getTicketArray(payload);

  if (options.openOnly !== false) {
    tickets = tickets.filter((ticket) => isOpenTicket(ticket));
  }

  if (currentAgentId) {
    tickets = tickets.filter((ticket) => isAssignedToAgent(ticket, currentAgentId));
  }

  return tickets
    .map((ticket) => toTicketSummary(ticket))
    .filter((ticket) => ticket.id || ticket.ticketNumber || ticket.summary);
}

function normalizeTicketsForLifecycle(payload, lifecycle) {
  const tickets = normalizeTickets(payload, null, { openOnly: lifecycle === "open" });
  if (lifecycle === "closed") {
    return tickets.filter((ticket) => ticket.lifecycle === "closed");
  }
  return tickets;
}

function promoteExactTicketMatches(tickets, query) {
  const expected = normalizeTicketNumber(query);
  const exactMatches = tickets.filter((ticket) => {
    return [ticket.id, ticket.ticketNumber]
      .map((value) => normalizeTicketNumber(value))
      .some((value) => value && value === expected);
  });

  return exactMatches.length
    ? [...exactMatches, ...tickets.filter((ticket) => !exactMatches.includes(ticket))]
    : tickets;
}

function normalizeTicketNumber(value) {
  return stringifyField(value)
    .replace(/^#/, "")
    .replace(/^0+(?=\d)/, "")
    .toLowerCase();
}

function getTicketArray(payload) {
  if (Array.isArray(payload)) {
    return payload;
  }

  if (!payload || typeof payload !== "object") {
    return [];
  }

  const possibleKeys = [
    "tickets",
    "Tickets",
    "items",
    "Items",
    "results",
    "Results",
    "data",
    "Data",
  ];
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

  if (closedValues.some((value) => isTruthyTicketFlag(value))) {
    return false;
  }

  const statusText = getTicketStatus(ticket).toLowerCase();
  if (!statusText) {
    return true;
  }

  return !["closed", "resolved", "complete", "completed", "cancelled", "canceled"].some((word) =>
    new RegExp(`\\b${word}\\b`, "i").test(statusText)
  );
}

function isTruthyTicketFlag(value) {
  if (value === true || value === 1) {
    return true;
  }
  return typeof value === "string" && ["1", "true"].includes(value.trim().toLowerCase());
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
    lifecycle: isOpenTicket(ticket) ? "open" : "closed",
    client: getNamedTicketField(
      ticket,
      ["client_name", "clientname", "clientName", "customer_name", "customername", "customerName"],
      ["client", "customer"]
    ),
    agent: getNamedTicketField(
      ticket,
      [
        "agent_name",
        "agentname",
        "agentName",
        "assigned_agent_name",
        "assignedagentname",
        "assignedAgentName",
        "owner_name",
        "ownername",
        "ownerName",
      ],
      ["agent", "assigned_agent", "assignedAgent", "owner"]
    ),
  };
}

function getNamedTicketField(ticket, flatKeys, referenceKeys) {
  const flatValue = stringifyField(getFirstField(ticket, flatKeys));
  if (flatValue) {
    return flatValue;
  }

  for (const key of referenceKeys) {
    const reference = ticket && ticket[key];
    if (typeof reference === "string" && reference.trim()) {
      return reference.trim();
    }
    if (reference && typeof reference === "object") {
      const name = stringifyField(
        getFirstField(reference, ["name", "label", "display_name", "displayname", "displayName"])
      );
      if (name) {
        return name;
      }
    }
  }

  return "";
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

function normalizeActionMode(value) {
  return value === "private-note" ? "private-note" : "email";
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

function normalizeHaloUrl(value, settingName = "Halo URL") {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error(`${settingName} must be set.`);
  }

  let url;
  try {
    url = new URL(value.trim());
  } catch (error) {
    throw new Error(`${settingName} must be a valid URL, including https://.`);
  }

  if (url.protocol !== "https:") {
    throw new Error(`${settingName} must use https://.`);
  }

  return url.origin;
}

function normalizeClientId(value, settingName = "Halo API application client ID") {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error(`${settingName} must be set.`);
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
  if (process.env.PUBLIC_BASE_URL) {
    return process.env.PUBLIC_BASE_URL.replace(/\/+$/, "");
  }

  const forwardedProto = req.headers["x-forwarded-proto"];
  const proto = Array.isArray(forwardedProto) ? forwardedProto[0] : forwardedProto;
  const socketProto = req.socket ? (req.socket.encrypted ? "https" : "http") : "https";
  const requestProto = proto || req.protocol || socketProto;
  return `${requestProto}://${req.headers.host || "localhost:3000"}`;
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

async function requireMicrosoftUser(req) {
  const user = await getMicrosoftUserFromRequest(req);

  if (!user) {
    throw new RequestError("Microsoft add-in authentication is required.", 401);
  }

  return user;
}

async function getMicrosoftUserFromRequest(req) {
  const token = getBearerToken(req);

  if (!token) {
    return null;
  }

  let claims;
  try {
    claims = await microsoftAuthVerifier.verify(token);
  } catch (error) {
    throw new RequestError(`Microsoft add-in authentication failed: ${publicError(error)}`, 401);
  }

  const tenantId = stringifyField(claims.tid);
  const objectId = stringifyField(claims.oid || claims.sub);

  if (!tenantId || !objectId) {
    throw new RequestError("Microsoft add-in authentication did not include a stable user.", 401);
  }

  return authStore.upsertUser({
    displayName: stringifyField(claims.name),
    email: stringifyField(claims.preferred_username || claims.email || claims.upn),
    objectId,
    tenantId,
  });
}

function getBearerToken(req) {
  const header = req.headers.authorization || req.headers.Authorization || "";
  const match = /^Bearer\s+(.+)$/i.exec(Array.isArray(header) ? header[0] : header);
  return match ? match[1].trim() : "";
}

async function getOrCreateSessionRecord(req, res) {
  const existingRecord = await getSessionRecord(req);
  if (existingRecord) {
    return existingRecord;
  }

  const user = await getMicrosoftUserFromRequest(req);
  if (!user) {
    return null;
  }

  const grant = await authStore.getGrantByUserId(user.id);
  if (!grant) {
    return null;
  }

  return (await createSessionForGrant(res, user.id, grant)).record;
}

async function getSessionOrBearerGrant(req) {
  return (await getSessionRecord(req)) || (await getBearerGrantRecord(req));
}

async function getBearerGrantRecord(req) {
  const user = await getMicrosoftUserFromRequest(req);
  if (!user) {
    return null;
  }

  const grant = await authStore.getGrantByUserId(user.id);
  if (!grant) {
    return null;
  }

  return {
    ...grant,
    expiresAt: Date.now() + SESSION_TTL_MS,
  };
}

async function createSessionForGrant(res, userId, grant) {
  const sessionId = randomBase64Url(32);
  const sessionHash = hashSessionId(sessionId);
  const expiresAt = Date.now() + SESSION_TTL_MS;

  await authStore.createSession({
    expiresAt,
    sessionHash,
    userId,
  });

  res.setHeader("Set-Cookie", serializeSessionCookie(sessionId, Math.floor(SESSION_TTL_MS / 1000)));

  return {
    backgroundSessionId: await createBackgroundSession(sessionHash, expiresAt),
    expiresAt,
    record: {
      ...grant,
      expiresAt,
      sessionHash,
      userId,
    },
    sessionId,
  };
}

function encryptJson(value) {
  return tokenCrypto.encryptJson(value);
}

function decryptJson(value) {
  return tokenCrypto.decryptJson(value);
}

async function getSessionRecord(req) {
  return getSessionRecordBySessionId(getSessionIdFromRequest(req));
}

async function getSessionRecordBySessionId(sessionId) {
  if (!sessionId) {
    return null;
  }

  const sessionHash = hashSessionId(sessionId);
  const record = await authStore.getSessionWithGrant(sessionHash);

  if (!record) {
    return null;
  }

  if (record.expiresAt <= Date.now()) {
    await authStore.deleteSession(sessionHash);
    return null;
  }

  return record;
}

async function createBackgroundSession(sessionHash, expiresAt) {
  const backgroundSessionId = randomBase64Url(32);
  await authStore.createBackgroundSession({
    backgroundSessionHash: hashBackgroundSessionId(backgroundSessionId),
    sessionHash,
    expiresAt,
  });

  return backgroundSessionId;
}

async function createBackgroundSessionForRequest(req) {
  const sessionId = getSessionIdFromRequest(req);
  const record = await getSessionRecordBySessionId(sessionId);

  if (!sessionId || !record) {
    return "";
  }

  return createBackgroundSession(hashSessionId(sessionId), record.expiresAt);
}

async function getBackgroundSessionRecord(backgroundSessionId) {
  const backgroundSessionHash = hashBackgroundSessionId(backgroundSessionId);
  const record = await authStore.getBackgroundSessionWithGrant(backgroundSessionHash);
  if (
    !record ||
    record.expiresAt <= Date.now() ||
    (record.backgroundExpiresAt && record.backgroundExpiresAt <= Date.now())
  ) {
    await authStore.cleanExpired(Date.now());
    return null;
  }

  return record;
}

async function deleteBackgroundSessionsForSessionHash(sessionHash) {
  await authStore.deleteBackgroundSessionsForSessionHash(sessionHash);
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
  return crypto
    .createHash("sha256")
    .update(backgroundSessionId || "")
    .digest("hex");
}

function randomBase64Url(byteLength) {
  return base64Url(crypto.randomBytes(byteLength));
}

function base64Url(buffer) {
  return buffer.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

async function cleanExpiredRecords() {
  if (cleaningExpiredRecords) {
    return;
  }
  cleaningExpiredRecords = true;
  const now = Date.now();
  try {
    deleteExpired(pendingStates, now);
    deleteExpired(handoffs, now);
    if (authStore) {
      await authStore.cleanExpired(now);
      await cleanupExpiredEmailAttachmentPrefetches(now);
    }
  } finally {
    cleaningExpiredRecords = false;
  }
}

async function cleanupExpiredEmailAttachmentPrefetches(now) {
  if (cleaningEmailAttachmentPrefetches || !authStore) {
    return;
  }
  cleaningEmailAttachmentPrefetches = true;
  try {
    const candidates = await authStore.getEmailAttachmentCleanupCandidates(now, 20);
    for (const candidate of candidates) {
      await authStore.markEmailAttachmentPrefetchForCleanup(candidate.prefetchKey);
      const legacyAttachments = candidate.items
        .filter((item) => item.haloAttachmentId)
        .map((item) => ({ attachmentKey: item.attachmentKey, id: item.haloAttachmentId }));
      if (candidate.stagingVersion === 2 && legacyAttachments.length === 0) {
        await authStore.deleteEmailAttachmentPrefetch(candidate.prefetchKey);
        continue;
      }
      const record = await authStore.getGrantByUserId(candidate.userId);
      if (!record) {
        continue;
      }
      const cleaned = await cleanupStoredHaloAttachments(
        record,
        candidate.prefetchKey,
        legacyAttachments
      );
      if (cleaned) {
        await authStore.deleteEmailAttachmentPrefetch(candidate.prefetchKey);
      }
    }
    const removedCandidates = await authStore.getEmailAttachmentRemovedCleanupCandidates(20);
    for (const candidate of removedCandidates) {
      const record = await authStore.getGrantByUserId(candidate.userId);
      if (!record) {
        continue;
      }
      await cleanupStoredHaloAttachments(
        record,
        candidate.prefetchKey,
        candidate.items
          .filter((item) => item.status === "removed" && item.haloAttachmentId)
          .map((item) => ({ attachmentKey: item.attachmentKey, id: item.haloAttachmentId }))
      );
    }
  } finally {
    cleaningEmailAttachmentPrefetches = false;
  }
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
  if (error instanceof RequestError) {
    return error.status;
  }
  if (error && error.isRequestError) {
    const status = Number(error.status);
    return Number.isInteger(status) && status >= 400 && status <= 599 ? status : fallbackStatus;
  }
  return fallbackStatus;
}

function getAttachmentFailureStatus(error) {
  return error && error.code === "attachments-not-ready" ? "attachments-not-ready" : "failed";
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

    return new HttpError(
      `${messagePrefix}: HTTP ${statusLabel} - ${responseError}`,
      responseDetails.status,
      {
        bodyExcerpt: getBodyExcerpt(responseDetails.bodyText),
        contentType: responseDetails.contentType || "(none)",
        endpoint: responseDetails.requestUrl,
        phase,
        status: responseDetails.status,
        statusText: responseDetails.statusText || "",
      }
    );
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
  _test: {
    buildEmailActionPayload,
    buildTicketsPath,
    createEmailAttachmentDiagnosticLogger,
    createSendDiagnosticLogger,
    extractHaloAttachmentId,
    extractHaloAttachmentUrl,
    fetchHaloAttachment,
    getTicketSearchQuery,
    normalizeTicketsForLifecycle,
    promoteExactTicketMatches,
    sanitizeEmailAttachmentDiagnosticMessage,
    toHaloActionAttachment,
    validateHaloInlineImageUrl,
  },
};
