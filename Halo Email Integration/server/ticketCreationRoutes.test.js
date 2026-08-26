const assert = require("node:assert");
const crypto = require("node:crypto");

process.env.NODE_ENV = "test";

const { registerHaloAuthRoutes } = require("./haloAuth");
const { createTestStore } = require("./testDatabase");
const { createTokenCrypto } = require("./tokenCrypto");

function createApp() {
  const routes = { DELETE: new Map(), GET: new Map(), PATCH: new Map(), POST: new Map() };
  return {
    locals: {},
    routes,
    get(path, handler) {
      routes.GET.set(path, handler);
    },
    post(path, handler) {
      routes.POST.set(path, handler);
    },
    patch(path, handler) {
      routes.PATCH.set(path, handler);
    },
    delete(path, handler) {
      routes.DELETE.set(path, handler);
    },
  };
}

function request(url, body, params = {}) {
  return {
    body,
    headers: { authorization: "Bearer microsoft-token", host: "localhost:3000" },
    originalUrl: url,
    params,
    url,
  };
}

function response() {
  return {
    body: undefined,
    statusCode: 200,
    status(code) {
      this.statusCode = code;
      return this;
    },
    json(body) {
      this.body = body;
      return this;
    },
    setHeader() {},
    send(body) {
      this.body = body;
      return this;
    },
  };
}

async function invoke(app, method, route, url, body, params) {
  const handler = app.routes[method].get(route);
  assert(handler, `Missing ${method} ${route}`);
  const res = response();
  await handler(request(url, body, params), res);
  return res;
}

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

async function run() {
  const key = Buffer.alloc(32, 9).toString("base64url");
  const tokenCrypto = createTokenCrypto({ HALO_TOKEN_ENCRYPTION_KEY: key });
  const store = await createTestStore();
  const saveConversationMapping = store.saveConversationMapping;
  const updateTicketCreationIntent = store.updateTicketCreationIntent;
  let failNextConversationMapping = false;
  let failActionReceiptUpdateForOperation = "";
  store.saveConversationMapping = (value) => {
    if (failNextConversationMapping) {
      failNextConversationMapping = false;
      throw new Error("Temporary mapping failure");
    }
    return saveConversationMapping(value);
  };
  store.updateTicketCreationIntent = (operationId, updates, scope) => {
    if (
      operationId === failActionReceiptUpdateForOperation &&
      updates &&
      updates.actionId
    ) {
      failActionReceiptUpdateForOperation = "";
      throw new Error("Temporary action receipt persistence failure");
    }
    return updateTicketCreationIntent(operationId, updates, scope);
  };
  const user = await store.upsertUser({
    displayName: "Support User",
    email: "support@example.com",
    objectId: "object-id",
    tenantId: "tenant-id",
  });
  await store.saveHaloGrant({
    clientId: "halo-client-id",
    encryptedToken: tokenCrypto.encryptJson({
      access_token: "halo-token",
      expires_at: Date.now() + 60 * 60 * 1000,
    }),
    haloUrl: "https://customer.halopsa.com",
    scope: "all",
    userId: user.id,
  });

  const app = createApp();
  registerHaloAuthRoutes(app, {
    env: {
      HALO_CLIENT_ID: "halo-client-id",
      HALO_TOKEN_ENCRYPTION_KEY: key,
      HALO_URL: "https://customer.halopsa.com",
      NODE_ENV: "test",
    },
    microsoftAuth: { clientId: "addin-client-id" },
    microsoftAuthVerifier: {
      async verify() {
        return {
          aud: "addin-client-id",
          email: "support@example.com",
          name: "Support User",
          oid: "object-id",
          preferred_username: "support@example.com",
          tid: "tenant-id",
        };
      },
    },
    store,
    tokenCrypto,
  });

  const originalFetch = global.fetch;
  let ticketCreates = 0;
  let actionCreates = 0;
  let attachmentCreates = 0;
  const createdSummaries = [];
  let actionFailuresRemaining = 0;
  let ticketTypeReads = 0;
  let schemaReads = 0;
  let fieldReads = 0;
  let optionReads = 0;
  let failTicketTypeReads = false;
  let failSchemaReads = false;
  let failFieldReads = false;
  let addNewMandatoryField = false;
  let duplicateRequesters = false;
  global.fetch = async (requestUrl, options = {}) => {
    const url = String(requestUrl);
    if (url.includes("/api/TicketType?")) {
      ticketTypeReads += 1;
      if (failTicketTypeReads) {
        return json({ message: "Temporary ticket-type failure" }, 503);
      }
      return json([{ id: 12, name: "Project Engineer", cancreate: true }]);
    }
    if (url.includes("/api/TicketType/12?")) {
      schemaReads += 1;
      if (failSchemaReads) {
        return json({ message: "Temporary schema failure" }, 503);
      }
      const fields = [
        {
          id: 9100,
          fieldid: 100,
          fieldname: "CFProjectCode",
          mandatory: true,
          technew: 3,
          fieldinfo: {
            id: 100,
            custom: 1,
            label: "Project code",
            name: "CFProjectCode",
            type: 0,
          },
        },
        {
          fieldid: 3,
          seq: 2,
          technew: 2,
          fieldinfo: {
            id: 3,
            custom: 0,
            label: "Details",
            name: "symptom2",
            type: 1,
          },
        },
        {
          fieldid: 28,
          seq: 3,
          technew: 1,
          fieldinfo: {
            id: 28,
            custom: 0,
            label: "Urgency",
            name: "urgency",
            type: 1,
          },
        },
        {
          fieldid: 310,
          seq: 4,
          technew: 1,
          fieldinfo: {
            id: 310,
            custom: 1,
            label: "Weeklookup",
            name: "CFWeeklookup",
            type: 2,
          },
        },
        {
          fieldid: 4,
          seq: 5,
          technew: 1,
          fieldinfo: { id: 4, custom: 0, label: "Asset", name: "N/A", type: -1 },
        },
        {
          fieldid: 13,
          seq: 6,
          technew: 1,
          fieldinfo: { id: 13, custom: 0, label: "Team", name: "sectio_", type: -1 },
        },
        {
          fieldid: 16,
          seq: 7,
          technew: 1,
          fieldinfo: {
            id: 16,
            custom: 0,
            label: "Estimated Time",
            name: "estimate",
            type: -1,
          },
        },
      ];
      if (addNewMandatoryField) {
        fields.push({
          id: 101,
          iscustom: true,
          label: "New mandatory value",
          fieldname: "CFNewMandatoryValue",
          fieldtype: "Text",
          mandatory: true,
        });
      }
      return json([{ id: 12, name: "Project Engineer", fields }]);
    }
    if (url.includes("/api/TicketTypeField?")) {
      fieldReads += 1;
      if (failFieldReads) {
        return json({ message: "Temporary ticket-field failure" }, 503);
      }
      return json([]);
    }
    if (url.includes("/api/Lookup?lookupid=27")) {
      optionReads += 1;
      return json([
        { id: 0, name: "Unknown" },
        { id: 1, name: "1. High" },
        { id: 2, name: "2. Medium" },
        { id: 3, name: "3. Low" },
      ]);
    }
    if (url.includes("/api/FieldInfo/310?")) {
      optionReads += 1;
      return json([
        { id: "02/01/2000 12:00:00 AM", name: "-1" },
        { id: "03/01/2000 12:00:00 AM", name: "0" },
        { id: "04/01/2000 12:00:00 AM", name: "1" },
      ]);
    }
    if (url.includes("/api/Users?")) {
      const users = [
        {
          id: 50,
          name: "Customer Contact",
          email: "customer@example.com",
          client_id: 60,
        },
      ];
      if (duplicateRequesters) {
        users.push({
          id: 51,
          name: "Duplicate Customer Contact",
          email: "customer@example.com",
          client_id: 60,
        });
      }
      return json({ users });
    }
    if (url.includes("/api/Asset?")) {
      return json({ assets: [{ id: 70, name: "Laptop", inventory_number: "LT-70" }] });
    }
    if (url.includes("/api/Team?")) {
      return json({ teams: [{ id: 80, name: "Helpdesk" }] });
    }
    if (url.endsWith("/api/Tickets") && options.method === "POST") {
      ticketCreates += 1;
      const payload = JSON.parse(options.body)[0];
      createdSummaries.push(payload.summary);
      assert.strictEqual(payload.tickettype_id, 12);
      assert.strictEqual(payload.user_id, 50);
      assert.match(payload.details, /Customer-provided detail/);
      assert.match(payload.details, /Created from an Outlook email/);
      assert.deepStrictEqual(payload.customfields, [
        { id: 100, value: "PROJ-42" },
        { id: 310, value: "04/01/2000 12:00:00 AM" },
      ]);
      return json({ data: { tickets: [{ id: 2001, ticketnumber: "T2001" }] } }, 201);
    }
    if (url.endsWith("/api/Attachment") && options.method === "POST") {
      attachmentCreates += 1;
      const payload = JSON.parse(options.body)[0];
      return json(
        [
          {
            id: 4000 + attachmentCreates,
            filename: payload.filename,
            filesize: payload.filesize,
          },
        ],
        201
      );
    }
    if (url.endsWith("/api/Actions") && options.method === "POST") {
      actionCreates += 1;
      if (actionFailuresRemaining > 0) {
        actionFailuresRemaining -= 1;
        return json({ message: "Temporary action failure" }, 503);
      }
      const payload = JSON.parse(options.body)[0];
      assert.strictEqual(payload.ticket_id, 2001);
      assert.match(payload.note_html, /Composed email/);
      return json({ actions: [{ id: 3001 }] }, 201);
    }
    throw new Error(`Unexpected Halo request: ${url}`);
  };

  try {
    const types = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/types",
      "/api/halo/ticket-creation/types"
    );
    assert.strictEqual(types.statusCode, 200);
    assert.strictEqual(types.body.types[0].name, "Project Engineer");
    assert.strictEqual(ticketTypeReads, 1);

    const cachedTypes = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/types",
      "/api/halo/ticket-creation/types"
    );
    assert.strictEqual(cachedTypes.body.cached, true);
    assert.strictEqual(ticketTypeReads, 1, "Fresh ticket types should be served from cache.");

    failTicketTypeReads = true;
    const staleTypes = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/types",
      "/api/halo/ticket-creation/types?refresh=1"
    );
    assert.strictEqual(staleTypes.statusCode, 200);
    assert.strictEqual(staleTypes.body.stale, true);
    assert.strictEqual(ticketTypeReads, 3, "A transient metadata call should be retried once.");
    failTicketTypeReads = false;

    const schema = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/types/:typeId/schema",
      "/api/halo/ticket-creation/types/12/schema",
      undefined,
      { typeId: "12" }
    );
    assert.strictEqual(schema.statusCode, 200);
    assert.strictEqual(schema.body.schema.available, true);
    assert.strictEqual(schema.body.schema.fields[0].label, "Project code");
    assert.deepStrictEqual(
      schema.body.schema.fields.find((field) => field.key === "core:urgency").options[0],
      { value: "0", label: "Unknown" }
    );
    assert.deepStrictEqual(
      schema.body.schema.fields
        .find((field) => field.key === "custom:310")
        .options.map((option) => option.label),
      ["-1", "0", "1"]
    );
    for (const key of ["core:asset_id", "core:team_id", "core:estimate"]) {
      assert.strictEqual(
        schema.body.schema.fields.find((field) => field.key === key).required,
        false,
        `${key} should remain optional for this ticket type.`
      );
    }
    assert.strictEqual(
      schema.body.schema.fields.find((field) => field.key === "core:details").recommended,
      true
    );
    assert.strictEqual(schemaReads, 1);
    assert.strictEqual(optionReads, 2);

    const cachedSchema = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/types/:typeId/schema",
      "/api/halo/ticket-creation/types/12/schema",
      undefined,
      { typeId: "12" }
    );
    assert.strictEqual(cachedSchema.body.cached, true);
    assert.strictEqual(schemaReads, 1, "Fresh field metadata should be served from cache.");

    failSchemaReads = true;
    const staleSchema = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/types/:typeId/schema",
      "/api/halo/ticket-creation/types/12/schema?refresh=1",
      undefined,
      { typeId: "12" }
    );
    assert.strictEqual(staleSchema.statusCode, 200);
    assert.strictEqual(staleSchema.body.stale, true);
    assert.strictEqual(schemaReads, 3);
    failSchemaReads = false;

    failFieldReads = true;
    const staleSchemaFromFieldFailure = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/types/:typeId/schema",
      "/api/halo/ticket-creation/types/12/schema?refresh=1",
      undefined,
      { typeId: "12" }
    );
    assert.strictEqual(staleSchemaFromFieldFailure.statusCode, 200);
    assert.strictEqual(staleSchemaFromFieldFailure.body.stale, true);
    assert.strictEqual(
      fieldReads,
      4,
      "Transient field metadata failures must not replace the cache."
    );
    failFieldReads = false;

    const requesters = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/requesters",
      "/api/halo/ticket-creation/requesters?query=customer%40example.com"
    );
    assert.strictEqual(requesters.body.requesters[0].id, "50");

    const assetLookup = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/lookups/:entity",
      "/api/halo/ticket-creation/lookups/asset?query=LT-70",
      undefined,
      { entity: "asset" }
    );
    assert.strictEqual(assetLookup.statusCode, 200);
    assert.deepStrictEqual(assetLookup.body.results[0], {
      id: "70",
      label: "Laptop",
      secondary: "LT-70",
    });

    const teamLookup = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/lookups/:entity",
      "/api/halo/ticket-creation/lookups/team?query=Helpdesk",
      undefined,
      { entity: "team" }
    );
    assert.strictEqual(teamLookup.statusCode, 200);
    assert.deepStrictEqual(teamLookup.body.results[0], {
      id: "80",
      label: "Helpdesk",
      secondary: "",
    });

    const operationId = "create-operation-1";
    const creation = {
      operationId,
      typeId: "12",
      schemaRevision: schema.body.schema.revision,
      summary: "Create project",
      summaryMode: "fixed",
      requesterMode: "explicit",
      requester: requesters.body.requesters[0],
      values: {
        "core:details": "Customer-provided detail",
        "core:urgency": "1",
        "custom:100": "PROJ-42",
        "custom:310": "04/01/2000 12:00:00 AM",
      },
      draftItemId: "draft-item-id",
    };
    const invalidIntent = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents",
      "/api/halo/ticket-creation/intents",
      { ...creation, operationId: "invalid-operation", values: {} }
    );
    assert.strictEqual(invalidIntent.statusCode, 400);
    assert.match(invalidIntent.body.error, /Project code is required/);

    const intent = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents",
      "/api/halo/ticket-creation/intents",
      creation
    );
    assert.strictEqual(intent.statusCode, 201);

    const email = {
      bodyHtml: "<p>Composed email</p>",
      bodyText: "",
      cc: [],
      conversationId: "conversation-1",
      dateTimeCreated: new Date().toISOString(),
      from: { displayName: "Support", emailAddress: "support@example.com" },
      inReplyToMessageIds: [],
      internetHeaders: "",
      internetMessageId: "",
      itemId: "draft-item-id",
      mailboxEmail: "support@example.com",
      normalizedSubject: "Create project",
      referenceMessageIds: [],
      subject: "Create project",
      timeZone: "Europe/London",
      to: [{ displayName: "Customer", emailAddress: "customer@example.com" }],
    };
    const created = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${operationId}/send`,
      email,
      { operationId }
    );
    assert.strictEqual(created.statusCode, 200, created.body.error);
    assert.strictEqual(created.body.ticketNumber, "T2001");
    assert.strictEqual(created.body.actionId, "3001");

    const retry = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${operationId}/send`,
      email,
      { operationId }
    );
    assert.strictEqual(retry.statusCode, 200);
    assert.strictEqual(retry.body.ticketNumber, "T2001");
    assert.strictEqual(ticketCreates, 1);
    assert.strictEqual(actionCreates, 1);

    const automaticSummaryOperationId = "create-operation-automatic-summary";
    await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents",
      "/api/halo/ticket-creation/intents",
      {
        ...creation,
        operationId: automaticSummaryOperationId,
        summaryMode: "auto",
      }
    );
    const automaticSummaryResult = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${automaticSummaryOperationId}/send`,
      { ...email, normalizedSubject: "", subject: "" },
      { operationId: automaticSummaryOperationId }
    );
    assert.strictEqual(automaticSummaryResult.statusCode, 200, automaticSummaryResult.body.error);
    assert.strictEqual(createdSummaries.at(-1), "(no subject)");

    const transientOperationId = "create-operation-transient";
    await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents",
      "/api/halo/ticket-creation/intents",
      { ...creation, operationId: transientOperationId }
    );
    actionFailuresRemaining = 1;
    const transientRetry = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${transientOperationId}/send`,
      email,
      { operationId: transientOperationId }
    );
    assert.strictEqual(transientRetry.statusCode, 200, transientRetry.body.error);
    assert.strictEqual(ticketCreates, 3);
    assert.strictEqual(
      actionCreates,
      4,
      "A transient Halo action response should be retried once."
    );

    const repairOperationId = "create-operation-repair";
    await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents",
      "/api/halo/ticket-creation/intents",
      { ...creation, operationId: repairOperationId }
    );
    actionFailuresRemaining = 2;
    const partial = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${repairOperationId}/send`,
      email,
      { operationId: repairOperationId }
    );
    assert.strictEqual(partial.statusCode, 502);
    assert.strictEqual(partial.body.ticketNumber, "T2001");
    const existingPartialIntent = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents",
      "/api/halo/ticket-creation/intents",
      { ...creation, operationId: repairOperationId, schemaRevision: "stale" }
    );
    assert.strictEqual(existingPartialIntent.statusCode, 200);
    assert.strictEqual(existingPartialIntent.body.status, "partial-failure");
    assert.strictEqual(existingPartialIntent.body.ticketNumber, "T2001");
    const partialAttachmentPrefetch = await invoke(
      app,
      "POST",
      "/api/halo/email-attachments/prefetch/start",
      "/api/halo/email-attachments/prefetch/start",
      {
        creationOperationId: repairOperationId,
        draftItemId: "repair-draft-item-id",
        emailAttachmentFingerprint: "a".repeat(64),
        emailAttachments: [
          {
            attachmentKey: "b".repeat(64),
            attachmentType: "file",
            contentType: "text/plain",
            name: "retry.txt",
            reportedSize: 5,
          },
        ],
        operationId: repairOperationId,
      }
    );
    assert.strictEqual(partialAttachmentPrefetch.statusCode, 200);
    assert.strictEqual(partialAttachmentPrefetch.body.status, "pending");
    assert.strictEqual(partialAttachmentPrefetch.body.ticketId, "0");
    const repaired = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${repairOperationId}/send`,
      email,
      { operationId: repairOperationId }
    );
    assert.strictEqual(repaired.statusCode, 200, repaired.body.error);
    assert.strictEqual(
      ticketCreates,
      4,
      "Repair must reuse the ticket created before the action failed."
    );
    assert.strictEqual(actionCreates, 7);

    const mappingRepairOperationId = "create-operation-mapping-repair";
    await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents",
      "/api/halo/ticket-creation/intents",
      { ...creation, operationId: mappingRepairOperationId }
    );
    failNextConversationMapping = true;
    const mappingPartial = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${mappingRepairOperationId}/send`,
      email,
      { operationId: mappingRepairOperationId }
    );
    assert.strictEqual(mappingPartial.statusCode, 502);
    const mappingRepair = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${mappingRepairOperationId}/send`,
      email,
      { operationId: mappingRepairOperationId }
    );
    assert.strictEqual(mappingRepair.statusCode, 200, mappingRepair.body.error);
    assert.strictEqual(ticketCreates, 5);
    assert.strictEqual(actionCreates, 8, "Mapping repair must not create a second Email action.");

    const ambiguousRequesterOperationId = "create-operation-ambiguous-requester";
    await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents",
      "/api/halo/ticket-creation/intents",
      {
        ...creation,
        operationId: ambiguousRequesterOperationId,
        requesterMode: "auto",
      }
    );
    duplicateRequesters = true;
    const ambiguousRequesterSend = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${ambiguousRequesterOperationId}/send`,
      email,
      { operationId: ambiguousRequesterOperationId }
    );
    duplicateRequesters = false;
    assert.strictEqual(ambiguousRequesterSend.statusCode, 409);
    assert.match(ambiguousRequesterSend.body.error, /uniquely match/i);
    assert.strictEqual(ticketCreates, 5);

    const changedSchemaOperationId = "create-operation-schema-change";
    const changedSchemaIntent = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents",
      "/api/halo/ticket-creation/intents",
      { ...creation, operationId: changedSchemaOperationId }
    );
    assert.strictEqual(changedSchemaIntent.statusCode, 201);
    addNewMandatoryField = true;
    const changedSchemaSend = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${changedSchemaOperationId}/send`,
      email,
      { operationId: changedSchemaOperationId }
    );
    assert.strictEqual(changedSchemaSend.statusCode, 409);
    assert.match(changedSchemaSend.body.error, /fields changed/i);
    assert.strictEqual(ticketCreates, 5, "Schema changes must be reviewed before ticket creation.");

    addNewMandatoryField = false;
    const restoredSchema = await invoke(
      app,
      "GET",
      "/api/halo/ticket-creation/types/:typeId/schema",
      "/api/halo/ticket-creation/types/12/schema?refresh=1",
      undefined,
      { typeId: "12" }
    );
    assert.strictEqual(restoredSchema.statusCode, 200);
    const attachmentReceiptOperationId = "create-operation-attachment-receipt";
    const attachmentReceiptIntent = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents",
      "/api/halo/ticket-creation/intents",
      {
        ...creation,
        operationId: attachmentReceiptOperationId,
        schemaRevision: restoredSchema.body.schema.revision,
      }
    );
    assert.strictEqual(attachmentReceiptIntent.statusCode, 201);
    const attachmentBytes = Buffer.from("hello");
    const attachmentHash = crypto.createHash("sha256").update(attachmentBytes).digest("hex");
    const attachmentKey = "d".repeat(64);
    const attachmentFingerprint = "e".repeat(64);
    const attachmentDraftItemId = "attachment-receipt-draft";
    const receiptPrefetch = await invoke(
      app,
      "POST",
      "/api/halo/email-attachments/prefetch/start",
      "/api/halo/email-attachments/prefetch/start",
      {
        creationOperationId: attachmentReceiptOperationId,
        draftItemId: attachmentDraftItemId,
        emailAttachmentFingerprint: attachmentFingerprint,
        emailAttachments: [
          {
            attachmentKey,
            attachmentType: "file",
            contentSha256: attachmentHash,
            contentType: "text/plain",
            name: "receipt.txt",
            reportedSize: attachmentBytes.length,
          },
        ],
        operationId: attachmentReceiptOperationId,
      }
    );
    assert.strictEqual(receiptPrefetch.statusCode, 200, receiptPrefetch.body.error);
    const receiptStage = await invoke(
      app,
      "POST",
      "/api/halo/email-attachments/prefetch/:prefetchKey/items",
      `/api/halo/email-attachments/prefetch/${receiptPrefetch.body.prefetchKey}/items`,
      {
        attachmentKey,
        contentBase64: attachmentBytes.toString("base64"),
        contentFormat: "base64",
        contentSha256: attachmentHash,
      },
      { prefetchKey: receiptPrefetch.body.prefetchKey }
    );
    assert.strictEqual(receiptStage.statusCode, 200, receiptStage.body.error);
    const receiptSendBody = {
      ...email,
      itemId: attachmentDraftItemId,
      emailAttachmentDraftItemId: attachmentDraftItemId,
      emailAttachmentFingerprint: attachmentFingerprint,
      emailAttachmentOperationId: attachmentReceiptOperationId,
      emailAttachmentPrefetchKey: receiptPrefetch.body.prefetchKey,
      emailAttachmentStagingVersion: 2,
      emailAttachmentSummary: {
        attached: 0,
        detected: 1,
        failed: 0,
        prepared: 1,
        selected: 1,
        skipped: 0,
        warnings: [],
      },
      includeEmailAttachments: true,
    };
    const actionsBeforeReceiptFailure = actionCreates;
    const ticketsBeforeReceiptFailure = ticketCreates;
    failActionReceiptUpdateForOperation = attachmentReceiptOperationId;
    const receiptFailure = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${attachmentReceiptOperationId}/send`,
      receiptSendBody,
      { operationId: attachmentReceiptOperationId }
    );
    assert.strictEqual(receiptFailure.statusCode, 502);
    assert.match(receiptFailure.body.error, /receipt persistence failure/);
    assert.strictEqual(actionCreates, actionsBeforeReceiptFailure + 1);
    assert.strictEqual(ticketCreates, ticketsBeforeReceiptFailure + 1);

    const receiptRetry = await invoke(
      app,
      "POST",
      "/api/halo/ticket-creation/intents/:operationId/send",
      `/api/halo/ticket-creation/intents/${attachmentReceiptOperationId}/send`,
      receiptSendBody,
      { operationId: attachmentReceiptOperationId }
    );
    assert.strictEqual(receiptRetry.statusCode, 200, receiptRetry.body.error);
    assert.strictEqual(actionCreates, actionsBeforeReceiptFailure + 1);
    assert.strictEqual(
      ticketCreates,
      ticketsBeforeReceiptFailure + 1,
      "Create-ticket retry must reuse both the ticket and the action receipt."
    );
  } finally {
    global.fetch = originalFetch;
    await store.close();
  }

  console.log("Ticket creation route tests passed");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
