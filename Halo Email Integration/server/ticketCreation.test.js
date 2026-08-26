const assert = require("node:assert");

const { createTestStore } = require("./testDatabase");
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

async function run() {
  const types = normalizeTicketTypes({
    requesttypes: [
      { id: 12, name: "Project Engineer", group_name: "Projects", cancreate: true },
      { id: 13, name: "Inactive", inactive: true, cancreate: true },
      { id: 14, name: "Not permitted", cancreate: false },
    ],
  });
  assert.deepStrictEqual(types, [
    {
      id: "12",
      name: "Project Engineer",
      group: "Projects",
      active: true,
      canCreate: true,
    },
  ]);
  assert.deepStrictEqual(
    normalizeTicketTypes({
      TicketTypes: [
        {
          tickettype_id: 22,
          tickettypename: "Service Request",
          groupname: "Support",
          isactive: 1,
          canagentscreate: 1,
        },
      ],
    })[0],
    {
      id: "22",
      name: "Service Request",
      group: "Support",
      active: true,
      canCreate: true,
    }
  );

  const schema = normalizeTicketTypeSchema(
    "12",
    "Project Engineer",
    [
      {
        id: 12,
        name: "Project Engineer",
        default_priority_id: 3,
        fields: [
          {
            id: 100,
            iscustom: true,
            fieldname: "CFProjectCode",
            label: "Project code",
            fieldtype: "Text",
            mandatory: true,
            sequence: 1,
          },
          {
            id: 101,
            iscustom: true,
            fieldname: "CFDeliveryModel",
            label: "Delivery model",
            fieldtype: "Single Selection",
            options: [
              { id: 1, name: "Remote" },
              { id: 2, name: "On site" },
            ],
            sequence: 2,
          },
          {
            fieldname: "priority_id",
            label: "Priority",
            fieldtype: "Number",
            mandatory: true,
            sequence: 3,
          },
          {
            fieldname: "client_id",
            label: "Client",
            fieldtype: "Lookup",
            sequence: 4,
          },
        ],
      },
    ],
    []
  );
  assert.strictEqual(schema.available, true);
  assert.strictEqual(schema.fields.length, 4);
  assert.strictEqual(schema.fields[0].key, "custom:100");
  assert.strictEqual(schema.fields[1].type, "select");
  assert.strictEqual(schema.fields[2].defaultValue, 3);
  assert.strictEqual(schema.fields[3].type, "client");
  assert.strictEqual(schema.fields[3].supported, true);
  assert.match(schema.revision, /^[a-f0-9]{64}$/);

  const validated = validateCreationInput(schema, {
    schemaRevision: schema.revision,
    summary: "Create project from Outlook",
    values: {
      "custom:100": "PROJ-42",
      "custom:101": "2",
    },
  });
  assert.strictEqual(validated.values["core:priority_id"], 3);
  const payload = buildTicketPayload({
    schema,
    summary: validated.summary,
    values: validated.values,
    requester: { id: "50", clientId: "60", siteId: "70" },
  });
  assert.deepStrictEqual(payload.customfields, [
    { id: 100, value: "PROJ-42" },
    { id: 101, value: "2" },
  ]);
  assert.strictEqual(payload.tickettype_id, 12);
  assert.strictEqual(payload.priority_id, 3);
  assert.strictEqual(payload.user_id, 50);
  assert.strictEqual(payload.client_id, 60);
  assert.strictEqual(payload.site_id, 70);

  assert.throws(
    () =>
      validateCreationInput(schema, {
        schemaRevision: schema.revision,
        summary: "Missing required value",
        values: {},
      }),
    /Project code is required/
  );
  assert.throws(
    () =>
      validateCreationInput(schema, {
        schemaRevision: "stale",
        summary: "Stale schema",
        values: { "custom:100": "PROJ-1" },
      }),
    /fields changed/
  );

  const unavailable = normalizeTicketTypeSchema(
    "99",
    "Unsupported",
    [
      {
        fields: [{ id: 999, iscustom: true, label: "Table", fieldtype: "Table", mandatory: true }],
      },
    ],
    []
  );
  assert.strictEqual(unavailable.available, false);
  assert.match(unavailable.unavailableReason, /Table/);

  const optionalUnsupported = normalizeTicketTypeSchema(
    "23",
    "Alternate response shape",
    { id: 23, name: "Alternate response shape" },
    {
      records: [
        {
          customfield_id: 500,
          is_custom: true,
          display_name: "Customer approved",
          default_value: "false",
          field_type: 6,
        },
        {
          customfield_id: 501,
          is_custom: true,
          display_name: "Unsupported optional table",
          field_type: "Table",
        },
        {
          customfield_id: 502,
          is_custom: true,
          display_name: "Preferred time",
          field_type: "Time",
        },
      ],
    }
  );
  assert.strictEqual(optionalUnsupported.available, true);
  assert.strictEqual(optionalUnsupported.fields[0].type, "boolean");
  assert.strictEqual(optionalUnsupported.fields[0].defaultValue, false);
  assert.strictEqual(optionalUnsupported.fields[2].type, "time");
  assert.match(optionalUnsupported.warnings[0], /Unsupported optional table/);

  const nestedFieldInfo = normalizeTicketTypeSchema(
    "24",
    "Halo fieldinfo response",
    {
      id: 24,
      fields: [
        {
          id: 9001,
          fieldid: 4,
          fieldname: "urgency",
          seq: 2,
          fieldinfo: {
            id: 4,
            name: "urgency",
            label: "Urgency",
            type: 2,
            custom: 0,
            values: [
              { id: 1, name: "Low" },
              { id: 2, name: "High" },
            ],
          },
        },
        {
          id: 9002,
          fieldid: 750,
          seq: 1,
          fieldinfo: {
            id: 750,
            name: "CFTescoBuild",
            label: "Tesco Build",
            type: 6,
            custom: 1,
          },
        },
      ],
    },
    {
      records: [
        // Halo can repeat the placement through TicketTypeField. The richer
        // fieldinfo row above must win and the field must not be duplicated.
        { id: 9901, fieldid: 4, seq: 2 },
      ],
    }
  );
  assert.strictEqual(nestedFieldInfo.fields.length, 2);
  assert.deepStrictEqual(
    nestedFieldInfo.fields.map((field) => ({
      key: field.key,
      label: field.label,
      type: field.type,
    })),
    [
      { key: "custom:750", label: "Tesco Build", type: "boolean" },
      { key: "core:urgency", label: "Urgency", type: "select" },
    ]
  );
  assert.deepStrictEqual(nestedFieldInfo.fields[1].options, [
    { value: "1", label: "Low" },
    { value: "2", label: "High" },
  ]);

  const haloIncident = hydrateTicketCreationFieldOptions(
    normalizeTicketTypeSchema(
      "1",
      "Incident",
      {
        default_agent: 1,
        id: 1,
        fields: [
          {
            fieldid: 2,
            seq: 1,
            technew: 3,
            fieldinfo: { id: 2, name: "symptom", label: "Summary", type: 0, custom: 0 },
          },
          {
            fieldid: 3,
            seq: 2,
            technew: 2,
            fieldinfo: { id: 3, name: "symptom2", label: "Details", type: 1, custom: 0 },
          },
          {
            fieldid: 28,
            seq: 3,
            technew: 1,
            fieldinfo: { id: 28, name: "urgency", label: "Urgency", type: 1, custom: 0 },
          },
          {
            fieldid: 27,
            seq: 4,
            technew: 1,
            fieldinfo: { id: 27, name: "impact", label: "Impact", type: 1, custom: 0 },
          },
          {
            fieldid: 4,
            seq: 5,
            technew: 1,
            fieldinfo: { id: 4, name: "N/A", label: "Asset", type: -1, custom: 0 },
          },
          {
            fieldid: 13,
            seq: 6,
            technew: 1,
            fieldinfo: { id: 13, name: "sectio_", label: "Team", type: -1, custom: 0 },
          },
          {
            fieldid: 14,
            seq: 7,
            technew: 1,
            fieldinfo: {
              id: 14,
              name: "assignedtoint",
              label: "Agent",
              type: -1,
              custom: 0,
            },
          },
          {
            fieldid: 16,
            seq: 8,
            technew: 1,
            fieldinfo: {
              id: 16,
              name: "estimate",
              label: "Estimated Time",
              type: -1,
              custom: 0,
            },
          },
          {
            fieldid: 310,
            seq: 9,
            technew: 1,
            fieldinfo: {
              id: 310,
              name: "CFWeeklookup",
              label: "Weeklookup",
              type: 2,
              custom: 1,
            },
          },
          {
            fieldid: 311,
            seq: 10,
            technew: 3,
            fieldinfo: { id: 311, name: "CFStartDate", label: "Start date", type: 4, custom: 1 },
          },
          {
            fieldid: 312,
            seq: 11,
            technew: 1,
            fieldinfo: { id: 312, name: "CFStartTime", label: "Start time", type: 5, custom: 1 },
          },
          {
            fieldid: 313,
            seq: 12,
            technew: 1,
            fieldinfo: { id: 313, name: "CFScorecard", label: "Score card", type: 7, custom: 1 },
          },
          {
            fieldid: 314,
            seq: 13,
            technew: 1,
            fieldinfo: {
              id: 314,
              name: "CFQuantity",
              label: "Quantity",
              type: 0,
              inputtype: 1,
              custom: 1,
            },
          },
          {
            fieldid: 315,
            seq: 14,
            technew: 0,
            fieldinfo: {
              id: 315,
              name: "CFHiddenControl",
              label: "Hidden control",
              type: 10,
              custom: 1,
            },
          },
        ],
      },
      []
    ),
    {
      "core:impact": [
        { id: 0, name: "Unknown" },
        { id: 1, name: "1. High" },
        { id: 2, name: "2. Medium" },
        { id: 3, name: "3. Low" },
      ],
      "core:urgency": [
        { id: 0, name: "Unknown" },
        { id: 1, name: "1. High" },
        { id: 2, name: "2. Medium" },
        { id: 3, name: "3. Low" },
      ],
      "custom:310": [
        { id: "02/01/2000 12:00:00 AM", name: "-1" },
        { id: "03/01/2000 12:00:00 AM", name: "0" },
        { id: "04/01/2000 12:00:00 AM", name: "1" },
      ],
    }
  );
  const incidentFields = Object.fromEntries(haloIncident.fields.map((field) => [field.key, field]));
  assert.strictEqual(incidentFields["core:summary"].managed, true);
  assert.strictEqual(incidentFields["core:summary"].required, true);
  assert.strictEqual(incidentFields["core:details"].type, "multiline");
  assert.strictEqual(incidentFields["core:details"].required, false);
  assert.strictEqual(incidentFields["core:details"].recommended, true);
  assert.strictEqual(incidentFields["core:asset_id"].type, "asset");
  assert.strictEqual(incidentFields["core:asset_id"].required, false);
  assert.strictEqual(incidentFields["core:team_id"].type, "team");
  assert.strictEqual(incidentFields["core:team_id"].required, false);
  assert.strictEqual(incidentFields["core:agent_id"].type, "agent");
  assert.strictEqual(incidentFields["core:agent_id"].required, true);
  assert.strictEqual(incidentFields["core:agent_id"].defaultValue, 1);
  assert.strictEqual(incidentFields["core:estimate"].type, "duration");
  assert.strictEqual(incidentFields["core:estimate"].required, false);
  assert.strictEqual(incidentFields["custom:311"].type, "date");
  assert.strictEqual(incidentFields["custom:311"].required, true);
  assert.strictEqual(incidentFields["custom:312"].type, "time");
  assert.strictEqual(incidentFields["custom:313"].supported, false);
  assert.strictEqual(incidentFields["custom:314"].type, "number");
  assert.strictEqual(incidentFields["custom:315"], undefined);
  assert(!haloIncident.warnings.some((warning) => /Hidden control/.test(warning)));
  assert.deepStrictEqual(incidentFields["core:urgency"].options[0], {
    value: "0",
    label: "Unknown",
  });
  assert.deepStrictEqual(
    incidentFields["custom:310"].options.map((option) => option.label),
    ["-1", "0", "1"]
  );
  assert(
    !haloIncident.warnings.some((warning) =>
      /Summary|Details|Urgency|Impact|Asset|Team|Agent/.test(warning)
    )
  );

  const incidentValues = validateCreationInput(haloIncident, {
    schemaRevision: haloIncident.revision,
    summary: "Email subject",
    values: {
      "core:details": "Customer-provided ticket detail.",
      "core:urgency": "1",
      "core:impact": "2",
      "core:estimate": "01:30",
      "custom:310": "04/01/2000 12:00:00 AM",
      "custom:311": "2026-08-21",
    },
  });
  const incidentPayload = buildTicketPayload({
    schema: haloIncident,
    summary: incidentValues.summary,
    values: incidentValues.values,
    requester: { id: "50" },
  });
  assert.match(incidentPayload.details, /^Customer-provided ticket detail\./);
  assert.match(incidentPayload.details, /Created from an Outlook email/);
  assert.strictEqual(incidentPayload.urgency, "1");
  assert.strictEqual(incidentPayload.impact, "2");
  assert.strictEqual(incidentPayload.estimate, 1.5);
  assert.strictEqual(incidentPayload.agent_id, 1);

  const perTypeField = (technew) => ({
    fieldid: 600,
    technew,
    fieldinfo: {
      id: 600,
      name: "CFProfileSpecific",
      label: "Profile-specific value",
      type: 0,
      custom: 1,
    },
  });
  const optionalProfile = normalizeTicketTypeSchema(
    "40",
    "Optional profile",
    { fields: [perTypeField(1)] },
    []
  );
  const requiredProfile = normalizeTicketTypeSchema(
    "41",
    "Required profile",
    { fields: [perTypeField(3)] },
    []
  );
  assert.strictEqual(optionalProfile.fields[0].required, false);
  assert.strictEqual(requiredProfile.fields[0].required, true);

  assert.deepStrictEqual(
    normalizeRequesters({ users: [{ id: 5, name: "Alex", email: "A@EXAMPLE.COM" }] }),
    [
      {
        id: "5",
        name: "Alex",
        emailAddress: "a@example.com",
        clientId: "",
        clientName: "",
        siteId: "",
        siteName: "",
      },
    ]
  );
  assert.deepStrictEqual(getCreatedTicket([{ id: 123, ticketnumber: "T123" }]), {
    id: "123",
    ticketNumber: "T123",
  });
  assert.deepStrictEqual(
    getCreatedTicket({ data: { tickets: [{ ticket_id: 124, ticket_number: "T124" }] } }),
    { id: "124", ticketNumber: "T124" }
  );
  assert.deepStrictEqual(
    normalizeLookupResults(
      { assets: [{ id: 44, name: "Laptop", inventory_number: "LT-44" }] },
      "asset"
    ),
    [{ id: "44", label: "Laptop", secondary: "LT-44" }]
  );

  await testPersistence();
  console.log("Ticket creation tests passed");
}

async function testPersistence() {
  const store = await createTestStore();
  const user = await store.upsertUser({ objectId: "object", tenantId: "tenant" });
  const crypto = createTokenCrypto({
    HALO_TOKEN_ENCRYPTION_KEY: Buffer.alloc(32, 4).toString("base64url"),
  });
  const encryptedIntent = crypto.encryptJson({ typeId: "12", summary: "Test" });
  await store.upsertTicketCreationIntent({
    encryptedIntent,
    expiresAt: Date.now() + 60_000,
    haloTenant: "https://example.halopsa.com",
    operationId: "operation-1",
    userId: user.id,
  });
  const stored = await store.getTicketCreationIntent("operation-1", {
    haloTenant: "https://example.halopsa.com",
    userId: user.id,
  });
  assert.deepStrictEqual(crypto.decryptJson(stored.encryptedIntent), {
    typeId: "12",
    summary: "Test",
  });
  const updated = await store.updateTicketCreationIntent(
    "operation-1",
    { status: "ticket-created", ticketId: 321, ticketNumber: "T321" },
    { haloTenant: "https://example.halopsa.com", userId: user.id }
  );
  assert.strictEqual(updated.status, "ticket-created");
  assert.strictEqual(updated.ticketId, 321);
  assert.strictEqual(
    await store.deleteTicketCreationIntent("operation-1", {
      haloTenant: "https://example.halopsa.com",
      userId: user.id,
    }),
    0,
    "An already-created ticket must not be deleted with a pending intent."
  );

  await store.saveTicketCreationMetadata({
    cacheKey: "types",
    expiresAt: Date.now() + 60_000,
    haloTenant: "https://example.halopsa.com",
    payload: { types: [{ id: "12" }] },
    userId: user.id,
  });
  assert.deepStrictEqual(
    (
      await store.getTicketCreationMetadata("types", {
      haloTenant: "https://example.halopsa.com",
      userId: user.id,
      })
    ).payload,
    { types: [{ id: "12" }] }
  );

  await store.upsertEmailAttachmentPrefetch(
    {
      attachmentFingerprint: "a".repeat(64),
      expectedBytes: 10,
      expectedCount: 1,
      expiresAt: Date.now() + 60_000,
      haloTenant: "https://example.halopsa.com",
      operationId: "operation-attachments",
      prefetchKey: "prefetch-1",
      ticketId: 0,
      userId: user.id,
    },
    [
      {
        attachmentKey: "b".repeat(64),
        attachmentType: "file",
        contentType: "text/plain",
        name: "notes.txt",
        reportedSize: 10,
      },
    ]
  );
  assert.strictEqual(
    await store.rebindEmailAttachmentPrefetch("prefetch-1", {
      haloTenant: "https://example.halopsa.com",
      ticketId: 321,
      userId: user.id,
    }),
    1
  );
  assert.strictEqual(
    (
      await store.getEmailAttachmentPrefetch("prefetch-1", {
      haloTenant: "https://example.halopsa.com",
      ticketId: 321,
      userId: user.id,
      })
    ).ticketId,
    321
  );
  await store.close();
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
