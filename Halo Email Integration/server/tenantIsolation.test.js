const assert = require("assert");
const crypto = require("crypto");
const { Pool } = require("pg");
const { createHaloStore, getDatabaseConfig } = require("./haloStore");

const databaseUrl = process.env.TEST_DATABASE_URL || process.env.DATABASE_URL;

async function run() {
  if (!databaseUrl) {
    throw new Error("TEST_DATABASE_URL or DATABASE_URL is required for tenant isolation tests.");
  }

  const env = { ...process.env, DATABASE_URL: databaseUrl };
  const store = createHaloStore({ env });
  await store.ready;
  await store.resetForTests();

  const tenantA = await store.registerOrganisation({
    companyName: "Alpha Service Desk",
    haloClientId: "alpha-client",
    haloUrl: "https://alpha.halopsa.com",
    microsoftTenantId: "microsoft-alpha",
    owner: {
      displayName: "Alpha Owner",
      email: "owner@alpha.example",
      objectId: "shared-object-id",
    },
  });
  const tenantB = await store.registerOrganisation({
    companyName: "Beta Service Desk",
    haloClientId: "beta-client",
    haloUrl: "https://beta.halopsa.com",
    microsoftTenantId: "microsoft-beta",
    owner: {
      displayName: "Beta Owner",
      email: "owner@beta.example",
      objectId: "shared-object-id",
    },
  });

  assert.notStrictEqual(tenantA.organisation.id, tenantB.organisation.id);
  assert.notStrictEqual(tenantA.user.id, tenantB.user.id);
  assert.strictEqual(tenantA.user.haloUrl, "https://alpha.halopsa.com");
  assert.strictEqual(tenantB.user.haloUrl, "https://beta.halopsa.com");

  const betaMember = await store.upsertUser({
    displayName: "Beta Member",
    email: "member@beta.example",
    objectId: "beta-member",
    tenantId: "microsoft-beta",
  });
  await assert.rejects(
    () =>
      store.registerOrganisation({
        companyName: "Hijacked Beta",
        haloClientId: "attacker-client",
        haloUrl: "https://attacker.halopsa.com",
        microsoftTenantId: "microsoft-beta",
        owner: {
          displayName: betaMember.displayName,
          email: betaMember.email,
          objectId: betaMember.objectId,
        },
      }),
    /owner or admin/i
  );

  const encryptedToken = { ciphertext: "cipher", iv: "iv", tag: "tag", version: 1 };
  await store.saveHaloGrant({
    clientId: "alpha-client",
    encryptedToken,
    haloUrl: "https://alpha.halopsa.com",
    organisationId: tenantA.organisation.id,
    scope: "all",
    userId: tenantA.user.id,
  });
  await store.saveHaloGrant({
    clientId: "beta-client",
    encryptedToken,
    haloUrl: "https://beta.halopsa.com",
    organisationId: tenantB.organisation.id,
    scope: "all",
    userId: tenantB.user.id,
  });

  const sharedSessionHash = crypto.createHash("sha256").update("same-session").digest("hex");
  await store.createSession({
    expiresAt: Date.now() + 60_000,
    organisationId: tenantA.organisation.id,
    sessionHash: sharedSessionHash,
    userId: tenantA.user.id,
  });
  await store.createSession({
    expiresAt: Date.now() + 60_000,
    organisationId: tenantB.organisation.id,
    sessionHash: sharedSessionHash,
    userId: tenantB.user.id,
  });

  const sessionA = await store.getSessionWithGrant(tenantA.organisation.id, sharedSessionHash);
  const sessionB = await store.getSessionWithGrant(tenantB.organisation.id, sharedSessionHash);
  assert.strictEqual(sessionA.haloUrl, "https://alpha.halopsa.com");
  assert.strictEqual(sessionB.haloUrl, "https://beta.halopsa.com");

  const mappingAId = crypto.randomUUID();
  const mappingBId = crypto.randomUUID();
  const baseMapping = {
    mailboxEmail: "support@shared.example",
    conversationId: "shared-conversation",
    normalizedSubject: "Shared subject",
    syncedMessageIds: new Set(),
    createdAt: Date.now(),
    updatedAt: Date.now(),
  };
  await store.saveConversationMapping(tenantA.organisation.id, {
    ...baseMapping,
    id: mappingAId,
    ticketId: 101,
    ticketNumber: "A-101",
  });
  await store.saveConversationMapping(tenantB.organisation.id, {
    ...baseMapping,
    id: mappingBId,
    ticketId: 202,
    ticketNumber: "B-202",
  });

  const mappingA = await store.getMappingByConversationId(
    tenantA.organisation.id,
    baseMapping.mailboxEmail,
    baseMapping.conversationId
  );
  const mappingB = await store.getMappingByConversationId(
    tenantB.organisation.id,
    baseMapping.mailboxEmail,
    baseMapping.conversationId
  );
  assert.strictEqual(mappingA.ticketNumber, "A-101");
  assert.strictEqual(mappingB.ticketNumber, "B-202");

  await assert.rejects(
    () =>
      store.saveMessageMapping(tenantB.organisation.id, {
        mailboxEmail: baseMapping.mailboxEmail,
        mappingId: mappingAId,
        messageIdKey: "<cross-tenant@example.com>",
      }),
    /foreign key|violates row-level security/i
  );

  const pool = new Pool(getDatabaseConfig(env));
  const role = await pool.query(
    "SELECT rolsuper, rolbypassrls FROM pg_roles WHERE rolname = current_user"
  );
  assert.strictEqual(role.rows[0].rolsuper, false, "The application DB role must not be superuser");
  assert.strictEqual(
    role.rows[0].rolbypassrls,
    false,
    "The application DB role must not bypass row-level security"
  );

  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    await client.query("SELECT set_config('app.current_organisation_id', $1, true)", [
      tenantA.organisation.id,
    ]);
    const visible = await client.query(
      "SELECT organisation_id, ticket_number FROM conversation_mappings ORDER BY ticket_number"
    );
    assert.deepStrictEqual(visible.rows, [
      { organisation_id: tenantA.organisation.id, ticket_number: "A-101" },
    ]);
    const forcedCrossTenantRead = await client.query(
      "SELECT ticket_number FROM conversation_mappings WHERE organisation_id = $1",
      [tenantB.organisation.id]
    );
    assert.deepStrictEqual(forcedCrossTenantRead.rows, []);
    await client.query("ROLLBACK");
  } finally {
    client.release();
    await pool.end();
  }

  await store.close();
  console.log("PostgreSQL tenant isolation tests passed.");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
