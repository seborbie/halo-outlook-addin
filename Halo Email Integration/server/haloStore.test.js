const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { createDatabasePool } = require("./database");
const { runMigrations } = require("./migrations");
const { createTestDatabase } = require("./testDatabase");

async function run() {
  assert.throws(() => createDatabasePool({ env: {} }), /DATABASE_URL must be set/);
  assert.throws(
    () => createDatabasePool({ connectionString: "https://database.example.com/haloaddin" }),
    /valid PostgreSQL connection URL/
  );
  const tlsPool = createDatabasePool({
    connectionString: "postgresql://database.example.com/haloaddin?ssl=false&sslmode=disable",
    password: "password",
    username: "application",
  });
  assert.deepEqual(tlsPool.options.ssl, { rejectUnauthorized: true });
  await tlsPool.end();

  const explicitlyPlaintextPool = createDatabasePool({
    connectionString: "postgresql://database.example.com/haloaddin?ssl=true&sslmode=require",
    env: { DATABASE_SSL: "false" },
    password: "password",
    username: "application",
  });
  assert.equal(explicitlyPlaintextPool.options.ssl, false);
  await explicitlyPlaintextPool.end();

  for (const authMode of [undefined, null, "", "psk", "PASSWORD", "usernamepassword"]) {
    const passwordPool = createDatabasePool({
      connectionString: "postgresql://database.example.com/haloaddin",
      env: { DATABASE_AUTH: authMode },
      password: "password",
      username: "application",
    });
    assert.equal(passwordPool.options.password, "password");
    assert.equal(passwordPool.options.user, "application");
    assert.equal(passwordPool.options.host, "database.example.com");
    assert.equal(passwordPool.options.database, "haloaddin");
    await passwordPool.end();
  }

  const requestedScopes = [];
  const entraPool = createDatabasePool({
    connectionString: "postgresql://database.example.com/haloaddin",
    env: { DATABASE_AUTH: " Entra " },
    tokenCredential: {
      async getToken(scope) {
        requestedScopes.push(scope);
        return { token: "managed-identity-token" };
      },
    },
    username: "haloaddin_webapp",
  });
  assert.equal(typeof entraPool.options.password, "function");
  assert.equal(await entraPool.options.password(), "managed-identity-token");
  assert.deepEqual(requestedScopes, ["https://ossrdbms-aad.database.windows.net/.default"]);
  assert.equal(entraPool.options.user, "haloaddin_webapp");
  await entraPool.end();

  assert.throws(
    () =>
      createDatabasePool({
        connectionString: "postgresql://database.example.com/haloaddin",
        env: { DATABASE_AUTH: "unknown" },
        password: "password",
        username: "application",
      }),
    /DATABASE_AUTH must be entra/
  );
  assert.throws(
    () =>
      createDatabasePool({
        connectionString: "postgresql://application:password@database.example.com/haloaddin",
        password: "password",
        username: "application",
      }),
    /DATABASE_URL must not include credentials/
  );

  assert.throws(
    () =>
      createDatabasePool({
        connectionString: "postgresql://database.example.com/haloaddin",
        password: "password",
      }),
    /DATABASE_USERNAME must be set/
  );
  assert.throws(
    () =>
      createDatabasePool({
        connectionString: "postgresql://database.example.com/haloaddin",
        username: "application",
      }),
    /DATABASE_PASSWORD must be set/
  );

  const database = await createTestDatabase();
  let store = await database.createStore();
  let inspectionPool = database.createPool();

  try {
    await store.initialize();
    const migrations = await inspectionPool.query(
      `SELECT name, checksum, COUNT(*)::int AS count
       FROM schema_migrations
       GROUP BY name, checksum`
    );
    assert.equal(migrations.rows.length, 5);
    assert.deepEqual(migrations.rows.map((row) => row.name).sort(), [
      "001_initial.sql",
      "002_encrypted_attachment_staging.sql",
      "003_expected_attachment_content_hash.sql",
      "004_conversation_action_mode.sql",
      "005_inline_image_visibility.sql",
    ]);
    migrations.rows.forEach((row) => {
      assert.equal(row.count, 1);
      assert.match(row.checksum, /^[a-f0-9]{64}$/);
    });

    const migrationTestDirectory = fs.mkdtempSync(path.join(os.tmpdir(), "halo-migrations-"));
    try {
      const migrationName = "001_initial.sql";
      const migrationSource = path.join(__dirname, "migrations", migrationName);
      const migrationCopy = path.join(migrationTestDirectory, migrationName);
      fs.copyFileSync(migrationSource, migrationCopy);
      fs.copyFileSync(
        path.join(__dirname, "migrations", "002_encrypted_attachment_staging.sql"),
        path.join(migrationTestDirectory, "002_encrypted_attachment_staging.sql")
      );
      fs.copyFileSync(
        path.join(__dirname, "migrations", "003_expected_attachment_content_hash.sql"),
        path.join(migrationTestDirectory, "003_expected_attachment_content_hash.sql")
      );
      fs.copyFileSync(
        path.join(__dirname, "migrations", "004_conversation_action_mode.sql"),
        path.join(migrationTestDirectory, "004_conversation_action_mode.sql")
      );
      fs.copyFileSync(
        path.join(__dirname, "migrations", "005_inline_image_visibility.sql"),
        path.join(migrationTestDirectory, "005_inline_image_visibility.sql")
      );
      await runMigrations(inspectionPool, { directory: migrationTestDirectory });
      fs.appendFileSync(migrationCopy, "\n-- checksum drift regression test\n");
      await assert.rejects(
        runMigrations(inspectionPool, { directory: migrationTestDirectory }),
        /has changed since it was applied/
      );
    } finally {
      fs.rmSync(migrationTestDirectory, { force: true, recursive: true });
    }

    const user = await store.upsertUser({
      displayName: "First Name",
      email: "first@example.com",
      objectId: "object-1",
      tenantId: "tenant-1",
    });
    const updatedUser = await store.upsertUser({
      displayName: "Updated Name",
      email: "updated@example.com",
      objectId: "object-1",
      tenantId: "tenant-1",
    });
    assert.equal(updatedUser.id, user.id);
    assert.equal(updatedUser.displayName, "Updated Name");

    const encryptedToken = { ciphertext: "token-ciphertext", iv: "token-iv", tag: "token-tag" };
    const grant = await store.saveHaloGrant({
      clientId: "client-1",
      encryptedToken,
      haloUrl: "https://customer.halopsa.com",
      scope: "all",
      userId: user.id,
    });
    assert.deepEqual(grant.encryptedToken, encryptedToken);

    const expiresAt = Date.now() + 60_000;
    await store.createSession({ expiresAt, sessionHash: "session-current", userId: user.id });
    await store.createSession({
      expiresAt: Date.now() - 1,
      sessionHash: "session-expired",
      userId: user.id,
    });
    await store.createBackgroundSession({
      backgroundSessionHash: "background-current",
      expiresAt,
      sessionHash: "session-current",
    });

    const mapping = {
      conversationId: "conversation-1",
      createdAt: Date.now(),
      id: "mapping-1",
      mailboxEmail: "support@example.com",
      normalizedSubject: "subject",
      ticketId: 101,
      ticketNumber: "T101",
      updatedAt: Date.now(),
    };
    await store.saveConversationMapping(mapping);
    await store.saveMessageMapping({
      mailboxEmail: mapping.mailboxEmail,
      mappingId: mapping.id,
      messageIdKey: "<message-1@example.com>",
    });
    const loadedMapping = await store.getMappingByMessageId(
      mapping.mailboxEmail,
      "<message-1@example.com>"
    );
    assert.equal(loadedMapping.ticketNumber, "T101");
    assert.equal(loadedMapping.actionMode, "email");
    assert(loadedMapping.syncedMessageIds.has("<message-1@example.com>"));
    mapping.actionMode = "private-note";
    await store.saveConversationMapping(mapping);
    assert.equal(
      (await store.getMappingByConversationId(mapping.mailboxEmail, mapping.conversationId))
        .actionMode,
      "private-note"
    );

    const encryptedIntent = {
      ciphertext: "intent-ciphertext",
      nested: { version: 1 },
    };
    await store.upsertTicketCreationIntent({
      encryptedIntent,
      expiresAt,
      haloTenant: grant.haloUrl,
      operationId: "intent-1",
      userId: user.id,
    });
    assert.deepEqual(
      (
        await store.getTicketCreationIntent("intent-1", {
          haloTenant: grant.haloUrl,
          userId: user.id,
        })
      ).encryptedIntent,
      encryptedIntent
    );

    await store.saveTicketCreationMetadata({
      cacheKey: "metadata-1",
      expiresAt,
      haloTenant: grant.haloUrl,
      payload: { fields: [{ id: 1, values: ["a", "b"] }] },
      userId: user.id,
    });
    assert.deepEqual(
      (
        await store.getTicketCreationMetadata("metadata-1", {
          haloTenant: grant.haloUrl,
          userId: user.id,
        })
      ).payload,
      { fields: [{ id: 1, values: ["a", "b"] }] }
    );

    await store.createBugReportSession({
      diagnostics: { addInVersion: "test" },
      expiresAt,
      sessionHash: "bug-session",
      userId: user.id,
    });
    const bugClaims = await Promise.all([
      store.claimBugReportSession("bug-session"),
      store.claimBugReportSession("bug-session"),
    ]);
    assert.equal(bugClaims.filter(Boolean).length, 1);

    await store.upsertEmailAttachmentPrefetch(
      {
        attachmentFingerprint: "a".repeat(64),
        expectedBytes: 20,
        expectedCount: 1,
        expiresAt,
        haloTenant: grant.haloUrl,
        operationId: "attachment-operation",
        prefetchKey: "attachment-prefetch",
        ticketId: 101,
        userId: user.id,
      },
      [
        {
          attachmentKey: "b".repeat(64),
          attachmentType: "file",
          contentSha256: "c".repeat(64),
          contentType: "text/plain",
          name: "notes.txt",
          reportedSize: 20,
        },
      ]
    );
    const attachmentClaims = await Promise.all([
      store.claimEmailAttachmentPrefetchItem(
        "attachment-prefetch",
        "b".repeat(64),
        20,
        "c".repeat(64),
        1024
      ),
      store.claimEmailAttachmentPrefetchItem(
        "attachment-prefetch",
        "b".repeat(64),
        20,
        "c".repeat(64),
        1024
      ),
    ]);
    assert.equal(
      attachmentClaims.filter((claim) => claim && claim.status === "preparing").length,
      1
    );
    const commitClaim = await store.claimEmailAttachmentPrefetchCommit("attachment-prefetch", {
      haloTenant: grant.haloUrl,
      ticketId: 101,
      userId: user.id,
    });
    assert.equal(commitClaim.status, "committing");
    const expiredCommitCandidates = await store.getEmailAttachmentCleanupCandidates(expiresAt + 1);
    assert.equal(
      expiredCommitCandidates.some((candidate) => candidate.prefetchKey === "attachment-prefetch"),
      true,
      "Expired commits must be purged instead of retaining ciphertext indefinitely."
    );
    assert.equal(await store.releaseEmailAttachmentPrefetchCommit("attachment-prefetch"), 1);

    await store.cleanExpired(Date.now());
    assert.equal(await store.getSessionWithGrant("session-expired"), null);
    assert((await store.getSessionWithGrant("session-current"))?.grantId);

    await store.close();
    store = await database.createStore();
    assert((await store.getSessionWithGrant("session-current"))?.grantId);
    assert.equal(
      (await store.getMappingByConversationId(mapping.mailboxEmail, mapping.conversationId))
        .ticketId,
      101
    );

    await inspectionPool.query("DELETE FROM conversation_mappings WHERE id = $1", [mapping.id]);
    const mappingChildren = await inspectionPool.query(
      "SELECT COUNT(*)::int AS count FROM message_mappings WHERE mapping_id = $1",
      [mapping.id]
    );
    assert.equal(mappingChildren.rows[0].count, 0);

    await inspectionPool.query("DELETE FROM users WHERE id = $1", [user.id]);
    assert.equal(await store.getGrantByUserId(user.id), null);
    assert.equal(await store.getSessionWithGrant("session-current"), null);
    const cascaded = await inspectionPool.query(
      `SELECT
         (SELECT COUNT(*)::int FROM halo_grants WHERE user_id = $1) AS grants,
         (SELECT COUNT(*)::int FROM sessions WHERE user_id = $1) AS sessions,
         (SELECT COUNT(*)::int FROM background_sessions) AS background_sessions,
         (SELECT COUNT(*)::int FROM email_attachment_prefetch WHERE user_id = $1) AS prefetches,
         (SELECT COUNT(*)::int FROM email_attachment_prefetch_items) AS prefetch_items,
         (SELECT COUNT(*)::int FROM ticket_creation_intents WHERE user_id = $1) AS intents,
         (SELECT COUNT(*)::int FROM ticket_creation_metadata_cache WHERE user_id = $1) AS metadata`,
      [user.id]
    );
    assert.deepEqual(cascaded.rows[0], {
      background_sessions: 0,
      grants: 0,
      intents: 0,
      metadata: 0,
      prefetch_items: 0,
      prefetches: 0,
      sessions: 0,
    });
  } finally {
    await inspectionPool.end();
    await store.close();
    await database.close();
  }

  console.log("PostgreSQL store tests passed");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
