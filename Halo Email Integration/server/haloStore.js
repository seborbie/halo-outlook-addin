const crypto = require("crypto");
const { Pool } = require("pg");
const { runMigrations } = require("./migrations");

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

class OrganisationNotConfiguredError extends Error {
  constructor() {
    super(
      "This Microsoft organisation has not been configured. Ask an administrator to complete InboxLink signup."
    );
    this.name = "OrganisationNotConfiguredError";
    this.status = 403;
  }
}

class OrganisationRegistrationError extends Error {
  constructor(message) {
    super(message);
    this.name = "OrganisationRegistrationError";
    this.status = 403;
  }
}

function getDatabaseConfig(env = process.env) {
  const connectionString = String(env.DATABASE_URL || "").trim();
  if (!connectionString) {
    throw new Error("DATABASE_URL must be set to a PostgreSQL connection string.");
  }

  const config = { connectionString };
  if (String(env.DATABASE_SSL || "").toLowerCase() === "require") {
    config.ssl = { rejectUnauthorized: true };
  }
  return config;
}

function createHaloStore(options = {}) {
  const ownsPool = !options.pool;
  const pool = options.pool || new Pool(options.databaseConfig || getDatabaseConfig(options.env));
  const ready = (options.skipMigrations ? Promise.resolve() : runMigrations(pool)).then(() =>
    assertTenantSafeRole(pool)
  );

  return {
    ready,
    cleanExpired,
    claimBugReportSession,
    close,
    consumeBugReportSession,
    createBackgroundSession,
    createBugReportSession,
    createSession,
    deleteBackgroundSessionsForSessionHash,
    deleteSession,
    deleteSessionsForUser,
    getBackgroundSessionWithGrant,
    getGrantByUserId,
    getMappingByConversationId,
    getMappingByMessageId,
    getOrganisationByMicrosoftTenantId,
    getSessionWithGrant,
    invalidateGrantById,
    invalidateGrantForUser,
    registerOrganisation,
    releaseBugReportSession,
    resetForTests,
    saveConversationMapping,
    saveHaloGrant,
    saveMessageMapping,
    updateGrantToken,
    upsertUser,
  };

  async function withTransaction(settings, callback) {
    await ready;
    const client = await pool.connect();
    try {
      await client.query("BEGIN");
      if (settings.microsoftTenantId) {
        await client.query("SELECT set_config('app.current_microsoft_tenant_id', $1, true)", [
          settings.microsoftTenantId,
        ]);
      }
      if (settings.organisationId) {
        assertOrganisationId(settings.organisationId);
        await client.query("SELECT set_config('app.current_organisation_id', $1, true)", [
          settings.organisationId,
        ]);
      }
      const result = await callback(client);
      await client.query("COMMIT");
      return result;
    } catch (error) {
      await client.query("ROLLBACK").catch(() => {});
      throw error;
    } finally {
      client.release();
    }
  }

  function withTenant(organisationId, callback) {
    return withTransaction({ organisationId }, callback);
  }

  function withMicrosoftTenant(microsoftTenantId, callback) {
    const normalizedTenantId = String(microsoftTenantId || "").trim();
    if (!normalizedTenantId) {
      throw new Error("A Microsoft tenant ID is required.");
    }
    return withTransaction({ microsoftTenantId: normalizedTenantId }, callback);
  }

  async function registerOrganisation({
    companyName,
    haloClientId,
    haloUrl,
    microsoftTenantId,
    owner,
  }) {
    return withMicrosoftTenant(microsoftTenantId, async (client) => {
      const now = Date.now();
      const previous = await client.query(
        `SELECT id, halo_url, halo_client_id
         FROM organisations
         WHERE microsoft_tenant_id = $1`,
        [microsoftTenantId]
      );
      const organisationId = previous.rows[0]?.id || crypto.randomUUID();
      const slug = uniqueSlug(companyName, organisationId);
      let ownerRole = "owner";
      if (previous.rows[0]) {
        await client.query("SELECT set_config('app.current_organisation_id', $1, true)", [
          organisationId,
        ]);
        const existingUser = await client.query(
          `SELECT role FROM users
           WHERE organisation_id = $1 AND object_id = $2`,
          [organisationId, owner.objectId]
        );
        ownerRole = existingUser.rows[0]?.role;
        if (!ownerRole || !["owner", "admin"].includes(ownerRole)) {
          throw new OrganisationRegistrationError(
            "This organisation is already registered. An InboxLink owner or admin must change its connection settings."
          );
        }
      }
      const result = await client.query(
        `INSERT INTO organisations
           (id, microsoft_tenant_id, slug, name, halo_url, halo_client_id, status, created_at, updated_at)
         VALUES ($1, $2, $3, $4, $5, $6, 'active', $7, $7)
         ON CONFLICT (microsoft_tenant_id) DO UPDATE SET
           name = EXCLUDED.name,
           halo_url = EXCLUDED.halo_url,
           halo_client_id = EXCLUDED.halo_client_id,
           status = 'active',
           updated_at = EXCLUDED.updated_at
         RETURNING id, microsoft_tenant_id, slug, name, halo_url, halo_client_id, status`,
        [organisationId, microsoftTenantId, slug, companyName, haloUrl, haloClientId, now]
      );
      const organisation = result.rows[0];

      await client.query("SELECT set_config('app.current_organisation_id', $1, true)", [
        organisation.id,
      ]);

      const connectionChanged = Boolean(
        previous.rows[0] &&
          (previous.rows[0].halo_url !== haloUrl || previous.rows[0].halo_client_id !== haloClientId)
      );
      if (connectionChanged) {
        await client.query(
          `UPDATE halo_grants
           SET invalidated_at = $2, updated_at = $2
           WHERE organisation_id = $1 AND invalidated_at IS NULL`,
          [organisation.id, now]
        );
      }

      const user = await upsertUserWithClient(client, organisation, {
        ...owner,
        role: ownerRole,
      });
      return {
        organisation: rowToOrganisation(organisation),
        user,
      };
    });
  }

  async function getOrganisationByMicrosoftTenantId(microsoftTenantId) {
    return withMicrosoftTenant(microsoftTenantId, async (client) => {
      const result = await client.query(
        `SELECT id, microsoft_tenant_id, slug, name, halo_url, halo_client_id, status
         FROM organisations
         WHERE microsoft_tenant_id = $1 AND status = 'active'`,
        [microsoftTenantId]
      );
      return rowToOrganisation(result.rows[0]);
    });
  }

  async function upsertUser({ displayName = "", email = "", objectId, tenantId }) {
    return withMicrosoftTenant(tenantId, async (client) => {
      const organisationResult = await client.query(
        `SELECT id, microsoft_tenant_id, slug, name, halo_url, halo_client_id, status
         FROM organisations
         WHERE microsoft_tenant_id = $1 AND status = 'active'`,
        [tenantId]
      );
      const organisation = organisationResult.rows[0];
      if (!organisation) {
        throw new OrganisationNotConfiguredError();
      }

      await client.query("SELECT set_config('app.current_organisation_id', $1, true)", [
        organisation.id,
      ]);
      return upsertUserWithClient(client, organisation, {
        displayName,
        email,
        objectId,
      });
    });
  }

  async function upsertUserWithClient(client, organisation, user) {
    const now = Date.now();
    const result = await client.query(
      `INSERT INTO users
         (id, organisation_id, object_id, email, display_name, role, created_at, updated_at)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $7)
       ON CONFLICT (organisation_id, object_id) DO UPDATE SET
         email = EXCLUDED.email,
         display_name = EXCLUDED.display_name,
         role = CASE WHEN users.role = 'owner' THEN users.role ELSE EXCLUDED.role END,
         updated_at = EXCLUDED.updated_at
       RETURNING id, organisation_id, object_id, email, display_name, role`,
      [
        crypto.randomUUID(),
        organisation.id,
        user.objectId,
        user.email || "",
        user.displayName || "",
        user.role || "member",
        now,
      ]
    );
    return rowToUser(result.rows[0], organisation);
  }

  async function saveHaloGrant({
    clientId,
    encryptedToken,
    haloUrl,
    organisationId,
    scope,
    userId,
  }) {
    return withTenant(organisationId, async (client) => {
      const now = Date.now();
      await client.query(
        `INSERT INTO halo_grants
           (id, organisation_id, user_id, halo_url, client_id, scope, encrypted_token_json,
            invalidated_at, created_at, updated_at)
         VALUES ($1, $2, $3, $4, $5, $6, $7::jsonb, NULL, $8, $8)
         ON CONFLICT (organisation_id, user_id) DO UPDATE SET
           halo_url = EXCLUDED.halo_url,
           client_id = EXCLUDED.client_id,
           scope = EXCLUDED.scope,
           encrypted_token_json = EXCLUDED.encrypted_token_json,
           invalidated_at = NULL,
           updated_at = EXCLUDED.updated_at`,
        [
          crypto.randomUUID(),
          organisationId,
          userId,
          haloUrl,
          clientId,
          scope,
          JSON.stringify(encryptedToken),
          now,
        ]
      );
      return getGrantByUserIdWithClient(client, organisationId, userId);
    });
  }

  function getGrantByUserId(organisationId, userId) {
    return withTenant(organisationId, (client) =>
      getGrantByUserIdWithClient(client, organisationId, userId)
    );
  }

  async function getGrantByUserIdWithClient(client, organisationId, userId) {
    const result = await client.query(
      `SELECT id, organisation_id, user_id, halo_url, client_id, scope, encrypted_token_json
       FROM halo_grants
       WHERE organisation_id = $1 AND user_id = $2 AND invalidated_at IS NULL`,
      [organisationId, userId]
    );
    return rowToGrant(result.rows[0]);
  }

  function updateGrantToken(organisationId, grantId, encryptedToken) {
    return withTenant(organisationId, (client) =>
      client.query(
        `UPDATE halo_grants
         SET encrypted_token_json = $3::jsonb, updated_at = $4
         WHERE organisation_id = $1 AND id = $2 AND invalidated_at IS NULL`,
        [organisationId, grantId, JSON.stringify(encryptedToken), Date.now()]
      )
    );
  }

  function invalidateGrantById(organisationId, grantId) {
    const now = Date.now();
    return withTenant(organisationId, (client) =>
      client.query(
        `UPDATE halo_grants SET invalidated_at = $3, updated_at = $3
         WHERE organisation_id = $1 AND id = $2 AND invalidated_at IS NULL`,
        [organisationId, grantId, now]
      )
    );
  }

  function invalidateGrantForUser(organisationId, userId) {
    const now = Date.now();
    return withTenant(organisationId, (client) =>
      client.query(
        `UPDATE halo_grants SET invalidated_at = $3, updated_at = $3
         WHERE organisation_id = $1 AND user_id = $2 AND invalidated_at IS NULL`,
        [organisationId, userId, now]
      )
    );
  }

  function createSession({ expiresAt, organisationId, sessionHash, userId }) {
    const now = Date.now();
    return withTenant(organisationId, (client) =>
      client.query(
        `INSERT INTO sessions
           (organisation_id, session_hash, user_id, expires_at, created_at, updated_at)
         VALUES ($1, $2, $3, $4, $5, $5)
         ON CONFLICT (organisation_id, session_hash) DO UPDATE SET
           user_id = EXCLUDED.user_id,
           expires_at = EXCLUDED.expires_at,
           updated_at = EXCLUDED.updated_at`,
        [organisationId, sessionHash, userId, expiresAt, now]
      )
    );
  }

  function createBugReportSession({ diagnostics, expiresAt, organisationId, sessionHash, userId }) {
    return withTenant(organisationId, (client) =>
      client.query(
        `INSERT INTO bug_report_sessions
           (organisation_id, session_hash, user_id, diagnostics_json, expires_at,
            claimed_at, consumed_at, created_at)
         VALUES ($1, $2, $3, $4::jsonb, $5, NULL, NULL, $6)`,
        [organisationId, sessionHash, userId, JSON.stringify(diagnostics || {}), expiresAt, Date.now()]
      )
    );
  }

  function claimBugReportSession(organisationId, sessionHash, now = Date.now()) {
    return withTenant(organisationId, async (client) => {
      const result = await client.query(
        `UPDATE bug_report_sessions
         SET claimed_at = $3
         WHERE organisation_id = $1 AND session_hash = $2 AND expires_at > $3
           AND consumed_at IS NULL AND claimed_at IS NULL
         RETURNING session_hash, organisation_id, user_id, diagnostics_json, expires_at`,
        [organisationId, sessionHash, now]
      );
      const row = result.rows[0];
      return row
        ? {
            diagnostics: row.diagnostics_json || {},
            expiresAt: numberValue(row.expires_at),
            organisationId: row.organisation_id,
            sessionHash: row.session_hash,
            userId: row.user_id,
          }
        : null;
    });
  }

  function releaseBugReportSession(organisationId, sessionHash) {
    return withTenant(organisationId, (client) =>
      client.query(
        `UPDATE bug_report_sessions SET claimed_at = NULL
         WHERE organisation_id = $1 AND session_hash = $2 AND consumed_at IS NULL`,
        [organisationId, sessionHash]
      )
    );
  }

  function consumeBugReportSession(organisationId, sessionHash, now = Date.now()) {
    return withTenant(organisationId, async (client) => {
      const result = await client.query(
        `UPDATE bug_report_sessions SET claimed_at = NULL, consumed_at = $3
         WHERE organisation_id = $1 AND session_hash = $2 AND consumed_at IS NULL`,
        [organisationId, sessionHash, now]
      );
      return result.rowCount;
    });
  }

  function getSessionWithGrant(organisationId, sessionHash) {
    return withTenant(organisationId, async (client) => {
      const result = await client.query(
        `SELECT s.organisation_id, s.session_hash, s.user_id, s.expires_at,
                g.id AS grant_id, g.halo_url, g.client_id, g.scope, g.encrypted_token_json
         FROM sessions s
         JOIN halo_grants g
           ON g.organisation_id = s.organisation_id
          AND g.user_id = s.user_id
          AND g.invalidated_at IS NULL
         WHERE s.organisation_id = $1 AND s.session_hash = $2`,
        [organisationId, sessionHash]
      );
      return rowToSessionRecord(result.rows[0]);
    });
  }

  function deleteSession(organisationId, sessionHash) {
    return withTenant(organisationId, (client) =>
      client.query("DELETE FROM sessions WHERE organisation_id = $1 AND session_hash = $2", [
        organisationId,
        sessionHash,
      ])
    );
  }

  function deleteSessionsForUser(organisationId, userId) {
    return withTenant(organisationId, (client) =>
      client.query("DELETE FROM sessions WHERE organisation_id = $1 AND user_id = $2", [
        organisationId,
        userId,
      ])
    );
  }

  function createBackgroundSession({
    backgroundSessionHash,
    expiresAt,
    organisationId,
    sessionHash,
  }) {
    const now = Date.now();
    return withTenant(organisationId, (client) =>
      client.query(
        `INSERT INTO background_sessions
           (organisation_id, background_session_hash, session_hash, expires_at, created_at, updated_at)
         VALUES ($1, $2, $3, $4, $5, $5)
         ON CONFLICT (organisation_id, background_session_hash) DO UPDATE SET
           session_hash = EXCLUDED.session_hash,
           expires_at = EXCLUDED.expires_at,
           updated_at = EXCLUDED.updated_at`,
        [organisationId, backgroundSessionHash, sessionHash, expiresAt, now]
      )
    );
  }

  function getBackgroundSessionWithGrant(organisationId, backgroundSessionHash) {
    return withTenant(organisationId, async (client) => {
      const result = await client.query(
        `SELECT b.organisation_id, b.background_session_hash, b.session_hash,
                b.expires_at AS background_expires_at, s.user_id, s.expires_at,
                g.id AS grant_id, g.halo_url, g.client_id, g.scope, g.encrypted_token_json
         FROM background_sessions b
         JOIN sessions s
           ON s.organisation_id = b.organisation_id AND s.session_hash = b.session_hash
         JOIN halo_grants g
           ON g.organisation_id = s.organisation_id
          AND g.user_id = s.user_id
          AND g.invalidated_at IS NULL
         WHERE b.organisation_id = $1 AND b.background_session_hash = $2`,
        [organisationId, backgroundSessionHash]
      );
      return rowToSessionRecord(result.rows[0]);
    });
  }

  function deleteBackgroundSessionsForSessionHash(organisationId, sessionHash) {
    return withTenant(organisationId, (client) =>
      client.query(
        "DELETE FROM background_sessions WHERE organisation_id = $1 AND session_hash = $2",
        [organisationId, sessionHash]
      )
    );
  }

  function saveConversationMapping(organisationId, mapping) {
    return withTenant(organisationId, (client) =>
      client.query(
        `INSERT INTO conversation_mappings
           (organisation_id, id, mailbox_email, ticket_id, ticket_number, conversation_id,
            normalized_subject, created_at, updated_at)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
         ON CONFLICT (organisation_id, id) DO UPDATE SET
           mailbox_email = EXCLUDED.mailbox_email,
           ticket_id = EXCLUDED.ticket_id,
           ticket_number = EXCLUDED.ticket_number,
           conversation_id = EXCLUDED.conversation_id,
           normalized_subject = EXCLUDED.normalized_subject,
           updated_at = EXCLUDED.updated_at`,
        [
          organisationId,
          mapping.id,
          mapping.mailboxEmail,
          mapping.ticketId,
          mapping.ticketNumber,
          mapping.conversationId,
          mapping.normalizedSubject,
          mapping.createdAt,
          mapping.updatedAt,
        ]
      )
    );
  }

  function saveMessageMapping(organisationId, { mailboxEmail, mappingId, messageIdKey }) {
    return withTenant(organisationId, (client) =>
      client.query(
        `INSERT INTO message_mappings
           (organisation_id, mailbox_email, message_id_key, mapping_id, created_at)
         VALUES ($1, $2, $3, $4, $5)
         ON CONFLICT (organisation_id, mailbox_email, message_id_key) DO UPDATE SET
           mapping_id = EXCLUDED.mapping_id`,
        [organisationId, mailboxEmail, messageIdKey, mappingId, Date.now()]
      )
    );
  }

  function getMappingByMessageId(organisationId, mailboxEmail, messageIdKey) {
    return withTenant(organisationId, async (client) => {
      const result = await client.query(
        `SELECT cm.*
         FROM message_mappings mm
         JOIN conversation_mappings cm
           ON cm.organisation_id = mm.organisation_id AND cm.id = mm.mapping_id
         WHERE mm.organisation_id = $1 AND mm.mailbox_email = $2 AND mm.message_id_key = $3`,
        [organisationId, mailboxEmail, messageIdKey]
      );
      return rowToMapping(client, organisationId, result.rows[0]);
    });
  }

  function getMappingByConversationId(organisationId, mailboxEmail, conversationId) {
    return withTenant(organisationId, async (client) => {
      const result = await client.query(
        `SELECT * FROM conversation_mappings
         WHERE organisation_id = $1 AND mailbox_email = $2 AND conversation_id = $3
         ORDER BY updated_at DESC LIMIT 1`,
        [organisationId, mailboxEmail, conversationId]
      );
      return rowToMapping(client, organisationId, result.rows[0]);
    });
  }

  function cleanExpired(organisationId, now = Date.now()) {
    return withTenant(organisationId, async (client) => {
      await client.query(
        "DELETE FROM bug_report_sessions WHERE organisation_id = $1 AND expires_at <= $2",
        [organisationId, now]
      );
      await client.query(
        "DELETE FROM background_sessions WHERE organisation_id = $1 AND expires_at <= $2",
        [organisationId, now]
      );
      await client.query("DELETE FROM sessions WHERE organisation_id = $1 AND expires_at <= $2", [
        organisationId,
        now,
      ]);
    });
  }

  async function close() {
    await ready.catch(() => {});
    if (ownsPool) await pool.end();
  }

  async function resetForTests() {
    await ready;
    await pool.query(`
      TRUNCATE TABLE message_mappings, conversation_mappings, bug_report_sessions,
        background_sessions, sessions, halo_grants, users, organisations CASCADE
    `);
  }
}

async function assertTenantSafeRole(pool) {
  const result = await pool.query(
    "SELECT rolsuper, rolbypassrls FROM pg_roles WHERE rolname = current_user"
  );
  const role = result.rows[0];
  if (!role || role.rolsuper || role.rolbypassrls) {
    throw new Error(
      "DATABASE_URL must use a non-superuser PostgreSQL role without BYPASSRLS so tenant isolation cannot be bypassed."
    );
  }
}

function uniqueSlug(companyName, organisationId) {
  const base =
    String(companyName || "company")
      .toLowerCase()
      .normalize("NFKD")
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "")
      .slice(0, 42) || "company";
  return `${base}-${organisationId.replace(/-/g, "").slice(0, 8)}`;
}

function assertOrganisationId(value) {
  if (!UUID_PATTERN.test(String(value || ""))) {
    throw new Error("A valid organisation ID is required for tenant-scoped storage.");
  }
}

function rowToOrganisation(row) {
  if (!row) return null;
  return {
    haloClientId: row.halo_client_id,
    haloUrl: row.halo_url,
    id: row.id,
    microsoftTenantId: row.microsoft_tenant_id,
    name: row.name,
    slug: row.slug,
    status: row.status,
  };
}

function rowToUser(row, organisation) {
  if (!row) return null;
  return {
    displayName: row.display_name,
    email: row.email,
    haloClientId: organisation.halo_client_id,
    haloUrl: organisation.halo_url,
    id: row.id,
    objectId: row.object_id,
    organisationId: row.organisation_id,
    organisationName: organisation.name,
    role: row.role,
    tenantId: organisation.microsoft_tenant_id,
  };
}

function rowToGrant(row) {
  if (!row) return null;
  return {
    clientId: row.client_id,
    encryptedToken: row.encrypted_token_json,
    grantId: row.id,
    haloUrl: row.halo_url,
    organisationId: row.organisation_id,
    scope: row.scope,
    userId: row.user_id,
  };
}

function rowToSessionRecord(row) {
  if (!row) return null;
  return {
    backgroundExpiresAt: numberValue(row.background_expires_at) || null,
    clientId: row.client_id,
    encryptedToken: row.encrypted_token_json,
    expiresAt: numberValue(row.expires_at),
    grantId: row.grant_id,
    haloUrl: row.halo_url,
    organisationId: row.organisation_id,
    scope: row.scope,
    sessionHash: row.session_hash,
    userId: row.user_id,
  };
}

async function rowToMapping(client, organisationId, row) {
  if (!row) return null;
  const messageResult = await client.query(
    `SELECT message_id_key FROM message_mappings
     WHERE organisation_id = $1 AND mapping_id = $2`,
    [organisationId, row.id]
  );
  return {
    conversationId: row.conversation_id || "",
    createdAt: numberValue(row.created_at),
    id: row.id,
    mailboxEmail: row.mailbox_email,
    normalizedSubject: row.normalized_subject || "",
    organisationId,
    syncedMessageIds: new Set(messageResult.rows.map((messageRow) => messageRow.message_id_key)),
    ticketId: numberValue(row.ticket_id),
    ticketNumber: row.ticket_number,
    updatedAt: numberValue(row.updated_at),
  };
}

function numberValue(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

module.exports = {
  OrganisationNotConfiguredError,
  OrganisationRegistrationError,
  createHaloStore,
  getDatabaseConfig,
};
