const fs = require("fs");
const path = require("path");
const Database = require("better-sqlite3");

const DEFAULT_DB_PATH = path.join(process.cwd(), "data", "halo.sqlite");

function createHaloStore(options = {}) {
  const dbPath = options.dbPath || process.env.HALO_DB_PATH || DEFAULT_DB_PATH;
  const db = openDatabase(dbPath);

  initializeSchema(db);

  return {
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
    getSessionWithGrant,
    invalidateGrantById,
    invalidateGrantForUser,
    releaseBugReportSession,
    saveConversationMapping,
    saveHaloGrant,
    saveMessageMapping,
    updateGrantToken,
    upsertUser,
  };

  function upsertUser({ displayName = "", email = "", objectId, tenantId }) {
    const now = Date.now();
    db.prepare(
      `INSERT INTO users (tenant_id, object_id, email, display_name, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, ?)
       ON CONFLICT (tenant_id, object_id) DO UPDATE SET
         email = excluded.email,
         display_name = excluded.display_name,
         updated_at = excluded.updated_at`
    ).run(tenantId, objectId, email, displayName, now, now);

    return db
      .prepare(
        `SELECT id, tenant_id AS tenantId, object_id AS objectId, email, display_name AS displayName
         FROM users
         WHERE tenant_id = ? AND object_id = ?`
      )
      .get(tenantId, objectId);
  }

  function saveHaloGrant({ clientId, encryptedToken, haloUrl, scope, userId }) {
    const now = Date.now();
    db.prepare(
      `INSERT INTO halo_grants
         (user_id, halo_url, client_id, scope, encrypted_token_json, invalidated_at, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, NULL, ?, ?)
       ON CONFLICT (user_id) DO UPDATE SET
         halo_url = excluded.halo_url,
         client_id = excluded.client_id,
         scope = excluded.scope,
         encrypted_token_json = excluded.encrypted_token_json,
         invalidated_at = NULL,
         updated_at = excluded.updated_at`
    ).run(userId, haloUrl, clientId, scope, JSON.stringify(encryptedToken), now, now);

    return getGrantByUserId(userId);
  }

  function getGrantByUserId(userId) {
    const row = db
      .prepare(
        `SELECT id, user_id AS userId, halo_url AS haloUrl, client_id AS clientId, scope,
                encrypted_token_json AS encryptedTokenJson
         FROM halo_grants
         WHERE user_id = ? AND invalidated_at IS NULL`
      )
      .get(userId);

    return rowToGrant(row);
  }

  function updateGrantToken(grantId, encryptedToken) {
    db.prepare(
      `UPDATE halo_grants
       SET encrypted_token_json = ?, updated_at = ?
       WHERE id = ? AND invalidated_at IS NULL`
    ).run(JSON.stringify(encryptedToken), Date.now(), grantId);
  }

  function invalidateGrantById(grantId) {
    db.prepare(
      `UPDATE halo_grants
       SET invalidated_at = ?, updated_at = ?
       WHERE id = ? AND invalidated_at IS NULL`
    ).run(Date.now(), Date.now(), grantId);
  }

  function invalidateGrantForUser(userId) {
    db.prepare(
      `UPDATE halo_grants
       SET invalidated_at = ?, updated_at = ?
       WHERE user_id = ? AND invalidated_at IS NULL`
    ).run(Date.now(), Date.now(), userId);
  }

  function createSession({ expiresAt, sessionHash, userId }) {
    const now = Date.now();
    db.prepare(
      `INSERT INTO sessions (session_hash, user_id, expires_at, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?)
       ON CONFLICT (session_hash) DO UPDATE SET
         user_id = excluded.user_id,
         expires_at = excluded.expires_at,
         updated_at = excluded.updated_at`
    ).run(sessionHash, userId, expiresAt, now, now);
  }

  function createBugReportSession({ diagnostics, expiresAt, sessionHash, userId }) {
    const now = Date.now();
    db.prepare(
      `INSERT INTO bug_report_sessions
         (session_hash, user_id, diagnostics_json, expires_at, claimed_at, consumed_at, created_at)
       VALUES (?, ?, ?, ?, NULL, NULL, ?)`
    ).run(sessionHash, userId, JSON.stringify(diagnostics || {}), expiresAt, now);
  }

  function claimBugReportSession(sessionHash, now = Date.now()) {
    const transaction = db.transaction(() => {
      const result = db
        .prepare(
          `UPDATE bug_report_sessions
           SET claimed_at = ?
           WHERE session_hash = ?
             AND expires_at > ?
             AND consumed_at IS NULL
             AND claimed_at IS NULL`
        )
        .run(now, sessionHash, now);

      if (!result.changes) {
        return null;
      }

      return db
        .prepare(
          `SELECT brs.session_hash AS sessionHash, brs.user_id AS userId,
                  brs.diagnostics_json AS diagnosticsJson, brs.expires_at AS expiresAt
           FROM bug_report_sessions brs
           WHERE brs.session_hash = ?`
        )
        .get(sessionHash);
    });

    const row = transaction();
    if (!row) {
      return null;
    }

    return {
      diagnostics: JSON.parse(row.diagnosticsJson || "{}"),
      expiresAt: row.expiresAt,
      sessionHash: row.sessionHash,
      userId: row.userId,
    };
  }

  function releaseBugReportSession(sessionHash) {
    db.prepare(
      `UPDATE bug_report_sessions
       SET claimed_at = NULL
       WHERE session_hash = ? AND consumed_at IS NULL`
    ).run(sessionHash);
  }

  function consumeBugReportSession(sessionHash, now = Date.now()) {
    return db
      .prepare(
        `UPDATE bug_report_sessions
         SET claimed_at = NULL, consumed_at = ?
         WHERE session_hash = ? AND consumed_at IS NULL`
      )
      .run(now, sessionHash).changes;
  }

  function getSessionWithGrant(sessionHash) {
    const row = db
      .prepare(
        `SELECT s.session_hash AS sessionHash, s.user_id AS userId, s.expires_at AS expiresAt,
                g.id AS grantId, g.halo_url AS haloUrl, g.client_id AS clientId, g.scope,
                g.encrypted_token_json AS encryptedTokenJson
         FROM sessions s
         JOIN halo_grants g ON g.user_id = s.user_id AND g.invalidated_at IS NULL
         WHERE s.session_hash = ?`
      )
      .get(sessionHash);

    return rowToSessionRecord(row);
  }

  function deleteSession(sessionHash) {
    db.prepare("DELETE FROM sessions WHERE session_hash = ?").run(sessionHash);
  }

  function deleteSessionsForUser(userId) {
    const sessionRows = db
      .prepare("SELECT session_hash AS sessionHash FROM sessions WHERE user_id = ?")
      .all(userId);

    const transaction = db.transaction(() => {
      sessionRows.forEach((row) => {
        deleteBackgroundSessionsForSessionHash(row.sessionHash);
      });
      db.prepare("DELETE FROM sessions WHERE user_id = ?").run(userId);
    });

    transaction();
  }

  function createBackgroundSession({ backgroundSessionHash, expiresAt, sessionHash }) {
    const now = Date.now();
    db.prepare(
      `INSERT INTO background_sessions
         (background_session_hash, session_hash, expires_at, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?)
       ON CONFLICT (background_session_hash) DO UPDATE SET
         session_hash = excluded.session_hash,
         expires_at = excluded.expires_at,
         updated_at = excluded.updated_at`
    ).run(backgroundSessionHash, sessionHash, expiresAt, now, now);
  }

  function getBackgroundSessionWithGrant(backgroundSessionHash) {
    const row = db
      .prepare(
        `SELECT b.background_session_hash AS backgroundSessionHash,
                b.session_hash AS sessionHash,
                b.expires_at AS backgroundExpiresAt,
                s.user_id AS userId,
                s.expires_at AS expiresAt,
                g.id AS grantId,
                g.halo_url AS haloUrl,
                g.client_id AS clientId,
                g.scope,
                g.encrypted_token_json AS encryptedTokenJson
         FROM background_sessions b
         JOIN sessions s ON s.session_hash = b.session_hash
         JOIN halo_grants g ON g.user_id = s.user_id AND g.invalidated_at IS NULL
         WHERE b.background_session_hash = ?`
      )
      .get(backgroundSessionHash);

    return rowToSessionRecord(row);
  }

  function deleteBackgroundSessionsForSessionHash(sessionHash) {
    db.prepare("DELETE FROM background_sessions WHERE session_hash = ?").run(sessionHash);
  }

  function saveConversationMapping(mapping) {
    db.prepare(
      `INSERT INTO conversation_mappings
         (id, mailbox_email, ticket_id, ticket_number, conversation_id, normalized_subject, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)
       ON CONFLICT (id) DO UPDATE SET
         mailbox_email = excluded.mailbox_email,
         ticket_id = excluded.ticket_id,
         ticket_number = excluded.ticket_number,
         conversation_id = excluded.conversation_id,
         normalized_subject = excluded.normalized_subject,
         updated_at = excluded.updated_at`
    ).run(
      mapping.id,
      mapping.mailboxEmail,
      mapping.ticketId,
      mapping.ticketNumber,
      mapping.conversationId,
      mapping.normalizedSubject,
      mapping.createdAt,
      mapping.updatedAt
    );
  }

  function saveMessageMapping({ mailboxEmail, mappingId, messageIdKey }) {
    const now = Date.now();
    db.prepare(
      `INSERT INTO message_mappings
         (mailbox_email, message_id_key, mapping_id, created_at)
       VALUES (?, ?, ?, ?)
       ON CONFLICT (mailbox_email, message_id_key) DO UPDATE SET
         mapping_id = excluded.mapping_id`
    ).run(mailboxEmail, messageIdKey, mappingId, now);
  }

  function getMappingByMessageId(mailboxEmail, messageIdKey) {
    const row = db
      .prepare(
        `SELECT cm.*
         FROM message_mappings mm
         JOIN conversation_mappings cm ON cm.id = mm.mapping_id
         WHERE mm.mailbox_email = ? AND mm.message_id_key = ?`
      )
      .get(mailboxEmail, messageIdKey);

    return rowToMapping(row);
  }

  function getMappingByConversationId(mailboxEmail, conversationId) {
    const row = db
      .prepare(
        `SELECT *
         FROM conversation_mappings
         WHERE mailbox_email = ? AND conversation_id = ?
         ORDER BY updated_at DESC
         LIMIT 1`
      )
      .get(mailboxEmail, conversationId);

    return rowToMapping(row);
  }

  function cleanExpired(now = Date.now()) {
    db.prepare("DELETE FROM bug_report_sessions WHERE expires_at <= ?").run(now);
    db.prepare("DELETE FROM background_sessions WHERE expires_at <= ?").run(now);
    db.prepare("DELETE FROM sessions WHERE expires_at <= ?").run(now);
  }

  function close() {
    db.close();
  }

  function rowToGrant(row) {
    if (!row) {
      return null;
    }

    return {
      clientId: row.clientId,
      encryptedToken: JSON.parse(row.encryptedTokenJson),
      grantId: row.id,
      haloUrl: row.haloUrl,
      scope: row.scope,
      userId: row.userId,
    };
  }

  function rowToSessionRecord(row) {
    if (!row) {
      return null;
    }

    return {
      backgroundExpiresAt: row.backgroundExpiresAt || null,
      clientId: row.clientId,
      encryptedToken: JSON.parse(row.encryptedTokenJson),
      expiresAt: row.expiresAt,
      grantId: row.grantId,
      haloUrl: row.haloUrl,
      scope: row.scope,
      sessionHash: row.sessionHash,
      userId: row.userId,
    };
  }

  function rowToMapping(row) {
    if (!row) {
      return null;
    }

    const messageRows = db
      .prepare(
        `SELECT message_id_key AS messageIdKey
         FROM message_mappings
         WHERE mapping_id = ?`
      )
      .all(row.id);

    return {
      id: row.id,
      mailboxEmail: row.mailbox_email,
      ticketId: row.ticket_id,
      ticketNumber: row.ticket_number,
      conversationId: row.conversation_id || "",
      normalizedSubject: row.normalized_subject || "",
      syncedMessageIds: new Set(messageRows.map((messageRow) => messageRow.messageIdKey)),
      createdAt: row.created_at,
      updatedAt: row.updated_at,
    };
  }
}

function openDatabase(dbPath) {
  if (dbPath !== ":memory:") {
    fs.mkdirSync(path.dirname(path.resolve(dbPath)), { recursive: true });
  }

  const db = new Database(dbPath);
  db.pragma("foreign_keys = ON");
  if (dbPath !== ":memory:") {
    db.pragma("journal_mode = WAL");
  }
  return db;
}

function initializeSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      tenant_id TEXT NOT NULL,
      object_id TEXT NOT NULL,
      email TEXT NOT NULL DEFAULT '',
      display_name TEXT NOT NULL DEFAULT '',
      created_at INTEGER NOT NULL,
      updated_at INTEGER NOT NULL,
      UNIQUE (tenant_id, object_id)
    );

    CREATE TABLE IF NOT EXISTS halo_grants (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER NOT NULL UNIQUE,
      halo_url TEXT NOT NULL,
      client_id TEXT NOT NULL,
      scope TEXT NOT NULL,
      encrypted_token_json TEXT NOT NULL,
      invalidated_at INTEGER,
      created_at INTEGER NOT NULL,
      updated_at INTEGER NOT NULL,
      FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS sessions (
      session_hash TEXT PRIMARY KEY,
      user_id INTEGER NOT NULL,
      expires_at INTEGER NOT NULL,
      created_at INTEGER NOT NULL,
      updated_at INTEGER NOT NULL,
      FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS background_sessions (
      background_session_hash TEXT PRIMARY KEY,
      session_hash TEXT NOT NULL,
      expires_at INTEGER NOT NULL,
      created_at INTEGER NOT NULL,
      updated_at INTEGER NOT NULL,
      FOREIGN KEY (session_hash) REFERENCES sessions(session_hash) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS bug_report_sessions (
      session_hash TEXT PRIMARY KEY,
      user_id INTEGER NOT NULL,
      diagnostics_json TEXT NOT NULL DEFAULT '{}',
      expires_at INTEGER NOT NULL,
      claimed_at INTEGER,
      consumed_at INTEGER,
      created_at INTEGER NOT NULL,
      FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS conversation_mappings (
      id TEXT PRIMARY KEY,
      mailbox_email TEXT NOT NULL,
      ticket_id INTEGER NOT NULL,
      ticket_number TEXT NOT NULL,
      conversation_id TEXT NOT NULL DEFAULT '',
      normalized_subject TEXT NOT NULL DEFAULT '',
      created_at INTEGER NOT NULL,
      updated_at INTEGER NOT NULL
    );

    CREATE TABLE IF NOT EXISTS message_mappings (
      mailbox_email TEXT NOT NULL,
      message_id_key TEXT NOT NULL,
      mapping_id TEXT NOT NULL,
      created_at INTEGER NOT NULL,
      PRIMARY KEY (mailbox_email, message_id_key),
      FOREIGN KEY (mapping_id) REFERENCES conversation_mappings(id) ON DELETE CASCADE
    );

    CREATE INDEX IF NOT EXISTS idx_sessions_user_id ON sessions(user_id);
    CREATE INDEX IF NOT EXISTS idx_sessions_expires_at ON sessions(expires_at);
    CREATE INDEX IF NOT EXISTS idx_background_sessions_expires_at ON background_sessions(expires_at);
    CREATE INDEX IF NOT EXISTS idx_bug_report_sessions_expires_at
      ON bug_report_sessions(expires_at);
    CREATE INDEX IF NOT EXISTS idx_conversation_mappings_conversation
      ON conversation_mappings(mailbox_email, conversation_id);
    CREATE INDEX IF NOT EXISTS idx_message_mappings_mapping_id ON message_mappings(mapping_id);
  `);
}

module.exports = {
  createHaloStore,
  initializeSchema,
};
