const { checkDatabaseReady, createDatabasePool } = require("./database");
const { runMigrations } = require("./migrations");

function createHaloStore(options = {}) {
  const pool = options.pool || createDatabasePool(options);
  const ownsPool = !options.pool;
  let initializePromise = null;
  let closed = false;

  const store = {
    cleanExpired,
    claimBugReportSession,
    claimEmailAttachmentPrefetchCommit,
    claimEmailAttachmentPrefetchItem,
    close,
    consumeBugReportSession,
    consumeEmailAttachmentPrefetch,
    createBackgroundSession,
    createBugReportSession,
    createSession,
    deleteBackgroundSessionsForSessionHash,
    deleteEmailAttachmentPrefetch,
    deleteSession,
    deleteSessionsForUser,
    deleteTicketCreationIntent,
    getBackgroundSessionWithGrant,
    getComposeInlineImagePrefetch,
    getEmailAttachmentCleanupCandidates,
    getEmailAttachmentPrefetch,
    getEmailAttachmentRemovedCleanupCandidates,
    getGrantByUserId,
    getInlineImageCacheEntries,
    getMappingByConversationId,
    getMappingByMessageId,
    getSessionWithGrant,
    getTicketCreationIntent,
    getTicketCreationMetadata,
    initialize,
    invalidateGrantById,
    invalidateGrantForUser,
    isReady,
    markEmailAttachmentPrefetchForCleanup,
    markEmailAttachmentPrefetchActionCreated,
    markEmailAttachmentPrefetchItemCleaned,
    rebindEmailAttachmentPrefetch,
    releaseEmailAttachmentPrefetchCommit,
    resetEmailAttachmentPrefetchItem,
    releaseBugReportSession,
    saveComposeInlineImagePrefetch,
    saveConversationMapping,
    saveEmailAttachmentPrefetchItemFailure,
    saveEmailAttachmentPrefetchItemPrepared,
    saveHaloGrant,
    saveMessageMapping,
    saveTicketCreationMetadata,
    touchInlineImageCacheEntry,
    updateGrantToken,
    updateTicketCreationIntent,
    upsertEmailAttachmentPrefetch,
    upsertInlineImageCacheEntry,
    upsertTicketCreationIntent,
    upsertUser,
  };

  return store;

  async function initialize() {
    if (!initializePromise) {
      initializePromise = (async () => {
        await checkDatabaseReady(pool);
        await runMigrations(pool, options.migrations);
        await checkDatabaseReady(pool);
        return store;
      })();
    }
    return initializePromise;
  }

  async function isReady() {
    if (closed) {
      return false;
    }
    try {
      await checkDatabaseReady(pool);
      return true;
    } catch {
      return false;
    }
  }

  async function close() {
    if (closed) {
      return;
    }
    closed = true;
    if (ownsPool) {
      await pool.end();
    }
  }

  async function upsertUser({ displayName = "", email = "", objectId, tenantId }) {
    const now = Date.now();
    const result = await pool.query(
      `INSERT INTO users (tenant_id, object_id, email, display_name, created_at, updated_at)
       VALUES ($1, $2, $3, $4, $5, $5)
       ON CONFLICT (tenant_id, object_id) DO UPDATE SET
         email = EXCLUDED.email,
         display_name = EXCLUDED.display_name,
         updated_at = EXCLUDED.updated_at
       RETURNING id, tenant_id, object_id, email, display_name`,
      [tenantId, objectId, email, displayName, now]
    );
    return rowToUser(result.rows[0]);
  }

  async function saveHaloGrant({ clientId, encryptedToken, haloUrl, scope, userId }) {
    const now = Date.now();
    await pool.query(
      `INSERT INTO halo_grants
         (user_id, halo_url, client_id, scope, encrypted_token_json, invalidated_at,
          created_at, updated_at)
       VALUES ($1, $2, $3, $4, $5, NULL, $6, $6)
       ON CONFLICT (user_id) DO UPDATE SET
         halo_url = EXCLUDED.halo_url,
         client_id = EXCLUDED.client_id,
         scope = EXCLUDED.scope,
         encrypted_token_json = EXCLUDED.encrypted_token_json,
         invalidated_at = NULL,
         updated_at = EXCLUDED.updated_at`,
      [userId, haloUrl, clientId, scope, encryptedToken, now]
    );
    return getGrantByUserId(userId);
  }

  async function getGrantByUserId(userId) {
    const result = await pool.query(
      `SELECT g.id, g.user_id, g.halo_url, g.client_id, g.scope,
              g.encrypted_token_json, u.email AS user_email
       FROM halo_grants g
       JOIN users u ON u.id = g.user_id
       WHERE g.user_id = $1 AND g.invalidated_at IS NULL`,
      [userId]
    );
    return rowToGrant(result.rows[0]);
  }

  async function updateGrantToken(grantId, encryptedToken) {
    const result = await pool.query(
      `UPDATE halo_grants
       SET encrypted_token_json = $1, updated_at = $2
       WHERE id = $3 AND invalidated_at IS NULL`,
      [encryptedToken, Date.now(), grantId]
    );
    return result.rowCount;
  }

  async function invalidateGrantById(grantId) {
    const now = Date.now();
    const result = await pool.query(
      `UPDATE halo_grants SET invalidated_at = $1, updated_at = $1
       WHERE id = $2 AND invalidated_at IS NULL`,
      [now, grantId]
    );
    return result.rowCount;
  }

  async function invalidateGrantForUser(userId) {
    const now = Date.now();
    const result = await pool.query(
      `UPDATE halo_grants SET invalidated_at = $1, updated_at = $1
       WHERE user_id = $2 AND invalidated_at IS NULL`,
      [now, userId]
    );
    return result.rowCount;
  }

  async function createSession({ expiresAt, sessionHash, userId }) {
    const now = Date.now();
    await pool.query(
      `INSERT INTO sessions (session_hash, user_id, expires_at, created_at, updated_at)
       VALUES ($1, $2, $3, $4, $4)
       ON CONFLICT (session_hash) DO UPDATE SET
         user_id = EXCLUDED.user_id, expires_at = EXCLUDED.expires_at,
         updated_at = EXCLUDED.updated_at`,
      [sessionHash, userId, expiresAt, now]
    );
  }

  async function createBugReportSession({ diagnostics, expiresAt, sessionHash, userId }) {
    await pool.query(
      `INSERT INTO bug_report_sessions
         (session_hash, user_id, diagnostics_json, expires_at, claimed_at, consumed_at, created_at)
       VALUES ($1, $2, $3, $4, NULL, NULL, $5)`,
      [sessionHash, userId, diagnostics || {}, expiresAt, Date.now()]
    );
  }

  async function claimBugReportSession(sessionHash, now = Date.now()) {
    const result = await pool.query(
      `UPDATE bug_report_sessions SET claimed_at = $1
       WHERE session_hash = $2 AND expires_at > $1
         AND consumed_at IS NULL AND claimed_at IS NULL
       RETURNING session_hash, user_id, diagnostics_json, expires_at`,
      [now, sessionHash]
    );
    const row = result.rows[0];
    return row
      ? {
          diagnostics: jsonValue(row.diagnostics_json, {}),
          expiresAt: numberValue(row.expires_at),
          sessionHash: row.session_hash,
          userId: row.user_id,
        }
      : null;
  }

  async function releaseBugReportSession(sessionHash) {
    const result = await pool.query(
      `UPDATE bug_report_sessions SET claimed_at = NULL
       WHERE session_hash = $1 AND consumed_at IS NULL`,
      [sessionHash]
    );
    return result.rowCount;
  }

  async function consumeBugReportSession(sessionHash, now = Date.now()) {
    const result = await pool.query(
      `UPDATE bug_report_sessions SET claimed_at = NULL, consumed_at = $1
       WHERE session_hash = $2 AND consumed_at IS NULL`,
      [now, sessionHash]
    );
    return result.rowCount;
  }

  async function getSessionWithGrant(sessionHash) {
    const result = await pool.query(
      `SELECT s.session_hash, s.user_id, s.expires_at, g.id AS grant_id,
              g.halo_url, g.client_id, g.scope, g.encrypted_token_json,
              u.email AS user_email
       FROM sessions s
       JOIN halo_grants g ON g.user_id = s.user_id AND g.invalidated_at IS NULL
       JOIN users u ON u.id = s.user_id
       WHERE s.session_hash = $1`,
      [sessionHash]
    );
    return rowToSessionRecord(result.rows[0]);
  }

  async function deleteSession(sessionHash) {
    const result = await pool.query("DELETE FROM sessions WHERE session_hash = $1", [sessionHash]);
    return result.rowCount;
  }

  async function deleteSessionsForUser(userId) {
    const result = await pool.query("DELETE FROM sessions WHERE user_id = $1", [userId]);
    return result.rowCount;
  }

  async function createBackgroundSession({ backgroundSessionHash, expiresAt, sessionHash }) {
    const now = Date.now();
    await pool.query(
      `INSERT INTO background_sessions
         (background_session_hash, session_hash, expires_at, created_at, updated_at)
       VALUES ($1, $2, $3, $4, $4)
       ON CONFLICT (background_session_hash) DO UPDATE SET
         session_hash = EXCLUDED.session_hash, expires_at = EXCLUDED.expires_at,
         updated_at = EXCLUDED.updated_at`,
      [backgroundSessionHash, sessionHash, expiresAt, now]
    );
  }

  async function getBackgroundSessionWithGrant(backgroundSessionHash) {
    const result = await pool.query(
      `SELECT b.background_session_hash, b.session_hash,
              b.expires_at AS background_expires_at, s.user_id, s.expires_at,
              g.id AS grant_id, g.halo_url, g.client_id, g.scope,
              g.encrypted_token_json, u.email AS user_email
       FROM background_sessions b
       JOIN sessions s ON s.session_hash = b.session_hash
       JOIN halo_grants g ON g.user_id = s.user_id AND g.invalidated_at IS NULL
       JOIN users u ON u.id = s.user_id
       WHERE b.background_session_hash = $1`,
      [backgroundSessionHash]
    );
    return rowToSessionRecord(result.rows[0]);
  }

  async function deleteBackgroundSessionsForSessionHash(sessionHash) {
    const result = await pool.query(
      "DELETE FROM background_sessions WHERE session_hash = $1",
      [sessionHash]
    );
    return result.rowCount;
  }

  async function saveConversationMapping(mapping) {
    await pool.query(
      `INSERT INTO conversation_mappings
         (id, mailbox_email, ticket_id, ticket_number, conversation_id,
          normalized_subject, action_mode, created_at, updated_at)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
       ON CONFLICT (id) DO UPDATE SET
         mailbox_email = EXCLUDED.mailbox_email, ticket_id = EXCLUDED.ticket_id,
         ticket_number = EXCLUDED.ticket_number,
         conversation_id = EXCLUDED.conversation_id,
         normalized_subject = EXCLUDED.normalized_subject,
         action_mode = EXCLUDED.action_mode,
         updated_at = EXCLUDED.updated_at`,
      [
        mapping.id,
        mapping.mailboxEmail,
        mapping.ticketId,
        mapping.ticketNumber,
        mapping.conversationId,
        mapping.normalizedSubject,
        mapping.actionMode === "private-note" ? "private-note" : "email",
        mapping.createdAt,
        mapping.updatedAt,
      ]
    );
  }

  async function saveMessageMapping({ mailboxEmail, mappingId, messageIdKey }) {
    await pool.query(
      `INSERT INTO message_mappings (mailbox_email, message_id_key, mapping_id, created_at)
       VALUES ($1, $2, $3, $4)
       ON CONFLICT (mailbox_email, message_id_key) DO UPDATE SET mapping_id = EXCLUDED.mapping_id`,
      [mailboxEmail, messageIdKey, mappingId, Date.now()]
    );
  }

  async function getMappingByMessageId(mailboxEmail, messageIdKey) {
    const result = await pool.query(
      `SELECT cm.* FROM message_mappings mm
       JOIN conversation_mappings cm ON cm.id = mm.mapping_id
       WHERE mm.mailbox_email = $1 AND mm.message_id_key = $2`,
      [mailboxEmail, messageIdKey]
    );
    return rowToMapping(result.rows[0]);
  }

  async function getMappingByConversationId(mailboxEmail, conversationId) {
    const result = await pool.query(
      `SELECT * FROM conversation_mappings
       WHERE mailbox_email = $1 AND conversation_id = $2
       ORDER BY updated_at DESC LIMIT 1`,
      [mailboxEmail, conversationId]
    );
    return rowToMapping(result.rows[0]);
  }

  async function getInlineImageCacheEntries(haloTenant, hashes, showForUsers = true) {
    const normalizedHashes = Array.from(
      new Set((hashes || []).map((value) => String(value || "").toLowerCase()))
    ).filter(Boolean);
    if (!normalizedHashes.length) {
      return [];
    }
    const result = await pool.query(
      `SELECT * FROM inline_image_cache
       WHERE halo_tenant = $1 AND sha256 = ANY($2::text[]) AND show_for_users = $3`,
      [haloTenant, normalizedHashes, Boolean(showForUsers)]
    );
    return result.rows.map(rowToInlineImageCacheEntry);
  }

  async function upsertInlineImageCacheEntry(entry) {
    const now = Date.now();
    const result = await pool.query(
      `INSERT INTO inline_image_cache
         (halo_tenant, sha256, halo_attachment_id, renderable_url, media_type,
          byte_length, filename, show_for_users, created_at, updated_at, last_used_at, usage_count)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $9, $9, 1)
       ON CONFLICT (halo_tenant, sha256, show_for_users) DO UPDATE SET
         halo_attachment_id = EXCLUDED.halo_attachment_id,
         renderable_url = EXCLUDED.renderable_url, media_type = EXCLUDED.media_type,
         byte_length = EXCLUDED.byte_length, filename = EXCLUDED.filename,
         updated_at = EXCLUDED.updated_at, last_used_at = EXCLUDED.last_used_at,
         usage_count = inline_image_cache.usage_count + 1
       RETURNING *`,
      [
        entry.haloTenant,
        String(entry.sha256 || "").toLowerCase(),
        String(entry.haloAttachmentId || ""),
        entry.renderableUrl,
        entry.mediaType,
        entry.byteLength,
        entry.filename,
        entry.showForUsers !== false,
        now,
      ]
    );
    return rowToInlineImageCacheEntry(result.rows[0]);
  }

  async function touchInlineImageCacheEntry(haloTenant, sha256, showForUsers = true) {
    const now = Date.now();
    const result = await pool.query(
      `UPDATE inline_image_cache
       SET last_used_at = $1, updated_at = $1, usage_count = usage_count + 1
       WHERE halo_tenant = $2 AND sha256 = $3 AND show_for_users = $4`,
      [now, haloTenant, String(sha256 || "").toLowerCase(), Boolean(showForUsers)]
    );
    return result.rowCount;
  }

  async function saveComposeInlineImagePrefetch(prefetch) {
    const now = Date.now();
    await pool.query(
      `INSERT INTO compose_inline_image_prefetch
         (prefetch_key, compose_operation_id, attachment_fingerprint, cid_hash_json,
          halo_tenant, expires_at, created_at, updated_at)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $7)
       ON CONFLICT (prefetch_key) DO UPDATE SET
         compose_operation_id = EXCLUDED.compose_operation_id,
         attachment_fingerprint = EXCLUDED.attachment_fingerprint,
         cid_hash_json = EXCLUDED.cid_hash_json, halo_tenant = EXCLUDED.halo_tenant,
         expires_at = EXCLUDED.expires_at, updated_at = EXCLUDED.updated_at`,
      [
        prefetch.prefetchKey,
        prefetch.composeOperationId,
        prefetch.attachmentFingerprint,
        prefetch.cidHash || {},
        prefetch.haloTenant,
        prefetch.expiresAt,
        now,
      ]
    );
  }

  async function getComposeInlineImagePrefetch(prefetchKey, haloTenant, now = Date.now()) {
    const result = await pool.query(
      `SELECT * FROM compose_inline_image_prefetch
       WHERE prefetch_key = $1 AND halo_tenant = $2 AND expires_at > $3`,
      [prefetchKey, haloTenant, now]
    );
    const row = result.rows[0];
    return row
      ? {
          attachmentFingerprint: row.attachment_fingerprint,
          cidHash: objectValue(row.cid_hash_json),
          cidHashJson: JSON.stringify(row.cid_hash_json),
          composeOperationId: row.compose_operation_id,
          expiresAt: numberValue(row.expires_at),
          haloTenant: row.halo_tenant,
          prefetchKey: row.prefetch_key,
        }
      : null;
  }

  async function upsertEmailAttachmentPrefetch(prefetch, descriptors) {
    return withTransaction(async (client) => {
      const now = Date.now();
      const existingResult = await client.query(
        `SELECT prefetch_key, expires_at, staging_version, status
         FROM email_attachment_prefetch
         WHERE user_id = $1 AND operation_id = $2 AND halo_tenant = $3 AND ticket_id = $4
         FOR UPDATE`,
        [prefetch.userId, prefetch.operationId, prefetch.haloTenant, prefetch.ticketId]
      );
      const existing = existingResult.rows[0];
      const restartPreparation = Boolean(
        existing &&
          existing.status !== "consumed" &&
          (Number(existing.staging_version || 1) !== 2 || numberValue(existing.expires_at) <= now)
      );
      const parentResult = await client.query(
        `INSERT INTO email_attachment_prefetch
           (prefetch_key, user_id, halo_tenant, ticket_id, operation_id,
            attachment_fingerprint, expected_count, expected_bytes, status,
            expires_at, consumed_at, created_at, updated_at, staging_version, draft_item_id)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, 'active', $9, NULL, $10, $10, 2, $11)
         ON CONFLICT (user_id, operation_id, halo_tenant, ticket_id) DO UPDATE SET
           attachment_fingerprint = CASE WHEN email_attachment_prefetch.status = 'consumed'
             OR (email_attachment_prefetch.status = 'committing'
               AND email_attachment_prefetch.expires_at > $10)
             THEN email_attachment_prefetch.attachment_fingerprint
             ELSE EXCLUDED.attachment_fingerprint END,
           expected_count = CASE WHEN email_attachment_prefetch.status = 'consumed'
             OR (email_attachment_prefetch.status = 'committing'
               AND email_attachment_prefetch.expires_at > $10)
             THEN email_attachment_prefetch.expected_count ELSE EXCLUDED.expected_count END,
           expected_bytes = CASE WHEN email_attachment_prefetch.status = 'consumed'
             OR (email_attachment_prefetch.status = 'committing'
               AND email_attachment_prefetch.expires_at > $10)
             THEN email_attachment_prefetch.expected_bytes ELSE EXCLUDED.expected_bytes END,
           staging_version = CASE WHEN email_attachment_prefetch.status = 'consumed'
             OR (email_attachment_prefetch.status = 'committing'
               AND email_attachment_prefetch.expires_at > $10)
             THEN email_attachment_prefetch.staging_version ELSE 2 END,
           draft_item_id = CASE WHEN email_attachment_prefetch.status = 'consumed'
             OR (email_attachment_prefetch.status = 'committing'
               AND email_attachment_prefetch.expires_at > $10)
             THEN email_attachment_prefetch.draft_item_id ELSE EXCLUDED.draft_item_id END,
           status = CASE WHEN email_attachment_prefetch.status = 'consumed'
             OR (email_attachment_prefetch.status = 'committing'
               AND email_attachment_prefetch.expires_at > $10)
             THEN email_attachment_prefetch.status ELSE 'active' END,
           expires_at = CASE WHEN email_attachment_prefetch.status = 'consumed'
             OR (email_attachment_prefetch.status = 'committing'
               AND email_attachment_prefetch.expires_at > $10)
             THEN email_attachment_prefetch.expires_at ELSE EXCLUDED.expires_at END,
           updated_at = CASE WHEN email_attachment_prefetch.status = 'consumed'
             OR (email_attachment_prefetch.status = 'committing'
               AND email_attachment_prefetch.expires_at > $10)
             THEN email_attachment_prefetch.updated_at ELSE EXCLUDED.updated_at END
         RETURNING *`,
        [
          prefetch.prefetchKey,
          prefetch.userId,
          prefetch.haloTenant,
          prefetch.ticketId,
          prefetch.operationId,
          prefetch.attachmentFingerprint,
          prefetch.expectedCount,
          prefetch.expectedBytes,
          prefetch.expiresAt,
          now,
          prefetch.draftItemId || "",
        ]
      );
      const parent = parentResult.rows[0];
      if (["consumed", "committing"].includes(parent.status)) {
        return rowToEmailAttachmentPrefetch(parent);
      }

      if (restartPreparation) {
        await client.query(
          `UPDATE email_attachment_prefetch_items
           SET status = 'pending', decoded_size = 0, failure_code = '',
               content_ciphertext = NULL, content_iv = NULL, content_tag = NULL,
               content_key_id = NULL, content_sha256 = NULL,
               expected_content_sha256 = NULL, prepared_at = NULL,
               updated_at = $1
           WHERE prefetch_key = $2 AND status != 'removed'`,
          [now, parent.prefetch_key]
        );
      }

      const keys = [];
      for (const descriptor of descriptors) {
        keys.push(descriptor.attachmentKey);
        await client.query(
          `INSERT INTO email_attachment_prefetch_items
             (prefetch_key, attachment_key, filename, content_type, attachment_type,
              reported_size, decoded_size, status, halo_attachment_id, halo_filename,
              halo_filesize, halo_type, show_for_users, failure_code, created_at, updated_at,
              expected_content_sha256)
           VALUES ($1, $2, $3, $4, $5, $6, 0, 'pending', NULL, '', 0, 0, TRUE, '', $7, $7,
             NULLIF($8, ''))
           ON CONFLICT (prefetch_key, attachment_key) DO UPDATE SET
             filename = EXCLUDED.filename, content_type = EXCLUDED.content_type,
             attachment_type = EXCLUDED.attachment_type,
             reported_size = EXCLUDED.reported_size,
             status = CASE WHEN email_attachment_prefetch_items.status = 'prepared'
               AND email_attachment_prefetch_items.content_sha256 = $8
               THEN 'prepared'
               WHEN email_attachment_prefetch_items.status = 'preparing'
                 AND email_attachment_prefetch_items.expected_content_sha256 = $8
               THEN 'preparing' ELSE 'pending' END,
             decoded_size = CASE WHEN (
                 email_attachment_prefetch_items.status = 'prepared'
                 AND email_attachment_prefetch_items.content_sha256 = $8
               ) OR (
                 email_attachment_prefetch_items.status = 'preparing'
                 AND email_attachment_prefetch_items.expected_content_sha256 = $8
               )
               THEN email_attachment_prefetch_items.decoded_size ELSE 0 END,
             content_ciphertext = CASE WHEN email_attachment_prefetch_items.status = 'prepared'
               AND email_attachment_prefetch_items.content_sha256 = $8
               THEN email_attachment_prefetch_items.content_ciphertext ELSE NULL END,
             content_iv = CASE WHEN email_attachment_prefetch_items.status = 'prepared'
               AND email_attachment_prefetch_items.content_sha256 = $8
               THEN email_attachment_prefetch_items.content_iv ELSE NULL END,
             content_tag = CASE WHEN email_attachment_prefetch_items.status = 'prepared'
               AND email_attachment_prefetch_items.content_sha256 = $8
               THEN email_attachment_prefetch_items.content_tag ELSE NULL END,
             content_key_id = CASE WHEN email_attachment_prefetch_items.status = 'prepared'
               AND email_attachment_prefetch_items.content_sha256 = $8
               THEN email_attachment_prefetch_items.content_key_id ELSE NULL END,
             content_sha256 = CASE WHEN email_attachment_prefetch_items.status = 'prepared'
               AND email_attachment_prefetch_items.content_sha256 = $8
               THEN email_attachment_prefetch_items.content_sha256 ELSE NULL END,
             prepared_at = CASE WHEN email_attachment_prefetch_items.status = 'prepared'
               AND email_attachment_prefetch_items.content_sha256 = $8
               THEN email_attachment_prefetch_items.prepared_at ELSE NULL END,
             expected_content_sha256 = NULLIF($8, ''),
             failure_code = '', updated_at = EXCLUDED.updated_at`,
          [
            parent.prefetch_key,
            descriptor.attachmentKey,
            descriptor.name,
            descriptor.contentType,
            descriptor.attachmentType,
            descriptor.reportedSize,
            now,
            descriptor.contentSha256 || "",
          ]
        );
      }
      await client.query(
        `UPDATE email_attachment_prefetch_items
         SET status = 'removed', content_ciphertext = NULL, content_iv = NULL,
             content_tag = NULL, content_key_id = NULL, content_sha256 = NULL,
             expected_content_sha256 = NULL, prepared_at = NULL, updated_at = $1
         WHERE prefetch_key = $2 AND NOT (attachment_key = ANY($3::text[]))
           AND status != 'cleaned'`,
        [now, parent.prefetch_key, keys]
      );
      return rowToEmailAttachmentPrefetch(parent);
    });
  }

  async function getEmailAttachmentPrefetch(
    prefetchKey,
    { haloTenant, ticketId, userId } = {},
    now = Date.now()
  ) {
    const conditions = ["prefetch_key = $1", "expires_at > $2"];
    const parameters = [prefetchKey, now];
    if (userId !== undefined && userId !== null) {
      parameters.push(userId);
      conditions.push(`user_id = $${parameters.length}`);
    }
    if (haloTenant) {
      parameters.push(haloTenant);
      conditions.push(`halo_tenant = $${parameters.length}`);
    }
    if (ticketId !== undefined && ticketId !== null) {
      parameters.push(ticketId);
      conditions.push(`ticket_id = $${parameters.length}`);
    }
    const result = await pool.query(
      `SELECT * FROM email_attachment_prefetch WHERE ${conditions.join(" AND ")}`,
      parameters
    );
    return result.rows[0] ? loadEmailAttachmentPrefetch(result.rows[0]) : null;
  }

  async function claimEmailAttachmentPrefetchItem(
    prefetchKey,
    attachmentKey,
    decodedSize,
    expectedContentSha256,
    maximumTotalBytes,
    now = Date.now()
  ) {
    return withTransaction(async (client) => {
      const staleBefore = now - 5 * 60 * 1000;
      await client.query(
        `UPDATE email_attachment_prefetch_items
         SET status = 'failed', failure_code = 'stale-preparation', updated_at = $1
         WHERE prefetch_key = $2 AND status = 'preparing' AND updated_at < $3`,
        [now, prefetchKey, staleBefore]
      );
      const parentResult = await client.query(
        `SELECT status, expires_at FROM email_attachment_prefetch
         WHERE prefetch_key = $1 FOR UPDATE`,
        [prefetchKey]
      );
      const itemResult = await client.query(
        `SELECT * FROM email_attachment_prefetch_items
         WHERE prefetch_key = $1 AND attachment_key = $2 FOR UPDATE`,
        [prefetchKey, attachmentKey]
      );
      const parent = parentResult.rows[0];
      let item = itemResult.rows[0];
      if (!parent || parent.status !== "active" || numberValue(parent.expires_at) <= now || !item) {
        return item || null;
      }
      if (item.status === "prepared") {
        return item;
      }
      if (
        !["pending", "failed"].includes(item.status) ||
        item.expected_content_sha256 !== expectedContentSha256
      ) {
        return null;
      }
      const reservedResult = await client.query(
        `SELECT COALESCE(SUM(decoded_size), 0)::bigint AS total
         FROM email_attachment_prefetch_items
         WHERE prefetch_key = $1 AND attachment_key != $2
           AND status IN ('preparing', 'prepared')`,
        [prefetchKey, attachmentKey]
      );
      const reserved = numberValue(reservedResult.rows[0].total);
      const updateResult =
        reserved + decodedSize > maximumTotalBytes
          ? await client.query(
              `UPDATE email_attachment_prefetch_items
               SET status = 'failed', failure_code = 'total-size-limit', updated_at = $1
               WHERE prefetch_key = $2 AND attachment_key = $3 RETURNING *`,
              [now, prefetchKey, attachmentKey]
            )
          : await client.query(
              `UPDATE email_attachment_prefetch_items
               SET status = 'preparing', decoded_size = $1, failure_code = '', updated_at = $2
               WHERE prefetch_key = $3 AND attachment_key = $4 RETURNING *`,
              [decodedSize, now, prefetchKey, attachmentKey]
            );
      item = updateResult.rows[0];
      return item || null;
    });
  }

  async function saveEmailAttachmentPrefetchItemPrepared(prefetchKey, attachmentKey, encrypted) {
    const result = await pool.query(
      `UPDATE email_attachment_prefetch_items
       SET decoded_size = $1, status = 'prepared', content_ciphertext = $2,
           content_iv = $3, content_tag = $4, content_key_id = $5,
           content_sha256 = $6, prepared_at = $7, failure_code = '', updated_at = $7,
           halo_attachment_id = NULL, halo_filename = '', halo_filesize = 0, halo_type = 0
       WHERE prefetch_key = $8 AND attachment_key = $9 AND status = 'preparing'
         AND expected_content_sha256 = $6
         AND EXISTS (SELECT 1 FROM email_attachment_prefetch p
           WHERE p.prefetch_key = email_attachment_prefetch_items.prefetch_key
             AND p.status = 'active' AND p.staging_version = 2)`,
      [
        encrypted.decodedSize,
        encrypted.ciphertext,
        encrypted.iv,
        encrypted.tag,
        encrypted.keyId,
        encrypted.contentSha256,
        encrypted.preparedAt || Date.now(),
        prefetchKey,
        attachmentKey,
      ]
    );
    return result.rowCount;
  }

  async function saveEmailAttachmentPrefetchItemFailure(
    prefetchKey,
    attachmentKey,
    expectedContentSha256,
    failureCode
  ) {
    const result = await pool.query(
      `UPDATE email_attachment_prefetch_items
       SET status = 'failed', failure_code = $1, content_ciphertext = NULL,
           content_iv = NULL, content_tag = NULL, content_key_id = NULL,
           content_sha256 = NULL, prepared_at = NULL, updated_at = $2
       WHERE prefetch_key = $3 AND attachment_key = $4 AND status = 'preparing'
         AND expected_content_sha256 = $5`,
      [
        String(failureCode || "preparation-failed").slice(0, 80),
        Date.now(),
        prefetchKey,
        attachmentKey,
        expectedContentSha256,
      ]
    );
    return result.rowCount;
  }

  async function resetEmailAttachmentPrefetchItem(prefetchKey, attachmentKey) {
    const result = await pool.query(
      `UPDATE email_attachment_prefetch_items
       SET status = 'pending', decoded_size = 0, failure_code = '',
           content_ciphertext = NULL, content_iv = NULL, content_tag = NULL,
           content_key_id = NULL, content_sha256 = NULL, prepared_at = NULL,
           updated_at = $1
       WHERE prefetch_key = $2 AND attachment_key = $3 AND status = 'prepared'
         AND EXISTS (SELECT 1 FROM email_attachment_prefetch p
           WHERE p.prefetch_key = email_attachment_prefetch_items.prefetch_key
             AND p.status = 'active' AND p.staging_version = 2)`,
      [Date.now(), prefetchKey, attachmentKey]
    );
    return result.rowCount;
  }

  async function consumeEmailAttachmentPrefetch(prefetchKey, now = Date.now()) {
    return withTransaction(async (client) => {
      const result = await client.query(
        `UPDATE email_attachment_prefetch SET status = 'consumed', consumed_at = $1, updated_at = $1
         WHERE prefetch_key = $2 AND status IN ('active', 'committing')`,
        [now, prefetchKey]
      );
      if (result.rowCount) {
        await client.query(
          `UPDATE email_attachment_prefetch_items
           SET content_ciphertext = NULL, content_iv = NULL, content_tag = NULL,
               content_key_id = NULL, content_sha256 = NULL, prepared_at = NULL,
               status = CASE WHEN status = 'prepared' THEN 'consumed' ELSE status END,
               updated_at = $1 WHERE prefetch_key = $2`,
          [now, prefetchKey]
        );
      }
      return result.rowCount;
    });
  }

  async function claimEmailAttachmentPrefetchCommit(
    prefetchKey,
    { haloTenant, ticketId, userId },
    now = Date.now()
  ) {
    const result = await pool.query(
      `UPDATE email_attachment_prefetch SET status = 'committing', updated_at = $1
       WHERE prefetch_key = $2 AND user_id = $3 AND halo_tenant = $4 AND ticket_id = $5
         AND staging_version = 2
         AND (status = 'active' OR (status = 'committing' AND updated_at < $6))
         AND expires_at > $1
       RETURNING *`,
      [now, prefetchKey, userId, haloTenant, ticketId, now - 5 * 60 * 1000]
    );
    return result.rows[0] ? loadEmailAttachmentPrefetch(result.rows[0]) : null;
  }

  async function releaseEmailAttachmentPrefetchCommit(prefetchKey, now = Date.now()) {
    const result = await pool.query(
      `UPDATE email_attachment_prefetch SET status = 'active', updated_at = $1
       WHERE prefetch_key = $2 AND status = 'committing' AND expires_at > $1`,
      [now, prefetchKey]
    );
    return result.rowCount;
  }

  async function rebindEmailAttachmentPrefetch(prefetchKey, { haloTenant, ticketId, userId }) {
    const normalizedTicketId = Number(ticketId);
    if (!Number.isSafeInteger(normalizedTicketId) || normalizedTicketId <= 0) {
      return 0;
    }
    const result = await pool.query(
      `UPDATE email_attachment_prefetch SET ticket_id = $1, updated_at = $2
       WHERE prefetch_key = $3 AND user_id = $4 AND halo_tenant = $5
         AND ticket_id = 0 AND status IN ('active', 'committing')`,
      [normalizedTicketId, Date.now(), prefetchKey, userId, haloTenant]
    );
    return result.rowCount;
  }

  async function getEmailAttachmentCleanupCandidates(now = Date.now(), limit = 50) {
    const result = await pool.query(
      `SELECT * FROM email_attachment_prefetch
       WHERE status IN ('cancelled', 'cleanup')
          OR (status IN ('active', 'committing') AND expires_at <= $1)
       ORDER BY updated_at ASC LIMIT $2`,
      [now, limit]
    );
    return Promise.all(result.rows.map(loadEmailAttachmentPrefetch));
  }

  async function getEmailAttachmentRemovedCleanupCandidates(limit = 50) {
    const result = await pool.query(
      `SELECT DISTINCT p.* FROM email_attachment_prefetch p
       JOIN email_attachment_prefetch_items i ON i.prefetch_key = p.prefetch_key
       WHERE p.status = 'active' AND i.status = 'removed'
         AND i.halo_attachment_id IS NOT NULL
       ORDER BY p.updated_at ASC LIMIT $1`,
      [limit]
    );
    return Promise.all(result.rows.map(loadEmailAttachmentPrefetch));
  }

  async function markEmailAttachmentPrefetchForCleanup(prefetchKey) {
    const result = await pool.query(
      `UPDATE email_attachment_prefetch SET status = 'cleanup', updated_at = $1
       WHERE prefetch_key = $2 AND status != 'consumed'`,
      [Date.now(), prefetchKey]
    );
    return result.rowCount;
  }

  async function markEmailAttachmentPrefetchActionCreated(prefetchKey, actionId, now = Date.now()) {
    const result = await pool.query(
      `UPDATE email_attachment_prefetch
       SET halo_action_id = $1, action_created_at = $2, updated_at = $2
       WHERE prefetch_key = $3 AND staging_version = 2 AND status = 'committing'`,
      [String(actionId || "created").slice(0, 200), now, prefetchKey]
    );
    return result.rowCount;
  }

  async function markEmailAttachmentPrefetchItemCleaned(prefetchKey, attachmentKey) {
    const result = await pool.query(
      `UPDATE email_attachment_prefetch_items
       SET status = CASE WHEN status = 'removed' THEN 'cleaned' ELSE status END,
           halo_attachment_id = NULL, updated_at = $1
       WHERE prefetch_key = $2 AND attachment_key = $3
         AND (status = 'removed' OR EXISTS (
           SELECT 1 FROM email_attachment_prefetch p
           WHERE p.prefetch_key = email_attachment_prefetch_items.prefetch_key
             AND (p.status IN ('cleanup', 'cancelled') OR p.staging_version = 2)))`,
      [Date.now(), prefetchKey, attachmentKey]
    );
    return result.rowCount;
  }

  async function deleteEmailAttachmentPrefetch(prefetchKey) {
    const result = await pool.query(
      "DELETE FROM email_attachment_prefetch WHERE prefetch_key = $1",
      [prefetchKey]
    );
    return result.rowCount;
  }

  async function saveTicketCreationMetadata({ cacheKey, expiresAt, haloTenant, payload, userId }) {
    const now = Date.now();
    await pool.query(
      `INSERT INTO ticket_creation_metadata_cache
         (cache_key, user_id, halo_tenant, payload_json, fetched_at,
          expires_at, created_at, updated_at)
       VALUES ($1, $2, $3, $4, $5, $6, $5, $5)
       ON CONFLICT (cache_key, user_id, halo_tenant) DO UPDATE SET
         payload_json = EXCLUDED.payload_json, fetched_at = EXCLUDED.fetched_at,
         expires_at = EXCLUDED.expires_at, updated_at = EXCLUDED.updated_at`,
      [cacheKey, userId, haloTenant, payload, now, expiresAt]
    );
    return getTicketCreationMetadata(cacheKey, { haloTenant, userId });
  }

  async function getTicketCreationMetadata(cacheKey, { haloTenant, userId }) {
    const result = await pool.query(
      `SELECT cache_key, payload_json, fetched_at, expires_at
       FROM ticket_creation_metadata_cache
       WHERE cache_key = $1 AND user_id = $2 AND halo_tenant = $3`,
      [cacheKey, userId, haloTenant]
    );
    const row = result.rows[0];
    return row
      ? {
          cacheKey: row.cache_key,
          expiresAt: numberValue(row.expires_at),
          fetchedAt: numberValue(row.fetched_at),
          payload: jsonValue(row.payload_json, null),
          payloadJson: JSON.stringify(row.payload_json),
        }
      : null;
  }

  async function upsertTicketCreationIntent({
    encryptedIntent,
    expiresAt,
    haloTenant,
    operationId,
    userId,
  }) {
    const now = Date.now();
    await pool.query(
      `INSERT INTO ticket_creation_intents
         (operation_id, user_id, halo_tenant, encrypted_intent_json, status,
          ticket_id, ticket_number, action_id, last_error, expires_at, created_at, updated_at)
       VALUES ($1, $2, $3, $4, 'pending', NULL, '', '', '', $5, $6, $6)
       ON CONFLICT (operation_id, user_id, halo_tenant) DO UPDATE SET
         encrypted_intent_json = EXCLUDED.encrypted_intent_json,
         expires_at = EXCLUDED.expires_at, updated_at = EXCLUDED.updated_at`,
      [operationId, userId, haloTenant, encryptedIntent, expiresAt, now]
    );
    return getTicketCreationIntent(operationId, { haloTenant, userId });
  }

  async function getTicketCreationIntent(operationId, { haloTenant, userId }, now = Date.now()) {
    const result = await pool.query(
      `SELECT * FROM ticket_creation_intents
       WHERE operation_id = $1 AND user_id = $2 AND halo_tenant = $3 AND expires_at > $4`,
      [operationId, userId, haloTenant, now]
    );
    return rowToTicketCreationIntent(result.rows[0]);
  }

  async function updateTicketCreationIntent(
    operationId,
    { actionId, encryptedIntent, lastError, status, ticketId, ticketNumber },
    { haloTenant, userId }
  ) {
    return withTransaction(async (client) => {
      const currentResult = await client.query(
        `SELECT * FROM ticket_creation_intents
         WHERE operation_id = $1 AND user_id = $2 AND halo_tenant = $3
           AND expires_at > $4 FOR UPDATE`,
        [operationId, userId, haloTenant, Date.now()]
      );
      const current = rowToTicketCreationIntent(currentResult.rows[0]);
      if (!current) {
        return null;
      }
      const result = await client.query(
        `UPDATE ticket_creation_intents
         SET encrypted_intent_json = $1, status = $2, ticket_id = $3,
             ticket_number = $4, action_id = $5, last_error = $6, updated_at = $7
         WHERE operation_id = $8 AND user_id = $9 AND halo_tenant = $10 RETURNING *`,
        [
          encryptedIntent || current.encryptedIntent,
          status || current.status,
          ticketId === undefined ? current.ticketId : ticketId,
          ticketNumber === undefined ? current.ticketNumber : String(ticketNumber || ""),
          actionId === undefined ? current.actionId : String(actionId || ""),
          lastError === undefined ? current.lastError : String(lastError || "").slice(0, 500),
          Date.now(),
          operationId,
          userId,
          haloTenant,
        ]
      );
      return rowToTicketCreationIntent(result.rows[0]);
    });
  }

  async function deleteTicketCreationIntent(operationId, { haloTenant, userId }) {
    const result = await pool.query(
      `DELETE FROM ticket_creation_intents
       WHERE operation_id = $1 AND user_id = $2 AND halo_tenant = $3 AND ticket_id IS NULL`,
      [operationId, userId, haloTenant]
    );
    return result.rowCount;
  }

  async function loadEmailAttachmentPrefetch(row) {
    const record = rowToEmailAttachmentPrefetch(row);
    const result = await pool.query(
      `SELECT * FROM email_attachment_prefetch_items
       WHERE prefetch_key = $1 ORDER BY created_at, attachment_key`,
      [record.prefetchKey]
    );
    record.items = result.rows.map(rowToEmailAttachmentPrefetchItem);
    return record;
  }

  async function cleanExpired(now = Date.now()) {
    return withTransaction(async (client) => {
      const statements = [
        ["DELETE FROM bug_report_sessions WHERE expires_at <= $1", [now]],
        ["DELETE FROM background_sessions WHERE expires_at <= $1", [now]],
        ["DELETE FROM sessions WHERE expires_at <= $1", [now]],
        ["DELETE FROM compose_inline_image_prefetch WHERE expires_at <= $1", [now]],
        [
          "DELETE FROM email_attachment_prefetch WHERE status = 'consumed' AND expires_at <= $1",
          [now],
        ],
        ["DELETE FROM ticket_creation_intents WHERE expires_at <= $1", [now]],
        [
          "DELETE FROM ticket_creation_metadata_cache WHERE expires_at <= $1",
          [now - 24 * 60 * 60 * 1000],
        ],
      ];
      let deleted = 0;
      for (const [sql, parameters] of statements) {
        const result = await client.query(sql, parameters);
        deleted += result.rowCount;
      }
      return deleted;
    });
  }

  async function rowToMapping(row) {
    if (!row) {
      return null;
    }
    const result = await pool.query(
      "SELECT message_id_key FROM message_mappings WHERE mapping_id = $1",
      [row.id]
    );
    return {
      conversationId: row.conversation_id || "",
      createdAt: numberValue(row.created_at),
      id: row.id,
      mailboxEmail: row.mailbox_email,
      normalizedSubject: row.normalized_subject || "",
      actionMode: row.action_mode === "private-note" ? "private-note" : "email",
      syncedMessageIds: new Set(result.rows.map((messageRow) => messageRow.message_id_key)),
      ticketId: row.ticket_id,
      ticketNumber: row.ticket_number,
      updatedAt: numberValue(row.updated_at),
    };
  }

  async function withTransaction(callback) {
    const client = await pool.connect();
    let releaseError;
    try {
      await client.query("BEGIN");
      const value = await callback(client);
      await client.query("COMMIT");
      return value;
    } catch (error) {
      try {
        await client.query("ROLLBACK");
      } catch (rollbackError) {
        error.rollbackError = rollbackError;
        releaseError = rollbackError;
      }
      throw error;
    } finally {
      client.release(releaseError);
    }
  }
}

function rowToUser(row) {
  return row
    ? {
        displayName: row.display_name,
        email: row.email,
        id: row.id,
        objectId: row.object_id,
        tenantId: row.tenant_id,
      }
    : null;
}

function rowToGrant(row) {
  return row
    ? {
        clientId: row.client_id,
        encryptedToken: jsonValue(row.encrypted_token_json, null),
        grantId: row.id,
        haloUrl: row.halo_url,
        scope: row.scope,
        userEmail: row.user_email || "",
        userId: row.user_id,
      }
    : null;
}

function rowToSessionRecord(row) {
  return row
    ? {
        backgroundExpiresAt: nullableNumber(row.background_expires_at),
        clientId: row.client_id,
        encryptedToken: jsonValue(row.encrypted_token_json, null),
        expiresAt: numberValue(row.expires_at),
        grantId: row.grant_id,
        haloUrl: row.halo_url,
        scope: row.scope,
        sessionHash: row.session_hash,
        userEmail: row.user_email || "",
        userId: row.user_id,
      }
    : null;
}

function rowToInlineImageCacheEntry(row) {
  return {
    byteLength: row.byte_length,
    createdAt: numberValue(row.created_at),
    filename: row.filename,
    haloAttachmentId: row.halo_attachment_id,
    haloTenant: row.halo_tenant,
    lastUsedAt: numberValue(row.last_used_at),
    mediaType: row.media_type,
    renderableUrl: row.renderable_url,
    sha256: row.sha256,
    showForUsers: row.show_for_users !== false,
    updatedAt: numberValue(row.updated_at),
    usageCount: row.usage_count,
  };
}

function rowToEmailAttachmentPrefetch(row) {
  return {
    attachmentFingerprint: row.attachment_fingerprint,
    actionCreatedAt: nullableNumber(row.action_created_at),
    consumedAt: nullableNumber(row.consumed_at),
    createdAt: numberValue(row.created_at),
    expectedBytes: row.expected_bytes,
    expectedCount: row.expected_count,
    draftItemId: row.draft_item_id || "",
    expiresAt: numberValue(row.expires_at),
    haloTenant: row.halo_tenant,
    haloActionId: row.halo_action_id || "",
    operationId: row.operation_id,
    prefetchKey: row.prefetch_key,
    status: row.status,
    stagingVersion: Number(row.staging_version || 1),
    ticketId: row.ticket_id,
    updatedAt: numberValue(row.updated_at),
    userId: row.user_id,
  };
}

function rowToEmailAttachmentPrefetchItem(row) {
  return {
    attachmentKey: row.attachment_key,
    attachmentType: row.attachment_type,
    contentType: row.content_type,
    contentCiphertext: row.content_ciphertext || null,
    contentIv: row.content_iv || null,
    contentKeyId: row.content_key_id || "",
    contentSha256: row.content_sha256 || "",
    contentTag: row.content_tag || null,
    createdAt: numberValue(row.created_at),
    decodedSize: row.decoded_size,
    failureCode: row.failure_code,
    expectedContentSha256: row.expected_content_sha256 || "",
    filename: row.filename,
    haloAttachmentId: row.halo_attachment_id,
    haloFilename: row.halo_filename,
    haloFilesize: row.halo_filesize,
    haloType: row.halo_type,
    reportedSize: row.reported_size,
    preparedAt: nullableNumber(row.prepared_at),
    showForUsers: row.show_for_users,
    status: row.status,
    updatedAt: numberValue(row.updated_at),
  };
}

function rowToTicketCreationIntent(row) {
  return row
    ? {
        actionId: row.action_id,
        createdAt: numberValue(row.created_at),
        encryptedIntent: jsonValue(row.encrypted_intent_json, null),
        encryptedIntentJson: JSON.stringify(row.encrypted_intent_json),
        expiresAt: numberValue(row.expires_at),
        haloTenant: row.halo_tenant,
        lastError: row.last_error,
        operationId: row.operation_id,
        status: row.status,
        ticketId: row.ticket_id,
        ticketNumber: row.ticket_number,
        updatedAt: numberValue(row.updated_at),
        userId: row.user_id,
      }
    : null;
}

function jsonValue(value, fallback) {
  if (value === null || value === undefined) {
    return fallback;
  }
  if (typeof value !== "string") {
    return value;
  }
  try {
    return JSON.parse(value);
  } catch {
    return fallback;
  }
}

function objectValue(value) {
  const parsed = jsonValue(value, {});
  return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
}

function numberValue(value) {
  return Number(value);
}

function nullableNumber(value) {
  return value === null || value === undefined ? null : numberValue(value);
}

module.exports = { createHaloStore };
