const crypto = require("crypto");
const path = require("path");
const { DatabaseSync } = require("node:sqlite");
require("dotenv").config({ path: path.join(__dirname, "..", ".env") });

const { createHaloStore } = require("./haloStore");

function getArgument(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : "";
}

function required(value, name) {
  const normalized = String(value || "").trim();
  if (!normalized) throw new Error(`${name} must be set for the SQLite import.`);
  return normalized;
}

async function main() {
  const sqlitePath = path.resolve(required(getArgument("--sqlite"), "--sqlite"));
  const haloUrl = httpsOrigin(required(process.env.HALO_URL, "HALO_URL"), "HALO_URL");
  const haloClientId = required(process.env.HALO_CLIENT_ID, "HALO_CLIENT_ID");
  const defaultTenantId = String(process.env.LEGACY_DEFAULT_MICROSOFT_TENANT_ID || "").trim();
  const configuredName = String(process.env.LEGACY_ORGANISATION_NAME || "").trim();
  const sqlite = new DatabaseSync(sqlitePath, { readOnly: true });
  const store = createHaloStore();

  try {
    await store.ready;
    const users = sqlite.prepare("SELECT * FROM users ORDER BY id").all();
    if (!users.length) {
      console.log("The legacy SQLite database contains no users; nothing was imported.");
      return;
    }

    const usersByTenant = Map.groupBy(users, (user) => String(user.tenant_id));
    const organisationByTenant = new Map();
    const userIdMap = new Map();

    for (const [tenantId, tenantUsers] of usersByTenant) {
      const firstUser = tenantUsers[0];
      const registration = await store.registerOrganisation({
        companyName:
          configuredName && usersByTenant.size === 1
            ? configuredName
            : `Imported organisation ${tenantId.slice(0, 8)}`,
        haloClientId,
        haloUrl,
        microsoftTenantId: tenantId,
        owner: {
          displayName: firstUser.display_name || "Imported owner",
          email: firstUser.email || "",
          objectId: firstUser.object_id,
        },
      });
      organisationByTenant.set(tenantId, registration.organisation);
      userIdMap.set(String(firstUser.id), registration.user);

      for (const legacyUser of tenantUsers.slice(1)) {
        const user = await store.upsertUser({
          displayName: legacyUser.display_name || "",
          email: legacyUser.email || "",
          objectId: legacyUser.object_id,
          tenantId,
        });
        userIdMap.set(String(legacyUser.id), user);
      }
    }

    const grants = sqlite.prepare("SELECT * FROM halo_grants ORDER BY id").all();
    for (const grant of grants) {
      if (grant.invalidated_at != null) continue;
      const user = userIdMap.get(String(grant.user_id));
      if (!user) throw new Error(`Legacy grant ${grant.id} references a missing user.`);
      await store.saveHaloGrant({
        clientId: grant.client_id,
        encryptedToken: JSON.parse(grant.encrypted_token_json),
        haloUrl: httpsOrigin(grant.halo_url, `legacy grant ${grant.id} Halo URL`),
        organisationId: user.organisationId,
        scope: grant.scope,
        userId: user.id,
      });
    }

    const mappings = sqlite.prepare("SELECT * FROM conversation_mappings ORDER BY created_at").all();
    let mappingOrganisation = null;
    if (mappings.length) {
      if (usersByTenant.size === 1) {
        mappingOrganisation = organisationByTenant.values().next().value;
      } else {
        mappingOrganisation = organisationByTenant.get(defaultTenantId);
        if (!mappingOrganisation) {
          throw new Error(
            "LEGACY_DEFAULT_MICROSOFT_TENANT_ID must identify the owner of legacy email mappings when SQLite contains multiple tenants."
          );
        }
      }
    }

    const mappingIdMap = new Map();
    for (const mapping of mappings) {
      const id = crypto.randomUUID();
      mappingIdMap.set(mapping.id, id);
      await store.saveConversationMapping(mappingOrganisation.id, {
        conversationId: mapping.conversation_id || "",
        createdAt: Number(mapping.created_at),
        id,
        mailboxEmail: mapping.mailbox_email,
        normalizedSubject: mapping.normalized_subject || "",
        syncedMessageIds: new Set(),
        ticketId: Number(mapping.ticket_id),
        ticketNumber: mapping.ticket_number,
        updatedAt: Number(mapping.updated_at),
      });
    }

    const messageMappings = sqlite.prepare("SELECT * FROM message_mappings ORDER BY created_at").all();
    for (const mapping of messageMappings) {
      const mappingId = mappingIdMap.get(mapping.mapping_id);
      if (!mappingId) throw new Error(`Legacy message mapping references ${mapping.mapping_id}.`);
      await store.saveMessageMapping(mappingOrganisation.id, {
        mailboxEmail: mapping.mailbox_email,
        mappingId,
        messageIdKey: mapping.message_id_key,
      });
    }

    const skippedSessions = countRows(sqlite, "sessions") + countRows(sqlite, "background_sessions");
    const skippedReports = countRows(sqlite, "bug_report_sessions");
    console.log(
      `Imported ${users.length} users, ${grants.length} grants, ${mappings.length} conversations, and ${messageMappings.length} message mappings.`
    );
    console.log(
      `Skipped ${skippedSessions} unrecoverable active session handles and ${skippedReports} temporary bug-report sessions.`
    );
  } finally {
    sqlite.close();
    await store.close();
  }
}

function countRows(sqlite, table) {
  try {
    return Number(sqlite.prepare(`SELECT COUNT(*) AS count FROM ${table}`).get().count);
  } catch {
    return 0;
  }
}

function httpsOrigin(value, name) {
  const url = new URL(value);
  if (url.protocol !== "https:") throw new Error(`${name} must use https://.`);
  return url.origin;
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
