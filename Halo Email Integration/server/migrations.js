const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const MIGRATION_LOCK_ID = 194253117;
const DEFAULT_MIGRATIONS_DIRECTORY = path.join(__dirname, "migrations");

async function runMigrations(pool, options = {}) {
  const directory = options.directory || DEFAULT_MIGRATIONS_DIRECTORY;
  const migrations = fs
    .readdirSync(directory, { withFileTypes: true })
    .filter((entry) => entry.isFile() && /^\d+.*\.sql$/i.test(entry.name))
    .map((entry) => entry.name)
    .sort((left, right) => left.localeCompare(right));
  if (!migrations.length) {
    throw new Error(`No PostgreSQL migrations were found in ${directory}.`);
  }

  const client = await pool.connect();
  let lockAcquired = false;
  let operationError = null;
  try {
    await client.query("SELECT pg_advisory_lock($1)", [MIGRATION_LOCK_ID]);
    lockAcquired = true;
    await client.query(`
      CREATE TABLE IF NOT EXISTS schema_migrations (
        name TEXT PRIMARY KEY,
        checksum TEXT,
        applied_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      )
    `);
    await client.query("ALTER TABLE schema_migrations ADD COLUMN IF NOT EXISTS checksum TEXT");
    const appliedResult = await client.query("SELECT name, checksum FROM schema_migrations");
    const applied = new Map(appliedResult.rows.map((row) => [row.name, row.checksum]));
    const unknownMigrations = appliedResult.rows
      .map((row) => row.name)
      .filter((name) => !migrations.includes(name));
    if (unknownMigrations.length) {
      throw new Error(
        `The database contains migrations unknown to this release: ${unknownMigrations.join(", ")}.`
      );
    }

    for (const name of migrations) {
      const sql = fs.readFileSync(path.join(directory, name), "utf8");
      const checksum = crypto.createHash("sha256").update(sql).digest("hex");
      if (applied.has(name)) {
        const appliedChecksum = applied.get(name);
        if (appliedChecksum && appliedChecksum !== checksum) {
          throw new Error(`PostgreSQL migration ${name} has changed since it was applied.`);
        }
        if (!appliedChecksum) {
          await client.query(
            "UPDATE schema_migrations SET checksum = $1 WHERE name = $2 AND checksum IS NULL",
            [checksum, name]
          );
        }
        continue;
      }
      await client.query("BEGIN");
      try {
        await client.query(sql);
        await client.query(
          "INSERT INTO schema_migrations (name, checksum) VALUES ($1, $2)",
          [name, checksum]
        );
        await client.query("COMMIT");
      } catch (error) {
        try {
          await client.query("ROLLBACK");
        } catch (rollbackError) {
          error.rollbackError = rollbackError;
        }
        throw new Error(`PostgreSQL migration ${name} failed: ${error.message}`, { cause: error });
      }
    }
    await client.query("ALTER TABLE schema_migrations ALTER COLUMN checksum SET NOT NULL");
  } catch (error) {
    operationError = error;
    throw error;
  } finally {
    let unlockError = null;
    try {
      if (lockAcquired) {
        await client.query("SELECT pg_advisory_unlock($1)", [MIGRATION_LOCK_ID]);
      }
    } catch (error) {
      unlockError = error;
    } finally {
      client.release(unlockError || undefined);
    }
    if (unlockError) {
      if (operationError) {
        console.error("PostgreSQL migration advisory lock could not be released.", unlockError);
      } else {
        throw unlockError;
      }
    }
  }
}

module.exports = { runMigrations };
