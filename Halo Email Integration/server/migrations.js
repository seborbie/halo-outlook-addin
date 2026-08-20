const fs = require("fs");
const path = require("path");

const MIGRATIONS_PATH = path.join(__dirname, "migrations");
const MIGRATION_LOCK_KEY = 714_265_991;

async function runMigrations(pool) {
  const client = await pool.connect();
  try {
    await client.query("SELECT pg_advisory_lock($1)", [MIGRATION_LOCK_KEY]);
    await client.query(`
      CREATE TABLE IF NOT EXISTS schema_migrations (
        name TEXT PRIMARY KEY,
        applied_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      )
    `);

    const applied = new Set(
      (await client.query("SELECT name FROM schema_migrations")).rows.map((row) => row.name)
    );
    const migrationFiles = fs
      .readdirSync(MIGRATIONS_PATH)
      .filter((file) => /^\d+.*\.sql$/.test(file))
      .sort();

    for (const file of migrationFiles) {
      if (applied.has(file)) continue;
      await client.query("BEGIN");
      try {
        await client.query(fs.readFileSync(path.join(MIGRATIONS_PATH, file), "utf8"));
        await client.query("INSERT INTO schema_migrations (name) VALUES ($1)", [file]);
        await client.query("COMMIT");
      } catch (error) {
        await client.query("ROLLBACK");
        throw error;
      }
    }
  } finally {
    await client.query("SELECT pg_advisory_unlock($1)", [MIGRATION_LOCK_KEY]).catch(() => {});
    client.release();
  }
}

module.exports = { runMigrations };
