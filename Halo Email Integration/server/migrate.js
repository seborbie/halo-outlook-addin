const path = require("path");

require("dotenv").config({ path: path.join(__dirname, "..", ".env") });

const { createDatabasePool } = require("./database");
const { runMigrations } = require("./migrations");

async function main() {
  const pool = createDatabasePool();
  try {
    await runMigrations(pool);
    console.log("PostgreSQL migrations completed.");
  } finally {
    await pool.end();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
