const path = require("path");
require("dotenv").config({ path: path.join(__dirname, "..", ".env") });

const { Pool } = require("pg");
const { getDatabaseConfig } = require("./haloStore");
const { runMigrations } = require("./migrations");

async function main() {
  const pool = new Pool(getDatabaseConfig(process.env));
  try {
    await runMigrations(pool);
    console.log("PostgreSQL migrations are up to date.");
  } finally {
    await pool.end();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
