const { DefaultAzureCredential } = require("@azure/identity");
const { Pool } = require("pg");

const DEFAULT_POOL_MAX = 5;
const DEFAULT_CONNECTION_TIMEOUT_MS = 10000;
const DEFAULT_IDLE_TIMEOUT_MS = 30000;
const POSTGRES_ENTRA_SCOPE = "https://ossrdbms-aad.database.windows.net/.default";
const PASSWORD_AUTH_MODES = new Set(["", "password", "psk", "usernamepassword"]);

function createDatabasePool(options = {}) {
  const env = options.env || process.env;
  const connectionString = String(options.connectionString || env.DATABASE_URL || "").trim();
  if (!connectionString) {
    throw new Error("DATABASE_URL must be set.");
  }

  const parsedUrl = parseDatabaseUrl(connectionString);
  if (parsedUrl.username || parsedUrl.password) {
    throw new Error(
      "DATABASE_URL must not include credentials; use DATABASE_USERNAME and DATABASE_PASSWORD."
    );
  }
  const hostname = parsedUrl.hostname.toLowerCase();
  const local = hostname === "localhost" || hostname === "127.0.0.1" || hostname === "::1";
  const sslEnabled = getBooleanSetting(env.DATABASE_SSL, !local);
  const rejectUnauthorized = getBooleanSetting(env.DATABASE_SSL_REJECT_UNAUTHORIZED, true);
  const authMode = normalizeDatabaseAuthMode(options.authMode ?? env.DATABASE_AUTH);
  const username = String(options.username ?? env.DATABASE_USERNAME ?? "").trim();
  if (!username) {
    throw new Error("DATABASE_USERNAME must be set.");
  }
  for (const parameter of ["ssl", "sslmode", "sslcert", "sslkey", "sslrootcert"]) {
    parsedUrl.searchParams.delete(parameter);
  }

  const poolOptions = {
    database: decodeURIComponent(parsedUrl.pathname.replace(/^\//, "")),
    host: parsedUrl.hostname,
    port: parsedUrl.port ? Number(parsedUrl.port) : 5432,
    connectionTimeoutMillis: getPositiveInteger(
      env.DATABASE_CONNECTION_TIMEOUT_MS,
      DEFAULT_CONNECTION_TIMEOUT_MS
    ),
    idleTimeoutMillis: getPositiveInteger(env.DATABASE_IDLE_TIMEOUT_MS, DEFAULT_IDLE_TIMEOUT_MS),
    max: getPositiveInteger(env.DATABASE_POOL_MAX, DEFAULT_POOL_MAX),
    user: username,
    ssl: sslEnabled ? { rejectUnauthorized } : false,
  };
  const connectionOptions = parsedUrl.searchParams.get("options");
  if (connectionOptions) {
    poolOptions.options = connectionOptions;
  }

  if (authMode === "entra") {
    poolOptions.password = createEntraPasswordProvider(
      options.tokenCredential || new DefaultAzureCredential()
    );
  } else {
    const password = String(options.password ?? env.DATABASE_PASSWORD ?? "");
    if (!password) {
      throw new Error("DATABASE_PASSWORD must be set for password authentication.");
    }
    poolOptions.password = password;
  }

  const pool = new Pool(poolOptions);

  pool.on("error", (error) => {
    console.error("Unexpected PostgreSQL pool error.", error);
  });

  return pool;
}

function normalizeDatabaseAuthMode(value) {
  const normalized = String(value ?? "")
    .trim()
    .toLowerCase();
  if (PASSWORD_AUTH_MODES.has(normalized)) {
    return "password";
  }
  if (normalized === "entra") {
    return normalized;
  }
  throw new Error(
    `DATABASE_AUTH must be entra, password, psk, usernamepassword, or unset; received ${JSON.stringify(
      value
    )}.`
  );
}

function createEntraPasswordProvider(credential) {
  return async () => {
    const accessToken = await credential.getToken(POSTGRES_ENTRA_SCOPE);
    if (!accessToken || !accessToken.token) {
      throw new Error("Microsoft Entra did not return a PostgreSQL access token.");
    }
    return accessToken.token;
  };
}

async function checkDatabaseReady(pool) {
  await pool.query("SELECT 1");
  return true;
}

function parseDatabaseUrl(connectionString) {
  try {
    const url = new URL(connectionString);
    if (!["postgres:", "postgresql:"].includes(url.protocol) || !url.hostname) {
      throw new Error("Unsupported PostgreSQL URL.");
    }
    return url;
  } catch {
    throw new Error("DATABASE_URL must be a valid PostgreSQL connection URL.");
  }
}

function getBooleanSetting(value, fallback) {
  if (value === undefined || value === null || value === "") {
    return fallback;
  }
  return !["0", "false", "no", "off"].includes(String(value).trim().toLowerCase());
}

function getPositiveInteger(value, fallback) {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
}

module.exports = {
  checkDatabaseReady,
  createDatabasePool,
};
