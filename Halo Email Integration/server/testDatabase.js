const crypto = require("node:crypto");
const { createDatabasePool } = require("./database");
const { createHaloStore } = require("./haloStore");

const LOCAL_TEST_DATABASE_URL = "postgresql://127.0.0.1:5432/haloaddin";
const LOCAL_TEST_DATABASE_USERNAME = "haloaddin";
const LOCAL_TEST_DATABASE_PASSWORD = "haloaddin_local";

async function createTestDatabase(options = {}) {
  const baseConnectionString = String(
    options.connectionString || process.env.TEST_DATABASE_URL || LOCAL_TEST_DATABASE_URL
  );
  const username = String(
    options.username || process.env.TEST_DATABASE_USERNAME || LOCAL_TEST_DATABASE_USERNAME
  );
  const password = String(
    options.password || process.env.TEST_DATABASE_PASSWORD || LOCAL_TEST_DATABASE_PASSWORD
  );
  const schemaName = `test_${process.pid}_${crypto.randomBytes(8).toString("hex")}`;
  const adminPool = createDatabasePool({
    authMode: "password",
    connectionString: baseConnectionString,
    env: { ...process.env, DATABASE_POOL_MAX: "1" },
    password,
    username,
  });
  try {
    await adminPool.query(`CREATE SCHEMA ${quoteIdentifier(schemaName)}`);
  } finally {
    await adminPool.end();
  }

  const connectionString = withSearchPath(baseConnectionString, schemaName);
  let closePromise = null;
  let removed = false;
  return {
    connectionString,
    createPool() {
      if (removed || closePromise) {
        throw new Error("The test database schema has already been removed.");
      }
      return createDatabasePool({
        authMode: "password",
        connectionString,
        password,
        username,
      });
    },
    async close() {
      if (removed) {
        return;
      }
      if (closePromise) {
        return closePromise;
      }
      closePromise = removeSchema();
      try {
        await closePromise;
        removed = true;
      } catch (error) {
        closePromise = null;
        throw error;
      }
    },
    async createStore() {
      if (removed || closePromise) {
        throw new Error("The test database schema has already been removed.");
      }
      const store = createHaloStore({
        authMode: "password",
        connectionString,
        password,
        username,
      });
      try {
        await store.initialize();
        return store;
      } catch (error) {
        await store.close();
        throw error;
      }
    },
    schemaName,
  };

  async function removeSchema() {
    const cleanupPool = createDatabasePool({
      authMode: "password",
      connectionString: baseConnectionString,
      env: { ...process.env, DATABASE_POOL_MAX: "1" },
      password,
      username,
    });
    try {
      await cleanupPool.query(`DROP SCHEMA ${quoteIdentifier(schemaName)} CASCADE`);
    } finally {
      await cleanupPool.end();
    }
  }
}

async function createTestStore(options = {}) {
  const database = await createTestDatabase(options);
  let store;
  try {
    store = await database.createStore();
  } catch (error) {
    try {
      await database.close();
    } catch (cleanupError) {
      error.cleanupError = cleanupError;
    }
    throw error;
  }
  const closeStore = store.close.bind(store);
  let closePromise = null;
  store.close = async () => {
    if (!closePromise) {
      closePromise = closeStoreAndDatabase(closeStore, database);
    }
    return closePromise;
  };
  return store;
}

async function closeStoreAndDatabase(closeStore, database) {
  let operationError = null;
  try {
    await closeStore();
  } catch (error) {
    operationError = error;
  }
  try {
    await database.close();
  } catch (error) {
    if (operationError) {
      operationError.cleanupError = error;
    } else {
      operationError = error;
    }
  }
  if (operationError) {
    throw operationError;
  }
}

function withSearchPath(connectionString, schemaName) {
  const url = new URL(connectionString);
  url.searchParams.set("options", `-c search_path=${schemaName}`);
  return url.toString();
}

function quoteIdentifier(value) {
  return `"${String(value).replace(/"/g, '""')}"`;
}

module.exports = { createTestDatabase, createTestStore };
