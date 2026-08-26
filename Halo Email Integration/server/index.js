const path = require("path");

require("dotenv").config({ path: path.join(__dirname, "..", ".env") });

const express = require("express");
const { registerHaloAuthRoutes } = require("./haloAuth");
const { createHaloStore } = require("./haloStore");
const { registerStatusRoute } = require("./statusRoute");

const port = Number(process.env.PORT || 3000);
const distPath = path.join(__dirname, "..", "dist");

async function startServer() {
  const store = createHaloStore();
  let app;
  let server;
  try {
    await store.initialize();
    app = createApp(store);
    server = await listenHttpServer(app, port);
    console.log(`Halo Outlook add-in server listening on port ${port}`);
  } catch (error) {
    try {
      await store.close();
    } catch (closeError) {
      error.closeError = closeError;
    }
    throw error;
  }

  let shuttingDown = false;
  const shutdown = async (signal) => {
    if (shuttingDown) {
      return;
    }
    shuttingDown = true;
    console.log(`Received ${signal}; shutting down.`);
    try {
      await closeHttpServer(server);
    } finally {
      await store.close();
    }
  };
  for (const signal of ["SIGTERM", "SIGINT"]) {
    process.once(signal, () => {
      void shutdown(signal)
        .then(() => process.exit(0))
        .catch((error) => {
          console.error("Graceful shutdown failed.", error);
          process.exit(1);
        });
    });
  }
  return { app, server, store };
}

function createApp(store) {
  const app = express();
  app.disable("x-powered-by");
  app.use((req, res, next) => {
    if (
      req.path === "/commands.html" ||
      req.path === "/classic-send-runtime.js" ||
      req.path === "/public/classic-send-runtime.js"
    ) {
      res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate");
      res.setHeader("Expires", "0");
      res.setHeader("Pragma", "no-cache");
    }
    next();
  });
  app.use("/api/halo/email-attachments/prefetch", express.json({ limit: "36mb" }));
  app.use(express.json({ limit: "10mb" }));

  registerHaloAuthRoutes(app, { store });
  registerStatusRoute(app, { checkReady: () => store.isReady() });

  app.get("/bugreport", (_req, res) => {
    res.sendFile(path.join(distPath, "bugreport.html"));
  });
  app.use(express.static(distPath));
  app.use("/public", express.static(distPath));
  app.get(/.*/, (_req, res) => {
    res.sendFile(path.join(distPath, "taskpane.html"));
  });
  return app;
}

function listenHttpServer(app, listenPort) {
  return new Promise((resolve, reject) => {
    const server = app.listen(listenPort);
    const onError = (error) => {
      server.off("listening", onListening);
      reject(error);
    };
    const onListening = () => {
      server.off("error", onError);
      resolve(server);
    };
    server.once("error", onError);
    server.once("listening", onListening);
  });
}

function closeHttpServer(server) {
  return new Promise((resolve, reject) => {
    server.close((error) => (error ? reject(error) : resolve()));
  });
}

if (require.main === module) {
  startServer().catch((error) => {
    console.error("Halo Outlook add-in server failed to start.", error);
    process.exitCode = 1;
  });
}

module.exports = { _test: { listenHttpServer }, startServer };
