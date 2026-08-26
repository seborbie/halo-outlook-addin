const assert = require("node:assert/strict");
const { EventEmitter } = require("node:events");
const test = require("node:test");

const { _test: serverTest } = require("./index");
const { registerStatusRoute } = require("./statusRoute");

function createHarness(checkReady) {
  const routes = new Map();
  const app = {
    get(path, handler) {
      routes.set(path, handler);
    },
  };
  registerStatusRoute(app, { checkReady });
  return routes;
}

function createResponse() {
  return {
    body: undefined,
    contentType: undefined,
    statusCode: undefined,
    json(value) {
      this.body = value;
      return this;
    },
    send(value) {
      this.body = value;
      return this;
    },
    status(value) {
      this.statusCode = value;
      return this;
    },
    type(value) {
      this.contentType = value;
      return this;
    },
  };
}

test("GET / reports that the add-in is up", () => {
  const route = createHarness(async () => true).get("/");
  const response = createResponse();
  route({}, response);
  assert.equal(response.statusCode, 200);
  assert.equal(response.contentType, "text/plain");
  assert.equal(response.body, "Halo Outlook add-in is up.");
});

test("GET /health/ready reflects PostgreSQL readiness", async () => {
  const readyResponse = createResponse();
  await createHarness(async () => true).get("/health/ready")({}, readyResponse);
  assert.equal(readyResponse.statusCode, 200);
  assert.deepEqual(readyResponse.body, { ready: true });

  const unavailableResponse = createResponse();
  await createHarness(async () => false).get("/health/ready")({}, unavailableResponse);
  assert.equal(unavailableResponse.statusCode, 503);
  assert.deepEqual(unavailableResponse.body, { ready: false });

  const errorResponse = createResponse();
  await createHarness(async () => {
    throw new Error("database unavailable");
  }).get("/health/ready")({}, errorResponse);
  assert.equal(errorResponse.statusCode, 503);
});

test("HTTP startup waits for listening and rejects bind errors", async () => {
  const listeningServer = new EventEmitter();
  const listeningApp = {
    listen(port) {
      assert.equal(port, 3100);
      queueMicrotask(() => listeningServer.emit("listening"));
      return listeningServer;
    },
  };
  assert.equal(await serverTest.listenHttpServer(listeningApp, 3100), listeningServer);

  const bindError = Object.assign(new Error("address already in use"), { code: "EADDRINUSE" });
  const failedServer = new EventEmitter();
  const failedApp = {
    listen() {
      queueMicrotask(() => failedServer.emit("error", bindError));
      return failedServer;
    },
  };
  await assert.rejects(serverTest.listenHttpServer(failedApp, 3100), bindError);
});
