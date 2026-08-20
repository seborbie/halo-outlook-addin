const assert = require("node:assert/strict");
const test = require("node:test");

const { registerStatusRoute } = require("./statusRoute");

test("GET /api/health reports that the service is up", () => {
  let route;
  const app = {
    get(path, handler) {
      assert.equal(path, "/api/health");
      route = handler;
    },
  };

  registerStatusRoute(app);

  const response = {
    statusCode: undefined,
    body: undefined,
    status(value) {
      this.statusCode = value;
      return this;
    },
    json(value) {
      this.body = value;
      return this;
    },
  };

  route({}, response);

  assert.equal(response.statusCode, 200);
  assert.deepEqual(response.body, { ok: true, service: "inboxlink" });
});
