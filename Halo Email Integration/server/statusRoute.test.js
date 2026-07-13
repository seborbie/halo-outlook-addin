const assert = require("node:assert/strict");
const test = require("node:test");

const { registerStatusRoute } = require("./statusRoute");

test("GET / reports that the add-in is up", () => {
  let route;
  const app = {
    get(path, handler) {
      assert.equal(path, "/");
      route = handler;
    },
  };

  registerStatusRoute(app);

  const response = {
    statusCode: undefined,
    contentType: undefined,
    body: undefined,
    status(value) {
      this.statusCode = value;
      return this;
    },
    type(value) {
      this.contentType = value;
      return this;
    },
    send(value) {
      this.body = value;
      return this;
    },
  };

  route({}, response);

  assert.equal(response.statusCode, 200);
  assert.equal(response.contentType, "text/plain");
  assert.equal(response.body, "Halo Outlook add-in is up.");
});
