const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const { _test } = require("./haloAuth");

const projectRoot = path.resolve(__dirname, "..");

function request(url) {
  return { headers: { host: "localhost:3000" }, originalUrl: url, url };
}

function run() {
  assert.strictEqual(
    _test.buildTicketsPath({ ownership: "mine", lifecycle: "open" }),
    "/api/Tickets?count=50&open_only=true&mine=true&includeagent=true&includestatus=true"
  );
  assert.strictEqual(
    _test.buildTicketsPath({
      query: "printer [warehouse]",
      ownership: "all",
      lifecycle: "all",
    }),
    "/api/Tickets?count=50&search=printer+%5Bwarehouse%5D&includeagent=true&includestatus=true"
  );

  assert.strictEqual(
    _test.getTicketSearchQuery(
      request("/api/halo/tickets/search?query=printer%20%5Bwarehouse%5D")
    ),
    "printer [warehouse]"
  );
  assert.strictEqual(
    _test.getTicketSearchQuery(
      request("/api/halo/tickets/search?ticketNumber=%5BID%3A%20002200%5D")
    ),
    "002200"
  );

  const tickets = _test.normalizeTicketsForLifecycle(
    {
      tickets: [
        {
          id: 1001,
          ticketnumber: "T1001",
          summary: "Incomplete configuration",
          status: "Incomplete",
          client: { name: "Nested Customer" },
          assigned_agent: { display_name: "Nested Agent" },
        },
        {
          id: 1002,
          ticketnumber: "T1002",
          summary: "Closed by explicit flag",
          status: "In Progress",
          closed: "1",
        },
      ],
    },
    "all"
  );

  assert.deepStrictEqual(tickets, [
    {
      id: "1001",
      ticketNumber: "T1001",
      summary: "Incomplete configuration",
      status: "Incomplete",
      lifecycle: "open",
      client: "Nested Customer",
      agent: "Nested Agent",
    },
    {
      id: "1002",
      ticketNumber: "T1002",
      summary: "Closed by explicit flag",
      status: "In Progress",
      lifecycle: "closed",
      client: "",
      agent: "",
    },
  ]);

  const promoted = _test.promoteExactTicketMatches(tickets, "1002");
  assert.deepStrictEqual(
    promoted.map((ticket) => ticket.id),
    ["1002", "1001"]
  );
  assert.strictEqual(promoted.length, tickets.length);

  const taskpaneSource = fs.readFileSync(
    path.join(projectRoot, "src", "taskpane", "outlook.ts"),
    "utf8"
  );
  const actionBusy = getFunctionBlock(taskpaneSource, "setTicketsBusy");
  assert.match(actionBusy, /cancelTicketSearchDebounce\(\)/);
  assert.match(actionBusy, /ticketRequestRevision \+= 1/);
  assert.match(actionBusy, /ticketQueryIsBusy = false/);
  assert.match(actionBusy, /updateTicketBusyState\(\)/);

  const applyFilters = getFunctionBlock(taskpaneSource, "applyTicketFilters");
  assert.match(
    applyFilters,
    /renderTicketList\([\s\S]*updateTicketBusyState\(\)/,
    "Rerendered ticket buttons must retain the active query/action busy state."
  );

  console.log("Ticket search tests passed");
}

function getFunctionBlock(source, functionName) {
  const start = source.indexOf(`function ${functionName}`);
  assert(start >= 0, `Missing ${functionName}.`);
  const nextFunction = source.indexOf("\nfunction ", start + 1);
  return source.slice(start, nextFunction < 0 ? source.length : nextFunction);
}

run();
