const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function run() {
  const isBuiltRuntime = Boolean(process.argv[2]);
  const runtimePath = isBuiltRuntime
    ? path.resolve(__dirname, "..", process.argv[2])
    : path.join(__dirname, "..", "src", "commands", "classic-send-runtime.js");
  const commandsHtmlPath = isBuiltRuntime
    ? path.join(__dirname, "..", "dist", "commands.html")
    : path.join(__dirname, "..", "src", "commands", "commands.html");
  const runtime = fs.readFileSync(runtimePath, "utf8");
  const commandsHtml = fs.readFileSync(commandsHtmlPath, "utf8");
  const timers = [];
  let sendHandler = null;

  const officeReadyIndex = commandsHtml.indexOf("Office.onReady");
  const runtimeScriptIndex = commandsHtml.indexOf("/public/classic-send-runtime.js");
  assert(officeReadyIndex >= 0, "The HTML runtime must initialize Office.js.");
  assert(
    runtimeScriptIndex > officeReadyIndex,
    "The HTML runtime must initialize Office.js before loading the send handlers."
  );

  if (isBuiltRuntime) {
    assert.doesNotMatch(
      runtime,
      /\?/,
      "The classic Outlook runtime must not be minified into unsupported conditional expressions."
    );
  }

  const Office = {
    actions: {
      associate(name, handler) {
        assert.strictEqual(name, "onHaloMessageSend");
        sendHandler = handler;
      },
    },
    AsyncResultStatus: { Succeeded: "succeeded" },
    CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: { ItemType: { Message: "message" } },
    context: {
      mailbox: {
        item: {
          body: {
            // Simulate an Outlook API callback that never arrives.
            getAsync() {},
          },
          conversationId: "conversation-id",
          itemType: "message",
        },
        userProfile: {},
      },
      roamingSettings: {
        get() {
          return "";
        },
      },
    },
  };

  const context = vm.createContext({
    Array,
    Date,
    Error,
    Intl,
    JSON,
    Object,
    Office,
    String,
    XMLHttpRequest: function XMLHttpRequest() {},
    clearTimeout(id) {
      const timer = timers.find((entry) => entry.id === id);
      if (timer) {
        timer.cleared = true;
      }
    },
    setTimeout(callback, delay) {
      const timer = { callback, cleared: false, delay, id: timers.length + 1 };
      timers.push(timer);
      return timer.id;
    },
  });

  vm.runInContext(runtime, context, { filename: runtimePath });
  assert.strictEqual(typeof sendHandler, "function");

  const completions = [];
  sendHandler({
    completed(options) {
      completions.push(options);
    },
  });

  assert.deepStrictEqual(completions, []);
  const watchdog = timers.find((entry) => entry.delay === 3000);
  assert(watchdog, "The send handler must register a three-second fail-open watchdog.");

  watchdog.callback();
  assert.strictEqual(completions.length, 1);
  assert.strictEqual(completions[0].allowEvent, true);

  watchdog.callback();
  assert.strictEqual(completions.length, 1, "The send event must only be completed once.");

  console.log("Classic Outlook send runtime tests passed");
}

run();
