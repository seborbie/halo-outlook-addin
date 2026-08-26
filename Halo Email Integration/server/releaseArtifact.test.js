const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const projectRoot = path.resolve(__dirname, "..");
const packageInfo = require(path.join(projectRoot, "package.json"));

function readDistFile(name) {
  const filePath = path.join(projectRoot, "dist", name);
  assert(fs.existsSync(filePath), `Missing production artifact: dist/${name}`);
  return fs.readFileSync(filePath, "utf8");
}

function run() {
  const manifest = readDistFile("manifest.xml");
  const runtime = readDistFile("classic-send-runtime.js");
  const taskpane = readDistFile("taskpane.html");
  const commands = readDistFile("commands.html");

  assert.strictEqual(packageInfo.version, "2026.8.24");
  assert.strictEqual(
    fs.existsSync(path.join(projectRoot, "dist", "manifest.debug.xml")),
    false,
    "Production builds must not ship the localhost diagnostics manifest."
  );
  assert.match(manifest, /<Id>55bbcff2-8191-4411-aec6-f9d2f9b4b5e8<\/Id>/);
  assert.match(manifest, /<Version>2026\.8\.24\.0<\/Version>/);
  assert.match(manifest, /<Permissions>ReadWriteMailbox<\/Permissions>/);
  assert.match(manifest, /classic-send-runtime\.js\?v=2026\.8\.24\.0/);
  assert.doesNotMatch(manifest, /localhost|your-addin-host|__HALO_|DEV -|HaloPSA DEV|DEV:/i);
  assert.doesNotMatch(runtime, /__HALO_PUBLIC_BASE_URL__/);
  assert.match(runtime, /https:\/\/[^"']+\/api\/halo\/tickets\//);
  assert.match(commands, /classic-send-runtime\.js\?v=2026\.8\.24\.0/);
  assert.doesNotMatch(commands, /diagnostics-2/);
  assert.match(taskpane, /v2026\.8\.24/);

  const manifestUrls = Array.from(
    manifest.matchAll(/DefaultValue="(https:\/\/[^\"]+)"/g),
    (match) => new URL(match[1]).origin
  );
  assert(manifestUrls.length > 0, "The production manifest must contain HTTPS asset URLs.");
  assert.strictEqual(
    new Set(manifestUrls).size,
    1,
    "Every production manifest URL must use the configured public origin."
  );
  const publicOrigin = manifestUrls[0];
  for (const apiPath of [
    "/api/diagnostics/send-event",
    "/api/halo/email/send-auto-attach",
    "/api/halo/email/recover-mapping",
    "/api/halo/email-attachments/prefetch/start",
    "/api/halo/email-attachments/prefetch/",
    "/api/halo/ticket-creation/intents/",
    "/api/halo/tickets/",
  ]) {
    assert(
      runtime.includes(`${publicOrigin}${apiPath}`),
      `The production event runtime must use the manifest origin for ${apiPath}.`
    );
  }

  console.log("Production release artifact tests passed");
}

run();
