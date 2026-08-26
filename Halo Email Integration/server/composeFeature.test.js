const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const projectRoot = path.resolve(__dirname, "..");

function read(relativePath) {
  return fs.readFileSync(path.join(projectRoot, relativePath), "utf8");
}

function run() {
  const manifest = read("manifest.xml");
  const packageInfo = JSON.parse(read("package.json"));
  const commandsHtml = read(path.join("src", "commands", "commands.html"));
  const taskpaneHtml = read(path.join("src", "taskpane", "taskpane.html"));
  const taskpaneSource = read(path.join("src", "taskpane", "outlook.ts"));
  const webpackSource = read("webpack.config.js");
  const webpackTest = require(path.join(projectRoot, "webpack.config.js"))._test;

  assert.match(manifest, /<Set Name="Mailbox" MinVersion="1\.12"\/>/);
  assert.match(manifest, /<Permissions>ReadWriteMailbox<\/Permissions>/);
  assert.match(manifest, /<Id>55bbcff2-8191-4411-aec6-f9d2f9b4b5e8<\/Id>/);
  assert.match(manifest, /<Version>2026\.8\.24\.0<\/Version>/);
  assert.match(manifest, /commands\.html\?v=2026\.8\.24\.0/);
  assert.match(manifest, /classic-send-runtime\.js\?v=2026\.8\.24\.0/);
  assert.doesNotMatch(manifest, /DEV - HaloPSA|HaloPSA DEV|DEV: Attach/);
  assert.match(manifest, /ExtensionPoint xsi:type="MessageComposeCommandSurface"/);
  assert.match(
    manifest,
    /<LaunchEvent Type="OnMessageSend" FunctionName="onHaloMessageSend" SendMode="PromptUser"\/>/
  );
  assert.match(
    manifest,
    /<LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onHaloMessageAttachmentsChanged"\/>/
  );
  assert.doesNotMatch(manifest, /SupportsSharedFolders|MobileFormFactor/);
  assert.match(commandsHtml, /__HALO_RUNTIME_CACHE_TOKEN__/);

  const diagnosticsManifest = webpackTest.createDevelopmentManifest(manifest);
  assert.match(
    diagnosticsManifest,
    new RegExp(`<Id>${escapeRegExp(webpackTest.developmentManifestId)}</Id>`)
  );
  assert.match(
    diagnosticsManifest,
    /<DisplayName DefaultValue="LOCAL DIAGNOSTICS - HaloPSA Outlook Add-in"\/>/
  );
  assert.match(diagnosticsManifest, /https:\/\/localhost:3000\/public\/classic-send-runtime\.js/);
  assert.match(
    diagnosticsManifest,
    new RegExp(`\\?v=${escapeRegExp(webpackTest.developmentRuntimeCacheToken)}`)
  );
  assert.doesNotMatch(diagnosticsManifest, /55bbcff2-8191-4411-aec6-f9d2f9b4b5e8/);
  assert.strictEqual(
    packageInfo.scripts.start,
    "office-addin-debugging start dist/manifest.debug.xml"
  );

  assert.match(
    webpackSource,
    /from: "src\/commands\/classic-send-runtime\.js"/,
    "The classic-compatible script must remain the sole copied event runtime."
  );
  assert.strictEqual(
    fs.existsSync(path.join(projectRoot, "src", "commands", "commands.outlook.ts")),
    false,
    "Do not restore the unreachable duplicate TypeScript event runtime."
  );

  const referencedIds = Array.from(
    taskpaneSource.matchAll(/document\.getElementById\("([^"]+)"\)/g),
    (match) => match[1]
  );
  const missingIds = Array.from(new Set(referencedIds)).filter(
    (id) => !new RegExp(`id=["']${escapeRegExp(id)}["']`).test(taskpaneHtml)
  );
  assert.deepStrictEqual(
    missingIds,
    [],
    `Task pane source references missing HTML IDs: ${missingIds}`
  );
  assert.match(taskpaneHtml, /id="create-ticket-tab"/);
  assert.match(taskpaneHtml, /id="create-ticket-form"/);
  assert.match(taskpaneHtml, /id="private-note-toggle"/);
  assert.match(taskpaneHtml, /id="action-mode-visibility"/);
  assert.match(taskpaneHtml, /role="switch"/);
  assert.doesNotMatch(taskpaneHtml, /Email visibility|action-mode-heading/);
  assert.match(
    taskpaneHtml,
    /id="tickets-heading"[\s\S]*id="ticket-action-mode-slot"[\s\S]*id="action-mode-control"/
  );
  assert.match(taskpaneSource, /placeActionModeControl\(destination\)/);
  assert.match(taskpaneSource, /"Customer hidden" : "Customer visible"/);
  assert.match(taskpaneSource, /destinationKind: "create-ticket"/);
  assert.match(taskpaneSource, /\/api\/halo\/ticket-creation\/from-email/);

  const createFieldControl = getFunctionBlock(taskpaneSource, "createTicketCreationFieldControl");
  assert.match(createFieldControl, /field\.required && field\.type !== "boolean"/);
  assert.match(createFieldControl, /empty\.disabled = field\.required/);
  const renderCreationSchema = getFunctionBlock(taskpaneSource, "renderTicketCreationSchema");
  assert.match(renderCreationSchema, /field\.required \|\| field\.recommended/);
  assert.doesNotMatch(renderCreationSchema, /field\.required \|\| field\.core/);

  const composeCreation = getFunctionBlock(taskpaneSource, "selectComposeTicketCreation");
  assert(
    composeCreation.indexOf("saveComposeItem(item)") <
      composeCreation.indexOf("collectEmailAttachmentMetadata(item, bodyHtml)"),
    "Create-on-send drafts must be saved before Outlook attachment IDs are collected."
  );
  assert.match(composeCreation, /saveComposeSessionMarker\(item, previousMarker\)/);
  assert.match(composeCreation, /clearComposeAttachMetadata\(item\)/);
  assert.match(composeCreation, /removeServerTicketCreationIntent\(operationId\)/);

  const composeSelection = getFunctionBlock(taskpaneSource, "selectComposeAttachTicket");
  assert.match(composeSelection, /saveComposeItem\(item\)/);
  assert(
    composeSelection.indexOf("saveComposeItem(item)") <
      composeSelection.indexOf("collectEmailAttachmentMetadata(item, bodyHtml)"),
    "A new compose item must be saved before Outlook attachment IDs are collected."
  );
  assert.match(composeSelection, /saveComposeSessionMarker\(item, marker\)/);
  assert.match(composeSelection, /saveComposeCustomMarker\(item, marker\)/);
  assert.match(composeSelection, /setComposeIdentityHeader\(item, marker\.composeAttachId\)/);
  assert.match(composeSelection, /actionMode: activeActionMode/);
  assert.match(composeSelection, /Will attach to \$\{marker\.ticketNumber\} when sent/);

  const restoreSelection = getFunctionBlock(taskpaneSource, "restoreComposeAttachSelection");
  assert.match(restoreSelection, /marker\.draftItemId !== currentItemId/);
  assert.match(restoreSelection, /clearComposeAttachMetadata\(item\)/);
  assert.match(
    restoreSelection,
    /clearComposeAttachMetadata\(item\)[\s\S]*saveComposeItem\(item\)/
  );
  assert.match(restoreSelection, /applyComposeMarkerRuntimeContext\(marker\)/);

  const saveSessionMarker = getFunctionBlock(taskpaneSource, "saveComposeSessionMarker");
  assert.match(saveSessionMarker, /applyComposeMarkerRuntimeContext\(marker\)/);
  const serializeCustomMarker = getFunctionBlock(taskpaneSource, "serializeComposeCustomMarker");
  assert.match(serializeCustomMarker, /delete persistentMarker\.backgroundSessionId/);

  const removeSelection = getFunctionBlock(taskpaneSource, "removeComposeAttachSelection");
  assert.match(removeSelection, /clearComposeAttachMetadata\(item\)/);
  assert.match(removeSelection, /setComposeIdentityHeader\(item, ""\)/);

  const reconcileAttachments = getFunctionBlock(taskpaneSource, "reconcileComposeEmailAttachments");
  assert.match(reconcileAttachments, /marker\.emailAttachmentDecision === "exclude"/);
  assert.match(reconcileAttachments, /preserveExclusion \? "exclude" : undefined/);

  const prefetchAttachments = getFunctionBlock(taskpaneSource, "prefetchComposeEmailAttachments");
  assert.match(prefetchAttachments, /collectEmailAttachmentMetadata\(item, bodyHtml\)/);
  assert.match(prefetchAttachments, /prepareEmailAttachmentsForHalo\(/);
  assert(
    prefetchAttachments.indexOf("collectEmailAttachmentMetadata(item, bodyHtml)") <
      prefetchAttachments.indexOf("prepareEmailAttachmentsForHalo("),
    "Compose attachment metadata must be refreshed immediately before prefetch."
  );
  assert.match(prefetchAttachments, /composeEmailAttachmentPreparationQueue\.then/);
  const runAttachmentPreparation = getFunctionBlock(
    taskpaneSource,
    "runComposeEmailAttachmentPreparation"
  );
  assert.match(runAttachmentPreparation, /clearComposeEmailAttachmentState\(item\)/);
  assert.match(runAttachmentPreparation, /saveComposeSessionMarker\(item, marker\)/);
  assert.match(runAttachmentPreparation, /saveComposeCustomMarker\(item, marker\)/);
  assert.match(runAttachmentPreparation, /failedSummary\.failed = Math\.max\(1/);

  const autoAttach = getFunctionBlock(taskpaneSource, "autoAttachCurrentEmail");
  assert.match(autoAttach, /recoverReadModeComposeAttachIds\(snapshot\.email\)/);
  assert.match(autoAttach, /composeAttachIds,/);
  assert.equal(
    (autoAttach.match(/if \(!match\.ok\)/g) || []).length,
    2,
    "Both the initial lookup and recovered lookup must reject server failures."
  );
  assert(
    autoAttach.indexOf("recoverReadModeComposeAttachIds(snapshot.email)") <
      autoAttach.lastIndexOf('fetchJson<HaloEmailMatchResponse>("/api/halo/email/match"'),
    "Read-mode recovery must validate recovered compose IDs through the authenticated match route."
  );

  const readModeRecovery = getFunctionBlock(taskpaneSource, "recoverReadModeComposeAttachIds");
  assert.match(readModeRecovery, /email\.inReplyToMessageIds/);
  assert.match(readModeRecovery, /email\.referenceMessageIds/);
  assert.match(readModeRecovery, /recoverReadModeComposeAttachIdsFromEws/);
  assert.doesNotMatch(
    readModeRecovery,
    /email\.(?:subject|dateTimeCreated|from|to)\b/,
    "Read-mode recovery must not use fuzzy message attributes."
  );

  const readModeFind = getFunctionBlock(taskpaneSource, "buildReadModeFindItemRequest");
  assert.match(readModeFind, /message:InternetMessageId/);
  assert.match(readModeFind, /DistinguishedFolderId/);
  const readModeMetadata = getFunctionBlock(taskpaneSource, "buildReadModeMetadataGetItemRequest");
  assert.match(readModeMetadata, /item:InternetMessageHeaders/);
  assert.match(readModeMetadata, /COMPOSE_CUSTOM_PROPERTY_NAME/);

  console.log("Compose feature wiring tests passed");
}

function getFunctionBlock(source, functionName) {
  const start = source.indexOf(`function ${functionName}`);
  assert(start >= 0, `Missing ${functionName}.`);
  const nextFunction = source.indexOf("\nfunction ", start + 1);
  return source.slice(start, nextFunction < 0 ? source.length : nextFunction);
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

run();
