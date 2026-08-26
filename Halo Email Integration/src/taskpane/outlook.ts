/* global document, window, Office, fetch, RequestInit, Response, URLSearchParams, HTMLInputElement, HTMLButtonElement, HTMLFormElement, HTMLSelectElement, HTMLElement, SVGElement */
import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";
import {
  clearInlineImagePreparationCache,
  PreparedInlineImages,
  prepareOutlookInlineImages,
} from "../shared/inlineImages";
import {
  collectEmailAttachmentMetadata,
  EmailAttachmentDescriptor,
  EmailAttachmentSummary,
  MAX_EMAIL_ATTACHMENTS,
  MAX_EMAIL_ATTACHMENT_BYTES,
  MAX_EMAIL_ATTACHMENT_TOTAL_BYTES,
  PreparedEmailAttachmentMetadata,
  prepareOutlookEmailAttachments,
} from "../shared/emailAttachments";

declare const __HALO_ADD_IN_VERSION__: string;

const BACKGROUND_SESSION_STORAGE_KEY = "halo-auth-background-session-v1";
const COMPOSE_ATTACH_STORAGE_KEY = "halo.composeAttach.v1";
const COMPOSE_ATTACH_HEADER_NAME = "X-Halo-Compose-Id";
const COMPOSE_CUSTOM_PROPERTY_NAME = "cecp-55bbcff2-8191-4411-aec6-f9d2f9b4b5e8";
const COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY = "halo.emailAttachmentPrefetch.v2";
const ADD_IN_VERSION = __HALO_ADD_IN_VERSION__;
const EMAIL_ATTACHMENT_READ_RETRY_DELAYS_MS = [200, 500];
const EMAIL_ATTACHMENT_READ_TIMEOUT_MS = 15_000;
const READ_MAPPING_RECOVERY_TIMEOUT_MS = 8_000;

type AuthConfigResponse = {
  authority: string;
  clientId: string;
  scopes: string[];
  ssoEnabled: boolean;
};

type AuthStartResponse = {
  dialogUrl: string;
};

type AuthStatusResponse = {
  authenticated: boolean;
  backgroundSessionId?: string;
  haloUrl?: string;
  expiresAt?: string;
};

type AuthCompleteResponse = {
  authenticated: boolean;
  backgroundSessionId?: string;
  expiresAt?: string;
};

type BugReportSessionResponse = {
  expiresAt: string;
  url: string;
};

type AuthDialogMessage = {
  type: "halo-auth";
  status: "success" | "failed";
  message?: string;
  error?: string;
  debug?: unknown;
  handoffCode?: string;
};

type HaloPingResponse = {
  ok: boolean;
  message: string;
  error?: string;
  debug?: unknown;
};

type HaloTicket = {
  id: string;
  ticketNumber: string;
  summary: string;
  status: string;
  lifecycle: "open" | "closed";
  client: string;
  agent: string;
};

type TicketOwnership = "mine" | "all";
type TicketLifecycle = "open" | "closed" | "all";
type ActionMode = "email" | "private-note";

type HaloTicketsResponse = {
  ok: boolean;
  tickets: HaloTicket[];
  message?: string;
  error?: string;
  debug?: unknown;
};

type TicketCreationType = {
  id: string;
  name: string;
  group: string;
  active: boolean;
  canCreate: boolean;
};

type TicketCreationOption = {
  value: string;
  label: string;
};

type TicketCreationField = {
  key: string;
  id: string;
  property: string;
  label: string;
  type: string;
  entity: string;
  required: boolean;
  supported: boolean;
  defaultValue: unknown;
  options: TicketCreationOption[];
  optionSource: string;
  order: number;
  core: boolean;
  managed: boolean;
  recommended: boolean;
  visibleOnCreate: boolean;
};

type TicketCreationSchema = {
  typeId: string;
  typeName: string;
  revision: string;
  available: boolean;
  unavailableReason: string;
  warnings: string[];
  defaults: Record<string, unknown>;
  fields: TicketCreationField[];
};

type TicketCreationTypesResponse = {
  ok: boolean;
  types: TicketCreationType[];
  stale?: boolean;
  error?: string;
};

type TicketCreationSchemaResponse = {
  ok: boolean;
  schema: TicketCreationSchema;
  stale?: boolean;
  error?: string;
};

type HaloRequester = {
  id: string;
  name: string;
  emailAddress: string;
  clientId: string;
  clientName: string;
  siteId: string;
  siteName: string;
};

type HaloRequestersResponse = {
  ok: boolean;
  requesters: HaloRequester[];
  error?: string;
};

type HaloLookupResponse = {
  ok: boolean;
  results: Array<{ id: string; label: string; secondary: string }>;
  error?: string;
};

type TicketCreationResult = HaloAttachEmailResponse & {
  ticketId?: string;
  ticketNumber?: string;
  operationId?: string;
};

type TicketCreationContext = {
  subject: string;
  requesterEmail: string;
  composeMode: boolean;
};

type TicketCreationFormValue = {
  typeId: string;
  schemaRevision: string;
  summary: string;
  summaryMode: "auto" | "fixed";
  values: Record<string, unknown>;
  requesterMode: "auto" | "explicit";
  requester: HaloRequester;
};

type TicketCreationFormControl = HTMLElement & {
  value: string;
  required: boolean;
};

type TicketCreationOptionGroup = HTMLElement & {
  label: string;
};

type EmailAddressPayload = {
  displayName: string;
  emailAddress: string;
};

type OutlookEmailPayload = {
  actionMode?: ActionMode;
  bodyHtml: string;
  bodyText: string;
  cc: EmailAddressPayload[];
  conversationId: string;
  dateTimeCreated: string;
  from: EmailAddressPayload | null;
  inReplyToMessageIds: string[];
  internetHeaders: string;
  internetMessageId: string;
  itemId: string;
  mailboxEmail: string;
  normalizedSubject: string;
  referenceMessageIds: string[];
  subject: string;
  timeZone: string;
  to: EmailAddressPayload[];
  inlineImageRefs: PreparedInlineImages["inlineImageRefs"];
  inlineImageUploads: PreparedInlineImages["inlineImageUploads"];
  inlineImageFingerprint: string;
  inlineImageTimings: PreparedInlineImages["inlineImageTimings"];
  includeEmailAttachments?: boolean;
  emailAttachmentPrefetchKey?: string;
  emailAttachmentFingerprint?: string;
  emailAttachmentDraftItemId?: string;
  emailAttachmentOperationId?: string;
  emailAttachmentStagingVersion?: number;
  emailAttachmentSummary?: EmailAttachmentSummary;
};

type InlineImageSummary = {
  referenced: number;
  cacheHits: number;
  uploaded: number;
  failed: number;
  warnings: string[];
};

type InlineImagePrefetchResponse = {
  ok: boolean;
  inlineImagePrefetchKey?: string;
  summary?: InlineImageSummary;
  error?: string;
};

type HaloAttachEmailResponse = {
  ok: boolean;
  attachMode?: "full-chain" | "latest-reply";
  message: string;
  actionId?: string;
  backgroundSessionId?: string;
  error?: string;
  debug?: unknown;
  status?:
    | "attached"
    | "attached-with-image-warnings"
    | "attached-with-attachment-warnings"
    | "attached-with-warnings";
  emailAttachments?: EmailAttachmentSummary;
  inlineImages?: InlineImageSummary;
};

type HaloAutoAttachResponse = {
  ok: boolean;
  status:
    | "attached"
    | "attached-with-image-warnings"
    | "attached-with-attachment-warnings"
    | "attached-with-warnings"
    | "already-attached"
    | "no-match";
  ticketId?: string;
  ticketNumber?: string;
  message?: string;
  actionId?: string;
  error?: string;
  debug?: unknown;
  inlineImages?: InlineImageSummary;
  emailAttachments?: EmailAttachmentSummary;
};

type HaloEmailMatchResponse = {
  ok: boolean;
  status: "matched" | "already-attached" | "no-match";
  ticketId?: string;
  ticketNumber?: string;
  actionMode?: ActionMode;
  message?: string;
  error?: string;
  debug?: unknown;
};

type OutlookEmailSnapshot = {
  email: OutlookEmailPayload;
  emailAttachments: Promise<PreparedEmailAttachmentMetadata>;
  inlineImages: Promise<PreparedInlineImages>;
  itemRevision: number;
};

type HaloAuthError = Error & {
  debug?: unknown;
  operationId?: string;
  statusCode?: number;
  ticketNumber?: string;
};

type ComposeAttachMarker = {
  version: 1 | 2 | 3 | 4;
  destinationKind?: "existing-ticket" | "create-ticket";
  backgroundSessionId?: string;
  composeAttachId: string;
  ticketId: string;
  ticketNumber: string;
  ticketSummary: string;
  draftItemId: string;
  creationOperationId?: string;
  ticketTypeId?: string;
  ticketTypeName?: string;
  inlineImagePrefetchKey?: string;
  inlineImageFingerprint?: string;
  emailAttachmentDecision?: "include" | "exclude";
  emailAttachmentFingerprint?: string;
  emailAttachmentPrefetchKey?: string;
  emailAttachmentStagingVersion?: number;
  emailAttachmentSummary?: EmailAttachmentSummary;
  mailboxEmail?: string;
  actionMode?: ActionMode;
  composeIdentityHeader?: boolean;
};

type ComposeEmailAttachmentPrefetchState = {
  version: 2;
  draftItemId: string;
  ticketId: string;
  operationId: string;
  emailAttachmentDecision: "include" | "exclude";
  emailAttachmentFingerprint: string;
  emailAttachmentPrefetchKey: string;
  emailAttachmentStagingVersion: 2;
  emailAttachmentSummary: EmailAttachmentSummary;
};

type EmailAttachmentPrefetchStartResponse = {
  ok: boolean;
  status: "pending" | "ready" | "no-match" | "already-attached";
  stagingVersion?: number;
  aggregate?: { pending: number; prepared: number; failed: number; selected: number };
  prefetchKey?: string;
  pendingAttachmentKeys?: string[];
  ticketId?: string;
  ticketNumber?: string;
  error?: string;
};

type ComposeSessionData = {
  getAsync: (name: string, callback: (result: Office.AsyncResult<string>) => void) => void;
  removeAsync: (name: string, callback: (result: Office.AsyncResult<void>) => void) => void;
  setAsync: (
    name: string,
    value: string,
    callback: (result: Office.AsyncResult<void>) => void
  ) => void;
};

type ComposeCustomProperties = {
  get: (name: string) => unknown;
  remove: (name: string) => void;
  saveAsync: (callback: (result: Office.AsyncResult<void>) => void) => void;
  set: (name: string, value: string) => void;
};

type ComposeMetadataItem = {
  body?: {
    getAsync: (
      coercionType: Office.CoercionType,
      callback: (result: Office.AsyncResult<string>) => void
    ) => void;
    setAsync?: (
      data: string,
      options: { coercionType: Office.CoercionType },
      callback: (result: Office.AsyncResult<void>) => void
    ) => void;
  };
  attachments?: unknown[];
  getAttachmentsAsync?: (callback: (result: Office.AsyncResult<unknown[]>) => void) => void;
  getAttachmentContentAsync?: (
    attachmentId: string,
    callback: (result: Office.AsyncResult<{ content: string; format: unknown }>) => void
  ) => void;
  itemType?: Office.MailboxEnums.ItemType | string;
  itemId?: string;
  conversationId?: string;
  inReplyTo?: string;
  internetMessageId?: string;
  internetHeaders?: {
    setAsync: (
      headers: Record<string, string>,
      callback: (result: Office.AsyncResult<void>) => void
    ) => void;
    removeAsync: (names: string[], callback: (result: Office.AsyncResult<void>) => void) => void;
  };
  loadCustomPropertiesAsync?: (
    callback: (result: Office.AsyncResult<ComposeCustomProperties>) => void
  ) => void;
  getItemIdAsync?: (callback: (result: Office.AsyncResult<string>) => void) => void;
  saveAsync?: (callback: (result: Office.AsyncResult<string>) => void) => void;
  sessionData?: ComposeSessionData;
};

let currentDialog: Office.Dialog | null = null;
let bugReportDialog: Office.Dialog | null = null;
let waitingForDialog = false;
let checkingSession = false;
let sessionCheckRequested = false;
let authConfigPromise: Promise<AuthConfigResponse> | null = null;
let msalInstancePromise: Promise<unknown> | null = null;
let activeBackgroundSessionId = "";
let activeComposeAttachMarker: ComposeAttachMarker | null = null;
let composeMarkerRevision = 0;
let composeMetadataWriteQueue: Promise<unknown> = Promise.resolve();
let composeEmailAttachmentPreparationQueue: Promise<void> = Promise.resolve();
let composeEmailAttachmentPreparationRevision = 0;
let composeInlineImagePrefetchRevision = 0;
let outlookItemRevision = 0;
let activeAttachmentPromptResolve: ((choice: boolean | null) => void) | null = null;
let loadedTickets: HaloTicket[] = [];
let lastExecutedTicketQuery = "";
let ticketRequestRevision = 0;
let ticketSearchDebounceId: number | null = null;
let ticketCreationTypes: TicketCreationType[] = [];
let activeTicketCreationSchema: TicketCreationSchema | null = null;
let activeTicketCreationRequester: HaloRequester | null = null;
let ticketCreationRequesters: HaloRequester[] = [];
let ticketCreationRequesterMode: "auto" | "explicit" = "auto";
let ticketCreationInitialSummary = "";
let ticketCreationContext: TicketCreationContext | null = null;
let ticketCreationIsBusy = false;
let readTicketCreationOperationId = "";
let mappedConversationTicket: {
  ticketId: string;
  ticketNumber: string;
  actionMode: ActionMode;
} | null = null;
let activeActionMode: ActionMode = "email";
let actionModeLocked = false;
let appIsBusy = false;
let ticketActionIsBusy = false;
let ticketQueryIsBusy = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    showApp();
    bindControls();
    checkExistingSession();
  }
});

function showApp() {
  const sideloadMessage = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");

  if (sideloadMessage) {
    sideloadMessage.style.display = "none";
  }

  if (appBody) {
    appBody.style.display = "flex";
  }
}

function bindControls() {
  getLoginButton().onclick = startHaloLogin;
  getLogoutButton().onclick = logout;
  getReportBugButton().onclick = () => void openBugReport();
  getRefreshTicketsButton().onclick = () => void loadTickets();
  getTicketSearchForm().onsubmit = (event) => {
    event.preventDefault();
    cancelTicketSearchDebounce();
    void searchTickets({ allowShortQuery: true });
  };
  getTicketQueryInput().oninput = scheduleTicketSearch;
  getClearSearchButton().onclick = () => void clearTicketSearch();
  getTicketOwnershipSelect().onchange = () => void loadTickets();
  getTicketLifecycleSelect().onchange = () => void loadTickets();
  getTicketCustomerSelect().onchange = onTicketFacetFilterChanged;
  getTicketAssigneeSelect().onchange = onTicketFacetFilterChanged;
  getResetTicketFiltersButton().onclick = () => void resetTicketFilters();
  getRemoveComposeSelectionButton().onclick = () => void removeComposeAttachSelection();
  getPrivateNoteToggle().onchange = () => void onActionModeChanged();
  getAttachMappedEmailButton().onclick = () => void attachMappedCurrentEmail();
  getAttachExistingTab().onclick = () => showTicketDestination("existing");
  getCreateTicketTab().onclick = () => void showTicketDestination("create");
  getRefreshTicketTypesButton().onclick = () => void loadTicketCreationTypes(true);
  getCreateTicketTypeSelect().onchange = () =>
    void loadTicketCreationSchema(getCreateTicketTypeSelect().value);
  getCreateTicketSummaryInput().oninput = updateTicketCreationSubmitState;
  getCreateTicketForm().onsubmit = (event) => {
    event.preventDefault();
    void submitTicketCreation();
  };
  getSearchRequestersButton().onclick = () => void searchTicketCreationRequesters();
  getCreateTicketRequesterQuery().onkeydown = (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      void searchTicketCreationRequesters();
    }
  };
  getCreateTicketRequesterSelect().onchange = () => {
    activeTicketCreationRequester = readSelectedRequester();
    ticketCreationRequesterMode = "explicit";
    updateTicketCreationSubmitState();
  };
  registerItemChangedHandler();
  registerComposeAttachmentsChangedHandler();
}

function normalizeActionMode(value: unknown): ActionMode {
  return value === "private-note" ? "private-note" : "email";
}

function setActionMode(mode: ActionMode, options: { disabled?: boolean } = {}) {
  activeActionMode = normalizeActionMode(mode);
  const toggle = getPrivateNoteToggle();
  const visibility = getActionModeVisibility();
  toggle.checked = activeActionMode === "private-note";
  actionModeLocked = Boolean(options.disabled);
  toggle.disabled = actionModeLocked;
  visibility.dataset.visibility = activeActionMode;
  visibility.textContent =
    activeActionMode === "private-note" ? "Customer hidden" : "Customer visible";
  getActionModeHelp().textContent =
    activeActionMode === "private-note"
      ? "Private note — agents only."
      : "Email action — visible to customer.";
}

async function onActionModeChanged() {
  setActionMode(getPrivateNoteToggle().checked ? "private-note" : "email");
  composeInlineImagePrefetchRevision += 1;
  const marker = activeComposeAttachMarker;
  const item = getComposeMetadataItem();
  if (!marker || !item) {
    return;
  }

  marker.actionMode = activeActionMode;
  marker.inlineImagePrefetchKey = undefined;
  marker.inlineImageFingerprint = undefined;
  await enqueueComposeMetadataWrite(async () => {
    await saveComposeSessionMarker(item, marker);
    await saveComposeCustomMarker(item, marker);
    await saveComposeItem(item);
  }).catch((error) =>
    setFailed(error, { hideLogout: false, message: "Email visibility could not be saved" })
  );

  if (isCreateTicketMarker(marker)) {
    await fetchJson(
      `/api/halo/ticket-creation/intents/${encodeURIComponent(
        marker.creationOperationId || marker.composeAttachId
      )}`,
      {
        method: "POST",
        body: JSON.stringify({ actionMode: activeActionMode }),
      }
    ).catch((error) =>
      setFailed(error, { hideLogout: false, message: "Email visibility could not be saved" })
    );
  } else {
    void prefetchComposeInlineImages(item, marker);
  }
}

function renderMappedEmailConfirmation() {
  const confirmation = getMappedEmailConfirmation();
  const show = Boolean(mappedConversationTicket && !getComposeMetadataItem());
  confirmation.hidden = !show;
  if (!show || !mappedConversationTicket) {
    getMappedEmailTicket().textContent = "";
    return;
  }
  getMappedEmailTicket().textContent = `Mapped to ticket ${mappedConversationTicket.ticketNumber}. Review visibility, then attach.`;
}

async function openBugReport() {
  const button = getReportBugButton();
  button.disabled = true;

  try {
    const diagnostics = Office.context.diagnostics;
    const mailboxDiagnostics = Office.context.mailbox.diagnostics;
    const session = await fetchJson<BugReportSessionResponse>("/api/bug-reports/session", {
      body: JSON.stringify({
        diagnostics: {
          addInVersion: ADD_IN_VERSION,
          officeVersion: diagnostics.version || mailboxDiagnostics.hostVersion || "",
          outlookHost: String(diagnostics.host || mailboxDiagnostics.hostName || "Outlook"),
          outlookPlatform: String(diagnostics.platform || ""),
        },
      }),
      method: "POST",
    });

    await openBugReportPage(session.url);
  } catch (error) {
    setStatus(
      "failed",
      "Could not open bug reporting",
      error instanceof Error ? error.message : "Bug reporting is temporarily unavailable."
    );
  } finally {
    button.disabled = false;
  }
}

function openBugReportPage(url: string): Promise<void> {
  const openBrowserWindow = Office.context.ui.openBrowserWindow;

  if (typeof openBrowserWindow === "function") {
    try {
      openBrowserWindow.call(Office.context.ui, url);
      return Promise.resolve();
    } catch {
      // Some Outlook hosts expose the API without supporting it at runtime.
    }
  }

  return openBugReportDialog(url);
}

function openBugReportDialog(url: string): Promise<void> {
  closeBugReportDialog();

  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      url,
      { height: 80, width: 60, displayInIframe: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          const message =
            asyncResult.error.code === 12009
              ? "Allow the bug report dialog when Outlook prompts you, then try again."
              : asyncResult.error.message || "Could not open the bug report window.";
          reject(new Error(message));
          return;
        }

        const dialog = asyncResult.value;
        bugReportDialog = dialog;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          if (bugReportDialog === dialog) {
            bugReportDialog = null;
          }
        });
        resolve();
      }
    );
  });
}

function closeBugReportDialog() {
  const dialog = bugReportDialog;
  bugReportDialog = null;
  if (dialog) {
    dialog.close();
  }
}

async function checkExistingSession() {
  if (checkingSession || waitingForDialog) {
    if (checkingSession) {
      sessionCheckRequested = true;
    }
    return;
  }

  checkingSession = true;
  sessionCheckRequested = false;

  try {
    const status = await fetchJson<AuthStatusResponse>(
      "/api/auth/status",
      {},
      { allowMissingAuth: true, interactive: false }
    );

    if (!status.authenticated) {
      setSignedOut();
      return;
    }

    await saveBackgroundSessionId(status.backgroundSessionId || "");
    setBusy(true);
    setStatus("loading", "Checking Halo API auth...", status.haloUrl || "");
    await refreshBackgroundSessionId();
    await pingHalo();
    void warmCurrentInlineImages();
    await restoreComposeAttachSelection();
    await prepareMappedComposeAttachmentsOnOpen();
    if (!(await autoAttachCurrentEmail())) {
      await loadTickets();
    }
  } catch (error) {
    setFailed(error);
  } finally {
    checkingSession = false;
    setBusy(false);
    if (sessionCheckRequested && !waitingForDialog) {
      sessionCheckRequested = false;
      void checkExistingSession();
    }
  }
}

async function startHaloLogin() {
  try {
    setBusy(true);
    setStatus("loading", "Opening Halo login...", "A Halo sign-in dialog will open.");

    const authStart = await fetchJson<AuthStartResponse>("/api/auth/start", {
      method: "POST",
    });

    openHaloDialog(authStart.dialogUrl);
  } catch (error) {
    setBusy(false);
    setFailed(error);
  }
}

function openHaloDialog(dialogUrl: string) {
  closeBugReportDialog();

  if (currentDialog) {
    currentDialog.close();
    currentDialog = null;
  }

  waitingForDialog = true;

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 60, width: 40, displayInIframe: false },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        waitingForDialog = false;
        setBusy(false);
        setStatus("failed", "Halo API Auth failed", asyncResult.error.message);
        return;
      }

      currentDialog = asyncResult.value;
      currentDialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        (arg) => void onDialogMessageReceived(arg as { message: string })
      );
      currentDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) =>
        onDialogEventReceived(arg as { error: number })
      );
    }
  );
}

async function onDialogMessageReceived(arg: { message: string }) {
  let message: AuthDialogMessage;

  try {
    message = JSON.parse(arg.message);
  } catch {
    closeDialog();
    setBusy(false);
    setStatus("failed", "Halo API Auth failed", "Halo login returned an invalid response.");
    return;
  }

  if (message.type !== "halo-auth") {
    return;
  }

  closeDialog();

  if (message.status !== "success" || !message.handoffCode) {
    setBusy(false);
    setStatus(
      "failed",
      message.message || "Halo API Auth failed",
      message.error || "",
      message.debug
    );
    return;
  }

  try {
    setStatus("loading", "Completing Halo login...", "");
    const complete = await fetchJson<AuthCompleteResponse>("/api/auth/complete", {
      method: "POST",
      body: JSON.stringify({ handoffCode: message.handoffCode }),
    });
    await saveBackgroundSessionId(complete.backgroundSessionId || "");
    await pingHalo();
    void warmCurrentInlineImages();
    await restoreComposeAttachSelection();
    await prepareMappedComposeAttachmentsOnOpen();
    if (!(await autoAttachCurrentEmail())) {
      await loadTickets();
    }
  } catch (error) {
    setFailed(error);
  } finally {
    setBusy(false);
  }
}

function onDialogEventReceived(arg: { error: number }) {
  if (!currentDialog || !waitingForDialog) {
    return;
  }

  currentDialog = null;
  waitingForDialog = false;
  setBusy(false);

  if (arg.error === 12006) {
    setStatus("failed", "Halo API Auth failed", "Halo login was cancelled.");
    return;
  }

  setStatus(
    "failed",
    "Halo API Auth failed",
    `The Halo login dialog closed unexpectedly (${arg.error}).`
  );
}

async function pingHalo() {
  const result = await fetchJson<HaloPingResponse>("/api/halo/ping");

  if (!result.ok) {
    throw createHaloAuthError(result.error || result.message, result.debug);
  }

  setConnectionState(true);
  setStatus("success", "Connected to HaloPSA", "Choose a ticket to attach the open email.");
}

async function loadTickets() {
  const ticketsPanel = getTicketsPanel();
  ticketsPanel.hidden = false;
  return requestTickets(getTicketQueryInput().value.trim());
}

async function showTicketDestination(destination: "existing" | "create") {
  const createMode = destination === "create";
  placeActionModeControl(destination);
  getAttachExistingTab().setAttribute("aria-selected", String(!createMode));
  getCreateTicketTab().setAttribute("aria-selected", String(createMode));
  getCreateTicketPanel().hidden = !createMode;
  document.querySelectorAll(".halo-attach-existing-panel").forEach((element) => {
    (element as HTMLElement).hidden = createMode;
  });
  if (!createMode) {
    updateTicketBusyState();
    return;
  }
  renderCreateTicketConversationWarning();
  try {
    ticketCreationContext = await readTicketCreationContext();
    ticketCreationInitialSummary = ticketCreationContext.subject || "(no subject)";
    getCreateTicketSummaryInput().value = ticketCreationInitialSummary;
    getSubmitCreateTicketButton().textContent = ticketCreationContext.composeMode
      ? "Create when sent"
      : "Create ticket";
    await loadTicketCreationTypes(false);
    if (ticketCreationContext.requesterEmail) {
      getCreateTicketRequesterQuery().value = ticketCreationContext.requesterEmail;
      await searchTicketCreationRequesters(true);
    }
  } catch (error) {
    setFailed(error, { hideLogout: false, message: "Ticket creation could not be loaded" });
  }
}

async function loadTicketCreationTypes(refresh: boolean) {
  setTicketCreationBusy(true);
  try {
    setStatus(
      "loading",
      refresh ? "Refreshing Halo ticket types..." : "Loading Halo ticket types...",
      "Building creation profiles from the current Halo configuration."
    );
    const result = await fetchJson<TicketCreationTypesResponse>(
      `/api/halo/ticket-creation/types${refresh ? "?refresh=1" : ""}`
    );
    if (!result.ok) {
      throw createHaloAuthError(result.error || "Halo ticket type discovery failed.");
    }
    ticketCreationTypes = result.types || [];
    const select = getCreateTicketTypeSelect();
    const previous = select.value;
    while (select.firstChild) {
      select.removeChild(select.firstChild);
    }
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = ticketCreationTypes.length
      ? "Choose a ticket type"
      : "No creatable ticket types available";
    select.appendChild(placeholder);
    const groups = new Map<string, TicketCreationOptionGroup>();
    ticketCreationTypes.forEach((type) => {
      const option = document.createElement("option");
      option.value = type.id;
      option.textContent = type.name;
      if (type.group) {
        let group = groups.get(type.group);
        if (!group) {
          group = document.createElement("optgroup");
          group.label = type.group;
          groups.set(type.group, group);
          select.appendChild(group);
        }
        group.appendChild(option);
      } else {
        select.appendChild(option);
      }
    });
    select.value = ticketCreationTypes.some((type) => type.id === previous) ? previous : "";
    if (!select.value && ticketCreationTypes.length === 1) {
      select.value = ticketCreationTypes[0].id;
    }
    if (select.value) {
      await loadTicketCreationSchema(select.value, refresh);
    } else {
      activeTicketCreationSchema = null;
      clearTicketCreationFields();
    }
    setStatus(
      result.stale ? "warning" : "success",
      `${ticketCreationTypes.length} ticket type${ticketCreationTypes.length === 1 ? "" : "s"} available`,
      result.stale
        ? "Using recently cached Halo configuration because it could not be refreshed."
        : "Choose a Halo ticket type to load its required fields."
    );
  } finally {
    setTicketCreationBusy(false);
  }
}

async function loadTicketCreationSchema(typeId: string, refresh = false) {
  activeTicketCreationSchema = null;
  clearTicketCreationFields();
  if (!typeId) {
    updateTicketCreationSubmitState();
    return;
  }
  setTicketCreationBusy(true);
  try {
    setStatus(
      "loading",
      "Loading Halo ticket fields...",
      "Reading the selected type configuration."
    );
    const result = await fetchJson<TicketCreationSchemaResponse>(
      `/api/halo/ticket-creation/types/${encodeURIComponent(typeId)}/schema${
        refresh ? "?refresh=1" : ""
      }`
    );
    if (!result.ok || !result.schema) {
      throw createHaloAuthError(result.error || "Halo ticket field discovery failed.");
    }
    activeTicketCreationSchema = result.schema;
    renderTicketCreationSchema(result.schema);
    setStatus(
      result.schema.available ? (result.stale ? "warning" : "success") : "warning",
      result.schema.available
        ? `${result.schema.typeName} fields loaded`
        : "Ticket type unavailable",
      result.schema.available
        ? "Complete the required fields, then create the ticket."
        : result.schema.unavailableReason
    );
  } catch (error) {
    activeTicketCreationSchema = null;
    clearTicketCreationFields();
    setFailed(error, { hideLogout: false, message: "Ticket fields could not be loaded" });
  } finally {
    setTicketCreationBusy(false);
  }
}

function renderTicketCreationSchema(schema: TicketCreationSchema) {
  clearTicketCreationFields();
  const selectedTypeOption = Array.from(getCreateTicketTypeSelect().options).find(
    (option) => option.value === schema.typeId
  );
  if (selectedTypeOption) {
    selectedTypeOption.disabled = !schema.available;
    selectedTypeOption.title = schema.available ? "" : schema.unavailableReason;
  }
  const requiredContainer = getCreateTicketRequiredFields();
  const optionalContainer = getCreateTicketOptionalFields();
  schema.fields.forEach((field) => {
    if (
      field.managed ||
      !field.supported ||
      ["user_id", "client_id", "site_id"].includes(field.property)
    ) {
      return;
    }
    const control = createTicketCreationFieldControl(field, schema.defaults[field.key]);
    (field.required || field.recommended ? requiredContainer : optionalContainer).appendChild(
      control
    );
  });
  getCreateTicketOptionalDetails().hidden = !optionalContainer.childElementCount;
  const warning = getCreateTicketSchemaWarning();
  const warnings = [schema.unavailableReason, ...schema.warnings].filter(Boolean);
  warning.textContent = warnings.join(" ");
  warning.hidden = !warnings.length;
  updateTicketCreationSubmitState();
}

function createTicketCreationFieldControl(
  field: TicketCreationField,
  typeDefault: unknown
): HTMLElement {
  const label = document.createElement("label");
  label.className = "halo-form-field";
  const id = `create-field-${field.key.replace(/[^A-Za-z0-9_-]/g, "-")}`;
  label.htmlFor = id;
  const caption = document.createElement("span");
  caption.textContent = `${field.label}${field.required ? " *" : ""}`;
  label.appendChild(caption);
  const defaultValue = isEmptyTicketCreationValue(field.defaultValue)
    ? typeDefault
    : field.defaultValue;
  let control: TicketCreationFormControl;
  if (["agent", "asset", "client", "site", "team", "user"].includes(field.type)) {
    const search = document.createElement("div");
    search.className = "halo-requester-search";
    const query = document.createElement("input");
    query.type = "search";
    query.placeholder = `Search Halo ${field.label.toLowerCase()}`;
    const find = document.createElement("button");
    find.type = "button";
    find.className = "halo-button";
    find.textContent = "Find";
    const select = document.createElement("select");
    const empty = document.createElement("option");
    empty.value = "";
    empty.textContent = `Choose ${field.label.toLowerCase()}`;
    select.appendChild(empty);
    if (!isEmptyTicketCreationValue(defaultValue)) {
      const configured = document.createElement("option");
      configured.value = String(defaultValue);
      configured.textContent = `Halo default (${defaultValue})`;
      configured.selected = true;
      select.appendChild(configured);
    }
    const runSearch = () => void searchTicketCreationLookup(field, query.value, select);
    find.onclick = runSearch;
    query.onkeydown = (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        runSearch();
      }
    };
    search.appendChild(query);
    search.appendChild(find);
    label.appendChild(search);
    control = select;
  } else if (field.type === "multiline") {
    control = document.createElement("textarea");
    control.value = String(defaultValue || "");
  } else if (field.type === "select" || field.type === "multiselect") {
    const select = document.createElement("select");
    select.multiple = field.type === "multiselect";
    if (!select.multiple) {
      const empty = document.createElement("option");
      empty.value = "";
      empty.textContent = field.required
        ? `Choose ${field.label.toLowerCase()}`
        : "Use Halo default";
      empty.disabled = field.required;
      empty.selected = isEmptyTicketCreationValue(defaultValue);
      select.appendChild(empty);
    }
    field.options.forEach((optionValue) => {
      const option = document.createElement("option");
      option.value = optionValue.value;
      option.textContent = optionValue.label;
      option.selected = Array.isArray(defaultValue)
        ? defaultValue.map(String).includes(optionValue.value)
        : String(defaultValue ?? "") === optionValue.value;
      select.appendChild(option);
    });
    control = select;
  } else {
    const input = document.createElement("input");
    if (field.type === "boolean") {
      input.type = "checkbox";
      input.checked = Boolean(defaultValue);
      if (!field.required && isEmptyTicketCreationValue(defaultValue)) {
        input.dataset.ticketFieldOmitDefault = "true";
      }
    } else if (field.type === "duration") {
      input.type = "text";
      input.inputMode = "numeric";
      input.pattern = "[0-9]{1,4}:[0-5][0-9]";
      input.placeholder = "HH:MM";
      input.value = formatTicketCreationDuration(defaultValue);
    } else if (field.type === "number") {
      input.type = "number";
      input.step = "any";
      input.value = String(defaultValue ?? "");
    } else if (field.type === "date") {
      input.type = "date";
      input.value = String(defaultValue || "").slice(0, 10);
    } else if (field.type === "datetime") {
      input.type = "datetime-local";
      input.value = String(defaultValue || "").slice(0, 16);
    } else if (field.type === "time") {
      input.type = "time";
      input.value = String(defaultValue || "").slice(0, 8);
    } else {
      input.type = "text";
      input.value = String(defaultValue ?? "");
    }
    control = input;
  }
  control.id = id;
  control.dataset.ticketFieldKey = field.key;
  control.dataset.ticketFieldType = field.type;
  control.required = field.required && field.type !== "boolean";
  control.onchange = () => {
    control.dataset.ticketFieldOmitDefault = "false";
    updateTicketCreationSubmitState();
  };
  control.oninput = updateTicketCreationSubmitState;
  label.appendChild(control);
  return label;
}

function formatTicketCreationDuration(value: unknown): string {
  if (isEmptyTicketCreationValue(value)) {
    return "";
  }
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric < 0) {
    return String(value);
  }
  let hours = Math.floor(numeric);
  let minutes = Math.round((numeric - hours) * 60);
  if (minutes === 60) {
    hours += 1;
    minutes = 0;
  }
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

async function searchTicketCreationLookup(
  field: TicketCreationField,
  query: string,
  select: HTMLSelectElement
) {
  const normalizedQuery = query.trim();
  if (!normalizedQuery) {
    return;
  }
  select.disabled = true;
  try {
    const result = await fetchJson<HaloLookupResponse>(
      `/api/halo/ticket-creation/lookups/${encodeURIComponent(
        field.type
      )}?query=${encodeURIComponent(normalizedQuery)}`
    );
    if (!result.ok) {
      throw createHaloAuthError(result.error || "Halo lookup failed.");
    }
    while (select.firstChild) {
      select.removeChild(select.firstChild);
    }
    const empty = document.createElement("option");
    empty.value = "";
    empty.textContent = result.results.length
      ? `Choose ${field.label.toLowerCase()}`
      : "No matching Halo records";
    select.appendChild(empty);
    result.results.forEach((value) => {
      const option = document.createElement("option");
      option.value = value.id;
      option.textContent = [value.label, value.secondary].filter(Boolean).join(" — ");
      select.appendChild(option);
    });
  } catch (error) {
    setFailed(error, { hideLogout: false, message: `${field.label} lookup failed` });
  } finally {
    select.disabled = false;
    updateTicketCreationSubmitState();
  }
}

function clearTicketCreationFields() {
  [getCreateTicketRequiredFields(), getCreateTicketOptionalFields()].forEach((container) => {
    while (container.firstChild) {
      container.removeChild(container.firstChild);
    }
  });
  getCreateTicketOptionalDetails().hidden = true;
  getCreateTicketSchemaWarning().hidden = true;
}

async function searchTicketCreationRequesters(automatic = false) {
  const query = getCreateTicketRequesterQuery().value.trim();
  if (!query) {
    return;
  }
  setTicketCreationBusy(true);
  try {
    const result = await fetchJson<HaloRequestersResponse>(
      `/api/halo/ticket-creation/requesters?query=${encodeURIComponent(query)}`
    );
    if (!result.ok) {
      throw createHaloAuthError(result.error || "Halo requester search failed.");
    }
    ticketCreationRequesters = result.requesters || [];
    renderTicketCreationRequesters(ticketCreationRequesters);
    const emailQuery = query.toLowerCase();
    const exact = ticketCreationRequesters.filter(
      (requester) => requester.emailAddress.toLowerCase() === emailQuery
    );
    if (automatic && exact.length === 1) {
      activeTicketCreationRequester = exact[0];
      ticketCreationRequesterMode = "auto";
      getCreateTicketRequesterSelect().value = exact[0].id;
      getCreateTicketRequesterHelp().textContent = `Matched ${exact[0].name || exact[0].emailAddress} from the email.`;
    } else {
      activeTicketCreationRequester = null;
      ticketCreationRequesterMode = "explicit";
      getCreateTicketRequesterHelp().textContent = ticketCreationRequesters.length
        ? "Choose the requester to use for this ticket."
        : "No Halo requester matched. Try another email address or name.";
    }
    updateTicketCreationSubmitState();
  } catch (error) {
    setFailed(error, { hideLogout: false, message: "Requester search failed" });
  } finally {
    setTicketCreationBusy(false);
  }
}

function renderTicketCreationRequesters(requesters: HaloRequester[]) {
  const select = getCreateTicketRequesterSelect();
  while (select.firstChild) {
    select.removeChild(select.firstChild);
  }
  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = requesters.length ? "Choose a Halo requester" : "No requester found";
  select.appendChild(empty);
  requesters.forEach((requester) => {
    const option = document.createElement("option");
    option.value = requester.id;
    option.textContent = [requester.name || requester.emailAddress, requester.clientName]
      .filter(Boolean)
      .join(" — ");
    select.appendChild(option);
  });
}

function readSelectedRequester(): HaloRequester | null {
  const id = getCreateTicketRequesterSelect().value;
  return ticketCreationRequesters.find((requester) => requester.id === id) || null;
}

function collectTicketCreationForm(): TicketCreationFormValue {
  const schema = activeTicketCreationSchema;
  const requester = activeTicketCreationRequester || readSelectedRequester();
  const summary = getCreateTicketSummaryInput().value.trim();
  if (!schema || !schema.available) {
    throw createHaloAuthError(schema?.unavailableReason || "Choose an available ticket type.");
  }
  if (!requester) {
    throw createHaloAuthError("Choose a Halo requester before creating the ticket.");
  }
  const missingRequesterRelationship = schema.fields.find(
    (field) =>
      field.required &&
      ((field.property === "client_id" && !requester.clientId) ||
        (field.property === "site_id" && !requester.siteId))
  );
  if (missingRequesterRelationship) {
    throw createHaloAuthError(
      `${missingRequesterRelationship.label} is required. Choose a Halo requester with that relationship.`
    );
  }
  if (!summary) {
    throw createHaloAuthError("Enter a ticket summary.");
  }
  const values: Record<string, unknown> = {};
  document.querySelectorAll("[data-ticket-field-key]").forEach((element) => {
    const control = element as TicketCreationFormControl;
    const key = control.dataset.ticketFieldKey || "";
    const type = control.dataset.ticketFieldType || "";
    if (!key) {
      return;
    }
    if (type === "boolean" && control instanceof HTMLInputElement) {
      if (control.dataset.ticketFieldOmitDefault === "true") {
        return;
      }
      values[key] = control.checked;
    } else if (type === "multiselect" && control instanceof HTMLSelectElement) {
      values[key] = Array.from(control.selectedOptions).map((option) => option.value);
    } else if (control.value !== "") {
      values[key] = control.value;
    }
  });
  const composeMode = Boolean(ticketCreationContext && ticketCreationContext.composeMode);
  return {
    typeId: schema.typeId,
    schemaRevision: schema.revision,
    summary,
    summaryMode: composeMode && summary === ticketCreationInitialSummary ? "auto" : "fixed",
    values,
    requesterMode: composeMode ? ticketCreationRequesterMode : "explicit",
    requester,
  };
}

function updateTicketCreationSubmitState() {
  const schema = activeTicketCreationSchema;
  const form = getCreateTicketForm();
  getSubmitCreateTicketButton().disabled = Boolean(
    ticketCreationIsBusy ||
    !schema ||
    !schema.available ||
    !activeTicketCreationRequester ||
    !getCreateTicketSummaryInput().value.trim() ||
    !form.checkValidity()
  );
}

function setTicketCreationBusy(busy: boolean) {
  ticketCreationIsBusy = busy;
  [
    getCreateTicketTypeSelect(),
    getCreateTicketSummaryInput(),
    getCreateTicketRequesterQuery(),
    getCreateTicketRequesterSelect(),
    getSearchRequestersButton(),
    getRefreshTicketTypesButton(),
  ].forEach((control) => {
    control.disabled = busy;
  });
  updateTicketCreationSubmitState();
}

async function readTicketCreationContext(): Promise<TicketCreationContext> {
  const item = Office.context.mailbox.item as unknown as {
    subject?:
      string | { getAsync: (callback: (result: Office.AsyncResult<string>) => void) => void };
    from?: Office.EmailAddressDetails;
    to?:
      | Office.EmailAddressDetails[]
      | {
          getAsync: (
            callback: (result: Office.AsyncResult<Office.EmailAddressDetails[]>) => void
          ) => void;
        };
  };
  const composeMode = Boolean(getComposeMetadataItem());
  const subject = composeMode
    ? await readComposeStringProperty(item.subject)
    : String(item.subject || "");
  let requesterEmail = "";
  if (composeMode) {
    const recipients = await readComposeRecipientProperty(item.to);
    const mailboxEmail = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
    requesterEmail =
      recipients.find(
        (recipient) =>
          recipient.emailAddress && recipient.emailAddress.toLowerCase() !== mailboxEmail
      )?.emailAddress || "";
  } else {
    const mailboxEmail = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
    const senderEmail = String(item.from?.emailAddress || "").toLowerCase();
    if (senderEmail && senderEmail !== mailboxEmail) {
      requesterEmail = item.from?.emailAddress || "";
    } else {
      const recipients = await readComposeRecipientProperty(item.to);
      requesterEmail =
        recipients.find(
          (recipient) =>
            recipient.emailAddress && recipient.emailAddress.toLowerCase() !== mailboxEmail
        )?.emailAddress || "";
    }
  }
  return { subject: subject || "(no subject)", requesterEmail, composeMode };
}

function readComposeStringProperty(
  value?: string | { getAsync: (callback: (result: Office.AsyncResult<string>) => void) => void }
): Promise<string> {
  if (typeof value === "string") {
    return Promise.resolve(value);
  }
  return new Promise((resolve) => {
    if (!value || !value.getAsync) {
      resolve("");
      return;
    }
    value.getAsync((result) =>
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "")
    );
  });
}

function readComposeRecipientProperty(
  value?:
    | Office.EmailAddressDetails[]
    | {
        getAsync: (
          callback: (result: Office.AsyncResult<Office.EmailAddressDetails[]>) => void
        ) => void;
      }
): Promise<Office.EmailAddressDetails[]> {
  if (Array.isArray(value)) {
    return Promise.resolve(value);
  }
  return new Promise((resolve) => {
    if (!value || !value.getAsync) {
      resolve([]);
      return;
    }
    value.getAsync((result) =>
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || [] : [])
    );
  });
}

function isEmptyTicketCreationValue(value: unknown): boolean {
  return (
    value === undefined || value === null || value === "" || (Array.isArray(value) && !value.length)
  );
}

async function submitTicketCreation() {
  try {
    const creation = collectTicketCreationForm();
    if (getComposeMetadataItem()) {
      await selectComposeTicketCreation(creation);
    } else {
      await createTicketFromSelectedEmail(creation);
    }
  } catch (error) {
    setFailed(error, { hideLogout: false, message: "Ticket creation failed" });
  }
}

async function createTicketFromSelectedEmail(creation: TicketCreationFormValue) {
  const snapshot = await readCurrentOutlookEmailSnapshot();
  if (!snapshot) {
    throw createHaloAuthError("Open an existing received email before creating a ticket.");
  }
  const operationId = readTicketCreationOperationId || createComposeAttachId();
  readTicketCreationOperationId = operationId;
  let preparedAttachments: Awaited<ReturnType<typeof prepareEmailAttachmentsForHalo>> | null = null;
  setTicketCreationBusy(true);
  setTicketsBusy(true);
  try {
    setStatus("loading", "Preparing new Halo ticket...", creation.summary);
    await fetchJson("/api/halo/ticket-creation/intents", {
      method: "POST",
      body: JSON.stringify({ ...creation, actionMode: activeActionMode, operationId }),
    });
    const metadata = await snapshot.emailAttachments;
    const includeAttachments = metadata.attachments.length
      ? await promptForEmailAttachments(metadata)
      : false;
    if (includeAttachments === null || snapshot.itemRevision !== outlookItemRevision) {
      await removeServerTicketCreationIntent(operationId);
      readTicketCreationOperationId = "";
      return;
    }
    const attachmentPreparation = includeAttachments
      ? prepareEmailAttachmentsForHalo(
          Office.context.mailbox.item as unknown as ComposeMetadataItem,
          metadata,
          "0",
          operationId,
          operationId
        )
      : Promise.resolve(null);
    const [inlineImages, attachmentResult] = await Promise.all([
      snapshot.inlineImages.catch(() => emptyPreparedInlineImages()),
      attachmentPreparation,
    ]);
    preparedAttachments = attachmentResult;
    if (snapshot.itemRevision !== outlookItemRevision) {
      if (preparedAttachments?.emailAttachmentPrefetchKey) {
        void cancelEmailAttachmentPrefetch(preparedAttachments.emailAttachmentPrefetchKey);
      }
      await removeServerTicketCreationIntent(operationId);
      readTicketCreationOperationId = "";
      return;
    }
    let email = applyPreparedInlineImages(snapshot.email, inlineImages);
    email = preparedAttachments
      ? applyPreparedEmailAttachments(email, preparedAttachments)
      : applyEmailOnlyChoice(email, metadata);
    email.actionMode = activeActionMode;
    const result = await fetchJson<TicketCreationResult>("/api/halo/ticket-creation/from-email", {
      method: "POST",
      body: JSON.stringify({
        ...email,
        creation: { ...creation, actionMode: activeActionMode },
        operationId,
      }),
    });
    if (!result.ok) {
      throw createHaloAuthError(result.error || result.message || "Halo ticket creation failed.");
    }
    await saveBackgroundSessionId(result.backgroundSessionId || "");
    const attachmentWarning = hasEmailAttachmentWarnings(result.emailAttachments);
    const imageWarning = Boolean(result.inlineImages && result.inlineImages.failed);
    setStatus(
      attachmentWarning || imageWarning ? "warning" : "success",
      `Created Halo ticket ${result.ticketNumber || result.ticketId}`,
      attachmentWarning || imageWarning
        ? buildEmailImportWarningDetail(result.inlineImages, result.emailAttachments)
        : result.message ||
            (activeActionMode === "private-note"
              ? "The selected email was attached as a private note."
              : "The selected email was added to the new ticket.")
    );
    actionModeLocked = true;
    getPrivateNoteToggle().disabled = true;
    getMappedEmailConfirmation().hidden = true;
    mappedConversationTicket = {
      ticketId: String(result.ticketId || ""),
      ticketNumber: String(result.ticketNumber || result.ticketId || ""),
      actionMode: activeActionMode,
    };
    renderCreateTicketConversationWarning();
    readTicketCreationOperationId = "";
    getSubmitCreateTicketButton().textContent = "Create ticket";
  } catch (error) {
    getSubmitCreateTicketButton().textContent = "Retry ticket creation";
    const ticketNumber =
      error instanceof Error ? String((error as HaloAuthError).ticketNumber || "") : "";
    if (!ticketNumber && preparedAttachments?.emailAttachmentPrefetchKey) {
      void cancelEmailAttachmentPrefetch(preparedAttachments.emailAttachmentPrefetchKey);
    }
    if (ticketNumber) {
      setStatus(
        "warning",
        `Halo ticket ${ticketNumber} needs attention`,
        `The ticket was created, but its ${
          activeActionMode === "private-note" ? "Private Note" : "Email action"
        } was not completed. Select Retry ticket creation to repair it without creating another ticket.`
      );
      return;
    }
    if (error instanceof Error && (error as HaloAuthError).statusCode === 409) {
      await loadTicketCreationSchema(creation.typeId, true);
      getSubmitCreateTicketButton().textContent = "Review and create";
      setStatus(
        "warning",
        "Halo ticket details changed",
        "Review the refreshed required fields before trying ticket creation again."
      );
      return;
    }
    throw error;
  } finally {
    setTicketCreationBusy(false);
    setTicketsBusy(false);
  }
}

async function selectComposeTicketCreation(creation: TicketCreationFormValue) {
  const item = getComposeMetadataItem();
  if (!item) {
    return;
  }
  const previousMarker = activeComposeAttachMarker;
  const operationId = createComposeAttachId();
  const markerRevision = ++composeMarkerRevision;
  let intentSaved = false;
  let preparedAttachments: Awaited<ReturnType<typeof prepareEmailAttachmentsForHalo>> | null = null;
  setTicketCreationBusy(true);
  setTicketsBusy(true);
  try {
    setStatus("loading", "Saving ticket creation details...", creation.summary);
    const draftItemId = await enqueueComposeMetadataWrite(() => saveComposeItem(item));
    const bodyHtml = await getBodyAsync(item, Office.CoercionType.Html).catch(() => "");
    const metadata = await collectEmailAttachmentMetadata(item, bodyHtml).catch(() =>
      emptyEmailAttachmentMetadata()
    );
    let emailAttachmentDecision: "include" | "exclude" | undefined;
    if (metadata.attachments.length) {
      const includeAttachments = await promptForEmailAttachments(metadata);
      if (includeAttachments === null || markerRevision !== composeMarkerRevision) {
        return;
      }
      emailAttachmentDecision = includeAttachments ? "include" : "exclude";
    }
    await fetchJson("/api/halo/ticket-creation/intents", {
      method: "POST",
      body: JSON.stringify({
        ...creation,
        actionMode: activeActionMode,
        draftItemId,
        emailAttachmentDecision,
        emailAttachmentFingerprint: metadata.emailAttachmentFingerprint,
        operationId,
      }),
    });
    intentSaved = true;
    preparedAttachments =
      emailAttachmentDecision === "include"
        ? await prepareEmailAttachmentsForHalo(item, metadata, "0", operationId, operationId)
        : null;
    if (markerRevision !== composeMarkerRevision) {
      if (preparedAttachments?.emailAttachmentPrefetchKey) {
        void cancelEmailAttachmentPrefetch(preparedAttachments.emailAttachmentPrefetchKey);
      }
      await removeServerTicketCreationIntent(operationId);
      return;
    }
    if (preparedAttachments) {
      await fetchJson(`/api/halo/ticket-creation/intents/${encodeURIComponent(operationId)}`, {
        method: "POST",
        body: JSON.stringify({
          emailAttachmentDecision,
          emailAttachmentFingerprint: preparedAttachments.emailAttachmentFingerprint,
          emailAttachmentPrefetchKey: preparedAttachments.emailAttachmentPrefetchKey,
          emailAttachmentStagingVersion: 2,
          emailAttachmentSummary: preparedAttachments.emailAttachmentSummary,
        }),
      });
    }
    const marker: ComposeAttachMarker = {
      version: 4,
      destinationKind: "create-ticket",
      composeAttachId: operationId,
      creationOperationId: operationId,
      ticketId: "",
      ticketNumber: `New ${activeTicketCreationSchema?.typeName || "ticket"}`,
      ticketSummary: creation.summary.slice(0, 500),
      ticketTypeId: creation.typeId,
      ticketTypeName: activeTicketCreationSchema?.typeName || "",
      draftItemId,
      actionMode: activeActionMode,
      composeIdentityHeader: true,
      emailAttachmentDecision,
      emailAttachmentFingerprint:
        preparedAttachments?.emailAttachmentFingerprint || metadata.emailAttachmentFingerprint,
      emailAttachmentPrefetchKey: preparedAttachments?.emailAttachmentPrefetchKey,
      emailAttachmentStagingVersion: 2,
      emailAttachmentSummary: preparedAttachments?.emailAttachmentSummary,
    };
    await enqueueComposeMetadataWrite(async () => {
      await setComposeIdentityHeader(item, marker.composeAttachId);
      await saveComposeSessionMarker(item, marker);
      await saveComposeCustomMarker(item, marker);
      const persistedItemId = await saveComposeItem(item);
      if (persistedItemId !== marker.draftItemId) {
        marker.draftItemId = persistedItemId;
        await saveComposeSessionMarker(item, marker);
        await saveComposeCustomMarker(item, marker);
        await saveComposeItem(item);
      }
    });
    activeComposeAttachMarker = marker;
    renderComposeAttachSelection();
    setStatus(
      isComposeAttachmentReady(marker) ? "success" : "failed",
      isComposeAttachmentReady(marker)
        ? `Will create ${marker.ticketTypeName} when sent`
        : "Attachment preparation failed — retry before sending",
      isComposeAttachmentReady(marker)
        ? "The final subject and recipients will be checked again when the email is sent."
        : "Open the Halo pane and let preparation finish, then try Send again."
    );
    if (preparedAttachments?.emailAttachmentPrefetchKey) {
      await saveComposeEmailAttachmentState(item, {
        version: 2,
        draftItemId: marker.draftItemId,
        ticketId: "0",
        operationId,
        emailAttachmentDecision: "include",
        emailAttachmentFingerprint: preparedAttachments.emailAttachmentFingerprint,
        emailAttachmentPrefetchKey: preparedAttachments.emailAttachmentPrefetchKey,
        emailAttachmentStagingVersion: 2,
        emailAttachmentSummary: preparedAttachments.emailAttachmentSummary,
      }).catch(() => undefined);
    }
    if (previousMarker) {
      if (isCreateTicketMarker(previousMarker)) {
        void removeServerTicketCreationIntent(
          previousMarker.creationOperationId || previousMarker.composeAttachId
        );
      }
      if (previousMarker.emailAttachmentPrefetchKey) {
        void cancelEmailAttachmentPrefetch(previousMarker.emailAttachmentPrefetchKey);
      }
    }
  } catch (error) {
    if (preparedAttachments?.emailAttachmentPrefetchKey) {
      void cancelEmailAttachmentPrefetch(preparedAttachments.emailAttachmentPrefetchKey);
    }
    if (intentSaved) {
      await removeServerTicketCreationIntent(operationId);
    }
    if (markerRevision === composeMarkerRevision) {
      if (previousMarker) {
        await enqueueComposeMetadataWrite(async () => {
          await restoreComposeIdentityHeader(item, previousMarker);
          await saveComposeSessionMarker(item, previousMarker);
          await saveComposeCustomMarker(item, previousMarker);
          await saveComposeItem(item);
        }).catch(() => undefined);
        activeComposeAttachMarker = previousMarker;
      } else {
        await enqueueComposeMetadataWrite(async () => {
          await setComposeIdentityHeader(item, "");
          await clearComposeAttachMetadata(item);
          await clearComposeEmailAttachmentState(item);
          await saveComposeItem(item);
        }).catch(() => undefined);
        activeComposeAttachMarker = null;
      }
      renderComposeAttachSelection();
    }
    throw error;
  } finally {
    setTicketCreationBusy(false);
    setTicketsBusy(false);
  }
}

async function removeServerTicketCreationIntent(operationId: string): Promise<void> {
  if (!operationId) {
    return;
  }
  await fetchJson(`/api/halo/ticket-creation/intents/${encodeURIComponent(operationId)}`, {
    method: "DELETE",
    body: JSON.stringify({}),
  }).catch(() => undefined);
}

function isCreateTicketMarker(marker: ComposeAttachMarker | null): boolean {
  return Boolean(marker && marker.version === 4 && marker.destinationKind === "create-ticket");
}

async function searchTickets(options: { allowShortQuery?: boolean } = {}) {
  const query = getTicketQueryInput().value.trim();
  updateClearSearchButton();

  if (!query) {
    await loadTickets();
    return;
  }

  if (!options.allowShortQuery && query.length < 2) {
    setStatus(
      "success",
      "Keep typing or press Enter",
      "Type at least two characters for automatic search, or press Enter to search now."
    );
    return;
  }

  await requestTickets(query);
}

function scheduleTicketSearch() {
  cancelTicketSearchDebounce();
  ticketRequestRevision += 1;
  setTicketQueryBusy(false);
  updateClearSearchButton();

  const query = getTicketQueryInput().value.trim();
  if (query.length === 1) {
    void searchTickets();
    return;
  }

  ticketSearchDebounceId = window.setTimeout(() => {
    ticketSearchDebounceId = null;
    void searchTickets();
  }, 350);
}

function cancelTicketSearchDebounce() {
  if (ticketSearchDebounceId !== null) {
    window.clearTimeout(ticketSearchDebounceId);
    ticketSearchDebounceId = null;
  }
}

async function clearTicketSearch() {
  cancelTicketSearchDebounce();
  getTicketQueryInput().value = "";
  updateClearSearchButton();
  await loadTickets();
  getTicketQueryInput().focus();
}

async function resetTicketFilters() {
  getTicketOwnershipSelect().value = "mine";
  getTicketLifecycleSelect().value = "open";
  getTicketCustomerSelect().value = "";
  getTicketAssigneeSelect().value = "";
  updateActiveFilterCount();
  await requestTickets(getTicketQueryInput().value.trim());
}

async function requestTickets(query: string) {
  const ownership = getTicketOwnershipSelect().value as TicketOwnership;
  const lifecycle = getTicketLifecycleSelect().value as TicketLifecycle;
  const params = new URLSearchParams({ ownership, lifecycle });
  const endpoint = query ? "/api/halo/tickets/search" : "/api/halo/tickets";
  if (query) {
    params.set("query", query);
  }

  const requestRevision = ++ticketRequestRevision;
  getTicketsPanel().hidden = false;
  getTicketsEmpty().hidden = true;
  updateActiveFilterCount();

  try {
    setTicketQueryBusy(true);
    setStatus(
      "loading",
      query ? `Searching for “${query}”...` : "Loading Halo tickets...",
      describeTicketScope(ownership, lifecycle)
    );

    const result = await fetchJson<HaloTicketsResponse>(`${endpoint}?${params.toString()}`);
    if (requestRevision !== ticketRequestRevision) {
      return;
    }
    if (!result.ok) {
      throw createHaloAuthError(
        result.error || result.message || "Halo ticket search failed.",
        result.debug
      );
    }

    loadedTickets = result.tickets || [];
    lastExecutedTicketQuery = query;
    updateTicketFacetOptions(loadedTickets);
    const displayedCount = applyTicketFilters();

    if (activeComposeAttachMarker) {
      setStatus(
        "success",
        `Will attach to ${activeComposeAttachMarker.ticketNumber} when sent`,
        `${displayedCount} ticket(s) shown. Select another ticket to replace it.`
      );
    } else if (displayedCount) {
      setStatus(
        "success",
        displayedCount === 1 ? "1 ticket found" : `${displayedCount} tickets found`,
        "Select a ticket below to attach the open email."
      );
    } else {
      setStatus(
        "signed-out",
        "No matching tickets",
        loadedTickets.length
          ? "Adjust the customer or assignee filter and try again."
          : "Try another search or broaden the ownership and lifecycle filters."
      );
    }
  } catch (error) {
    if (requestRevision !== ticketRequestRevision) {
      return;
    }
    setFailed(error, {
      hideLogout: false,
      message: query ? "Ticket search failed" : "Halo ticket list failed",
    });
  } finally {
    if (requestRevision === ticketRequestRevision) {
      setTicketQueryBusy(false);
    }
  }
}

async function autoAttachCurrentEmail(): Promise<boolean> {
  let snapshot: OutlookEmailSnapshot | null = null;

  try {
    snapshot = await readCurrentOutlookEmailSnapshot({ suppressUnsupported: true });
  } catch {
    return false;
  }

  if (!snapshot) {
    return false;
  }

  try {
    setStatus(
      "loading",
      "Checking Halo email mapping...",
      "Looking for an existing ticket link for this email chain."
    );

    let match = await fetchJson<HaloEmailMatchResponse>("/api/halo/email/match", {
      method: "POST",
      body: JSON.stringify({
        conversationId: snapshot.email.conversationId,
        inReplyToMessageIds: snapshot.email.inReplyToMessageIds,
        internetMessageId: snapshot.email.internetMessageId,
        itemId: snapshot.email.itemId,
        mailboxEmail: snapshot.email.mailboxEmail,
        referenceMessageIds: snapshot.email.referenceMessageIds,
      }),
    });

    if (!match.ok) {
      throw createHaloAuthError(
        match.error || match.message || "Email mapping lookup failed.",
        match.debug
      );
    }

    if (snapshot.itemRevision !== outlookItemRevision) {
      return true;
    }

    if (match.status === "no-match") {
      setStatus(
        "loading",
        "Recovering Halo email mapping...",
        "Checking exact message references against the original sent email."
      );
      const composeAttachIds = await recoverReadModeComposeAttachIds(snapshot.email);
      if (snapshot.itemRevision !== outlookItemRevision) {
        return true;
      }
      if (composeAttachIds.length) {
        match = await fetchJson<HaloEmailMatchResponse>("/api/halo/email/match", {
          method: "POST",
          body: JSON.stringify({
            composeAttachIds,
            conversationId: snapshot.email.conversationId,
            inReplyToMessageIds: snapshot.email.inReplyToMessageIds,
            internetMessageId: snapshot.email.internetMessageId,
            itemId: snapshot.email.itemId,
            mailboxEmail: snapshot.email.mailboxEmail,
            referenceMessageIds: snapshot.email.referenceMessageIds,
          }),
        });
        if (!match.ok) {
          throw createHaloAuthError(
            match.error || match.message || "Email mapping recovery failed.",
            match.debug
          );
        }
      }
      if (match.status === "no-match") {
        mappedConversationTicket = null;
        renderCreateTicketConversationWarning();
        return false;
      }
    }

    mappedConversationTicket = {
      ticketId: String(match.ticketId || ""),
      ticketNumber: String(match.ticketNumber || match.ticketId || ""),
      actionMode: normalizeActionMode(match.actionMode),
    };
    setActionMode(mappedConversationTicket.actionMode, {
      disabled: match.status === "already-attached",
    });
    renderCreateTicketConversationWarning();
    clearTickets();
    setConnectionState(true);

    if (match.status === "already-attached") {
      setStatus(
        "success",
        "This email is already attached to ticket",
        match.message ||
          `This email is already attached to ticket ${match.ticketNumber || match.ticketId}.`
      );
      getTicketsPanel().hidden = false;
      renderMappedEmailConfirmation();
      getMappedEmailConfirmation().hidden = true;
      return true;
    }

    setStatus(
      "success",
      `Mapped to ticket ${mappedConversationTicket.ticketNumber}`,
      "Review whether this email should be visible to the customer, then attach it."
    );
    getTicketsPanel().hidden = false;
    renderMappedEmailConfirmation();
    return true;
  } catch (error) {
    clearTickets();
    setFailed(error, { hideLogout: false, message: "Email mapping lookup failed" });
    return true;
  }
}

async function attachMappedCurrentEmail(): Promise<void> {
  const target = mappedConversationTicket;
  if (!target) {
    return;
  }
  const snapshot = await readCurrentOutlookEmailSnapshot({ suppressUnsupported: true });
  if (!snapshot) {
    return;
  }

  try {
    setTicketsBusy(true);
    setStatus("loading", "Attaching mapped email...", `Ticket ${target.ticketNumber}`);
    const metadata = await snapshot.emailAttachments;
    const includeAttachments = metadata.attachments.length
      ? await promptForEmailAttachments(metadata)
      : false;
    if (includeAttachments === null || snapshot.itemRevision !== outlookItemRevision) {
      return;
    }
    const attachmentPreparation = includeAttachments
      ? prepareEmailAttachmentsForHalo(
          Office.context.mailbox.item as unknown as ComposeMetadataItem,
          metadata,
          target.ticketId,
          createComposeAttachId()
        )
      : Promise.resolve(null);
    const [inlineImages, preparedAttachments] = await Promise.all([
      snapshot.inlineImages.catch(() => emptyPreparedInlineImages()),
      attachmentPreparation,
    ]);
    if (snapshot.itemRevision !== outlookItemRevision) {
      if (preparedAttachments?.emailAttachmentPrefetchKey) {
        void cancelEmailAttachmentPrefetch(preparedAttachments.emailAttachmentPrefetchKey);
      }
      return;
    }
    let email = applyPreparedInlineImages(snapshot.email, inlineImages);
    email = preparedAttachments
      ? applyPreparedEmailAttachments(email, preparedAttachments)
      : applyEmailOnlyChoice(email, metadata);
    email.actionMode = activeActionMode;
    const result = await fetchJson<HaloAutoAttachResponse>("/api/halo/email/auto-attach", {
      method: "POST",
      body: JSON.stringify(email),
    });

    if (!result.ok) {
      throw createHaloAuthError(
        result.error || result.message || "Email auto-attach failed.",
        result.debug
      );
    }

    const imageWarning = Boolean(result.inlineImages && result.inlineImages.failed > 0);
    const attachmentWarning = hasEmailAttachmentWarnings(result.emailAttachments);
    setStatus(
      imageWarning || attachmentWarning ? "warning" : "success",
      imageWarning || attachmentWarning
        ? "Email added to ticket with warnings"
        : activeActionMode === "private-note"
          ? "Email attached as a private note"
          : "Email attached as an Email action",
      imageWarning || attachmentWarning
        ? buildEmailImportWarningDetail(result.inlineImages, result.emailAttachments)
        : result.message || `Email added to ticket ${result.ticketNumber || result.ticketId}.`
    );
    mappedConversationTicket.actionMode = activeActionMode;
    actionModeLocked = true;
    getPrivateNoteToggle().disabled = true;
    getMappedEmailConfirmation().hidden = true;
    getTicketsPanel().hidden = false;
  } catch (error) {
    setFailed(error, { hideLogout: false, message: "Email auto-attach failed" });
  } finally {
    setTicketsBusy(false);
  }
}

async function logout() {
  try {
    setBusy(true);
    await fetchJson("/api/auth/logout", {
      method: "POST",
      body: JSON.stringify({}),
    });
    await clearBackgroundSessionId();
    clearTickets();
    setSignedOut();
  } catch (error) {
    setFailed(error);
  } finally {
    setBusy(false);
  }
}

function closeDialog() {
  const dialog = currentDialog;
  currentDialog = null;
  waitingForDialog = false;

  if (dialog) {
    dialog.close();
  }
}

function registerItemChangedHandler() {
  const mailbox = Office.context.mailbox as unknown as {
    addHandlerAsync?: (
      eventType: Office.EventType,
      handler: () => void,
      callback?: (result: Office.AsyncResult<void>) => void
    ) => void;
  };

  if (!mailbox.addHandlerAsync || !Office.EventType.ItemChanged) {
    return;
  }

  mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
    cancelAttachmentPrompt();
    clearInlineImagePreparationCache();
    outlookItemRevision += 1;
    composeMarkerRevision += 1;
    composeInlineImagePrefetchRevision += 1;
    activeComposeAttachMarker = null;
    ticketCreationContext = null;
    activeTicketCreationRequester = null;
    ticketCreationRequesters = [];
    readTicketCreationOperationId = "";
    mappedConversationTicket = null;
    setActionMode("email");
    renderMappedEmailConfirmation();
    renderCreateTicketConversationWarning();
    renderComposeAttachSelection();
    registerComposeAttachmentsChangedHandler();
    void checkExistingSession();
  });
}

function registerComposeAttachmentsChangedHandler() {
  const item = getComposeMetadataItem() as
    | (ComposeMetadataItem & {
        attachments?: unknown[];
        addHandlerAsync?: (
          eventType: Office.EventType | string,
          handler: () => void,
          callback?: (result: Office.AsyncResult<void>) => void
        ) => void;
      })
    | null;
  if (!item || !item.addHandlerAsync) {
    return;
  }
  item.addHandlerAsync(Office.EventType.AttachmentsChanged || "attachmentsChanged", () => {
    clearInlineImagePreparationCache();
    composeMarkerRevision += 1;
    void reconcileComposeEmailAttachments();
  });
}

function promptForEmailAttachments(
  metadata: PreparedEmailAttachmentMetadata
): Promise<boolean | null> {
  if (!metadata.attachments.length) {
    return Promise.resolve(false);
  }
  cancelAttachmentPrompt();
  const prompt = getAttachmentPrompt();
  const count = metadata.attachments.length;
  getAttachmentPromptDetail().textContent = `This email contains ${count} non-inline attachment${
    count === 1 ? "" : "s"
  } (${formatByteCount(metadata.reportedTotalBytes)} reported). Add eligible files to the Halo ticket action?`;
  const limits: string[] = [];
  if (metadata.reportedOversized) {
    limits.push(
      `${metadata.reportedOversized} exceed${metadata.reportedOversized === 1 ? "s" : ""} 25 MiB`
    );
  }
  if (metadata.overCount) {
    limits.push(`${metadata.overCount} exceed the 20 attachment limit`);
  }
  if (metadata.reportedTotalBytes > MAX_EMAIL_ATTACHMENT_TOTAL_BYTES) {
    limits.push("the reported total exceeds 50 MiB");
  }
  getAttachmentPromptLimits().textContent = limits.length
    ? `${limits.join("; ")}. Those files will be skipped.`
    : "Up to 25 MiB per file and 50 MiB total will be added.";
  prompt.hidden = false;
  const addButton = getAttachmentAddButton();
  const emailOnlyButton = getAttachmentEmailOnlyButton();
  const previousFocus = document.activeElement as HTMLElement | null;

  return new Promise((resolve) => {
    activeAttachmentPromptResolve = resolve;
    const finish = (choice: boolean | null) => {
      if (activeAttachmentPromptResolve !== resolve) {
        return;
      }
      activeAttachmentPromptResolve = null;
      prompt.hidden = true;
      addButton.onclick = null;
      emailOnlyButton.onclick = null;
      prompt.onkeydown = null;
      if (previousFocus && previousFocus.focus) {
        previousFocus.focus();
      }
      resolve(choice);
    };
    addButton.onclick = () => finish(true);
    emailOnlyButton.onclick = () => finish(false);
    prompt.onkeydown = (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        finish(null);
      } else if (event.key === "Tab") {
        if (event.shiftKey && document.activeElement === emailOnlyButton) {
          event.preventDefault();
          addButton.focus();
        } else if (!event.shiftKey && document.activeElement === addButton) {
          event.preventDefault();
          emailOnlyButton.focus();
        }
      }
    };
    addButton.focus();
  });
}

function cancelAttachmentPrompt() {
  const resolve = activeAttachmentPromptResolve;
  activeAttachmentPromptResolve = null;
  const prompt = document.getElementById("attachment-prompt");
  if (prompt) {
    prompt.hidden = true;
  }
  if (resolve) {
    resolve(null);
  }
}

function formatByteCount(value: number): string {
  if (!value) {
    return "unknown size";
  }
  if (value >= 1024 * 1024) {
    return `${(value / (1024 * 1024)).toFixed(1)} MiB`;
  }
  return `${Math.max(1, Math.round(value / 1024))} KiB`;
}

function sendEmailAttachmentClientDiagnostic(
  stage: "attachment-read-complete" | "attachment-read-retry" | "attachment-prefetch-complete",
  details: {
    attachmentCount: number;
    attachmentError?: string;
    attemptCount?: number;
    elapsedMs: number;
    failedCount?: number;
    outcome: "failed" | "ok" | "started";
    skippedCount?: number;
    uploadedCount?: number;
  }
): void {
  void sendJsonRequest(
    "/api/diagnostics/send-event",
    {
      method: "POST",
      body: JSON.stringify({ stage, ...details }),
    },
    ""
  ).catch(() => undefined);
}

async function fetchJson<T>(
  url: string,
  options: RequestInit = {},
  authOptions: { allowMissingAuth?: boolean; interactive?: boolean } = {}
): Promise<T> {
  let authHeader = await getMicrosoftAuthHeader(authOptions);
  let response = await sendJsonRequest(url, options, authHeader);
  let body = await response.json().catch(() => ({}));

  if (shouldRefreshMicrosoftToken(response.status, body)) {
    authHeader = await getMicrosoftAuthHeader({ ...authOptions, forceRefresh: true });
    response = await sendJsonRequest(url, options, authHeader);
    body = await response.json().catch(() => ({}));
  }

  if (!response.ok) {
    const error = createHaloAuthError(
      body.error || body.message || `Request failed with status ${response.status}.`,
      body.debug
    );
    error.operationId = String(body.operationId || "");
    error.statusCode = response.status;
    error.ticketNumber = String(body.ticketNumber || "");
    throw error;
  }

  return body as T;
}

function sendJsonRequest(url: string, options: RequestInit, authHeader: string): Promise<Response> {
  return fetch(url, {
    credentials: "same-origin",
    ...options,
    headers: {
      Accept: "application/json",
      ...(authHeader ? { Authorization: authHeader } : {}),
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
}

function shouldRefreshMicrosoftToken(status: number, body: unknown): boolean {
  const error =
    body && typeof body === "object" && "error" in body
      ? String((body as { error?: unknown }).error || "")
      : "";

  return status === 401 && /Microsoft add-in authentication failed/i.test(error);
}

async function getMicrosoftAuthHeader(
  options: {
    allowMissingAuth?: boolean;
    forceRefresh?: boolean;
    interactive?: boolean;
  } = {}
): Promise<string> {
  if (options.allowMissingAuth && !(await isSsoEnabled())) {
    return "";
  }

  try {
    const token = await acquireMicrosoftToken(
      options.interactive !== false,
      options.forceRefresh === true
    );
    return token ? `Bearer ${token}` : "";
  } catch (error) {
    if (options.allowMissingAuth) {
      return "";
    }

    throw error;
  }
}

async function isSsoEnabled(): Promise<boolean> {
  const config = await getAuthConfig();
  return config.ssoEnabled;
}

async function acquireMicrosoftToken(interactive: boolean, forceRefresh = false): Promise<string> {
  const config = await getAuthConfig();

  if (!config.ssoEnabled || !config.clientId || !config.scopes.length) {
    return "";
  }

  const msalInstance = (await getMsalInstance()) as {
    acquireTokenPopup: (request: unknown) => Promise<{ accessToken?: string }>;
    ssoSilent: (request: unknown) => Promise<{ accessToken?: string }>;
  };
  const request = {
    forceRefresh,
    scopes: config.scopes,
    loginHint: await getLoginHint(),
  };

  try {
    const result = await msalInstance.ssoSilent(request);
    return requireMicrosoftAccessToken(result);
  } catch (error) {
    if (!interactive || !(error instanceof InteractionRequiredAuthError)) {
      throw error;
    }

    const result = await msalInstance.acquireTokenPopup(request);
    return requireMicrosoftAccessToken(result);
  }
}

function requireMicrosoftAccessToken(result: { accessToken?: string }): string {
  if (!result.accessToken) {
    throw new Error("Microsoft did not return an access token for the add-in API.");
  }

  return result.accessToken;
}

async function getMsalInstance(): Promise<unknown> {
  if (msalInstancePromise) {
    return msalInstancePromise;
  }

  msalInstancePromise = getAuthConfig().then((config) =>
    createNestablePublicClientApplication({
      auth: {
        authority: config.authority,
        clientId: config.clientId,
      },
      cache: {
        cacheLocation: "localStorage",
      },
    })
  );

  return msalInstancePromise;
}

async function getAuthConfig(): Promise<AuthConfigResponse> {
  if (!authConfigPromise) {
    authConfigPromise = fetch("/api/auth/config", {
      credentials: "same-origin",
      headers: {
        Accept: "application/json",
      },
    }).then((response) => response.json());
  }

  return authConfigPromise;
}

async function getLoginHint(): Promise<string | undefined> {
  try {
    const officeAuth = (
      Office as unknown as {
        auth?: {
          getAuthContext?: () => Promise<{ userPrincipalName?: string }>;
        };
      }
    ).auth;
    const authContext =
      officeAuth && officeAuth.getAuthContext ? await officeAuth.getAuthContext() : null;
    return authContext && authContext.userPrincipalName ? authContext.userPrincipalName : undefined;
  } catch {
    return undefined;
  }
}

function setSignedOut() {
  setConnectionState(false);
  clearTickets();
  setStatus("signed-out", "Connect to HaloPSA to start", "Sign in with your HaloPSA account.");
}

function setFailed(error: unknown, options: { hideLogout?: boolean; message?: string } = {}) {
  if (options.hideLogout !== false) {
    setConnectionState(false);
  }

  const detail = error instanceof Error ? error.message : "Unexpected Halo auth error.";
  setStatus("failed", options.message || "Halo API Auth failed", detail, getErrorDebug(error));
}

function hasEmailAttachmentWarnings(summary?: EmailAttachmentSummary): boolean {
  return Boolean(summary && (summary.failed > 0 || summary.skipped > 0));
}

function buildEmailImportWarningDetail(
  inlineImages?: InlineImageSummary,
  emailAttachments?: EmailAttachmentSummary
): string {
  const details: string[] = [];
  if (inlineImages && inlineImages.failed) {
    details.push(`${inlineImages.failed} inline image(s) unavailable`);
  }
  if (emailAttachments && emailAttachments.failed) {
    details.push(`${emailAttachments.failed} attachment(s) unavailable`);
  }
  if (emailAttachments && emailAttachments.skipped) {
    details.push(`${emailAttachments.skipped} attachment(s) skipped by limits or format`);
  }
  return `${details.join("; ")}. The email was still attached.`;
}

function setStatus(
  state: "signed-out" | "loading" | "success" | "warning" | "failed",
  message: string,
  detail: string,
  debug?: unknown
) {
  const statusCard = document.getElementById("status-card");
  const statusMessage = document.getElementById("status-message");
  const statusDetail = document.getElementById("status-detail");
  const debugDetail = document.getElementById("debug-detail");

  if (statusCard) {
    statusCard.dataset.state = state;
  }

  if (statusMessage) {
    statusMessage.textContent = message;
  }

  if (statusDetail) {
    statusDetail.textContent = detail;
  }

  if (debugDetail) {
    const renderedDebug = renderDebug(debug);
    debugDetail.textContent = renderedDebug;
    debugDetail.hidden = !renderedDebug;
  }
}

function setBusy(isBusy: boolean) {
  appIsBusy = isBusy;
  getLoginButton().disabled = isBusy;
  getLogoutButton().disabled = isBusy;
  getReportBugButton().disabled = isBusy;
  getRemoveComposeSelectionButton().disabled = isBusy;
  getAttachExistingTab().disabled = isBusy;
  getCreateTicketTab().disabled = isBusy;
  updateTicketBusyState();
}

function setTicketsBusy(isBusy: boolean) {
  if (isBusy) {
    cancelTicketSearchDebounce();
    ticketRequestRevision += 1;
    ticketQueryIsBusy = false;
  }
  ticketActionIsBusy = isBusy;
  updateTicketBusyState();
}

function setTicketQueryBusy(isBusy: boolean) {
  ticketQueryIsBusy = isBusy;
  updateTicketBusyState();
}

function updateTicketBusyState() {
  const interactionBlocked = appIsBusy || ticketActionIsBusy;
  const resultsBlocked = interactionBlocked || ticketQueryIsBusy;
  getRefreshTicketsButton().disabled = resultsBlocked;
  getSearchTicketsButton().disabled = resultsBlocked;
  getClearSearchButton().disabled = interactionBlocked;
  getTicketQueryInput().disabled = interactionBlocked;
  getTicketOwnershipSelect().disabled = interactionBlocked;
  getTicketLifecycleSelect().disabled = interactionBlocked;
  getTicketCustomerSelect().disabled =
    interactionBlocked || getTicketCustomerSelect().options.length <= 1;
  getTicketAssigneeSelect().disabled =
    interactionBlocked || getTicketAssigneeSelect().options.length <= 1;
  getResetTicketFiltersButton().disabled = interactionBlocked;
  getPrivateNoteToggle().disabled = interactionBlocked || actionModeLocked;
  getAttachMappedEmailButton().disabled = interactionBlocked;
  getTicketsPanel().setAttribute("aria-busy", String(resultsBlocked));
  setTicketButtonsBusy(resultsBlocked);
}

function setConnectionState(isConnected: boolean) {
  getConnectionPanel().hidden = isConnected;
  getLogoutButton().hidden = !isConnected;
  const appBody = document.getElementById("app-body");
  if (appBody) {
    appBody.dataset.connected = isConnected ? "true" : "false";
  }
}

function setTicketButtonsBusy(isBusy: boolean) {
  const ticketButtons = document.querySelectorAll(".halo-ticket");
  for (let index = 0; index < ticketButtons.length; index += 1) {
    (ticketButtons[index] as HTMLButtonElement).disabled = isBusy;
  }
}

function saveBackgroundSessionId(backgroundSessionId: string): Promise<void> {
  if (!backgroundSessionId) {
    return Promise.resolve();
  }

  activeBackgroundSessionId = backgroundSessionId;
  if (!Office.context.roamingSettings) {
    return Promise.resolve();
  }

  try {
    Office.context.roamingSettings.set(BACKGROUND_SESSION_STORAGE_KEY, backgroundSessionId);
  } catch {
    return Promise.resolve();
  }
  return saveRoamingSettings();
}

async function refreshBackgroundSessionId(): Promise<void> {
  try {
    const result = await fetchJson<{ backgroundSessionId?: string; ok?: boolean }>(
      "/api/auth/background-session",
      {
        method: "POST",
        body: JSON.stringify({}),
      }
    );
    await saveBackgroundSessionId(result.backgroundSessionId || "");
  } catch {
    // The task pane can still use the normal cookie session; send events will no-op without a handle.
  }
}

function clearBackgroundSessionId(): Promise<void> {
  activeBackgroundSessionId = "";
  if (!Office.context.roamingSettings) {
    return Promise.resolve();
  }

  Office.context.roamingSettings.remove(BACKGROUND_SESSION_STORAGE_KEY);
  return saveRoamingSettings();
}

function saveRoamingSettings(): Promise<void> {
  return new Promise((resolve) => {
    Office.context.roamingSettings.saveAsync(() => resolve());
  });
}

function createHaloAuthError(message: string, debug?: unknown): HaloAuthError {
  const error = new Error(message) as HaloAuthError;
  error.debug = debug;
  return error;
}

function getErrorDebug(error: unknown): unknown {
  return error instanceof Error ? (error as HaloAuthError).debug : null;
}

function renderDebug(debug: unknown): string {
  if (!debug) {
    return "";
  }

  if (typeof debug === "string") {
    return debug;
  }

  try {
    return JSON.stringify(debug, null, 2);
  } catch {
    return String(debug);
  }
}

function applyTicketFilters(): number {
  const customer = getTicketCustomerSelect().value;
  const assignee = getTicketAssigneeSelect().value;
  const tickets = loadedTickets.filter((ticket) => {
    return (!customer || ticket.client === customer) && (!assignee || ticket.agent === assignee);
  });

  getTicketsPanel().hidden = false;
  clearTicketList();
  getTicketsEmpty().hidden = tickets.length > 0;
  getTicketsEmpty().textContent = getTicketEmptyMessage(tickets.length);
  getTicketResultsHeading().textContent = getTicketResultsTitle();
  getTicketResultsCount().textContent = formatTicketResultCount(
    tickets.length,
    loadedTickets.length
  );
  renderTicketList(tickets, getTicketsList());
  updateActiveFilterCount();
  updateTicketBusyState();
  return tickets.length;
}

function onTicketFacetFilterChanged() {
  const displayedCount = applyTicketFilters();
  setStatus(
    displayedCount ? "success" : "signed-out",
    displayedCount === 1 ? "1 ticket shown" : `${displayedCount} tickets shown`,
    displayedCount
      ? "Select a ticket below to attach the open email."
      : "Adjust the customer or assignee filter and try again."
  );
}

function formatTicketResultCount(displayedCount: number, totalCount: number): string {
  if (displayedCount === 1 && totalCount === 1) {
    return "1 ticket shown";
  }
  return `${displayedCount} of ${totalCount} tickets shown`;
}

function updateTicketFacetOptions(tickets: HaloTicket[]) {
  replaceTicketFilterOptions(
    getTicketCustomerSelect(),
    tickets.map((ticket) => ticket.client),
    "Any customer"
  );
  replaceTicketFilterOptions(
    getTicketAssigneeSelect(),
    tickets.map((ticket) => ticket.agent),
    "Any assignee"
  );
  updateTicketBusyState();
}

function replaceTicketFilterOptions(
  select: HTMLSelectElement,
  rawValues: string[],
  allLabel: string
) {
  const selectedValue = select.value;
  const values = Array.from(new Set(rawValues.filter(Boolean))).sort((left, right) =>
    left.localeCompare(right, undefined, { sensitivity: "base" })
  );

  while (select.firstChild) {
    select.removeChild(select.firstChild);
  }

  const allOption = document.createElement("option");
  allOption.value = "";
  allOption.textContent = allLabel;
  select.appendChild(allOption);

  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });

  select.value = values.includes(selectedValue) ? selectedValue : "";
  select.disabled = values.length === 0;
}

function getTicketResultsTitle(): string {
  if (lastExecutedTicketQuery) {
    return "Search results";
  }

  const ownership = getTicketOwnershipSelect().value as TicketOwnership;
  const lifecycle = getTicketLifecycleSelect().value as TicketLifecycle;
  const ownerLabel = ownership === "mine" ? "My" : "All";
  if (lifecycle === "all") {
    return `${ownerLabel} tickets`;
  }
  return `${ownerLabel} ${lifecycle} tickets`;
}

function getTicketEmptyMessage(displayedCount: number): string {
  if (displayedCount) {
    return "";
  }
  if (loadedTickets.length) {
    return "No loaded tickets match the selected customer and assignee.";
  }
  if (lastExecutedTicketQuery) {
    return `No tickets matched “${lastExecutedTicketQuery}”. Try different words or broader filters.`;
  }
  return "No tickets match the selected ownership and lifecycle.";
}

function updateActiveFilterCount() {
  let count = 0;
  if (getTicketOwnershipSelect().value !== "mine") {
    count += 1;
  }
  if (getTicketLifecycleSelect().value !== "open") {
    count += 1;
  }
  if (getTicketCustomerSelect().value) {
    count += 1;
  }
  if (getTicketAssigneeSelect().value) {
    count += 1;
  }

  const badge = getActiveFilterCount();
  badge.textContent = String(count);
  badge.hidden = count === 0;
}

function updateClearSearchButton() {
  getClearSearchButton().hidden = !getTicketQueryInput().value.trim();
}

function describeTicketScope(ownership: TicketOwnership, lifecycle: TicketLifecycle): string {
  const owner = ownership === "mine" ? "your" : "all accessible";
  const state = lifecycle === "all" ? "open and closed" : lifecycle;
  return `Fetching ${owner} ${state} tickets from HaloPSA.`;
}

function renderTicketList(tickets: HaloTicket[], container: HTMLElement) {
  const composeMode = Boolean(getComposeMetadataItem());

  tickets.forEach((ticket) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "halo-ticket";
    button.dataset.ticketId = ticket.id;
    button.dataset.selected = String(
      Boolean(activeComposeAttachMarker && activeComposeAttachMarker.ticketId === ticket.id)
    );
    button.setAttribute(
      "aria-label",
      composeMode
        ? `Attach email to ${formatTicketTitle(ticket)} when sent`
        : `Attach email to ${formatTicketTitle(ticket)}`
    );
    button.onclick = () => void attachEmailToTicket(ticket, button);

    const main = document.createElement("span");
    main.className = "halo-ticket__main";

    const number = document.createElement("span");
    number.className = "halo-ticket__number";
    number.textContent = ticket.ticketNumber || ticket.id || "Ticket";

    const summary = document.createElement("span");
    summary.className = "halo-ticket__summary";
    summary.textContent = ticket.summary || "Untitled ticket";

    main.appendChild(number);
    main.appendChild(summary);
    button.appendChild(main);

    if (ticket.status) {
      const status = document.createElement("span");
      status.className = "halo-ticket__status";
      status.textContent = ticket.status;
      button.appendChild(status);
    } else {
      button.appendChild(createTicketChevron());
    }

    const metaValues = [ticket.client, ticket.agent].filter(Boolean);
    if (metaValues.length) {
      const meta = document.createElement("span");
      meta.className = "halo-ticket__meta";

      metaValues.forEach((value) => {
        const item = document.createElement("span");
        item.className = "halo-ticket__meta-item";
        item.textContent = value;
        meta.appendChild(item);
      });

      button.appendChild(meta);
    }

    container.appendChild(button);
  });
}

function createTicketChevron(): SVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("class", "halo-ticket__chevron");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("aria-hidden", "true");
  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", "m9 18 6-6-6-6");
  svg.appendChild(path);
  return svg;
}

async function attachEmailToTicket(ticket: HaloTicket, selectedButton: HTMLElement) {
  if (!ticket.id) {
    setFailed(createHaloAuthError("This ticket does not include a Halo ticket ID."), {
      hideLogout: false,
      message: "Email attach failed",
    });
    return;
  }

  if (getComposeMetadataItem()) {
    await selectComposeAttachTicket(ticket, selectedButton);
    return;
  }

  selectTicketButton(selectedButton);

  try {
    setTicketsBusy(true);
    setStatus("loading", "Attaching email to Halo ticket...", formatTicketTitle(ticket));

    const snapshot = await readCurrentOutlookEmailSnapshot();
    if (!snapshot) {
      throw createHaloAuthError("Open an existing received email, then choose a Halo ticket.");
    }
    const metadata = await snapshot.emailAttachments;
    const includeAttachments = metadata.attachments.length
      ? await promptForEmailAttachments(metadata)
      : false;
    if (includeAttachments === null || snapshot.itemRevision !== outlookItemRevision) {
      return;
    }
    if (metadata.attachments.length) {
      setStatus(
        "loading",
        includeAttachments ? "Preparing email attachments..." : "Attaching email only...",
        formatTicketTitle(ticket)
      );
    }
    const attachmentPreparation = includeAttachments
      ? prepareEmailAttachmentsForHalo(
          Office.context.mailbox.item as unknown as ComposeMetadataItem,
          metadata,
          ticket.id,
          createComposeAttachId()
        )
      : Promise.resolve(null);
    const [inlineImages, preparedAttachments] = await Promise.all([
      snapshot.inlineImages.catch(() => emptyPreparedInlineImages()),
      attachmentPreparation,
    ]);
    if (snapshot.itemRevision !== outlookItemRevision) {
      if (preparedAttachments?.emailAttachmentPrefetchKey) {
        void cancelEmailAttachmentPrefetch(preparedAttachments.emailAttachmentPrefetchKey);
      }
      return;
    }
    let email = applyPreparedInlineImages(snapshot.email, inlineImages);
    email = preparedAttachments
      ? applyPreparedEmailAttachments(email, preparedAttachments)
      : applyEmailOnlyChoice(email, metadata);
    email.actionMode = activeActionMode;

    const result = await fetchJson<HaloAttachEmailResponse>(
      `/api/halo/tickets/${encodeURIComponent(ticket.id)}/email`,
      {
        method: "POST",
        body: JSON.stringify({
          ...email,
          ticketNumber: ticket.ticketNumber || ticket.id,
        }),
      }
    );

    if (!result.ok) {
      throw createHaloAuthError(
        result.error || result.message || "Email attach failed.",
        result.debug
      );
    }

    await saveBackgroundSessionId(result.backgroundSessionId || "");
    const successMessage =
      result.message ||
      (activeActionMode === "private-note"
        ? "Email attached as a private note"
        : "Email attached as an Email action");
    const imageWarning = Boolean(result.inlineImages && result.inlineImages.failed > 0);
    const attachmentWarning = hasEmailAttachmentWarnings(result.emailAttachments);
    setStatus(
      imageWarning || attachmentWarning ? "warning" : "success",
      imageWarning || attachmentWarning ? "Email attached with warnings" : successMessage,
      imageWarning || attachmentWarning
        ? buildEmailImportWarningDetail(result.inlineImages, result.emailAttachments)
        : `Attached ${email.subject || "selected email"} to ${ticket.ticketNumber || ticket.id}.`
    );
    actionModeLocked = true;
    getPrivateNoteToggle().disabled = true;
    mappedConversationTicket = {
      ticketId: ticket.id,
      ticketNumber: ticket.ticketNumber || ticket.id,
      actionMode: activeActionMode,
    };
    getMappedEmailConfirmation().hidden = true;
  } catch (error) {
    setFailed(error, { hideLogout: false, message: "Email attach failed" });
  } finally {
    setTicketsBusy(false);
  }
}

async function selectComposeAttachTicket(ticket: HaloTicket, selectedButton: HTMLElement) {
  const item = getComposeMetadataItem();
  if (!item) {
    return;
  }
  const previousMarker = activeComposeAttachMarker;
  composeMarkerRevision += 1;
  const markerRevision = composeMarkerRevision;

  try {
    setTicketsBusy(true);
    getRemoveComposeSelectionButton().disabled = true;
    setStatus("loading", "Saving ticket selection...", formatTicketTitle(ticket));

    // Outlook can replace compose attachment IDs when an unsaved message is first
    // persisted. Save before reading attachment metadata so content retrieval uses
    // IDs from the saved draft rather than transient compose IDs.
    const draftItemId = await enqueueComposeMetadataWrite(() => saveComposeItem(item));
    const bodyHtml = await getBodyAsync(item, Office.CoercionType.Html).catch(() => "");
    const attachmentMetadata = await collectEmailAttachmentMetadata(item, bodyHtml).catch(() =>
      emptyEmailAttachmentMetadata()
    );
    let emailAttachmentDecision: "include" | "exclude" | undefined;
    if (attachmentMetadata.attachments.length) {
      if (
        previousMarker &&
        previousMarker.emailAttachmentFingerprint ===
          attachmentMetadata.emailAttachmentFingerprint &&
        previousMarker.emailAttachmentDecision
      ) {
        emailAttachmentDecision = previousMarker.emailAttachmentDecision;
      } else {
        const includeAttachments = await promptForEmailAttachments(attachmentMetadata);
        if (includeAttachments === null || markerRevision !== composeMarkerRevision) {
          return;
        }
        emailAttachmentDecision = includeAttachments ? "include" : "exclude";
      }
    }

    const marker: ComposeAttachMarker = {
      version: 3,
      composeAttachId: createComposeAttachId(),
      ticketId: ticket.id,
      ticketNumber: ticket.ticketNumber || ticket.id,
      ticketSummary: (ticket.summary || "Untitled ticket").slice(0, 500),
      draftItemId,
      actionMode: activeActionMode,
      composeIdentityHeader: true,
      emailAttachmentDecision,
      emailAttachmentFingerprint: attachmentMetadata.emailAttachmentFingerprint,
      emailAttachmentStagingVersion: 2,
      emailAttachmentSummary:
        emailAttachmentDecision === "include"
          ? createPendingEmailAttachmentSummary(attachmentMetadata)
          : undefined,
    };

    await enqueueComposeMetadataWrite(async () => {
      if (markerRevision !== composeMarkerRevision) {
        throw new Error("The Outlook item changed while the ticket selection was being saved.");
      }
      await setComposeIdentityHeader(item, marker.composeAttachId);
      await saveComposeSessionMarker(item, marker);
      await saveComposeCustomMarker(item, marker);
      const persistedItemId = await saveComposeItem(item);
      if (persistedItemId !== marker.draftItemId) {
        marker.draftItemId = persistedItemId;
        await saveComposeSessionMarker(item, marker);
        await saveComposeCustomMarker(item, marker);
        await saveComposeItem(item);
      }
    });

    activeComposeAttachMarker = marker;
    selectTicketButton(selectedButton);
    renderComposeAttachSelection();
    setStatus(
      marker.emailAttachmentDecision === "include" ? "loading" : "success",
      marker.emailAttachmentDecision === "include"
        ? `Preparing ${attachmentMetadata.attachments.length} attachments...`
        : `Will attach to ${marker.ticketNumber} when sent`,
      marker.emailAttachmentDecision === "include"
        ? "Keep Outlook open while Halo prepares the selected files."
        : "You can close this pane, change the ticket, or remove the selection before sending."
    );
    void prefetchComposeInlineImages(item, marker);
    if (marker.emailAttachmentDecision === "include") {
      void prefetchComposeEmailAttachments(item, marker, markerRevision);
    } else {
      void clearComposeEmailAttachmentState(item);
    }
    if (previousMarker && previousMarker.emailAttachmentPrefetchKey) {
      void cancelEmailAttachmentPrefetch(previousMarker.emailAttachmentPrefetchKey);
    }
    if (previousMarker && isCreateTicketMarker(previousMarker)) {
      void removeServerTicketCreationIntent(
        previousMarker.creationOperationId || previousMarker.composeAttachId
      );
    }
  } catch (error) {
    if (markerRevision !== composeMarkerRevision) {
      return;
    }
    if (previousMarker) {
      await enqueueComposeMetadataWrite(async () => {
        await restoreComposeIdentityHeader(item, previousMarker);
        await saveComposeSessionMarker(item, previousMarker);
        await saveComposeCustomMarker(item, previousMarker);
        await saveComposeItem(item);
      }).catch(() => undefined);
      activeComposeAttachMarker = previousMarker;
    } else {
      await enqueueComposeMetadataWrite(async () => {
        await setComposeIdentityHeader(item, "");
        await clearComposeAttachMetadata(item);
        await saveComposeItem(item);
      }).catch(() => undefined);
      activeComposeAttachMarker = null;
    }
    renderComposeAttachSelection();
    if (previousMarker) {
      setStatus(
        "failed",
        "Could not change ticket selection",
        `The email is still set to attach to ${previousMarker.ticketNumber}. ${
          error instanceof Error ? error.message : "Try again."
        }`
      );
    } else {
      setFailed(error, { hideLogout: false, message: "Could not save ticket selection" });
    }
  } finally {
    setTicketsBusy(false);
    getRemoveComposeSelectionButton().disabled = false;
  }
}

function createPendingEmailAttachmentSummary(
  metadata: PreparedEmailAttachmentMetadata
): EmailAttachmentSummary {
  let total = 0;
  let selected = 0;
  let skipped = 0;
  metadata.attachments.forEach((attachment, index) => {
    const unsupportedType = ["cloud", "link", "reference", "url"].includes(
      attachment.attachmentType
    );
    if (
      index >= MAX_EMAIL_ATTACHMENTS ||
      unsupportedType ||
      attachment.reportedSize > MAX_EMAIL_ATTACHMENT_BYTES ||
      total + attachment.reportedSize > MAX_EMAIL_ATTACHMENT_TOTAL_BYTES
    ) {
      skipped += 1;
      return;
    }
    total += attachment.reportedSize;
    selected += 1;
  });
  return {
    attached: 0,
    detected: metadata.attachments.length,
    failed: 0,
    prepared: 0,
    selected,
    skipped,
    warnings: skipped ? ["Some email attachments are unsupported and will be skipped."] : [],
  };
}

async function removeComposeAttachSelection() {
  const item = getComposeMetadataItem();
  if (!item) {
    return;
  }
  const previousMarker = activeComposeAttachMarker;
  let removalRevision = composeMarkerRevision;

  try {
    composeMarkerRevision += 1;
    composeInlineImagePrefetchRevision += 1;
    removalRevision = composeMarkerRevision;
    setTicketsBusy(true);
    getRemoveComposeSelectionButton().disabled = true;
    setStatus(
      "loading",
      "Removing Halo selection...",
      "No Halo action will run when this email is sent."
    );
    await enqueueComposeMetadataWrite(async () => {
      await setComposeIdentityHeader(item, "");
      await clearComposeAttachMetadata(item);
      await clearComposeEmailAttachmentState(item);
      await saveComposeItem(item);
    });
    activeComposeAttachMarker = null;
    selectTicketById("");
    renderComposeAttachSelection();
    setStatus("success", "Ticket selection removed", "Choose another ticket if needed.");
    if (previousMarker && previousMarker.emailAttachmentPrefetchKey) {
      void cancelEmailAttachmentPrefetch(previousMarker.emailAttachmentPrefetchKey);
    }
    if (previousMarker && isCreateTicketMarker(previousMarker)) {
      void removeServerTicketCreationIntent(
        previousMarker.creationOperationId || previousMarker.composeAttachId
      );
    }
  } catch (error) {
    if (previousMarker && removalRevision === composeMarkerRevision) {
      await enqueueComposeMetadataWrite(async () => {
        await restoreComposeIdentityHeader(item, previousMarker);
        await saveComposeSessionMarker(item, previousMarker);
        await saveComposeCustomMarker(item, previousMarker);
        await saveComposeItem(item);
      }).catch(() => undefined);
      activeComposeAttachMarker = previousMarker;
      renderComposeAttachSelection();
    }
    setFailed(error, { hideLogout: false, message: "Could not remove ticket selection" });
  } finally {
    setTicketsBusy(false);
    getRemoveComposeSelectionButton().disabled = false;
  }
}

async function prefetchComposeInlineImages(
  item: ComposeMetadataItem,
  marker: ComposeAttachMarker
): Promise<void> {
  if (isCreateTicketMarker(marker)) {
    return;
  }
  const inlineRevision = ++composeInlineImagePrefetchRevision;
  const isCurrentSelection = () =>
    inlineRevision === composeInlineImagePrefetchRevision &&
    Boolean(
      activeComposeAttachMarker &&
      activeComposeAttachMarker.composeAttachId === marker.composeAttachId
    );
  try {
    const bodyHtml = await getBodyAsync(item, Office.CoercionType.Html);
    const prepared = await prepareOutlookInlineImages(item, bodyHtml, marker.draftItemId);
    if (!prepared.inlineImageRefs.length) {
      return;
    }

    const result = await fetchJson<InlineImagePrefetchResponse>(
      "/api/halo/inline-images/prefetch",
      {
        method: "POST",
        body: JSON.stringify({
          ...prepared,
          actionMode: normalizeActionMode(marker.actionMode),
          composeOperationId: marker.composeAttachId,
          ticketId: marker.ticketId,
        }),
      }
    );
    if (
      !result.ok ||
      !result.inlineImagePrefetchKey ||
      (result.summary && result.summary.failed > 0)
    ) {
      return;
    }
    if (!isCurrentSelection()) {
      return;
    }

    await enqueueComposeMetadataWrite(async () => {
      if (!isCurrentSelection()) {
        return;
      }
      marker.inlineImageFingerprint = prepared.inlineImageFingerprint;
      marker.inlineImagePrefetchKey = result.inlineImagePrefetchKey;
      await saveComposeSessionMarker(item, marker);
      await saveComposeCustomMarker(item, marker);
      const persistedItemId = await saveComposeItem(item);
      if (persistedItemId !== marker.draftItemId) {
        marker.draftItemId = persistedItemId;
        await saveComposeSessionMarker(item, marker);
        await saveComposeCustomMarker(item, marker);
        await saveComposeItem(item);
      }
    });
  } catch {
    // Ticket selection remains valid; send-time fallback will retry within its budget.
  }
}

function prefetchComposeEmailAttachments(
  item: ComposeMetadataItem,
  marker: ComposeAttachMarker,
  markerRevision: number
): Promise<void> {
  const preparationRevision = ++composeEmailAttachmentPreparationRevision;
  const operation = composeEmailAttachmentPreparationQueue.then(
    () => runComposeEmailAttachmentPreparation(item, marker, markerRevision, preparationRevision),
    () => runComposeEmailAttachmentPreparation(item, marker, markerRevision, preparationRevision)
  );
  composeEmailAttachmentPreparationQueue = operation.catch(() => undefined);
  return operation;
}

async function runComposeEmailAttachmentPreparation(
  item: ComposeMetadataItem,
  marker: ComposeAttachMarker,
  markerRevision: number,
  preparationRevision: number
): Promise<void> {
  const previousPrefetchKey = marker.emailAttachmentPrefetchKey || "";
  let metadata = emptyEmailAttachmentMetadata();
  const isCurrentSelection = () =>
    preparationRevision === composeEmailAttachmentPreparationRevision &&
    markerRevision === composeMarkerRevision &&
    marker.emailAttachmentDecision === "include" &&
    Boolean(
      activeComposeAttachMarker &&
      activeComposeAttachMarker.composeAttachId === marker.composeAttachId
    );
  try {
    if (!isCurrentSelection()) {
      return;
    }
    // Metadata captured before a save may contain attachment IDs Outlook has
    // already replaced. Always recollect immediately before reading content.
    const bodyHtml = await getBodyAsync(item, Office.CoercionType.Html).catch(() => "");
    metadata = await collectEmailAttachmentMetadata(item, bodyHtml);
    if (!isCurrentSelection()) {
      return;
    }
    const prepared = await prepareEmailAttachmentsForHalo(
      item,
      metadata,
      isCreateTicketMarker(marker) ? "0" : marker.ticketId,
      marker.composeAttachId,
      isCreateTicketMarker(marker) ? marker.creationOperationId || marker.composeAttachId : ""
    );
    if (!isCurrentSelection()) {
      const activeMarker = activeComposeAttachMarker;
      if (
        prepared.emailAttachmentPrefetchKey &&
        (!activeMarker || activeMarker.composeAttachId !== marker.composeAttachId)
      ) {
        void cancelEmailAttachmentPrefetch(prepared.emailAttachmentPrefetchKey);
      }
      return;
    }
    marker.emailAttachmentFingerprint = prepared.emailAttachmentFingerprint;
    marker.emailAttachmentPrefetchKey = prepared.emailAttachmentPrefetchKey;
    marker.emailAttachmentStagingVersion = prepared.emailAttachmentStagingVersion;
    marker.emailAttachmentSummary = prepared.emailAttachmentSummary;
    await enqueueComposeMetadataWrite(async () => {
      if (!isCurrentSelection()) {
        return;
      }
      await saveComposeSessionMarker(item, marker);
      await saveComposeCustomMarker(item, marker);
      if (prepared.emailAttachmentPrefetchKey) {
        await saveComposeEmailAttachmentState(item, {
          draftItemId: marker.draftItemId,
          emailAttachmentDecision: "include",
          emailAttachmentFingerprint: prepared.emailAttachmentFingerprint,
          emailAttachmentPrefetchKey: prepared.emailAttachmentPrefetchKey,
          emailAttachmentStagingVersion: 2,
          emailAttachmentSummary: prepared.emailAttachmentSummary,
          operationId: marker.composeAttachId,
          ticketId: isCreateTicketMarker(marker) ? "0" : marker.ticketId,
          version: 2,
        });
      } else {
        await clearComposeEmailAttachmentState(item);
      }
      await saveComposeItem(item);
    });
    if (!isCurrentSelection()) {
      return;
    }
    renderComposeAttachSelection();
    if (
      prepared.emailAttachmentSummary.failed > 0 ||
      prepared.emailAttachmentSummary.prepared < prepared.emailAttachmentSummary.selected
    ) {
      setStatus(
        "failed",
        "Attachment preparation failed — retry before sending",
        "Open the Halo pane and let preparation finish, then try Send again."
      );
    } else if (hasEmailAttachmentWarnings(prepared.emailAttachmentSummary)) {
      setStatus(
        "warning",
        `${prepared.emailAttachmentSummary.prepared} ready, ${prepared.emailAttachmentSummary.skipped} unsupported`,
        buildEmailImportWarningDetail(undefined, prepared.emailAttachmentSummary)
      );
    } else {
      setStatus(
        "success",
        `Will attach to ${marker.ticketNumber} when sent`,
        `${prepared.emailAttachmentSummary.prepared} attachments ready for Halo.`
      );
    }
  } catch {
    if (!isCurrentSelection()) {
      return;
    }
    const failedSummary = createPendingEmailAttachmentSummary(metadata);
    failedSummary.detected = Math.max(
      failedSummary.detected,
      marker.emailAttachmentSummary?.detected || 0
    );
    failedSummary.selected = Math.max(
      failedSummary.selected,
      marker.emailAttachmentSummary?.selected || 0
    );
    failedSummary.failed = Math.max(1, failedSummary.selected);
    marker.emailAttachmentFingerprint =
      metadata.emailAttachmentFingerprint || marker.emailAttachmentFingerprint || "";
    marker.emailAttachmentStagingVersion = 2;
    marker.emailAttachmentPrefetchKey = "";
    marker.emailAttachmentSummary = failedSummary;
    await enqueueComposeMetadataWrite(async () => {
      if (!isCurrentSelection()) {
        return;
      }
      await saveComposeSessionMarker(item, marker);
      await saveComposeCustomMarker(item, marker);
      await clearComposeEmailAttachmentState(item);
      await saveComposeItem(item);
    }).catch(() => undefined);
    if (previousPrefetchKey) {
      void cancelEmailAttachmentPrefetch(previousPrefetchKey);
    }
    if (!isCurrentSelection()) {
      return;
    }
    renderComposeAttachSelection();
    setStatus(
      "failed",
      "Attachment preparation failed — retry before sending",
      "Open the Halo pane and let preparation finish, then try Send again."
    );
  }
}

async function reconcileComposeEmailAttachments(): Promise<void> {
  const marker = activeComposeAttachMarker;
  const item = getComposeMetadataItem();
  if (!marker || !item) {
    return;
  }
  const markerRevision = composeMarkerRevision;
  const bodyHtml = await getBodyAsync(item, Office.CoercionType.Html).catch(() => "");
  const metadata = await collectEmailAttachmentMetadata(item, bodyHtml).catch(() =>
    emptyEmailAttachmentMetadata()
  );
  if (markerRevision !== composeMarkerRevision) {
    return;
  }
  if (marker.emailAttachmentFingerprint === metadata.emailAttachmentFingerprint) {
    if (marker.emailAttachmentDecision === "include" && metadata.attachments.length) {
      const ready = await refreshComposeAttachmentPreparationStatus(marker).catch(() => false);
      if (!ready && markerRevision === composeMarkerRevision) {
        marker.emailAttachmentPrefetchKey = "";
        marker.emailAttachmentStagingVersion = 2;
        renderComposeAttachSelection();
      }
      if (markerRevision === composeMarkerRevision) {
        // A matching metadata fingerprint cannot detect same-name/same-size
        // content replacement. Re-read and send content hashes so the server
        // can preserve unchanged ciphertext and restage only changed files.
        void prefetchComposeEmailAttachments(item, marker, markerRevision);
      }
    }
    return;
  }

  const previousPrefetchKey = marker.emailAttachmentPrefetchKey || "";
  const preserveExclusion = marker.emailAttachmentDecision === "exclude";
  if (!metadata.attachments.length) {
    marker.emailAttachmentDecision = preserveExclusion ? "exclude" : undefined;
    marker.emailAttachmentFingerprint = "";
    marker.emailAttachmentPrefetchKey = "";
    marker.emailAttachmentSummary = undefined;
  } else if (preserveExclusion) {
    marker.emailAttachmentFingerprint = metadata.emailAttachmentFingerprint;
    marker.emailAttachmentPrefetchKey = "";
    marker.emailAttachmentSummary = undefined;
  } else if (marker.emailAttachmentDecision === "include") {
    marker.emailAttachmentFingerprint = metadata.emailAttachmentFingerprint;
    marker.emailAttachmentPrefetchKey = "";
    marker.emailAttachmentSummary = undefined;
  } else {
    const includeAttachments = await promptForEmailAttachments(metadata);
    if (includeAttachments === null || markerRevision !== composeMarkerRevision) {
      return;
    }
    marker.emailAttachmentDecision = includeAttachments ? "include" : "exclude";
    marker.emailAttachmentFingerprint = metadata.emailAttachmentFingerprint;
    marker.emailAttachmentPrefetchKey = "";
    marker.emailAttachmentSummary = undefined;
  }

  await enqueueComposeMetadataWrite(async () => {
    if (markerRevision !== composeMarkerRevision) {
      return;
    }
    await saveComposeSessionMarker(item, marker);
    await saveComposeCustomMarker(item, marker);
    await clearComposeEmailAttachmentState(item);
    await saveComposeItem(item);
  });
  if (marker.emailAttachmentDecision === "include") {
    // Keep the same operation alive so the server can reconcile unchanged,
    // added, and removed attachments without uploading unchanged files again.
    void prefetchComposeEmailAttachments(item, marker, markerRevision);
  } else if (previousPrefetchKey) {
    void cancelEmailAttachmentPrefetch(previousPrefetchKey);
  }
}

async function refreshComposeAttachmentPreparationStatus(
  marker: ComposeAttachMarker
): Promise<boolean> {
  if (marker.emailAttachmentStagingVersion !== 2 || !marker.emailAttachmentPrefetchKey) {
    return false;
  }
  const result = await fetchJson<{
    ok: boolean;
    stagingVersion?: number;
    status?: string;
    aggregate?: { failed: number; pending: number; prepared: number; selected: number };
  }>(
    `/api/halo/email-attachments/prefetch/${encodeURIComponent(
      marker.emailAttachmentPrefetchKey
    )}/status`,
    { method: "GET" }
  );
  if (!result.ok || result.stagingVersion !== 2 || result.status !== "active") {
    return false;
  }
  const aggregate = result.aggregate;
  if (!aggregate) {
    return false;
  }
  marker.emailAttachmentSummary = {
    attached: 0,
    detected: marker.emailAttachmentSummary?.detected || aggregate.selected,
    failed: aggregate.failed,
    prepared: aggregate.prepared,
    selected: aggregate.selected,
    skipped: marker.emailAttachmentSummary?.skipped || 0,
    warnings: marker.emailAttachmentSummary?.warnings || [],
  };
  renderComposeAttachSelection();
  return (
    aggregate.failed === 0 && aggregate.pending === 0 && aggregate.prepared === aggregate.selected
  );
}

async function cancelEmailAttachmentPrefetch(prefetchKey: string): Promise<void> {
  if (!prefetchKey) {
    return;
  }
  await fetchJson(`/api/halo/email-attachments/prefetch/${encodeURIComponent(prefetchKey)}`, {
    method: "DELETE",
    body: JSON.stringify({}),
  }).catch(() => undefined);
}

async function prepareEmailAttachmentsForHalo(
  item: ComposeMetadataItem,
  metadata: PreparedEmailAttachmentMetadata,
  ticketId: string,
  operationId: string,
  creationOperationId = ""
): Promise<{
  includeEmailAttachments: boolean;
  emailAttachmentDraftItemId: string;
  emailAttachmentFingerprint: string;
  emailAttachmentOperationId: string;
  emailAttachmentPrefetchKey: string;
  emailAttachmentStagingVersion: 2;
  emailAttachmentSummary: EmailAttachmentSummary;
}> {
  const readStartedAt = Date.now();
  const draftItemId = await readComposeItemId(item);
  let attemptCount = 0;
  let currentMetadata = metadata;
  let prepared = await prepareOutlookEmailAttachments(
    item,
    currentMetadata,
    EMAIL_ATTACHMENT_READ_TIMEOUT_MS
  );
  attemptCount += 1;
  for (const retryDelayMs of EMAIL_ATTACHMENT_READ_RETRY_DELAYS_MS) {
    if (!prepared.failed) {
      break;
    }
    sendEmailAttachmentClientDiagnostic("attachment-read-retry", {
      attachmentCount: currentMetadata.attachments.length,
      attachmentError: prepared.failureCodes[0] || "read-failed",
      attemptCount,
      elapsedMs: Date.now() - readStartedAt,
      outcome: "started",
    });
    await waitForEmailAttachmentRetry(retryDelayMs);
    const bodyHtml = await getBodyAsync(item, Office.CoercionType.Html).catch(() => "");
    currentMetadata = await collectEmailAttachmentMetadata(item, bodyHtml).catch(
      () => currentMetadata
    );
    prepared = await prepareOutlookEmailAttachments(
      item,
      currentMetadata,
      EMAIL_ATTACHMENT_READ_TIMEOUT_MS
    );
    attemptCount += 1;
  }
  sendEmailAttachmentClientDiagnostic("attachment-read-complete", {
    attachmentCount: currentMetadata.attachments.length,
    attachmentError: prepared.failureCodes[0] || "none",
    attemptCount,
    elapsedMs: Date.now() - readStartedAt,
    failedCount: prepared.failed,
    outcome: prepared.failed ? "failed" : "ok",
    skippedCount: prepared.skipped,
  });
  const summary: EmailAttachmentSummary = {
    attached: 0,
    detected: currentMetadata.attachments.length,
    failed: prepared.failed,
    selected: prepared.descriptors.length,
    skipped: prepared.skipped,
    prepared: 0,
    warnings: prepared.warnings,
  };
  if (!prepared.descriptors.length) {
    sendEmailAttachmentClientDiagnostic("attachment-prefetch-complete", {
      attachmentCount: summary.detected,
      elapsedMs: Date.now() - readStartedAt,
      failedCount: summary.failed,
      outcome: summary.failed || summary.skipped ? "failed" : "ok",
      skippedCount: summary.skipped,
      uploadedCount: 0,
    });
    return {
      emailAttachmentDraftItemId: draftItemId,
      emailAttachmentFingerprint: currentMetadata.emailAttachmentFingerprint,
      emailAttachmentOperationId: operationId,
      emailAttachmentPrefetchKey: "",
      emailAttachmentStagingVersion: 2,
      emailAttachmentSummary: summary,
      includeEmailAttachments: true,
    };
  }

  const start = await fetchJson<EmailAttachmentPrefetchStartResponse>(
    "/api/halo/email-attachments/prefetch/start",
    {
      method: "POST",
      body: JSON.stringify({
        draftItemId,
        emailAttachmentFingerprint: currentMetadata.emailAttachmentFingerprint,
        emailAttachments: prepared.descriptors.map(toEmailAttachmentDescriptor),
        operationId,
        ...(creationOperationId ? { creationOperationId } : { ticketId }),
      }),
    }
  );
  if (
    !start.ok ||
    (start.status !== "ready" && start.status !== "pending") ||
    start.stagingVersion !== 2 ||
    !start.prefetchKey
  ) {
    summary.failed += prepared.uploads.length;
    summary.warnings = uniqueAttachmentWarnings([
      ...summary.warnings,
      "Email attachments could not be prepared for Halo.",
    ]);
    sendEmailAttachmentClientDiagnostic("attachment-prefetch-complete", {
      attachmentCount: summary.detected,
      elapsedMs: Date.now() - readStartedAt,
      failedCount: summary.failed,
      outcome: "failed",
      skippedCount: summary.skipped,
      uploadedCount: 0,
    });
    return {
      emailAttachmentDraftItemId: draftItemId,
      emailAttachmentFingerprint: currentMetadata.emailAttachmentFingerprint,
      emailAttachmentOperationId: operationId,
      emailAttachmentPrefetchKey: "",
      emailAttachmentStagingVersion: 2,
      emailAttachmentSummary: summary,
      includeEmailAttachments: true,
    };
  }

  const pending = new Set(start.pendingAttachmentKeys || []);
  const results = await mapWithAttachmentConcurrency(
    prepared.uploads.filter((upload) => pending.has(upload.attachmentKey)),
    3,
    async (upload) => {
      try {
        const result = await fetchJson<{ ok: boolean; status?: string }>(
          `/api/halo/email-attachments/prefetch/${encodeURIComponent(start.prefetchKey!)}/items`,
          {
            method: "POST",
            body: JSON.stringify({
              attachmentKey: upload.attachmentKey,
              contentBase64: upload.contentBase64,
              contentFormat: upload.contentFormat,
              contentSha256: upload.contentSha256,
            }),
          }
        );
        return result.ok && (result.status === "prepared" || result.status === "already-prepared");
      } catch {
        return false;
      }
    }
  );
  const successfulUploads = results.filter(Boolean).length;
  const preparedCount = Number(start.aggregate?.prepared || 0) + successfulUploads;
  const uploadFailures = Math.max(0, pending.size - successfulUploads);
  summary.prepared = preparedCount;
  summary.failed = uploadFailures;
  if (uploadFailures) {
    summary.warnings = uniqueAttachmentWarnings([
      ...summary.warnings,
      "Some email attachments could not be prepared for Halo.",
    ]);
  }
  sendEmailAttachmentClientDiagnostic("attachment-prefetch-complete", {
    attachmentCount: summary.detected,
    elapsedMs: Date.now() - readStartedAt,
    failedCount: summary.failed,
    outcome: summary.failed || summary.skipped ? "failed" : "ok",
    skippedCount: summary.skipped,
    uploadedCount: summary.prepared,
  });
  return {
    emailAttachmentDraftItemId: draftItemId,
    emailAttachmentFingerprint: currentMetadata.emailAttachmentFingerprint,
    emailAttachmentOperationId: operationId,
    emailAttachmentPrefetchKey: start.prefetchKey,
    emailAttachmentStagingVersion: 2,
    emailAttachmentSummary: summary,
    includeEmailAttachments: true,
  };
}

function waitForEmailAttachmentRetry(delayMs: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, delayMs));
}

function toEmailAttachmentDescriptor(upload: EmailAttachmentDescriptor): EmailAttachmentDescriptor {
  return {
    attachmentKey: upload.attachmentKey,
    attachmentType: upload.attachmentType,
    contentSha256: upload.contentSha256 || "",
    contentType: upload.contentType,
    name: upload.name,
    reportedSize: upload.reportedSize,
  };
}

async function mapWithAttachmentConcurrency<T, TResult>(
  values: T[],
  limit: number,
  worker: (value: T) => Promise<TResult>
): Promise<TResult[]> {
  const results = new Array<TResult>(values.length);
  let nextIndex = 0;
  async function run() {
    while (nextIndex < values.length) {
      const index = nextIndex++;
      results[index] = await worker(values[index]);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, values.length) }, () => run()));
  return results;
}

function uniqueAttachmentWarnings(values: string[]): string[] {
  return Array.from(new Set(values)).slice(0, 10);
}

function applyPreparedEmailAttachments(
  email: OutlookEmailPayload,
  prepared: {
    includeEmailAttachments: boolean;
    emailAttachmentDraftItemId: string;
    emailAttachmentFingerprint: string;
    emailAttachmentOperationId: string;
    emailAttachmentPrefetchKey: string;
    emailAttachmentStagingVersion: 2;
    emailAttachmentSummary: EmailAttachmentSummary;
  }
): OutlookEmailPayload {
  return { ...email, ...prepared };
}

function applyEmailOnlyChoice(
  email: OutlookEmailPayload,
  metadata: PreparedEmailAttachmentMetadata
): OutlookEmailPayload {
  return {
    ...email,
    emailAttachmentFingerprint: metadata.emailAttachmentFingerprint,
    emailAttachmentSummary: {
      attached: 0,
      detected: metadata.attachments.length,
      failed: 0,
      prepared: 0,
      selected: 0,
      skipped: 0,
      warnings: [],
    },
    includeEmailAttachments: false,
  };
}

async function restoreComposeAttachSelection() {
  const itemRevision = outlookItemRevision;
  const item = getComposeMetadataItem();
  updateComposeModeCopy(Boolean(item));

  if (!item) {
    activeComposeAttachMarker = null;
    renderComposeAttachSelection();
    return;
  }

  const currentItemId = await readComposeItemId(item);
  const sessionMarker = await readComposeSessionMarker(item);
  const customMarker = sessionMarker ? null : await readComposeCustomMarker(item);
  const marker = sessionMarker || customMarker;
  if (itemRevision !== outlookItemRevision) {
    return;
  }

  if (!marker) {
    activeComposeAttachMarker = null;
    renderComposeAttachSelection();
    return;
  }

  if (!currentItemId || marker.draftItemId !== currentItemId) {
    await enqueueComposeMetadataWrite(async () => {
      if (marker.composeIdentityHeader) {
        await setComposeIdentityHeader(item, "");
      }
      await clearComposeAttachMetadata(item);
      await saveComposeItem(item);
    }).catch(() => undefined);
    if (itemRevision !== outlookItemRevision) {
      return;
    }
    activeComposeAttachMarker = null;
    renderComposeAttachSelection();
    return;
  }

  applyComposeMarkerRuntimeContext(marker);
  activeComposeAttachMarker = marker;
  composeMarkerRevision += 1;
  await enqueueComposeMetadataWrite(async () => {
    if (marker.composeIdentityHeader) {
      await setComposeIdentityHeader(item, marker.composeAttachId);
    }
    await saveComposeSessionMarker(item, marker);
  }).catch(() => undefined);
  if (itemRevision !== outlookItemRevision) {
    return;
  }
  renderComposeAttachSelection();
  const attachmentsReady = isComposeAttachmentReady(marker);
  setStatus(
    marker.emailAttachmentDecision === "include" && !attachmentsReady ? "loading" : "success",
    marker.emailAttachmentDecision === "include" && !attachmentsReady
      ? `Preparing ${marker.emailAttachmentSummary?.selected || ""} attachments...`
      : isCreateTicketMarker(marker)
        ? `Will create ${marker.ticketTypeName || "a Halo ticket"} when sent`
        : `Will attach to ${marker.ticketNumber} when sent`,
    attachmentsReady
      ? "This selection and its attachments were restored with the Outlook draft."
      : "This selection was saved with the Outlook draft."
  );
  void reconcileComposeEmailAttachments();
}

async function prepareMappedComposeAttachmentsOnOpen(): Promise<void> {
  const item = getComposeMetadataItem();
  if (!item || activeComposeAttachMarker) {
    return;
  }
  const conversationId = String(item.conversationId || "");
  const inReplyToMessageIds = item.inReplyTo ? [String(item.inReplyTo)] : [];
  if (!conversationId && !inReplyToMessageIds.length) {
    return;
  }
  const draftItemId = await enqueueComposeMetadataWrite(() => saveComposeItem(item)).catch(
    () => ""
  );
  if (!draftItemId) {
    return;
  }
  const bodyHtml = await getBodyAsync(item, Office.CoercionType.Html).catch(() => "");
  const metadata = await collectEmailAttachmentMetadata(item, bodyHtml).catch(() =>
    emptyEmailAttachmentMetadata()
  );
  const match = await fetchJson<HaloEmailMatchResponse>("/api/halo/email/match", {
    method: "POST",
    body: JSON.stringify({
      conversationId,
      inReplyToMessageIds,
      internetMessageId: String(item.internetMessageId || ""),
      itemId: draftItemId,
      mailboxEmail: Office.context.mailbox.userProfile.emailAddress || "",
      referenceMessageIds: [],
    }),
  }).catch(() => null);
  if (!match || !match.ok || match.status !== "matched" || !match.ticketId) {
    return;
  }
  mappedConversationTicket = {
    ticketId: String(match.ticketId),
    ticketNumber: String(match.ticketNumber || match.ticketId),
    actionMode: normalizeActionMode(match.actionMode),
  };
  setActionMode(mappedConversationTicket.actionMode);
  renderCreateTicketConversationWarning();

  const previousState = await readComposeEmailAttachmentState(item);
  const reusable = Boolean(
    previousState &&
    previousState.draftItemId === draftItemId &&
    previousState.ticketId === String(match.ticketId)
  );
  const operationId = reusable ? previousState!.operationId : createComposeAttachId();
  const marker: ComposeAttachMarker = {
    version: 3,
    destinationKind: "existing-ticket",
    composeAttachId: operationId,
    ticketId: mappedConversationTicket.ticketId,
    ticketNumber: mappedConversationTicket.ticketNumber,
    ticketSummary: "Mapped Outlook conversation",
    draftItemId,
    actionMode: mappedConversationTicket.actionMode,
  };
  const persistMarker = async () => {
    await enqueueComposeMetadataWrite(async () => {
      await saveComposeSessionMarker(item, marker);
      await saveComposeCustomMarker(item, marker);
      await saveComposeItem(item);
    });
    activeComposeAttachMarker = marker;
    renderComposeAttachSelection();
  };
  await persistMarker();
  if (!metadata.attachments.length) {
    setStatus(
      "success",
      `Will attach to ${marker.ticketNumber} when sent`,
      marker.actionMode === "private-note"
        ? "This mapped conversation will be attached as a private note."
        : "This mapped conversation will be attached as an Email action."
    );
    return;
  }
  if (
    reusable &&
    previousState!.emailAttachmentFingerprint === metadata.emailAttachmentFingerprint &&
    previousState!.emailAttachmentDecision === "exclude"
  ) {
    marker.emailAttachmentDecision = "exclude";
    marker.emailAttachmentFingerprint = metadata.emailAttachmentFingerprint;
    await persistMarker();
    return;
  }
  if (
    reusable &&
    previousState!.emailAttachmentFingerprint === metadata.emailAttachmentFingerprint &&
    previousState!.emailAttachmentDecision === "include" &&
    previousState!.emailAttachmentPrefetchKey
  ) {
    const status = await fetchJson<{
      ok: boolean;
      status?: string;
      stagingVersion?: number;
      aggregate?: { failed: number; pending: number; prepared: number; selected: number };
    }>(
      `/api/halo/email-attachments/prefetch/${encodeURIComponent(
        previousState!.emailAttachmentPrefetchKey
      )}/status`,
      { method: "GET" }
    ).catch(() => null);
    if (
      status?.ok &&
      status.status === "active" &&
      status.stagingVersion === 2 &&
      status.aggregate?.failed === 0 &&
      status.aggregate.pending === 0 &&
      status.aggregate.prepared === status.aggregate.selected
    ) {
      setStatus(
        "loading",
        `Checking ${status.aggregate.prepared} prepared attachments...`,
        "Confirming that the saved draft contents have not changed."
      );
    }
  }

  let decision = reusable ? previousState!.emailAttachmentDecision : undefined;
  if (!decision) {
    const include = await promptForEmailAttachments(metadata);
    if (include === null) {
      return;
    }
    decision = include ? "include" : "exclude";
  }
  if (decision === "exclude") {
    if (previousState?.emailAttachmentPrefetchKey) {
      void cancelEmailAttachmentPrefetch(previousState.emailAttachmentPrefetchKey);
    }
    await saveComposeEmailAttachmentState(item, {
      version: 2,
      draftItemId,
      ticketId: String(match.ticketId),
      operationId,
      emailAttachmentDecision: "exclude",
      emailAttachmentFingerprint: metadata.emailAttachmentFingerprint,
      emailAttachmentPrefetchKey: "",
      emailAttachmentStagingVersion: 2,
      emailAttachmentSummary: createPendingEmailAttachmentSummary(metadata),
    });
    await saveComposeItem(item);
    marker.emailAttachmentDecision = "exclude";
    marker.emailAttachmentFingerprint = metadata.emailAttachmentFingerprint;
    marker.emailAttachmentSummary = undefined;
    await persistMarker();
    return;
  }

  setStatus(
    "loading",
    `Preparing ${metadata.attachments.length} attachments...`,
    `Preparing files for mapped ticket ${match.ticketNumber || match.ticketId}.`
  );
  const prepared = await prepareEmailAttachmentsForHalo(
    item,
    metadata,
    String(match.ticketId),
    operationId
  );
  await saveComposeEmailAttachmentState(item, {
    version: 2,
    draftItemId,
    ticketId: String(match.ticketId),
    operationId,
    emailAttachmentDecision: "include",
    emailAttachmentFingerprint: prepared.emailAttachmentFingerprint,
    emailAttachmentPrefetchKey: prepared.emailAttachmentPrefetchKey,
    emailAttachmentStagingVersion: 2,
    emailAttachmentSummary: prepared.emailAttachmentSummary,
  });
  await saveComposeItem(item);
  marker.emailAttachmentDecision = "include";
  marker.emailAttachmentFingerprint = prepared.emailAttachmentFingerprint;
  marker.emailAttachmentPrefetchKey = prepared.emailAttachmentPrefetchKey;
  marker.emailAttachmentStagingVersion = 2;
  marker.emailAttachmentSummary = prepared.emailAttachmentSummary;
  await persistMarker();
  const ready =
    prepared.emailAttachmentSummary.failed === 0 &&
    prepared.emailAttachmentSummary.prepared === prepared.emailAttachmentSummary.selected;
  setStatus(
    ready ? (prepared.emailAttachmentSummary.skipped ? "warning" : "success") : "failed",
    ready
      ? `${prepared.emailAttachmentSummary.prepared} attachments ready for Halo`
      : "Attachment preparation failed — retry before sending",
    ready
      ? `Mapped to ticket ${match.ticketNumber || match.ticketId}.`
      : "Keep the Halo pane open, retry preparation, then send again."
  );
}

function getComposeMetadataItem(): ComposeMetadataItem | null {
  const item = Office.context.mailbox.item as unknown as ComposeMetadataItem;
  if (!item || typeof item.saveAsync !== "function") {
    return null;
  }

  if (
    item.itemType &&
    item.itemType !== Office.MailboxEnums.ItemType.Message &&
    String(item.itemType).toLowerCase() !== "message"
  ) {
    return null;
  }

  return item;
}

function enqueueComposeMetadataWrite<T>(operation: () => Promise<T>): Promise<T> {
  const result = composeMetadataWriteQueue.then(operation, operation);
  composeMetadataWriteQueue = result.then(
    () => undefined,
    () => undefined
  );
  return result;
}

function saveComposeItem(item: ComposeMetadataItem): Promise<string> {
  return new Promise((resolve, reject) => {
    if (!item.saveAsync) {
      reject(new Error("Outlook could not save this draft."));
      return;
    }

    item.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
        return;
      }

      reject(new Error(result.error.message || "Outlook could not save this draft."));
    });
  });
}

function setComposeIdentityHeader(
  item: ComposeMetadataItem,
  composeAttachId: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    const headers = item.internetHeaders;
    if (!headers) {
      reject(new Error("Outlook does not support compose internet headers."));
      return;
    }
    const callback = (result: Office.AsyncResult<void>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error.message || "Could not save the Halo compose identity."));
      }
    };
    if (composeAttachId) {
      headers.setAsync({ [COMPOSE_ATTACH_HEADER_NAME]: composeAttachId }, callback);
    } else {
      headers.removeAsync([COMPOSE_ATTACH_HEADER_NAME], callback);
    }
  });
}

function restoreComposeIdentityHeader(
  item: ComposeMetadataItem,
  marker: ComposeAttachMarker | null
): Promise<void> {
  return setComposeIdentityHeader(
    item,
    marker?.composeIdentityHeader ? marker.composeAttachId : ""
  );
}

function readComposeItemId(item: ComposeMetadataItem): Promise<string> {
  return new Promise((resolve) => {
    if (!item.getItemIdAsync) {
      resolve(String(item.itemId || ""));
      return;
    }

    item.getItemIdAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "");
    });
  });
}

function saveComposeSessionMarker(
  item: ComposeMetadataItem,
  marker: ComposeAttachMarker
): Promise<void> {
  return new Promise((resolve, reject) => {
    if (!item.sessionData) {
      reject(new Error("Outlook does not support compose-session metadata."));
      return;
    }

    applyComposeMarkerRuntimeContext(marker);
    item.sessionData.setAsync(
      COMPOSE_ATTACH_STORAGE_KEY,
      serializeComposeSessionMarker(marker),
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        reject(new Error(result.error.message || "Could not save the ticket selection."));
      }
    );
  });
}

function readComposeSessionMarker(item: ComposeMetadataItem): Promise<ComposeAttachMarker | null> {
  return new Promise((resolve) => {
    if (!item.sessionData) {
      resolve(null);
      return;
    }

    item.sessionData.getAsync(COMPOSE_ATTACH_STORAGE_KEY, (result) => {
      resolve(
        result.status === Office.AsyncResultStatus.Succeeded
          ? parseComposeAttachMarker(result.value)
          : null
      );
    });
  });
}

function saveComposeCustomMarker(
  item: ComposeMetadataItem,
  marker: ComposeAttachMarker
): Promise<void> {
  applyComposeMarkerRuntimeContext(marker);
  return loadComposeCustomProperties(item).then(
    (properties) =>
      new Promise((resolve, reject) => {
        properties.set(COMPOSE_ATTACH_STORAGE_KEY, serializeComposeCustomMarker(marker));
        properties.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
            return;
          }

          reject(new Error(result.error.message || "Could not persist the ticket selection."));
        });
      })
  );
}

function applyComposeMarkerRuntimeContext(marker: ComposeAttachMarker): ComposeAttachMarker {
  const userProfile = Office.context.mailbox.userProfile;
  if (activeBackgroundSessionId) {
    marker.backgroundSessionId = activeBackgroundSessionId;
  }
  if (userProfile && userProfile.emailAddress) {
    marker.mailboxEmail = userProfile.emailAddress;
  }
  return marker;
}

function serializeComposeCustomMarker(marker: ComposeAttachMarker): string {
  const persistentMarker = { ...marker };
  delete persistentMarker.backgroundSessionId;
  persistentMarker.emailAttachmentSummary = attachmentSummaryForOutlookStorage(
    persistentMarker.emailAttachmentSummary
  );
  return JSON.stringify(persistentMarker);
}

function serializeComposeSessionMarker(marker: ComposeAttachMarker): string {
  return JSON.stringify({
    ...marker,
    emailAttachmentSummary: attachmentSummaryForOutlookStorage(marker.emailAttachmentSummary),
  });
}

function attachmentSummaryForOutlookStorage(
  summary: EmailAttachmentSummary | undefined
): EmailAttachmentSummary | undefined {
  return summary ? { ...summary, warnings: [] } : undefined;
}

async function readComposeCustomMarker(
  item: ComposeMetadataItem
): Promise<ComposeAttachMarker | null> {
  try {
    const properties = await loadComposeCustomProperties(item);
    return parseComposeAttachMarker(properties.get(COMPOSE_ATTACH_STORAGE_KEY));
  } catch {
    return null;
  }
}

async function clearComposeAttachMetadata(item: ComposeMetadataItem): Promise<void> {
  const operations: Promise<void>[] = [];

  if (item.sessionData) {
    operations.push(
      new Promise((resolve) => {
        item.sessionData!.removeAsync(COMPOSE_ATTACH_STORAGE_KEY, () => resolve());
      })
    );
  }

  if (item.loadCustomPropertiesAsync) {
    operations.push(
      loadComposeCustomProperties(item).then(
        (properties) =>
          new Promise((resolve, reject) => {
            properties.remove(COMPOSE_ATTACH_STORAGE_KEY);
            properties.saveAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
                return;
              }
              reject(new Error(result.error.message || "Could not clear the ticket selection."));
            });
          })
      )
    );
  }

  await Promise.all(operations);
}

async function saveComposeEmailAttachmentState(
  item: ComposeMetadataItem,
  state: ComposeEmailAttachmentPrefetchState
): Promise<void> {
  const value = JSON.stringify({
    ...state,
    emailAttachmentSummary: attachmentSummaryForOutlookStorage(state.emailAttachmentSummary),
  });
  const operations: Promise<void>[] = [];
  if (item.sessionData) {
    operations.push(
      new Promise((resolve, reject) => {
        item.sessionData!.setAsync(COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY, value, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error(result.error.message || "Could not save attachment preparation."));
          }
        });
      })
    );
  }
  if (item.loadCustomPropertiesAsync) {
    operations.push(
      loadComposeCustomProperties(item).then(
        (properties) =>
          new Promise((resolve, reject) => {
            properties.set(COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY, value);
            properties.saveAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
              } else {
                reject(
                  new Error(result.error.message || "Could not persist attachment preparation.")
                );
              }
            });
          })
      )
    );
  }
  await Promise.all(operations);
}

async function readComposeEmailAttachmentState(
  item: ComposeMetadataItem
): Promise<ComposeEmailAttachmentPrefetchState | null> {
  let value: unknown = null;
  if (item.sessionData) {
    value = await new Promise((resolve) => {
      item.sessionData!.getAsync(COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY, (result) =>
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null)
      );
    });
  }
  if (!value && item.loadCustomPropertiesAsync) {
    value = await loadComposeCustomProperties(item)
      .then((properties) => properties.get(COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY))
      .catch(() => null);
  }
  if (typeof value !== "string" || !value) {
    return null;
  }
  try {
    const state = JSON.parse(value) as Partial<ComposeEmailAttachmentPrefetchState>;
    if (
      state.version !== 2 ||
      !state.draftItemId ||
      !state.ticketId ||
      !state.operationId ||
      !state.emailAttachmentFingerprint ||
      (state.emailAttachmentDecision !== "include" && state.emailAttachmentDecision !== "exclude")
    ) {
      return null;
    }
    return {
      version: 2,
      draftItemId: String(state.draftItemId),
      ticketId: String(state.ticketId),
      operationId: String(state.operationId),
      emailAttachmentDecision: state.emailAttachmentDecision,
      emailAttachmentFingerprint: String(state.emailAttachmentFingerprint),
      emailAttachmentPrefetchKey: String(state.emailAttachmentPrefetchKey || ""),
      emailAttachmentStagingVersion: 2,
      emailAttachmentSummary:
        normalizeEmailAttachmentSummary(state.emailAttachmentSummary) ||
        createEmptyEmailAttachmentSummary(),
    };
  } catch {
    return null;
  }
}

async function clearComposeEmailAttachmentState(item: ComposeMetadataItem): Promise<void> {
  const operations: Promise<void>[] = [];
  if (item.sessionData) {
    operations.push(
      new Promise((resolve) => {
        item.sessionData!.removeAsync(COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY, () => resolve());
      })
    );
  }
  if (item.loadCustomPropertiesAsync) {
    operations.push(
      loadComposeCustomProperties(item).then(
        (properties) =>
          new Promise((resolve) => {
            properties.remove(COMPOSE_EMAIL_ATTACHMENT_STORAGE_KEY);
            properties.saveAsync(() => resolve());
          })
      )
    );
  }
  await Promise.all(operations);
}

function loadComposeCustomProperties(item: ComposeMetadataItem): Promise<ComposeCustomProperties> {
  return new Promise((resolve, reject) => {
    if (!item.loadCustomPropertiesAsync) {
      reject(new Error("Outlook does not support saved item metadata."));
      return;
    }

    item.loadCustomPropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }

      reject(new Error(result.error.message || "Could not load saved item metadata."));
    });
  });
}

function parseComposeAttachMarker(value: unknown): ComposeAttachMarker | null {
  if (typeof value !== "string" || !value) {
    return null;
  }

  try {
    const marker = JSON.parse(value) as Partial<ComposeAttachMarker>;
    const isCreation = marker.version === 4 && marker.destinationKind === "create-ticket";
    if (
      (marker.version !== 1 &&
        marker.version !== 2 &&
        marker.version !== 3 &&
        marker.version !== 4) ||
      !marker.composeAttachId ||
      !marker.draftItemId ||
      (isCreation
        ? !marker.creationOperationId || !marker.ticketTypeId
        : !marker.ticketId || !marker.ticketNumber)
    ) {
      return null;
    }

    return {
      version: marker.version,
      destinationKind: isCreation ? "create-ticket" : "existing-ticket",
      backgroundSessionId: String(marker.backgroundSessionId || ""),
      composeAttachId: String(marker.composeAttachId),
      ticketId: String(marker.ticketId || ""),
      ticketNumber: String(marker.ticketNumber || ""),
      ticketSummary: String(marker.ticketSummary || "Untitled ticket").slice(0, 500),
      draftItemId: String(marker.draftItemId),
      creationOperationId: String(marker.creationOperationId || ""),
      ticketTypeId: String(marker.ticketTypeId || ""),
      ticketTypeName: String(marker.ticketTypeName || ""),
      inlineImageFingerprint: String(marker.inlineImageFingerprint || ""),
      inlineImagePrefetchKey: String(marker.inlineImagePrefetchKey || ""),
      emailAttachmentDecision:
        marker.emailAttachmentDecision === "include" || marker.emailAttachmentDecision === "exclude"
          ? marker.emailAttachmentDecision
          : undefined,
      emailAttachmentFingerprint: String(marker.emailAttachmentFingerprint || ""),
      emailAttachmentPrefetchKey: String(marker.emailAttachmentPrefetchKey || ""),
      emailAttachmentStagingVersion:
        Number(marker.emailAttachmentStagingVersion) === 2 ? 2 : undefined,
      emailAttachmentSummary: normalizeEmailAttachmentSummary(marker.emailAttachmentSummary),
      mailboxEmail: String(marker.mailboxEmail || ""),
      actionMode: normalizeActionMode(marker.actionMode),
      composeIdentityHeader: marker.composeIdentityHeader === true,
    };
  } catch {
    return null;
  }
}

function normalizeEmailAttachmentSummary(value: unknown): EmailAttachmentSummary | undefined {
  if (!value || typeof value !== "object") {
    return undefined;
  }
  const summary = value as Partial<EmailAttachmentSummary>;
  const count = (entry: unknown) => {
    const number = Number(entry);
    return Number.isInteger(number) && number >= 0 ? number : 0;
  };
  return {
    attached: count(summary.attached),
    detected: count(summary.detected),
    failed: count(summary.failed),
    selected: count(summary.selected),
    skipped: count(summary.skipped),
    prepared: count(
      summary.prepared !== undefined
        ? summary.prepared
        : (summary as unknown as { uploaded?: unknown }).uploaded
    ),
    warnings: Array.isArray(summary.warnings)
      ? summary.warnings.map((warning) => String(warning)).slice(0, 10)
      : [],
  };
}

function createEmptyEmailAttachmentSummary(): EmailAttachmentSummary {
  return {
    attached: 0,
    detected: 0,
    failed: 0,
    prepared: 0,
    selected: 0,
    skipped: 0,
    warnings: [],
  };
}

function createComposeAttachId(): string {
  const randomPart = Math.random().toString(36).slice(2);
  return `${Date.now().toString(36)}-${randomPart}-${Math.random().toString(36).slice(2)}`;
}

function renderComposeAttachSelection() {
  const panel = getComposeSelectionPanel();
  const marker = activeComposeAttachMarker;
  panel.hidden = !marker;

  if (marker) {
    setActionMode(normalizeActionMode(marker.actionMode));
    getComposeSelectionPanel().querySelector(".halo-results__kicker")!.textContent =
      isCreateTicketMarker(marker) ? "Create on send" : "Attach on send";
    getComposeSelectionTicket().textContent = isCreateTicketMarker(marker)
      ? marker.ticketTypeName || "New Halo ticket"
      : marker.ticketNumber;
    getComposeSelectionSummary().textContent = marker.ticketSummary;
    renderComposeAttachmentStatus(marker);
    selectTicketById(marker.ticketId);
  } else {
    setActionMode("email");
    getComposeSelectionTicket().textContent = "No ticket selected";
    getComposeSelectionSummary().textContent = "";
    getComposeAttachmentStatus().hidden = true;
    selectTicketById("");
  }
}

function renderComposeAttachmentStatus(marker: ComposeAttachMarker) {
  const status = getComposeAttachmentStatus();
  const summary = marker.emailAttachmentSummary;
  if (marker.emailAttachmentDecision !== "include" || !summary || !summary.detected) {
    status.hidden = true;
    return;
  }
  status.hidden = false;
  if (summary.failed > 0) {
    status.dataset.state = "failed";
    status.textContent = "Attachment preparation failed — retry before sending";
    return;
  }
  if (
    marker.emailAttachmentStagingVersion !== 2 ||
    !marker.emailAttachmentPrefetchKey ||
    summary.prepared < summary.selected
  ) {
    status.dataset.state = "pending";
    status.textContent = `Preparing ${summary.selected} attachment${summary.selected === 1 ? "" : "s"}…`;
    return;
  }
  if (summary.skipped > 0) {
    status.dataset.state = "warning";
    status.textContent = `${summary.prepared} ready, ${summary.skipped} unsupported`;
    return;
  }
  status.dataset.state = "ready";
  status.textContent = `${summary.prepared} attachment${summary.prepared === 1 ? "" : "s"} ready for Halo`;
}

function isComposeAttachmentReady(marker: ComposeAttachMarker): boolean {
  const summary = marker.emailAttachmentSummary;
  if (marker.emailAttachmentDecision !== "include" || !summary || !summary.selected) {
    return true;
  }
  return Boolean(
    marker.emailAttachmentStagingVersion === 2 &&
    marker.emailAttachmentPrefetchKey &&
    summary.failed === 0 &&
    summary.prepared === summary.selected
  );
}

function updateComposeModeCopy(isComposeMode: boolean) {
  const subtitle = document.getElementById("page-subtitle");
  const searchHelp = document.getElementById("ticket-search-help");

  if (subtitle) {
    subtitle.textContent = isComposeMode
      ? "Choose a HaloPSA ticket, then send the email when it is ready."
      : "Add the open Outlook email to the right HaloPSA ticket.";
  }

  if (searchHelp) {
    searchHelp.textContent = isComposeMode
      ? "Search by ticket subject or reference, then select where this email should be attached when sent."
      : "Search by ticket subject, reference, or other searchable Halo content, then narrow the results with filters.";
  }
}

function selectTicketById(ticketId: string) {
  const ticketButtons = document.querySelectorAll(".halo-ticket");
  for (let index = 0; index < ticketButtons.length; index += 1) {
    const button = ticketButtons[index] as HTMLElement;
    button.dataset.selected = String(Boolean(ticketId && button.dataset.ticketId === ticketId));
  }
}

function selectTicketButton(selectedButton: HTMLElement) {
  const ticketButtons = document.querySelectorAll(".halo-ticket");
  for (let index = 0; index < ticketButtons.length; index += 1) {
    (ticketButtons[index] as HTMLElement).dataset.selected = "false";
  }

  selectedButton.dataset.selected = "true";
}

function formatTicketTitle(ticket: HaloTicket): string {
  const label = ticket.ticketNumber || ticket.id;
  const summary = ticket.summary || "Untitled ticket";
  return label ? `${label} - ${summary}` : summary;
}

function clearTickets() {
  cancelTicketSearchDebounce();
  ticketRequestRevision += 1;
  ticketQueryIsBusy = false;
  loadedTickets = [];
  lastExecutedTicketQuery = "";
  getTicketsPanel().hidden = true;
  getTicketsEmpty().hidden = true;
  clearTicketList();
  getTicketQueryInput().value = "";
  getTicketOwnershipSelect().value = "mine";
  getTicketLifecycleSelect().value = "open";
  updateTicketFacetOptions([]);
  updateClearSearchButton();
  updateActiveFilterCount();
  getTicketResultsCount().textContent = "";
  getTicketResultsHeading().textContent = "My open tickets";
  getCreateTicketPanel().hidden = true;
  placeActionModeControl("existing");
  getAttachExistingTab().setAttribute("aria-selected", "true");
  getCreateTicketTab().setAttribute("aria-selected", "false");
  document.querySelectorAll(".halo-attach-existing-panel").forEach((element) => {
    (element as HTMLElement).hidden = false;
  });
}

function clearTicketList() {
  const ticketsList = getTicketsList();
  while (ticketsList.firstChild) {
    ticketsList.removeChild(ticketsList.firstChild);
  }
}

async function readCurrentOutlookEmailSnapshot(
  options: { suppressUnsupported?: boolean } = {}
): Promise<OutlookEmailSnapshot | null> {
  const itemRevision = outlookItemRevision;
  const item = Office.context.mailbox.item as unknown as {
    body?: {
      getAsync: (
        coercionType: Office.CoercionType,
        callback: (result: Office.AsyncResult<string>) => void
      ) => void;
    };
    cc?: Office.EmailAddressDetails[];
    conversationId?: string;
    dateTimeCreated?: Date;
    from?: Office.EmailAddressDetails;
    getAttachmentsAsync?: (callback: (result: Office.AsyncResult<unknown[]>) => void) => void;
    getAttachmentContentAsync?: (
      attachmentId: string,
      callback: (result: Office.AsyncResult<{ content: string; format: unknown }>) => void
    ) => void;
    getAllInternetHeadersAsync?: (callback: (result: Office.AsyncResult<string>) => void) => void;
    internetMessageId?: string;
    itemId?: string;
    itemType?: Office.MailboxEnums.ItemType | string;
    normalizedSubject?: string;
    subject?: string;
    to?: Office.EmailAddressDetails[];
  };

  if (
    !item ||
    item.itemType !== Office.MailboxEnums.ItemType.Message ||
    !item.internetMessageId ||
    !item.itemId ||
    !item.body
  ) {
    if (options.suppressUnsupported) {
      return null;
    }

    throw createHaloAuthError("Open an existing received email, then choose a Halo ticket.");
  }

  const body = await readMessageBody(item);
  if (!body.bodyHtml && !body.bodyText) {
    if (options.suppressUnsupported) {
      return null;
    }

    throw createHaloAuthError("Could not read an email body to attach.");
  }

  const inlineImages = prepareOutlookInlineImages(item, body.bodyHtml, item.itemId).catch(() =>
    emptyPreparedInlineImages()
  );
  const emailAttachments = collectEmailAttachmentMetadata(item, body.bodyHtml).catch(() =>
    emptyEmailAttachmentMetadata()
  );
  const internetHeaders = await readInternetHeaders(item);
  if (itemRevision !== outlookItemRevision) {
    throw createHaloAuthError("The selected Outlook item changed while it was being read.");
  }
  const userProfile = Office.context.mailbox.userProfile;

  const email: OutlookEmailPayload = {
    ...body,
    ...emptyPreparedInlineImages(),
    cc: normalizeEmailAddressList(item.cc),
    conversationId: item.conversationId || "",
    dateTimeCreated: item.dateTimeCreated
      ? item.dateTimeCreated.toISOString()
      : new Date().toISOString(),
    from: normalizeEmailAddress(item.from),
    inReplyToMessageIds: extractHeaderMessageIds(internetHeaders, "In-Reply-To"),
    internetHeaders,
    internetMessageId: item.internetMessageId,
    itemId: item.itemId,
    mailboxEmail: userProfile && userProfile.emailAddress ? userProfile.emailAddress : "",
    normalizedSubject: item.normalizedSubject || "",
    referenceMessageIds: extractHeaderMessageIds(internetHeaders, "References"),
    subject: item.subject || item.normalizedSubject || "",
    timeZone: getClientTimeZone(),
    to: normalizeEmailAddressList(item.to),
  };

  return { email, emailAttachments, inlineImages, itemRevision };
}

function applyPreparedInlineImages(
  email: OutlookEmailPayload,
  inlineImages: PreparedInlineImages
): OutlookEmailPayload {
  return {
    ...email,
    inlineImageFingerprint: inlineImages.inlineImageFingerprint,
    inlineImageRefs: inlineImages.inlineImageRefs,
    inlineImageTimings: inlineImages.inlineImageTimings,
    inlineImageUploads: inlineImages.inlineImageUploads,
  };
}

async function warmCurrentInlineImages(): Promise<void> {
  const item = Office.context.mailbox.item as unknown as {
    body?: {
      getAsync: (
        coercionType: Office.CoercionType,
        callback: (result: Office.AsyncResult<string>) => void
      ) => void;
    };
    itemId?: string;
    itemType?: Office.MailboxEnums.ItemType | string;
  };
  if (
    !item ||
    item.itemType !== Office.MailboxEnums.ItemType.Message ||
    !item.itemId ||
    !item.body
  ) {
    return;
  }
  const bodyHtml = await getBodyAsync(item, Office.CoercionType.Html).catch(() => "");
  if (bodyHtml) {
    await prepareOutlookInlineImages(item, bodyHtml, item.itemId).catch(() => undefined);
  }
}

function emptyPreparedInlineImages(): PreparedInlineImages {
  return {
    inlineImageFingerprint: "",
    inlineImageRefs: [],
    inlineImageUploads: [],
    inlineImageWarnings: [],
    inlineImageTimings: { hashingMs: 0, outlookReadMs: 0 },
  };
}

function emptyEmailAttachmentMetadata(): PreparedEmailAttachmentMetadata {
  return {
    attachments: [],
    emailAttachmentFingerprint: "",
    overCount: 0,
    reportedOversized: 0,
    reportedTotalBytes: 0,
  };
}

function getClientTimeZone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || "";
  } catch {
    return "";
  }
}

async function readMessageBody(item: {
  body?: {
    getAsync: (
      coercionType: Office.CoercionType,
      callback: (result: Office.AsyncResult<string>) => void
    ) => void;
  };
}): Promise<{ bodyHtml: string; bodyText: string }> {
  try {
    const bodyHtml = await getBodyAsync(item, Office.CoercionType.Html);
    return { bodyHtml, bodyText: "" };
  } catch {
    const bodyText = await getBodyAsync(item, Office.CoercionType.Text);
    return { bodyHtml: "", bodyText };
  }
}

function getBodyAsync(
  item: {
    body?: {
      getAsync: (
        coercionType: Office.CoercionType,
        callback: (result: Office.AsyncResult<string>) => void
      ) => void;
    };
  },
  coercionType: Office.CoercionType
): Promise<string> {
  return new Promise((resolve, reject) => {
    if (!item.body) {
      reject(new Error("Could not read the selected email body."));
      return;
    }

    item.body.getAsync(coercionType, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
        return;
      }

      reject(new Error(result.error.message || "Could not read the selected email body."));
    });
  });
}

function readInternetHeaders(item: {
  getAllInternetHeadersAsync?: (callback: (result: Office.AsyncResult<string>) => void) => void;
}): Promise<string> {
  return new Promise((resolve) => {
    if (!item.getAllInternetHeadersAsync) {
      resolve("");
      return;
    }

    item.getAllInternetHeadersAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
        return;
      }

      resolve("");
    });
  });
}

async function recoverReadModeComposeAttachIds(email: OutlookEmailPayload): Promise<string[]> {
  const directIds = extractComposeAttachIdsFromRawHeaders(email.internetHeaders);
  const messageIds = uniqueRecoveryMessageIds([
    ...email.inReplyToMessageIds,
    ...email.referenceMessageIds,
  ]).slice(0, 12);

  if (!messageIds.length || !canMakeReadModeEwsRequest()) {
    return directIds;
  }

  const ewsRecovery = recoverReadModeComposeAttachIdsFromEws(messageIds).catch(() => []);
  const timeout = new Promise<string[]>((resolve) => {
    window.setTimeout(() => resolve([]), READ_MAPPING_RECOVERY_TIMEOUT_MS);
  });
  const recoveredIds = await Promise.race([ewsRecovery, timeout]);
  return uniqueOpaqueComposeAttachIds([...directIds, ...recoveredIds]);
}

async function recoverReadModeComposeAttachIdsFromEws(messageIds: string[]): Promise<string[]> {
  const findResponse = await makeReadModeEwsRequest(
    buildReadModeFindItemRequest(messageIds, ["sentitems", "inbox"])
  );
  const itemIds = extractReadModeEwsItemIds(findResponse).slice(0, 12);
  if (!itemIds.length) {
    return [];
  }

  const metadataResponse = await makeReadModeEwsRequest(
    buildReadModeMetadataGetItemRequest(itemIds)
  );
  return uniqueOpaqueComposeAttachIds([
    ...extractComposeAttachIdsFromEwsHeaders(metadataResponse),
    ...extractComposeAttachIdsFromEwsCustomProperties(metadataResponse),
  ]);
}

function canMakeReadModeEwsRequest(): boolean {
  const mailbox = Office.context.mailbox as unknown as {
    makeEwsRequestAsync?: (
      request: string,
      callback: (result: Office.AsyncResult<string>) => void
    ) => void;
  };
  return typeof mailbox.makeEwsRequestAsync === "function";
}

function makeReadModeEwsRequest(request: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const mailbox = Office.context.mailbox as unknown as {
      makeEwsRequestAsync?: (
        request: string,
        callback: (result: Office.AsyncResult<string>) => void
      ) => void;
    };
    if (!mailbox.makeEwsRequestAsync) {
      reject(new Error("EWS mapping recovery is unavailable."));
      return;
    }
    mailbox.makeEwsRequestAsync(request, (result) => {
      if (
        result.status !== Office.AsyncResultStatus.Succeeded ||
        !result.value ||
        String(result.value).length > 1024 * 1024
      ) {
        reject(new Error("EWS mapping recovery failed."));
        return;
      }
      resolve(String(result.value));
    });
  });
}

function buildReadModeEwsEnvelope(body: string): string {
  return (
    '<?xml version="1.0"?>' +
    '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
    `<soap:Body>${body}</soap:Body></soap:Envelope>`
  );
}

function buildReadModeFindItemRequest(messageIds: string[], folders: string[]): string {
  const restrictions = messageIds.map(
    (messageId) =>
      '<t:IsEqualTo><t:FieldURI FieldURI="message:InternetMessageId"/>' +
      `<t:FieldURIOrConstant><t:Constant Value="${escapeReadModeEwsXml(
        messageId
      )}"/></t:FieldURIOrConstant></t:IsEqualTo>`
  );
  const restriction =
    restrictions.length === 1 ? restrictions[0] : `<t:Or>${restrictions.join("")}</t:Or>`;
  const parentFolders = folders
    .map((folder) => `<t:DistinguishedFolderId Id="${escapeReadModeEwsXml(folder)}"/>`)
    .join("");
  return buildReadModeEwsEnvelope(
    '<m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape></m:ItemShape>' +
      `<m:Restriction>${restriction}</m:Restriction>` +
      `<m:ParentFolderIds>${parentFolders}</m:ParentFolderIds></m:FindItem>`
  );
}

function buildReadModeMetadataGetItemRequest(itemIds: string[]): string {
  const ids = itemIds.map((itemId) => `<t:ItemId Id="${escapeReadModeEwsXml(itemId)}"/>`).join("");
  return buildReadModeEwsEnvelope(
    "<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>" +
      '<t:FieldURI FieldURI="item:InternetMessageHeaders"/>' +
      '<t:ExtendedFieldURI DistinguishedPropertySetId="PublicStrings" PropertyName="' +
      COMPOSE_CUSTOM_PROPERTY_NAME +
      '" PropertyType="String"/>' +
      `</t:AdditionalProperties></m:ItemShape><m:ItemIds>${ids}</m:ItemIds></m:GetItem>`
  );
}

function extractReadModeEwsItemIds(xml: string): string[] {
  const values: string[] = [];
  const pattern = /<(?:\w+:)?ItemId\b[^>]*\bId="([^"]+)"/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) && values.length < 12) {
    values.push(unescapeReadModeEwsXml(match[1]));
  }
  return values;
}

function extractComposeAttachIdsFromRawHeaders(headers: string): string[] {
  const value = getInternetHeaderValue(headers, COMPOSE_ATTACH_HEADER_NAME);
  return uniqueOpaqueComposeAttachIds(value ? [value] : []);
}

function extractComposeAttachIdsFromEwsHeaders(xml: string): string[] {
  const values: string[] = [];
  const pattern =
    /<(?:\w+:)?InternetMessageHeader\b[^>]*HeaderName="X-Halo-Compose-Id"[^>]*>([\s\S]*?)<\/(?:\w+:)?InternetMessageHeader>/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) && values.length < 12) {
    values.push(unescapeReadModeEwsXml(match[1]).trim());
  }
  return uniqueOpaqueComposeAttachIds(values);
}

function extractComposeAttachIdsFromEwsCustomProperties(xml: string): string[] {
  const values: string[] = [];
  const pattern = new RegExp(
    '<(?:\\w+:)?ExtendedProperty>[\\s\\S]*?PropertyName="' +
      escapeRegExp(COMPOSE_CUSTOM_PROPERTY_NAME) +
      '"[\\s\\S]*?<(?:\\w+:)?Value>([\\s\\S]*?)<\\/(?:\\w+:)?Value>[\\s\\S]*?<\\/(?:\\w+:)?ExtendedProperty>',
    "gi"
  );
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(xml)) && values.length < 12) {
    try {
      const properties = JSON.parse(unescapeReadModeEwsXml(match[1])) as Record<string, unknown>;
      const markerValue = properties && properties[COMPOSE_ATTACH_STORAGE_KEY];
      const marker =
        typeof markerValue === "string"
          ? (JSON.parse(markerValue) as { composeAttachId?: unknown })
          : (markerValue as { composeAttachId?: unknown } | undefined);
      if (marker?.composeAttachId) {
        values.push(String(marker.composeAttachId));
      }
    } catch {
      // Ignore copied or malformed custom properties.
    }
  }
  return uniqueOpaqueComposeAttachIds(values);
}

function uniqueOpaqueComposeAttachIds(values: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const rawValue of values) {
    const value = String(rawValue || "").trim();
    if (value && value.length <= 200 && /^[A-Za-z0-9_-]+$/.test(value) && !seen.has(value)) {
      seen.add(value);
      result.push(value);
    }
    if (result.length >= 12) {
      break;
    }
  }
  return result;
}

function uniqueRecoveryMessageIds(values: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const rawValue of values) {
    let value = String(rawValue || "").trim();
    if (value && !value.startsWith("<")) {
      value = `<${value.replace(/^<|>$/g, "")}>`;
    }
    const key = value.toLowerCase();
    if (value && !seen.has(key)) {
      seen.add(key);
      result.push(value);
    }
    if (result.length >= 24) {
      break;
    }
  }
  return result;
}

function escapeReadModeEwsXml(value: string): string {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function unescapeReadModeEwsXml(value: string): string {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function extractHeaderMessageIds(headers: string, headerName: string): string[] {
  const headerValue = getInternetHeaderValue(headers, headerName);

  if (!headerValue) {
    return [];
  }

  const bracketedIds = headerValue.match(/<[^>]+>/g);
  const candidates = bracketedIds && bracketedIds.length ? bracketedIds : headerValue.split(/\s+/);
  const seen: { [key: string]: boolean } = {};
  const messageIds: string[] = [];

  candidates.forEach((candidate) => {
    const messageId = candidate.trim();
    const key = messageId.toLowerCase();

    if (messageId && !seen[key]) {
      seen[key] = true;
      messageIds.push(messageId);
    }
  });

  return messageIds;
}

function getInternetHeaderValue(headers: string, headerName: string): string {
  const unfoldedHeaders = headers.replace(/\r?\n[ \t]+/g, " ");
  const headerPattern = new RegExp(`^${escapeRegExp(headerName)}:\\s*(.*)$`, "im");
  const match = headerPattern.exec(unfoldedHeaders);

  return match ? match[1].trim() : "";
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function normalizeEmailAddressList(value?: Office.EmailAddressDetails[]): EmailAddressPayload[] {
  return (value || [])
    .map((entry) => normalizeEmailAddress(entry))
    .filter((entry): entry is EmailAddressPayload => Boolean(entry));
}

function normalizeEmailAddress(value?: Office.EmailAddressDetails): EmailAddressPayload | null {
  if (!value) {
    return null;
  }

  return {
    displayName: value.displayName || "",
    emailAddress: value.emailAddress || "",
  };
}

function getLoginButton(): HTMLButtonElement {
  return document.getElementById("login-button") as HTMLButtonElement;
}

function getLogoutButton(): HTMLButtonElement {
  return document.getElementById("logout-button") as HTMLButtonElement;
}

function getReportBugButton(): HTMLButtonElement {
  return document.getElementById("report-bug-button") as HTMLButtonElement;
}

function getRefreshTicketsButton(): HTMLButtonElement {
  return document.getElementById("refresh-tickets-button") as HTMLButtonElement;
}

function getConnectionPanel(): HTMLElement {
  return document.getElementById("connection-panel") as HTMLElement;
}

function getTicketSearchForm(): HTMLFormElement {
  return document.getElementById("ticket-search-form") as HTMLFormElement;
}

function getTicketQueryInput(): HTMLInputElement {
  return document.getElementById("ticket-query") as HTMLInputElement;
}

function getSearchTicketsButton(): HTMLButtonElement {
  return document.getElementById("search-tickets-button") as HTMLButtonElement;
}

function getClearSearchButton(): HTMLButtonElement {
  return document.getElementById("clear-search-button") as HTMLButtonElement;
}

function getTicketOwnershipSelect(): HTMLSelectElement {
  return document.getElementById("ticket-ownership") as HTMLSelectElement;
}

function getTicketLifecycleSelect(): HTMLSelectElement {
  return document.getElementById("ticket-lifecycle") as HTMLSelectElement;
}

function getTicketCustomerSelect(): HTMLSelectElement {
  return document.getElementById("ticket-customer") as HTMLSelectElement;
}

function getTicketAssigneeSelect(): HTMLSelectElement {
  return document.getElementById("ticket-assignee") as HTMLSelectElement;
}

function getResetTicketFiltersButton(): HTMLButtonElement {
  return document.getElementById("reset-ticket-filters-button") as HTMLButtonElement;
}

function getActiveFilterCount(): HTMLElement {
  return document.getElementById("active-filter-count") as HTMLElement;
}

function getTicketResultsCount(): HTMLElement {
  return document.getElementById("ticket-results-count") as HTMLElement;
}

function getTicketResultsHeading(): HTMLElement {
  return document.getElementById("tickets-heading") as HTMLElement;
}

function getTicketsPanel(): HTMLElement {
  return document.getElementById("tickets-panel") as HTMLElement;
}

function getTicketsEmpty(): HTMLElement {
  return document.getElementById("tickets-empty") as HTMLElement;
}

function getTicketsList(): HTMLElement {
  return document.getElementById("tickets-list") as HTMLElement;
}

function getComposeSelectionPanel(): HTMLElement {
  return document.getElementById("compose-selection-panel") as HTMLElement;
}

function getComposeSelectionTicket(): HTMLElement {
  return document.getElementById("compose-selection-ticket") as HTMLElement;
}

function getComposeSelectionSummary(): HTMLElement {
  return document.getElementById("compose-selection-summary") as HTMLElement;
}

function getComposeAttachmentStatus(): HTMLElement {
  return document.getElementById("compose-attachment-status") as HTMLElement;
}

function getRemoveComposeSelectionButton(): HTMLButtonElement {
  return document.getElementById("remove-compose-selection-button") as HTMLButtonElement;
}

function getPrivateNoteToggle(): HTMLInputElement {
  return document.getElementById("private-note-toggle") as HTMLInputElement;
}

function placeActionModeControl(destination: "existing" | "create") {
  const control = document.getElementById("action-mode-control") as HTMLElement;
  const slotId = destination === "create" ? "create-action-mode-slot" : "ticket-action-mode-slot";
  const slot = document.getElementById(slotId) as HTMLElement;
  if (control.parentElement !== slot) {
    slot.appendChild(control);
  }
}

function getActionModeHelp(): HTMLElement {
  return document.getElementById("action-mode-help") as HTMLElement;
}

function getActionModeVisibility(): HTMLElement {
  return document.getElementById("action-mode-visibility") as HTMLElement;
}

function getMappedEmailConfirmation(): HTMLElement {
  return document.getElementById("mapped-email-confirmation") as HTMLElement;
}

function getMappedEmailTicket(): HTMLElement {
  return document.getElementById("mapped-email-ticket") as HTMLElement;
}

function getAttachMappedEmailButton(): HTMLButtonElement {
  return document.getElementById("attach-mapped-email-button") as HTMLButtonElement;
}

function getAttachExistingTab(): HTMLButtonElement {
  return document.getElementById("attach-existing-tab") as HTMLButtonElement;
}

function getCreateTicketTab(): HTMLButtonElement {
  return document.getElementById("create-ticket-tab") as HTMLButtonElement;
}

function getCreateTicketPanel(): HTMLElement {
  return document.getElementById("create-ticket-panel") as HTMLElement;
}

function getRefreshTicketTypesButton(): HTMLButtonElement {
  return document.getElementById("refresh-ticket-types-button") as HTMLButtonElement;
}

function getCreateTicketForm(): HTMLFormElement {
  return document.getElementById("create-ticket-form") as HTMLFormElement;
}

function getCreateTicketTypeSelect(): HTMLSelectElement {
  return document.getElementById("create-ticket-type") as HTMLSelectElement;
}

function getCreateTicketSummaryInput(): HTMLInputElement {
  return document.getElementById("create-ticket-summary") as HTMLInputElement;
}

function getCreateTicketRequesterQuery(): HTMLInputElement {
  return document.getElementById("create-ticket-requester-query") as HTMLInputElement;
}

function getSearchRequestersButton(): HTMLButtonElement {
  return document.getElementById("search-requesters-button") as HTMLButtonElement;
}

function getCreateTicketRequesterSelect(): HTMLSelectElement {
  return document.getElementById("create-ticket-requester") as HTMLSelectElement;
}

function getCreateTicketRequesterHelp(): HTMLElement {
  return document.getElementById("create-ticket-requester-help") as HTMLElement;
}

function getCreateTicketRequiredFields(): HTMLElement {
  return document.getElementById("create-ticket-required-fields") as HTMLElement;
}

function getCreateTicketOptionalDetails(): HTMLElement {
  return document.getElementById("create-ticket-optional-details") as HTMLElement;
}

function getCreateTicketOptionalFields(): HTMLElement {
  return document.getElementById("create-ticket-optional-fields") as HTMLElement;
}

function getCreateTicketSchemaWarning(): HTMLElement {
  return document.getElementById("create-ticket-schema-warning") as HTMLElement;
}

function getCreateTicketConversationWarning(): HTMLElement {
  return document.getElementById("create-ticket-conversation-warning") as HTMLElement;
}

function renderCreateTicketConversationWarning() {
  const warning = getCreateTicketConversationWarning();
  getCreateTicketTab().textContent = mappedConversationTicket
    ? "Create another ticket"
    : "Create new ticket";
  warning.hidden = !mappedConversationTicket;
  warning.textContent = mappedConversationTicket
    ? `This conversation is mapped to ticket ${
        mappedConversationTicket.ticketNumber || mappedConversationTicket.ticketId
      }. Creating another ticket will move future conversation attachments to the new ticket.`
    : "";
}

function getSubmitCreateTicketButton(): HTMLButtonElement {
  return document.getElementById("submit-create-ticket-button") as HTMLButtonElement;
}

function getAttachmentPrompt(): HTMLElement {
  return document.getElementById("attachment-prompt") as HTMLElement;
}

function getAttachmentPromptDetail(): HTMLElement {
  return document.getElementById("attachment-prompt-detail") as HTMLElement;
}

function getAttachmentPromptLimits(): HTMLElement {
  return document.getElementById("attachment-prompt-limits") as HTMLElement;
}

function getAttachmentAddButton(): HTMLButtonElement {
  return document.getElementById("attachment-add-button") as HTMLButtonElement;
}

function getAttachmentEmailOnlyButton(): HTMLButtonElement {
  return document.getElementById("attachment-email-only-button") as HTMLButtonElement;
}
