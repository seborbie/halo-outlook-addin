/* global document, Office, fetch, localStorage, RequestInit, HTMLInputElement, HTMLButtonElement, HTMLElement */

const CONNECTION_STORAGE_KEY = "halo-auth-connection-v1";
const BACKGROUND_SESSION_STORAGE_KEY = "halo-auth-background-session-v1";

type AuthStartResponse = {
  dialogUrl: string;
};

type AuthStatusResponse = {
  authenticated: boolean;
  haloUrl?: string;
  expiresAt?: string;
};

type AuthCompleteResponse = {
  authenticated: boolean;
  backgroundSessionId?: string;
  expiresAt?: string;
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
  client: string;
  agent: string;
};

type HaloTicketsResponse = {
  ok: boolean;
  tickets: HaloTicket[];
  message?: string;
  error?: string;
  debug?: unknown;
};

type EmailAddressPayload = {
  displayName: string;
  emailAddress: string;
};

type OutlookEmailPayload = {
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
};

type HaloAttachEmailResponse = {
  ok: boolean;
  attachMode?: "full-chain" | "latest-reply";
  message: string;
  actionId?: string;
  backgroundSessionId?: string;
  error?: string;
  debug?: unknown;
};

type HaloAutoAttachResponse = {
  ok: boolean;
  status: "attached" | "already-attached" | "no-match";
  ticketId?: string;
  ticketNumber?: string;
  message?: string;
  actionId?: string;
  error?: string;
  debug?: unknown;
};

type HaloAuthError = Error & {
  debug?: unknown;
};

type StoredConnection = {
  haloUrl: string;
  clientId: string;
};

let currentDialog: Office.Dialog | null = null;
let waitingForDialog = false;
let checkingSession = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    showApp();
    bindControls();
    restoreConnectionSettings();
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
  getRefreshTicketsButton().onclick = () => void loadTickets();
  registerItemChangedHandler();
}

async function checkExistingSession() {
  if (checkingSession || waitingForDialog) {
    return;
  }

  checkingSession = true;

  try {
    const status = await fetchJson<AuthStatusResponse>("/api/auth/status");

    if (!status.authenticated) {
      setSignedOut();
      return;
    }

    setBusy(true);
    setStatus("loading", "Checking Halo API auth...", status.haloUrl || "");
    await refreshBackgroundSessionId();
    await pingHalo();
    if (!(await autoAttachCurrentEmail())) {
      await loadTickets();
    }
  } catch (error) {
    setFailed(error);
  } finally {
    checkingSession = false;
    setBusy(false);
  }
}

async function startHaloLogin() {
  const haloUrl = getHaloUrlInput().value.trim();
  const clientId = getClientIdInput().value.trim();

  if (!haloUrl) {
    setStatus("failed", "Halo API Auth failed", "Enter your Halo URL.");
    return;
  }

  if (!clientId) {
    setStatus("failed", "Halo API Auth failed", "Enter your Halo API application client ID.");
    return;
  }

  try {
    setBusy(true);
    setStatus("loading", "Opening Halo login...", "A Halo sign-in dialog will open.");

    const authStart = await fetchJson<AuthStartResponse>("/api/auth/start", {
      method: "POST",
      body: JSON.stringify({ haloUrl, clientId }),
    });

    openHaloDialog(authStart.dialogUrl);
  } catch (error) {
    setBusy(false);
    setFailed(error);
  }
}

function openHaloDialog(dialogUrl: string) {
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
    saveConnectionSettings();
    await pingHalo();
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

  getLogoutButton().hidden = false;
  setStatus("success", "Halo API Auth works", "The access token was accepted by the Halo API.");
}

async function loadTickets() {
  const ticketsPanel = getTicketsPanel();
  ticketsPanel.hidden = false;
  getTicketsEmpty().hidden = true;
  clearTicketList();

  try {
    setTicketsBusy(true);
    setStatus("loading", "Loading Halo tickets...", "Fetching your open tickets from Halo.");
    const result = await fetchJson<HaloTicketsResponse>("/api/halo/tickets");

    if (!result.ok) {
      throw createHaloAuthError(
        result.error || result.message || "Halo ticket list failed.",
        result.debug
      );
    }

    renderTickets(result.tickets || []);
    setStatus("success", "Halo API Auth works", `${result.tickets.length} open ticket(s) loaded.`);
  } catch (error) {
    if (!getTicketsList().childElementCount) {
      getTicketsPanel().hidden = true;
    }
    setFailed(error, { hideLogout: false });
  } finally {
    setTicketsBusy(false);
  }
}

async function autoAttachCurrentEmail(): Promise<boolean> {
  let email: OutlookEmailPayload | null = null;

  try {
    email = await readCurrentOutlookEmail({ suppressUnsupported: true });
  } catch {
    return false;
  }

  if (!email) {
    return false;
  }

  try {
    setStatus(
      "loading",
      "Checking Halo email mapping...",
      "Looking for an existing ticket link for this email chain."
    );

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

    if (result.status === "no-match") {
      return false;
    }

    clearTickets();
    getLogoutButton().hidden = false;

    if (result.status === "already-attached") {
      setStatus(
        "success",
        "This email is already attached to ticket",
        result.message ||
          `This email is already attached to ticket ${result.ticketNumber || result.ticketId}.`
      );
      return true;
    }

    setStatus(
      "success",
      "Email automatically added to ticket",
      result.message ||
        `Email automatically added to ticket ${result.ticketNumber || result.ticketId}.`
    );
    return true;
  } catch (error) {
    clearTickets();
    setFailed(error, { hideLogout: false, message: "Email auto-attach failed" });
    return true;
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

  mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => void checkExistingSession());
}

async function fetchJson<T>(url: string, options: RequestInit = {}): Promise<T> {
  const response = await fetch(url, {
    credentials: "same-origin",
    ...options,
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  const body = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw createHaloAuthError(
      body.error || body.message || `Request failed with status ${response.status}.`,
      body.debug
    );
  }

  return body as T;
}

function setSignedOut() {
  getLogoutButton().hidden = true;
  clearTickets();
  setStatus("signed-out", "Enter your Halo URL and client ID to start.", "");
}

function setFailed(error: unknown, options: { hideLogout?: boolean; message?: string } = {}) {
  if (options.hideLogout !== false) {
    getLogoutButton().hidden = true;
  }

  const detail = error instanceof Error ? error.message : "Unexpected Halo auth error.";
  setStatus("failed", options.message || "Halo API Auth failed", detail, getErrorDebug(error));
}

function setStatus(
  state: "signed-out" | "loading" | "success" | "failed",
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
  getLoginButton().disabled = isBusy;
  getLogoutButton().disabled = isBusy;
  getRefreshTicketsButton().disabled = isBusy;
  getHaloUrlInput().disabled = isBusy;
  getClientIdInput().disabled = isBusy;
}

function setTicketsBusy(isBusy: boolean) {
  getRefreshTicketsButton().disabled = isBusy;
}

function setTicketButtonsBusy(isBusy: boolean) {
  const ticketButtons = document.querySelectorAll(".halo-auth__ticket");
  for (let index = 0; index < ticketButtons.length; index += 1) {
    (ticketButtons[index] as HTMLButtonElement).disabled = isBusy;
  }
}

function restoreConnectionSettings() {
  try {
    const rawValue = localStorage.getItem(CONNECTION_STORAGE_KEY);
    if (!rawValue) {
      return;
    }

    const stored = JSON.parse(rawValue) as StoredConnection;
    if (stored.haloUrl) {
      getHaloUrlInput().value = stored.haloUrl;
    }

    if (stored.clientId) {
      getClientIdInput().value = stored.clientId;
    }
  } catch {
    localStorage.removeItem(CONNECTION_STORAGE_KEY);
  }
}

function saveConnectionSettings() {
  const stored: StoredConnection = {
    haloUrl: getHaloUrlInput().value.trim(),
    clientId: getClientIdInput().value.trim(),
  };

  localStorage.setItem(CONNECTION_STORAGE_KEY, JSON.stringify(stored));
}

function saveBackgroundSessionId(backgroundSessionId: string): Promise<void> {
  if (!backgroundSessionId || !Office.context.roamingSettings) {
    return Promise.resolve();
  }

  Office.context.roamingSettings.set(BACKGROUND_SESSION_STORAGE_KEY, backgroundSessionId);
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

function renderTickets(tickets: HaloTicket[]) {
  const ticketsPanel = getTicketsPanel();
  const ticketsList = getTicketsList();
  const ticketsEmpty = getTicketsEmpty();

  ticketsPanel.hidden = false;
  clearTicketList();
  ticketsEmpty.hidden = tickets.length > 0;

  tickets.forEach((ticket) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "halo-auth__ticket";
    button.dataset.selected = "false";
    button.onclick = () => void attachEmailToTicket(ticket, button);

    const title = document.createElement("span");
    title.className = "halo-auth__ticket-title";
    title.textContent = formatTicketTitle(ticket);

    const meta = document.createElement("span");
    meta.className = "halo-auth__ticket-meta";
    meta.textContent = formatTicketMeta(ticket);

    button.appendChild(title);
    if (meta.textContent) {
      button.appendChild(meta);
    }

    ticketsList.appendChild(button);
  });
}

async function attachEmailToTicket(ticket: HaloTicket, selectedButton: HTMLElement) {
  if (!ticket.id) {
    setFailed(createHaloAuthError("This ticket does not include a Halo ticket ID."), {
      hideLogout: false,
      message: "Email attach failed",
    });
    return;
  }

  selectTicketButton(selectedButton);

  try {
    setTicketsBusy(true);
    setTicketButtonsBusy(true);
    setStatus("loading", "Attaching email to Halo ticket...", formatTicketTitle(ticket));

    const email = await readCurrentOutlookEmail();
    if (!email) {
      throw createHaloAuthError("Open an existing received email, then choose a Halo ticket.");
    }

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
    const successMessage = result.message || "Email attached to Halo ticket";
    setStatus(
      "success",
      successMessage,
      `Attached ${email.subject || "selected email"} to ${ticket.ticketNumber || ticket.id}.`
    );
  } catch (error) {
    setFailed(error, { hideLogout: false, message: "Email attach failed" });
  } finally {
    setTicketsBusy(false);
    setTicketButtonsBusy(false);
  }
}

function selectTicketButton(selectedButton: HTMLElement) {
  const ticketButtons = document.querySelectorAll(".halo-auth__ticket");
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

function formatTicketMeta(ticket: HaloTicket): string {
  return [ticket.status, ticket.client, ticket.agent].filter(Boolean).join(" | ");
}

function clearTickets() {
  getTicketsPanel().hidden = true;
  getTicketsEmpty().hidden = true;
  clearTicketList();
}

function clearTicketList() {
  const ticketsList = getTicketsList();
  while (ticketsList.firstChild) {
    ticketsList.removeChild(ticketsList.firstChild);
  }
}

async function readCurrentOutlookEmail(
  options: { suppressUnsupported?: boolean } = {}
): Promise<OutlookEmailPayload | null> {
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

  const internetHeaders = await readInternetHeaders(item);
  const userProfile = Office.context.mailbox.userProfile;

  return {
    ...body,
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

function getHaloUrlInput(): HTMLInputElement {
  return document.getElementById("halo-url") as HTMLInputElement;
}

function getClientIdInput(): HTMLInputElement {
  return document.getElementById("halo-client-id") as HTMLInputElement;
}

function getLoginButton(): HTMLButtonElement {
  return document.getElementById("login-button") as HTMLButtonElement;
}

function getLogoutButton(): HTMLButtonElement {
  return document.getElementById("logout-button") as HTMLButtonElement;
}

function getRefreshTicketsButton(): HTMLButtonElement {
  return document.getElementById("refresh-tickets-button") as HTMLButtonElement;
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
