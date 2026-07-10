/* global document, Office, fetch, RequestInit, Response, HTMLInputElement, HTMLButtonElement, HTMLFormElement, HTMLElement, SVGElement */
import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

const BACKGROUND_SESSION_STORAGE_KEY = "halo-auth-background-session-v1";

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

let currentDialog: Office.Dialog | null = null;
let waitingForDialog = false;
let checkingSession = false;
let authConfigPromise: Promise<AuthConfigResponse> | null = null;
let msalInstancePromise: Promise<unknown> | null = null;

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
  getRefreshTicketsButton().onclick = () => void loadTickets();
  getTicketSearchForm().onsubmit = (event) => {
    event.preventDefault();
    void searchTickets();
  };
  getClearSearchButton().onclick = clearSearchResults;
  registerItemChangedHandler();
}

async function checkExistingSession() {
  if (checkingSession || waitingForDialog) {
    return;
  }

  checkingSession = true;

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

async function searchTickets() {
  const ticketNumber = getTicketNumberInput().value.trim();

  if (!ticketNumber) {
    setStatus(
      "failed",
      "Enter a ticket number",
      "Type a HaloPSA ticket number, then select Search."
    );
    getTicketNumberInput().focus();
    return;
  }

  const searchResults = getSearchResults();
  searchResults.hidden = false;
  getSearchEmpty().hidden = true;
  clearSearchList();

  try {
    setSearchBusy(true);
    setStatus(
      "loading",
      `Searching for ticket ${ticketNumber}...`,
      "Searching all tickets you can access in HaloPSA."
    );

    const result = await fetchJson<HaloTicketsResponse>(
      `/api/halo/tickets/search?ticketNumber=${encodeURIComponent(ticketNumber)}`
    );

    if (!result.ok) {
      throw createHaloAuthError(
        result.error || result.message || "Halo ticket search failed.",
        result.debug
      );
    }

    renderSearchResults(result.tickets || []);
    if (result.tickets.length) {
      setStatus(
        "success",
        result.tickets.length === 1 ? "Ticket found" : `${result.tickets.length} tickets found`,
        "Select a ticket below to attach the open email."
      );
    } else {
      setStatus("signed-out", "No matching ticket found", "Check the ticket number and try again.");
    }
  } catch (error) {
    setFailed(error, { hideLogout: false, message: "Ticket search failed" });
  } finally {
    setSearchBusy(false);
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
    setConnectionState(true);

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
    throw createHaloAuthError(
      body.error || body.message || `Request failed with status ${response.status}.`,
      body.debug
    );
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
  getSearchTicketsButton().disabled = isBusy;
  getClearSearchButton().disabled = isBusy;
  getTicketNumberInput().disabled = isBusy;
}

function setTicketsBusy(isBusy: boolean) {
  getRefreshTicketsButton().disabled = isBusy;
}

function setSearchBusy(isBusy: boolean) {
  getSearchTicketsButton().disabled = isBusy;
  getClearSearchButton().disabled = isBusy;
  getTicketNumberInput().disabled = isBusy;
  setTicketButtonsBusy(isBusy);
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

  ticketsPanel.hidden = false;
  clearTicketList();
  getTicketsEmpty().hidden = tickets.length > 0;
  renderTicketList(tickets, getTicketsList());
}

function renderSearchResults(tickets: HaloTicket[]) {
  getSearchResults().hidden = false;
  clearSearchList();
  getSearchEmpty().hidden = tickets.length > 0;
  renderTicketList(tickets, getSearchList());
}

function renderTicketList(tickets: HaloTicket[], container: HTMLElement) {
  tickets.forEach((ticket) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "halo-ticket";
    button.dataset.selected = "false";
    button.setAttribute("aria-label", `Attach email to ${formatTicketTitle(ticket)}`);
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
  getTicketsPanel().hidden = true;
  getTicketsEmpty().hidden = true;
  clearTicketList();
  clearSearchResults();
}

function clearTicketList() {
  const ticketsList = getTicketsList();
  while (ticketsList.firstChild) {
    ticketsList.removeChild(ticketsList.firstChild);
  }
}

function clearSearchResults() {
  getTicketNumberInput().value = "";
  getSearchResults().hidden = true;
  getSearchEmpty().hidden = true;
  clearSearchList();
}

function clearSearchList() {
  const searchList = getSearchList();
  while (searchList.firstChild) {
    searchList.removeChild(searchList.firstChild);
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

function getLoginButton(): HTMLButtonElement {
  return document.getElementById("login-button") as HTMLButtonElement;
}

function getLogoutButton(): HTMLButtonElement {
  return document.getElementById("logout-button") as HTMLButtonElement;
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

function getTicketNumberInput(): HTMLInputElement {
  return document.getElementById("ticket-number") as HTMLInputElement;
}

function getSearchTicketsButton(): HTMLButtonElement {
  return document.getElementById("search-tickets-button") as HTMLButtonElement;
}

function getClearSearchButton(): HTMLButtonElement {
  return document.getElementById("clear-search-button") as HTMLButtonElement;
}

function getSearchResults(): HTMLElement {
  return document.getElementById("search-results") as HTMLElement;
}

function getSearchEmpty(): HTMLElement {
  return document.getElementById("search-empty") as HTMLElement;
}

function getSearchList(): HTMLElement {
  return document.getElementById("search-list") as HTMLElement;
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
