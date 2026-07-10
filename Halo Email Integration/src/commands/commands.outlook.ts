/* global Office, fetch, setTimeout, clearTimeout, RequestInit, Response, window */

const BACKGROUND_SESSION_STORAGE_KEY = "halo-auth-background-session-v1";
const SEND_AUTO_ATTACH_URL = `${window.location.origin}/api/halo/email/send-auto-attach`;
const SEND_AUTO_ATTACH_TIMEOUT_MS = 4500;
const SEND_EVENT_TIMEOUT_MS = 3000;

type EmailAddressPayload = {
  displayName: string;
  emailAddress: string;
};

type OutlookSendEmailPayload = {
  backgroundSessionId: string;
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

type HaloSendAutoAttachResponse = {
  ok: boolean;
  status?: "attached" | "already-attached" | "failed" | "no-match" | "no-session";
  ticketId?: string;
  ticketNumber?: string;
  message?: string;
  error?: string;
  debug?: unknown;
};

type HaloSendEvent = {
  completed: (options?: { allowEvent?: boolean; errorMessage?: string }) => void;
};

type CompleteSendEvent = (options?: { allowEvent?: boolean; errorMessage?: string }) => void;

type ComposeItem = {
  body?: {
    getAsync: (
      coercionType: Office.CoercionType,
      callback: (result: Office.AsyncResult<string>) => void
    ) => void;
  };
  cc?: RecipientCollection | Office.EmailAddressDetails[];
  conversationId?: string;
  from?: SenderCollection | Office.EmailAddressDetails;
  getItemIdAsync?: (callback: (result: Office.AsyncResult<string>) => void) => void;
  inReplyTo?: string;
  internetMessageId?: string;
  itemId?: string;
  itemType?: Office.MailboxEnums.ItemType | string;
  normalizedSubject?: string;
  subject?: string | SubjectCollection;
  to?: RecipientCollection | Office.EmailAddressDetails[];
};

type RecipientCollection = {
  getAsync: (callback: (result: Office.AsyncResult<Office.EmailAddressDetails[]>) => void) => void;
};

type SenderCollection = {
  getAsync: (callback: (result: Office.AsyncResult<Office.EmailAddressDetails>) => void) => void;
};

type SubjectCollection = {
  getAsync: (callback: (result: Office.AsyncResult<string>) => void) => void;
};

export async function onHaloMessageSend(event: HaloSendEvent) {
  const complete = createSendEventCompletion(event);
  const watchdog = setTimeout(() => complete({ allowEvent: true }), SEND_EVENT_TIMEOUT_MS);

  try {
    const email = await readCurrentComposeEmail();

    if (!email) {
      completeAllow(complete);
      return;
    }

    const result = await sendAutoAttach(email);
    if (result.ok || result.status === "no-match" || result.status === "no-session") {
      completeAllow(complete);
      return;
    }

    if (result.ticketNumber || result.ticketId) {
      completeWithHaloWarning(complete, result);
      return;
    }

    completeAllow(complete);
  } catch {
    completeAllow(complete);
  } finally {
    clearTimeout(watchdog);
  }
}

async function readCurrentComposeEmail(): Promise<OutlookSendEmailPayload | null> {
  const item = Office.context.mailbox.item as unknown as ComposeItem;

  if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message || !item.body) {
    return null;
  }

  const inReplyToMessageIds = item.inReplyTo ? [item.inReplyTo] : [];
  const conversationId = item.conversationId || "";
  if (!inReplyToMessageIds.length && !conversationId) {
    return null;
  }

  const body = await readMessageBody(item);
  if (!body.bodyHtml && !body.bodyText) {
    return null;
  }

  const subject = await readSubject(item);
  const userProfile = Office.context.mailbox.userProfile;
  const mailboxEmail = userProfile && userProfile.emailAddress ? userProfile.emailAddress : "";

  return {
    ...body,
    backgroundSessionId: getBackgroundSessionId(),
    cc: await readRecipients(item.cc),
    conversationId,
    dateTimeCreated: new Date().toISOString(),
    from: (await readSender(item.from)) || getUserProfileAddress(),
    inReplyToMessageIds,
    internetHeaders: "",
    internetMessageId: item.internetMessageId || "",
    itemId: item.itemId || (await readItemId(item)),
    mailboxEmail,
    normalizedSubject: normalizeSubject(subject),
    referenceMessageIds: [],
    subject,
    timeZone: getClientTimeZone(),
    to: await readRecipients(item.to),
  };
}

function getClientTimeZone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || "";
  } catch {
    return "";
  }
}

async function sendAutoAttach(email: OutlookSendEmailPayload): Promise<HaloSendAutoAttachResponse> {
  const response = await fetchWithTimeout(
    SEND_AUTO_ATTACH_URL,
    {
      body: JSON.stringify(email),
      credentials: "include",
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      method: "POST",
    },
    SEND_AUTO_ATTACH_TIMEOUT_MS
  );

  return (await response.json().catch(() => ({
    ok: response.ok,
    status: response.ok ? "no-match" : "failed",
  }))) as HaloSendAutoAttachResponse;
}

function fetchWithTimeout(url: string, options: RequestInit, timeoutMs: number): Promise<Response> {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(
      () => reject(new Error("Halo send auto-attach timed out.")),
      timeoutMs
    );

    fetch(url, options)
      .then((response) => {
        clearTimeout(timeout);
        resolve(response);
      })
      .catch((error) => {
        clearTimeout(timeout);
        reject(error);
      });
  });
}

async function readMessageBody(item: ComposeItem): Promise<{ bodyHtml: string; bodyText: string }> {
  try {
    const bodyHtml = await getBodyAsync(item, Office.CoercionType.Html);
    return { bodyHtml, bodyText: "" };
  } catch {
    const bodyText = await getBodyAsync(item, Office.CoercionType.Text);
    return { bodyHtml: "", bodyText };
  }
}

function getBodyAsync(item: ComposeItem, coercionType: Office.CoercionType): Promise<string> {
  return new Promise((resolve, reject) => {
    if (!item.body) {
      reject(new Error("Could not read the compose body."));
      return;
    }

    item.body.getAsync(coercionType, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
        return;
      }

      reject(new Error(result.error.message || "Could not read the compose body."));
    });
  });
}

function readSubject(item: ComposeItem): Promise<string> {
  if (typeof item.subject === "string") {
    return Promise.resolve(item.subject);
  }

  if (!item.subject || !item.subject.getAsync) {
    return Promise.resolve(item.normalizedSubject || "");
  }

  return new Promise((resolve) => {
    (item.subject as SubjectCollection).getAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "");
    });
  });
}

function readRecipients(
  value?: RecipientCollection | Office.EmailAddressDetails[]
): Promise<EmailAddressPayload[]> {
  if (Array.isArray(value)) {
    return Promise.resolve(normalizeEmailAddressList(value));
  }

  if (!value || !value.getAsync) {
    return Promise.resolve([]);
  }

  return new Promise((resolve) => {
    value.getAsync((result) => {
      resolve(
        result.status === Office.AsyncResultStatus.Succeeded
          ? normalizeEmailAddressList(result.value)
          : []
      );
    });
  });
}

function readSender(
  value?: SenderCollection | Office.EmailAddressDetails
): Promise<EmailAddressPayload | null> {
  if (!value) {
    return Promise.resolve(null);
  }

  if ("emailAddress" in value || "displayName" in value) {
    return Promise.resolve(normalizeEmailAddress(value as Office.EmailAddressDetails));
  }

  return new Promise((resolve) => {
    (value as SenderCollection).getAsync((result) => {
      resolve(
        result.status === Office.AsyncResultStatus.Succeeded
          ? normalizeEmailAddress(result.value)
          : null
      );
    });
  });
}

function readItemId(item: ComposeItem): Promise<string> {
  if (!item.getItemIdAsync) {
    return Promise.resolve("");
  }

  return new Promise((resolve) => {
    item.getItemIdAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "");
    });
  });
}

function getBackgroundSessionId(): string {
  const value = Office.context.roamingSettings.get(BACKGROUND_SESSION_STORAGE_KEY);
  return typeof value === "string" ? value : "";
}

function getUserProfileAddress(): EmailAddressPayload | null {
  const userProfile = Office.context.mailbox.userProfile;
  if (!userProfile || !userProfile.emailAddress) {
    return null;
  }

  return {
    displayName: userProfile.displayName || "",
    emailAddress: userProfile.emailAddress,
  };
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

function normalizeSubject(value: string): string {
  return value.replace(/^(\s*(re|fw|fwd)\s*:\s*)+/i, "").trim();
}

function createSendEventCompletion(event: HaloSendEvent): CompleteSendEvent {
  let completed = false;

  return (options) => {
    if (completed) {
      return;
    }

    completed = true;
    event.completed(options);
  };
}

function completeAllow(complete: CompleteSendEvent) {
  complete({ allowEvent: true });
}

function completeWithHaloWarning(complete: CompleteSendEvent, result: HaloSendAutoAttachResponse) {
  const ticketLabel = result.ticketNumber || result.ticketId || "the mapped Halo ticket";
  complete({
    allowEvent: false,
    errorMessage: `Could not add this reply to Halo ticket ${ticketLabel}. Send anyway or try again.`,
  });
}

Office.actions.associate("onHaloMessageSend", onHaloMessageSend);
