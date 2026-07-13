/* global document, fetch, FormData, HTMLButtonElement, HTMLFormElement, HTMLInputElement, HTMLElement, location, history, sessionStorage, URLSearchParams */

const SESSION_STORAGE_KEY = "halo-bug-report-session";

type BugReportResponse = {
  error?: string;
  ok?: boolean;
  reference?: number;
};

document.addEventListener("DOMContentLoaded", () => {
  const token = loadSessionToken();
  const form = getForm();
  const summary = document.getElementById("summary") as HTMLInputElement;

  summary.addEventListener("input", () => {
    getSummaryCount().textContent = `${summary.value.length} / 120`;
  });
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    void submitReport(token);
  });

  if (!token) {
    setStatus(
      "error",
      "This report link is missing or has expired. Open a new report from the Outlook add-in."
    );
    setFormEnabled(false);
  }
});

function loadSessionToken(): string {
  const fragment = new URLSearchParams(location.hash.replace(/^#/, ""));
  const token = fragment.get("token") || sessionStorage.getItem(SESSION_STORAGE_KEY) || "";

  if (fragment.has("token")) {
    history.replaceState(null, "", `${location.pathname}${location.search}`);
  }
  if (token) {
    sessionStorage.setItem(SESSION_STORAGE_KEY, token);
  }
  return token;
}

async function submitReport(token: string) {
  if (!token) {
    return;
  }

  const form = getForm();
  const values = new FormData(form);
  setStatus("loading", "Submitting your bug report…");
  setFormEnabled(false);
  let sessionInvalid = false;

  try {
    const response = await fetch("/api/bug-reports", {
      body: JSON.stringify({
        additionalContext: values.get("additionalContext") || "",
        description: values.get("description") || "",
        expectedBehavior: values.get("expectedBehavior") || "",
        reproductionSteps: values.get("reproductionSteps") || "",
        summary: values.get("summary") || "",
      }),
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
        "X-Bug-Report-Session": token,
      },
      method: "POST",
    });
    const body = (await response.json().catch(() => ({}))) as BugReportResponse;

    if (!response.ok || !body.ok) {
      if (response.status === 401) {
        sessionInvalid = true;
        sessionStorage.removeItem(SESSION_STORAGE_KEY);
      }
      throw new Error(body.error || `Submission failed with status ${response.status}.`);
    }

    sessionStorage.removeItem(SESSION_STORAGE_KEY);
    getStatus().hidden = true;
    form.hidden = true;
    getReference().textContent = body.reference
      ? `Your reference is bug report #${body.reference}.`
      : "Your report has been recorded.";
    getSuccess().hidden = false;
  } catch (error) {
    setStatus(
      "error",
      error instanceof Error ? error.message : "Could not submit the bug report. Please try again."
    );
    setFormEnabled(!sessionInvalid);
  }
}

function setStatus(state: "error" | "loading", message: string) {
  const status = getStatus();
  status.dataset.state = state;
  status.textContent = message;
  status.hidden = false;
}

function setFormEnabled(enabled: boolean) {
  Array.from(getForm().elements).forEach((element) => {
    if ("disabled" in element) {
      (element as HTMLButtonElement | HTMLInputElement).disabled = !enabled;
    }
  });
}

function getForm(): HTMLFormElement {
  return document.getElementById("bug-report-form") as HTMLFormElement;
}

function getStatus(): HTMLElement {
  return document.getElementById("report-status") as HTMLElement;
}

function getSummaryCount(): HTMLElement {
  return document.getElementById("summary-count") as HTMLElement;
}

function getSuccess(): HTMLElement {
  return document.getElementById("report-success") as HTMLElement;
}

function getReference(): HTMLElement {
  return document.getElementById("report-reference") as HTMLElement;
}
