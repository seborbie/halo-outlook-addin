<script lang="ts">
  type State = "idle" | "connecting" | "submitting" | "success" | "error";

  let companyName = "";
  let workEmail = "";
  let haloUrl = "";
  let haloClientId = "";
  let state: State = "idle";
  let message = "";

  async function submitSignup() {
    state = "connecting";
    message = "Opening Microsoft sign-in…";

    try {
      const configResponse = await fetch("/api/auth/config");
      if (!configResponse.ok) {
        throw new Error("The signup service is not available yet. Try again shortly.");
      }

      const config = await configResponse.json();
      if (!config.ssoEnabled || !config.clientId || !config.scopes?.length) {
        throw new Error("Microsoft sign-in has not been configured for this deployment.");
      }

      const { PublicClientApplication } = await import("@azure/msal-browser");
      const msal = new PublicClientApplication({
        auth: {
          authority: config.authority,
          clientId: config.clientId,
          redirectUri: `${window.location.origin}/signup`,
        },
        cache: { cacheLocation: "sessionStorage" },
      });
      await msal.initialize();
      const authResult = await msal.loginPopup({
        prompt: "select_account",
        scopes: config.scopes,
      });
      const tokenResult = await msal.acquireTokenSilent({
        account: authResult.account,
        scopes: config.scopes,
      });

      state = "submitting";
      message = "Creating your organisation…";
      const response = await fetch("/api/organisations/register", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${tokenResult.accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ companyName, workEmail, haloUrl, haloClientId }),
      });
      const result = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(result.error || "We could not create your organisation.");
      }

      state = "success";
      message = `Your ${result.organisation?.name || companyName} workspace is ready. Continue with the installation guide.`;
    } catch (error) {
      state = "error";
      message = error instanceof Error ? error.message : "Signup failed. Please try again.";
    }
  }
</script>

<form class="signup-form" onsubmit={(event) => { event.preventDefault(); submitSignup(); }}>
  <div class="field-row">
    <label>
      <span>Company name</span>
      <input bind:value={companyName} name="companyName" autocomplete="organization" required placeholder="Acme IT" />
    </label>
    <label>
      <span>Work email</span>
      <input bind:value={workEmail} name="workEmail" type="email" autocomplete="email" required placeholder="you@company.com" />
    </label>
  </div>
  <label>
    <span>HaloPSA URL</span>
    <input bind:value={haloUrl} name="haloUrl" type="url" required placeholder="https://your-company.halopsa.com" />
    <small>The HTTPS origin of your company’s HaloPSA instance.</small>
  </label>
  <label>
    <span>Halo API application client ID</span>
    <input bind:value={haloClientId} name="haloClientId" required placeholder="Your native OAuth application client ID" />
    <small>This is public OAuth metadata—not a client secret.</small>
  </label>
  <label class="consent">
    <input type="checkbox" required />
    <span>I’m authorised to configure software for this Microsoft organisation.</span>
  </label>
  <button class="button button-wide" type="submit" disabled={state === "connecting" || state === "submitting"}>
    {state === "connecting" || state === "submitting" ? "Working…" : "Continue with Microsoft"}
    <span aria-hidden="true">→</span>
  </button>
  <p class:success={state === "success"} class:error={state === "error"} class="form-status" aria-live="polite">
    {message}
    {#if state === "success"}
      <a href="/docs/quickstart">Open the quickstart →</a>
    {/if}
  </p>
</form>
