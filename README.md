# Halo Outlook Add-in

The Outlook add-in lives in `Halo Email Integration`. It runs as an Office task pane locally, and can also be built and served by the production Node server for an Azure container deployment.

Current release: `v2026.7.10B` (Beta).

## Required Environment

Use `.env.example` as the template for the required values. The app does not load `.env` files by itself, so either export these values in your shell, use a dotenv loader in your local workflow, or configure them as Azure app/container environment variables.

Required for local development and production:

```text
HALO_TOKEN_ENCRYPTION_KEY=<32-byte base64/base64url key>
HALO_URL=https://your-company.halopsa.com
HALO_CLIENT_ID=<Halo API application client ID>
ADDIN_CLIENT_ID=<Microsoft Entra app registration client ID>
```

Generate a local encryption key with:

```powershell
node -e "console.log(require('crypto').randomBytes(32).toString('base64url'))"
```

Optional environment:

```text
HALO_DB_PATH=./data/halo.sqlite
ADDIN_AUTHORITY=https://login.microsoftonline.com/common
ADDIN_API_CLIENT_ID=<Microsoft Entra API app client ID>
ADDIN_API_AUDIENCE=api://<Microsoft Entra app client ID>
ADDIN_AUTH_SCOPES=api://<Microsoft Entra app client ID>/access_as_user
ADDIN_REQUIRED_SCOPE=access_as_user
PUBLIC_BASE_URL=https://your-addin-host.example.com
PORT=3000
BUG_REPORT_GITHUB_REPOSITORY=owner/private-bug-report-repository
BUG_REPORT_GITHUB_TOKEN=<fine-grained GitHub token>
BUG_REPORT_GITHUB_LABELS=bug,outlook-addin
BUG_REPORT_SESSION_TTL_MINUTES=15
```

`HALO_TOKEN_ENCRYPTION_KEY` protects the stored Halo OAuth tokens. Keep the same value across restarts, otherwise existing stored Halo grants cannot be decrypted. In Azure, configure this as an app secret now and later source the same setting from Azure Key Vault.

`HALO_DB_PATH` defaults to `./data/halo.sqlite` relative to the process working directory. When running from `Halo Email Integration`, the default database path is `Halo Email Integration/data/halo.sqlite`.

The bug-report settings are optional. Without them, the add-in continues to run and the report button displays a temporary-unavailability message. `BUG_REPORT_GITHUB_LABELS` defaults to `bug,outlook-addin`, and the report-session lifetime defaults to 15 minutes with a maximum of 60 minutes.

## Private GitHub Bug Reporting

The task pane's **Report a bug** button creates a short-lived, single-use report link and opens `/bugreport` in the user's external browser. Completed reports are stored only as issues in a dedicated private GitHub repository; SQLite stores only hashed temporary session identifiers until they expire.

Set up the private report dashboard as follows:

1. Create a private GitHub repository for add-in bug reports and enable Issues.
2. Add labels named `bug` and `outlook-addin` (or override `BUG_REPORT_GITHUB_LABELS`).
3. Create a fine-grained personal access token with access only to that repository and grant **Issues: Read and write**. No Contents permission is required.
4. Store the token as the Azure App Service secret `BUG_REPORT_GITHUB_TOKEN`, and set `BUG_REPORT_GITHUB_REPOSITORY` to `owner/repository`.
5. In the report repository, choose **Watch > Custom > Issues**. In GitHub notification settings, enable **Email** and **On GitHub** delivery so new issues notify the maintainers.
6. Rotate the fine-grained token before its configured expiry and update the Azure secret without rebuilding the container.

Each issue contains the authenticated user's name and email, the add-in version, Outlook host/platform, Office version, and the form contents. The add-in never includes the open email's subject, recipients, body, attachments, or Halo ticket data. Keep the report repository private because issues contain reporter identity and diagnostic context.

The public `/bugreport` page cannot submit by itself. `POST /api/bug-reports/session` requires the existing Microsoft add-in bearer token, and `POST /api/bug-reports` requires the resulting single-use session token. A failed GitHub request releases the session so the user can retry; a successful request consumes it to prevent duplicate submissions.

## Microsoft Add-In App Registration

The add-in uses Microsoft authentication to identify the Outlook user before reconnecting them to their stored Halo grant. Create this app registration in Microsoft Entra ID to get `ADDIN_CLIENT_ID`.

1. Open the Azure portal and go to Microsoft Entra ID > App registrations > New registration.
2. Name the app, for example `Halo Outlook Add-in`.
3. Choose the supported account type for your deployment. For internal company use, single tenant is usually the simplest option.
4. Add a Single-page application redirect URI for nested app authentication:

```text
brk-multihub://localhost:3000
```

5. For production, add another SPA redirect URI for the hosted add-in origin only:

```text
brk-multihub://your-addin-host.example.com
```

Do not include `/taskpane.html`, `/auth/callback`, or another path in the `brk-multihub://` redirect. Microsoft NAA expects the add-in origin.

6. Copy the Application (client) ID from the app registration and use it as `ADDIN_CLIENT_ID`.
7. Go to **Expose an API**, add an Application ID URI, and accept the default `api://<ADDIN_CLIENT_ID>` value.
8. Add a delegated scope named `access_as_user`. Allow admins and users to consent, enable the scope, and use descriptions explaining that it lets the add-in call its web API as the signed-in user.
9. In **Expose an API**, select **Add a client application**, enter the same `ADDIN_CLIENT_ID`, select `access_as_user`, and add the application. This preauthorizes the add-in to call its own API.
10. Go to **API permissions** > **Add a permission** > **My APIs**, choose this application, add the delegated `access_as_user` permission, and grant admin consent if required by your tenant.
11. Open the app registration's manifest, set `api.requestedAccessTokenVersion` to `2`, and save it.

The add-in derives the API client ID, API audience, and requested scope from `ADDIN_CLIENT_ID`. Keep `ADDIN_API_CLIENT_ID`, `ADDIN_API_AUDIENCE`, `ADDIN_AUTH_SCOPES`, and `ADDIN_REQUIRED_SCOPE` unset when using the values above. Those settings only need overriding if you deliberately use a separate API app registration, custom Application ID URI, or custom scope name.

Remove any legacy `ADDIN_AUTH_SCOPES=openid profile email User.Read` setting from local or Azure configuration. That requests a Microsoft Graph token, which is not valid for this add-in's web API. The server rejects configured scopes that do not target its own API so this cannot silently regress into repeated 401 responses.

Reference: [Microsoft nested app authentication for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in).

## Halo API Application

Create a Halo API application so the add-in can run the Halo OAuth flow and receive refreshable Halo tokens.

1. In Halo, open the API application setup area.
2. Create a new API application for this Outlook add-in.
3. Use the `Authorisation Code (Native Application)` authentication method.
4. Add the local redirect URI:

```text
https://localhost:3000/auth/callback
```

5. For production, add the redirect URI based on the hosted add-in URL:

```text
https://your-addin-host.example.com/auth/callback
```

6. Grant the application the Halo permissions needed to search tickets and create ticket actions/emails. The exact permission labels depend on your Halo configuration, but the add-in needs ticket read/search access and permission to add ticket actions.
7. Copy the Halo API application client ID.

The Halo client ID is public OAuth application metadata. The add-in does not ask for or use a Halo client secret. Configure the tenant origin as `HALO_URL` and the Halo API application client ID as `HALO_CLIENT_ID`; users no longer enter either value in the task pane.

## Local Development

Set the required environment variables, then start the Office add-in dev server:

```powershell
$env:HALO_TOKEN_ENCRYPTION_KEY="<generated key>"
$env:HALO_URL="https://your-company.halopsa.com"
$env:HALO_CLIENT_ID="<Halo API application client ID>"
$env:ADDIN_CLIENT_ID="<Microsoft Entra application client ID>"
cd ".\Halo Email Integration"
npm run dev-server
```

The local Halo callback URL is:

```text
https://localhost:3000/auth/callback
```

If you sideload the manifest into Outlook, the task pane will silently authenticate the Microsoft user, then offer Halo sign-in if that Microsoft user does not already have a stored Halo grant.

## Production-Style Run

For an Azure-style build, set the public HTTPS origin and run the production server:

```powershell
$env:HALO_TOKEN_ENCRYPTION_KEY="<generated key>"
$env:HALO_URL="https://your-company.halopsa.com"
$env:HALO_CLIENT_ID="<Halo API application client ID>"
$env:ADDIN_CLIENT_ID="<Microsoft Entra application client ID>"
$env:PUBLIC_BASE_URL="https://your-addin-host.example.com"
cd ".\Halo Email Integration"
npm run build
npm run serve
```

`PUBLIC_BASE_URL` is used for production asset URLs, manifest/runtime URLs, and the Halo OAuth callback base URL. The production Halo callback URL will be:

```text
https://your-addin-host.example.com/auth/callback
```

Production builds require `PUBLIC_BASE_URL` and fail rather than emitting a manifest that points to localhost. The value must be the HTTPS origin only, with no path. HTTPS uses public port 443 by default, so the value normally doesn't need an explicit port.

For a container image, pass the same origin as a Docker build argument because manifest and event-runtime URLs are generated while the image is built:

```powershell
cd ".\Halo Email Integration"
docker build `
  --build-arg PUBLIC_BASE_URL="https://your-addin-host.example.com" `
  --tag halo-outlook-addin:2026.7.10-beta `
  .
```

## GitHub Container Registry Publishing

The `.github/workflows/publish-container.yml` workflow runs on every commit to `main` and can also be started manually. It verifies the application, builds the Linux container, and publishes these tags to GitHub Container Registry:

```text
ghcr.io/seborbie/halo-outlook-addin:2026.7.10-beta
ghcr.io/seborbie/halo-outlook-addin:latest-beta
ghcr.io/seborbie/halo-outlook-addin:sha-<commit>
```

Before the first run, open GitHub repository **Settings > Secrets and variables > Actions > Variables** and create this repository variable:

```text
PUBLIC_BASE_URL=https://your-addin-host.example.com
```

This is public build metadata rather than a secret. It must match the App Service HTTPS origin and must not contain a path. The workflow uses the repository `GITHUB_TOKEN`; no registry password or personal access token is required. If organization or repository policy restricts the token, allow GitHub Actions read/write workflow permissions so the job's `packages: write` permission can publish.

The first GHCR package is private by default. After its first successful publication, open the package settings on GitHub and change its visibility to **Public** so Azure can pull it anonymously. GitHub doesn't allow a public package to be changed back to private.

## Azure App Service Container Ports and HTTPS

Azure App Service owns the public listeners on ports 80 and 443. Enable the App Service HTTPS-only setting so requests received on public HTTP port 80 are redirected to HTTPS port 443. TLS terminates at Azure's front end, and the Node container continues to listen on its internal HTTP port 3000; the container does not need to expose ports 80 or 443 or contain the public TLS certificate.

Configure the App Service with the container port, the runtime public origin, and HTTPS-only mode:

```powershell
az webapp config appsettings set `
  --resource-group "<resource-group>" `
  --name "<app-name>" `
  --settings `
    WEBSITES_PORT=3000 `
    PORT=3000 `
    HALO_URL="https://your-company.halopsa.com" `
    HALO_CLIENT_ID="<Halo API application client ID>" `
    BUG_REPORT_GITHUB_REPOSITORY="owner/private-bug-report-repository" `
    BUG_REPORT_GITHUB_TOKEN="<fine-grained GitHub token>" `
    PUBLIC_BASE_URL="https://your-addin-host.example.com"

az webapp update `
  --resource-group "<resource-group>" `
  --name "<app-name>" `
  --https-only true
```

The build argument and the runtime `PUBLIC_BASE_URL` must match. Rebuild the image if the public hostname changes, then update the Microsoft Entra and Halo redirect URIs to use the same HTTPS origin.

For Azure, mount or persist the SQLite database location if you need state to survive container replacement. Set `HALO_DB_PATH` to that mounted path when it differs from the default.
