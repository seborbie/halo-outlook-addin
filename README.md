# Halo Outlook Add-in

The Outlook add-in lives in `Halo Email Integration`. It runs as an Office task pane locally, and can also be built and served by the production Node server for an Azure container deployment.

Current release: `v2026.8.24`.

## Ticket Search

- The task pane opens with the current agent's open Halo tickets and uses one shared results list for
  browsing and searching.
- Search uses Halo's native ticket search, so queries can match ticket references, summaries/subjects,
  and any additional searchable ticket content enabled in the Halo tenant. Exact ticket references are
  promoted without hiding other relevant results.
- Results can be scoped to the current agent or all accessible agents and to open, closed, or all
  tickets. Customer and assignee filters narrow the loaded result set, which is capped at 50 tickets.
- Searches run after a short delay once two characters have been entered. Pressing Enter or selecting
  **Search** runs immediately, including for one-character ticket references.

This is Halo-native indexed search rather than an external embedding or vector-search service. It does
not copy ticket content into a separate search database or require another AI provider.

## Ticket Creation

- **Create new** discovers every active ticket type that the signed-in Halo agent may create. The
  add-in builds the form from Halo's current ticket-type and field metadata; there is no separate
  profile editor or hard-coded type allowlist.
- Required and core fields appear first. Supported optional fields appear under **Optional details**.
  Types with an unsupported mandatory control remain visible but disabled with the reason, while
  unsupported optional controls are omitted with a warning.
- Ticket-type schemas are cached separately for each Halo tenant and agent for 30 minutes. A manual
  refresh is available, and a cached schema may be used for up to 24 hours during a temporary Halo
  metadata outage. The server always revalidates current field requirements before creation.
- The summary is prefilled from the email subject. The requester is resolved from the incoming sender
  or the outgoing primary external recipient; missing or ambiguous matches require an explicit Halo
  user selection.
- From an open received email, creation happens immediately. From a draft, **Create when sent** saves
  an editable creation intent and creates the ticket during Outlook's `OnMessageSend` event. Choosing
  creation and choosing an existing ticket are mutually exclusive.
- The new ticket receives one Halo Email action or customer-hidden Private Note containing the
  selected or sent message's visible body, including quoted history, CID images, and any ordinary
  attachments the user chose to include. It does not retrieve separate conversation items or create
  an `.eml` export.

## Email Attachment Behaviour

- In read mode, selecting an existing Halo ticket immediately adds the open Outlook email to the
  ticket; choosing **Create new** creates a ticket and then adds the email to it.
- In compose mode, selecting a ticket marks the draft for attachment when it is sent. The selection
  is saved with the draft, can be changed or removed, and overrides an automatic conversation match.
- **Attach as private note** changes the import from a customer-visible Halo Email action to a
  customer-hidden Private Note. New conversations default to Email actions; mapped conversations
  remember the last successful visibility choice. A mapped received email waits for an explicit
  **Attach email** confirmation so the visibility can be reviewed before it is committed.
- The add-in creates the selected Halo action from the message subject, recipients, and HTML or text
  body. It uploads referenced CID images so they render in Halo. When an interactive import contains
  ordinary attachments, it asks whether to add them to the same Halo action or import the email only.
- For compose messages, ordinary attachments selected for Halo are encrypted and staged in PostgreSQL
  as they are added. The task pane shows preparation progress and does not show the green ready state
  until every eligible file has server acknowledgement. Removing a file or cancelling the selection
  immediately purges its staged ciphertext. Inline CID/signature images are classified separately and
  continue to use the existing image pipeline.
- Automatic sent-reply processing cannot display a custom prompt, so a mapped conversation includes
  new eligible ordinary attachments automatically. Explicit **Email only** exclusions are preserved.
- Ordinary attachments are limited to 20 files, 25 MiB per file, and 50 MiB decoded across one email.
  Outlook `.eml` and calendar attachment formats are preserved as `.eml` and `.ics`; cloud-link URL
  attachments are skipped with a visible warning. Unsupported files do not block Send, but a missing,
  pending, failed, stale, or expired preparation for an eligible included file produces a Smart Alerts
  warning before Outlook offers **Send Anyway**.

### Compose send semantics and support

The compose integration uses Outlook Smart Alerts (`OnMessageSend`) with `PromptUser`. When a user
selects an existing ticket or chooses **Create when sent**, that choice is saved with the draft. After
the user selects **Send**, the add-in reads the final compose subject, From, To, Cc, body, inline
images, and ordinary-attachment inventory. It does not read ordinary attachment contents during the
send event. Instead, the server verifies that the current draft identity, destination, operation,
inventory fingerprint, counts, and encrypted staged contents agree, then uploads those files to Halo
and creates the email action as one idempotent commit. It creates the requested ticket first when
necessary. If preparation is incomplete or Halo is unavailable, Outlook warns the user and offers the
normal **Send Anyway** path. Bcc recipients are intentionally not copied to Halo.

For later replies, the event runtime uses exact message identifiers and races the opaque
`X-Halo-Compose-Id` header against the add-in custom property on the referenced Sent Items message.
Every recovered compose ID is revalidated by the server before it can select a Halo mapping. A
no-match, EWS error, or recovery timeout sends normally without a Halo action; the add-in warns only
after a valid mapping has been established and the required Halo processing then fails.

Halo receives a pre-send compose snapshot, not an exact MIME copy from Sent Items. The integration
does not request Microsoft Graph mail permissions and cannot confirm final delivery or capture changes
made later by Exchange transport rules, such as server-side disclaimers. A rare failure after the
Smart Alerts handler succeeds can therefore leave a Halo action for a message Outlook did not
ultimately submit.

The first release supports standard primary-mailbox compose surfaces in Outlook on the web, new
Outlook for Windows, supported classic Outlook for Windows, and Outlook for Mac. Outlook mobile and
shared or delegated mailboxes are not supported. The feature also requires network connectivity when
the send event runs. See Microsoft's [Smart Alerts client and server support](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events).

## Required Environment

Use `.env.example` as the template for the required values. The local webpack and Node servers load `Halo Email Integration/.env`; process environment variables override the file. In Azure, configure the same values as app/container environment variables instead of deploying `.env`.

Required for local development and production:

```text
HALO_TOKEN_ENCRYPTION_KEY=<32-byte base64/base64url key>
HALO_URL=https://your-company.halopsa.com
HALO_CLIENT_ID=<Halo API application client ID>
ADDIN_CLIENT_ID=<Microsoft Entra app registration client ID>
DATABASE_URL=postgresql://127.0.0.1:5432/haloaddin
DATABASE_USERNAME=haloaddin
DATABASE_PASSWORD=haloaddin_local
```

Generate a local encryption key with:

```powershell
node -e "console.log(require('crypto').randomBytes(32).toString('base64url'))"
```

Optional environment:

```text
ADDIN_AUTHORITY=https://login.microsoftonline.com/common
ADDIN_API_CLIENT_ID=<Microsoft Entra API app client ID>
ADDIN_API_AUDIENCE=api://<Microsoft Entra app client ID>
ADDIN_AUTH_SCOPES=api://<Microsoft Entra app client ID>/access_as_user
ADDIN_REQUIRED_SCOPE=access_as_user
PUBLIC_BASE_URL=https://your-addin-host.example.com
PORT=3000
BUG_REPORT_GITHUB_REPOSITORY=seborbie/halo-outlook-addin
BUG_REPORT_GITHUB_TOKEN=<fine-grained GitHub token>
BUG_REPORT_SESSION_TTL_MINUTES=15
INLINE_IMAGE_DIAGNOSTICS=0
EMAIL_ATTACHMENT_DIAGNOSTICS=0
SEND_EVENT_DIAGNOSTICS=0
DATABASE_POOL_MAX=5
DATABASE_CONNECTION_TIMEOUT_MS=10000
DATABASE_IDLE_TIMEOUT_MS=30000
```

`HALO_TOKEN_ENCRYPTION_KEY` protects the stored Halo OAuth tokens. Keep the same value across restarts, otherwise existing stored Halo grants cannot be decrypted. In Azure, configure this as an app secret now and later source the same setting from Azure Key Vault.

The same key also encrypts staged ordinary attachment bytes with AES-256-GCM. Ciphertext is bound to
the staging operation, attachment, authenticated user, and Halo tenant. Staged bytes expire after
seven days and are purged immediately after a successful Halo action. Changing this key makes both
stored Halo grants and any unconsumed attachment stages unreadable, so rotate it only with a planned
reauthentication and draft-restaging window.

`DATABASE_URL` is required. Startup checks the connection and runs ordered migrations under a
PostgreSQL advisory lock before the HTTP server listens. `DATABASE_URL` must contain only the
PostgreSQL endpoint and database; credentials embedded in it are rejected. `DATABASE_USERNAME` is
always required. `DATABASE_AUTH` defaults to password authentication when it is unset or set to
`password`, `psk`, or `usernamepassword`, in which case `DATABASE_PASSWORD` is also required. Set
`DATABASE_AUTH=entra` to replace the password with Azure PostgreSQL tokens acquired through
`DefaultAzureCredential`. Local connections do not use TLS by default; remote connections,
including Azure Database for PostgreSQL, require certificate-verified TLS by default. The connection
pool uses five connections unless `DATABASE_POOL_MAX` is set.

### Local PostgreSQL

Start Docker Desktop, then run the PostgreSQL 17 service from `Halo Email Integration`:

```powershell
npm run db:up
npm run db:migrate
```

The service listens only on `127.0.0.1:5432` and uses a named Docker volume. The credentials in
`compose.yaml` are for local development only. Use `npm run db:down` to stop it without losing data,
or `npm run db:reset` to stop it and permanently remove the local database volume. Tests use
`TEST_DATABASE_URL`, `TEST_DATABASE_USERNAME`, and `TEST_DATABASE_PASSWORD` when set and otherwise
use the local Compose database; each store-backed suite creates and removes its own PostgreSQL
schema.

Inline-image and ordinary-attachment diagnostics are enabled by default and write only stage names,
safe outcomes, counts, and durations. They never log CIDs, HTML, Base64, image or file bytes,
attachment metadata or IDs, Halo response text, render tokens, or URLs. Set
`INLINE_IMAGE_DIAGNOSTICS=0` or `EMAIL_ATTACHMENT_DIAGNOSTICS=0` to disable the corresponding logs.

Outlook send diagnostics are also enabled by default. During Smart Alerts processing, the server CLI
prints `[halo-send]` entries containing only the client/server stage, elapsed time, safe outcome,
HTTP status, authentication mechanism, and attachment/image counts. These entries make the last
completed stage visible when Outlook reports that the add-in is taking too long. Message bodies,
subjects, recipients, ticket content, ticket or message IDs, attachment names/data, and session
values are never recorded. Set `SEND_EVENT_DIAGNOSTICS=0` to disable these logs.

For mapped replies with ordinary attachments, follow both prefixes in the CLI. Client
`[halo-send]` stages show `attachment-change-event`, `attachment-inventory`,
`attachment-staging`, and the Send-time `attachment-state`; server stages then show
`mapping-lookup`, `assets-prepare-*`, `halo-action-*`, and `response-complete`.
`[email-attachments]` entries identify the preparation or commit phase and use safe outcomes such
as `stage-missing`, `stage-unavailable`, `inventory-mismatch`, `preparation-incomplete`,
`integrity-failed`, or `halo-upload-failed` without revealing the message or attachment.

Referenced CID images are uploaded directly to Halo's native `POST /api/attachment/image`
endpoint. The email action is rewritten only with the returned same-tenant
`/api/attachment/image?token=...` URL. The add-in never persists image bytes or Base64; PostgreSQL
contains only SHA-256 deduplication data and Halo-owned attachment metadata.

Ordinary compose files use a two-phase pipeline. `OnMessageAttachmentsChanged` reads each eligible
file, calculates its SHA-256 hash, and sends it to the authenticated staging API. PostgreSQL stores
only AES-256-GCM ciphertext plus the IV, authentication tag, key ID, hash, decoded size, and safe
status metadata; plaintext and Base64 exist only in request memory and are never logged. On Send, the
server decrypts the matching stage, calls Halo's native `POST /api/Attachment` endpoint with bounded
concurrency, and includes the returned descriptors in the same `POST /api/Actions` request as the
email body. A partial upload or action failure deletes temporary Halo files and preserves ciphertext
for retry. The successful commit consumes the operation and purges its ciphertext. Legacy pre-upload
records are cleanup-only and are never considered prepared or reused.

Draft ticket-creation intents are stored as encrypted JSON in PostgreSQL for up to 30 days. Outlook stores
only the opaque operation ID, draft ID, and safe display summary. Creation intents never contain email
bodies or attachment bytes. Completed operations retain only the information needed to prevent a
retry from creating a duplicate ticket and are removed by the same expiry cleanup path.

The bug-report settings are optional. Without them, the add-in continues to run and the report button displays a temporary-unavailability message. Reports always use the `bug` label, and the report-session lifetime defaults to 15 minutes with a maximum of 60 minutes.

## GitHub Bug Reporting

The task pane's **Report a bug** button creates a short-lived, single-use report link and opens `/bugreport`. Classic Outlook uses the external browser API. New Outlook and Outlook on the web, which do not support that API, automatically use an Office dialog instead. Completed reports are stored as issues in the configured GitHub repository; PostgreSQL stores only hashed temporary session identifiers until they expire.

Set up the report dashboard as follows:

1. Enable Issues on the repository that will receive reports. For this project, set `BUG_REPORT_GITHUB_REPOSITORY=seborbie/halo-outlook-addin`.
2. Ensure the `bug` label exists.
3. Create a fine-grained personal access token with access only to the report repository and grant **Issues: Read and write**. No Contents permission is required.
4. Store the token as the Azure App Service secret `BUG_REPORT_GITHUB_TOKEN`.
5. In the report repository, choose **Watch > Custom > Issues**. In GitHub notification settings, enable **Email** and **On GitHub** delivery so new issues notify the maintainers.
6. Rotate the fine-grained token before its configured expiry and update the Azure secret without rebuilding the container.

Each issue contains the add-in version, Outlook host/platform, Office version, and the form contents. The authenticated user's name and email are not included. The add-in also excludes mailbox contents, attachments, and Halo ticket data.

The configured repository may be public. Anything a user types into the form will be visible wherever the resulting GitHub issue is visible, so users must not enter names, email addresses, customer data, credentials, or other sensitive information.

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

6. Grant the application the Halo permissions needed to search and create tickets, read ticket-type
   configuration and ticket-field metadata, search users/requesters and related client/site/asset
   lookups, create ticket actions/emails, and upload/read ticket attachments. The exact permission
   labels depend on your Halo version and configuration.
7. Copy the Halo API application client ID.

The Halo client ID is public OAuth application metadata. The add-in does not ask for or use a Halo client secret. Configure the tenant origin as `HALO_URL` and the Halo API application client ID as `HALO_CLIENT_ID`; users no longer enter either value in the task pane.

## Local Development

Copy `.env.example` to `.env`, set all four required values, then install Microsoft's trusted
localhost development certificate and start the Office add-in development server:

```powershell
cd ".\Halo Email Integration"
npx office-addin-dev-certs install
npm run dev-server
```

The local Halo callback URL is:

```text
https://localhost:3000/auth/callback
```

If you sideload the manifest into Outlook, the task pane will silently authenticate the Microsoft user, then offer Halo sign-in if that Microsoft user does not already have a stored Halo grant.

### Sideload the local manifest

1. Keep `npm run dev-server` running and verify that `https://localhost:3000` opens without a certificate warning.
2. In a second terminal run `npm start`. This builds and sideloads
   `Halo Email Integration/dist/manifest.debug.xml`. Running `npm run dev-server` alone serves the
   files but does not update Outlook's add-in registration.
3. If automatic sideloading is unavailable, open [the Outlook add-in sideload dialog](https://aka.ms/olksideload),
   select **My add-ins > Custom Addins > Add a custom add-in > Add from File**, and choose
   `Halo Email Integration/dist/manifest.debug.xml`.
4. Confirm Outlook shows **LOCAL DIAGNOSTICS - HaloPSA Outlook Add-in**. Any older add-in labelled
   **DEV - HaloPSA Outlook Add-in** is a stale registration and will not use the current local runtime.
5. Open an existing email or a compose window. Select **Apps** or the ribbon overflow if necessary,
   then open **HaloPSA > Attach or Create Ticket**.

The generated diagnostics manifest uses a separate development add-in ID, a cache-busted event
runtime URL, and localhost sources. It cannot collide with the production/admin-managed add-in.
The checked-in manifest retains the production add-in ID; a production build replaces its localhost
origin with `PUBLIC_BASE_URL` in `dist/manifest.xml`.

The same sideloaded add-in should become available in Outlook on the web and supported new/classic Outlook desktop clients. Classic Outlook can cache manual installations for up to 24 hours. Remove or replace it from the same **Custom Addins** section when testing a new manifest version.

### Deploy the send-event update

Production event-based activation should be deployed through **Microsoft 365 admin center > Settings >
Integrated apps**. Upload the production manifest and assign it to the intended users or groups. When
an existing deployment gains `OnMessageSend` or another automatic event, Microsoft marks the update
as pending until an administrator accepts the additional event-based capability. Test with an
admin-assigned pilot group before broad deployment, and verify the event-runtime URL is reachable over
public HTTPS from every supported Outlook client.

This release elevates the add-in-only permission from `ReadWriteItem` to `ReadWriteMailbox` so the
event runtime can perform bounded EWS metadata searches in Sent Items. That permission technically
allows the add-in to read or write any mailbox item or folder and to send mail, although this feature
uses exact Internet Message ID restrictions and reads only the recovery metadata it needs. An
administrator must explicitly approve the elevated permission, and EWS must be enabled for the target
Exchange mailboxes. No Microsoft Graph permission or consent is added.

The manifest deliberately uses `PromptUser`, which allows deployment through supported Marketplace
and admin-managed routes while letting users override a Halo failure. Restricted Marketplace listings
still require administrator deployment for event-based activation. See Microsoft's
[event-based add-in deployment guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation)
and [Marketplace listing options](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/autolaunch-store-options).

## Production-Style Run

For an Azure-style build, set the public HTTPS origin and run the production server:

```powershell
$env:HALO_TOKEN_ENCRYPTION_KEY="<generated key>"
$env:HALO_URL="https://your-company.halopsa.com"
$env:HALO_CLIENT_ID="<Halo API application client ID>"
$env:ADDIN_CLIENT_ID="<Microsoft Entra application client ID>"
$env:DATABASE_URL="postgresql://127.0.0.1:5432/haloaddin"
$env:DATABASE_USERNAME="haloaddin"
$env:DATABASE_PASSWORD="haloaddin_local"
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

### Release checklist

For the `2026.8.24` release:

1. Confirm the GitHub Actions `PUBLIC_BASE_URL` variable matches the deployed HTTPS origin.
2. Run the full CI workflow and publish the version and `sha-<commit>` container tags.
3. Deploy the immutable `sha-<commit>` tag and verify the container is running with the same
   `PUBLIC_BASE_URL`, encryption key, Halo OAuth settings, Microsoft Entra application ID, and
   TLS-required database endpoint, username, and password used for the release.
4. Download or build `dist/manifest.xml`; confirm it contains version `2026.8.24.1`, the production
   add-in ID, and no localhost or development labels.
5. Upload `dist/manifest.xml` through Microsoft 365 Integrated Apps and accept the event-based
   activation capability update. Assign it to a pilot group before broad deployment.
6. Test read-mode attach, attach-new-email-on-send, create-ticket-on-send, retry/idempotency, and
   **Send Anyway** behavior against the deployed service.

For a container image, pass the same origin as a Docker build argument because manifest and event-runtime URLs are generated while the image is built:

```powershell
cd ".\Halo Email Integration"
docker build `
  --build-arg PUBLIC_BASE_URL="https://your-addin-host.example.com" `
  --tag halo-outlook-addin:2026.8.24 `
  .
```

## GitHub Container Registry Publishing

The `.github/workflows/publish-container.yml` workflow runs on every commit to `main` and can also be started manually. It verifies the application, builds the Linux container, and publishes these tags to GitHub Container Registry:

```text
ghcr.io/seborbie/halo-outlook-addin:2026.8.24
ghcr.io/seborbie/halo-outlook-addin:latest
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
    DATABASE_AUTH="password" `
    DATABASE_URL="postgresql://<server>.postgres.database.azure.com:5432/haloaddin" `
    DATABASE_USERNAME="@Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/halo-db-username)" `
    DATABASE_PASSWORD="@Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/halo-db-password)" `
    BUG_REPORT_GITHUB_REPOSITORY="seborbie/halo-outlook-addin" `
    BUG_REPORT_GITHUB_TOKEN="<fine-grained GitHub token>" `
    PUBLIC_BASE_URL="https://your-addin-host.example.com"

az webapp update `
  --resource-group "<resource-group>" `
  --name "<app-name>" `
  --https-only true
```

The build argument and the runtime `PUBLIC_BASE_URL` must match. Rebuild the image if the public hostname changes, then update the Microsoft Entra and Halo redirect URIs to use the same HTTPS origin. Configure App Service Health check to use `/health/ready`; it reports 200 only while PostgreSQL accepts queries. `DATABASE_URL` is non-secret and contains no credentials. For password authentication, provide `DATABASE_USERNAME` and `DATABASE_PASSWORD` through Key Vault references, another runtime secret store, or the local environment, never in the image.

### Azure PostgreSQL production profile

Use an Azure Database for PostgreSQL Flexible Server in UK South with PostgreSQL 17,
`Standard_B1ms` Burstable compute, 32 GiB Premium LRS storage, storage autogrow enabled, high
availability and geo-redundant backup disabled, and seven-day locally redundant backup retention.
Add a cost budget and alert before production traffic is enabled.

Keep public network access enabled, but create a separate single-address firewall rule for every IP
in the App Service `possibleOutboundIpAddresses` property. Do not create the broad
`0.0.0.0` "Allow Azure services" rule. Recheck the possible outbound list whenever the App Service
plan or scale configuration changes.

Use the server administrator only for provisioning. Create a `haloaddin` database and a dedicated
password-authenticated application role. Because the application runs migrations during startup,
that role must own the application schema and its objects or inherit the existing owner role. Revoke
default public database/schema creation privileges and set the application role's search path to its
owned schema. Configure `DATABASE_AUTH=password`, store the application username and password as
separate Key Vault secrets, and use a credential-free `DATABASE_URL` with the server's
`*.postgres.database.azure.com:5432/haloaddin` endpoint and TLS.

Provision and test the database before changing the web app image. Deploy the immutable
`sha-<commit>` container tag, configure the App Service health check as `/health/ready`, and confirm
the migration/startup logs before smoke testing authentication, ticket lookup and creation,
attachments, inline images, mappings, and bug reports. This migration intentionally starts empty,
so users must authenticate again. Once PostgreSQL contains production writes, roll back only to a
PostgreSQL-capable image or apply a forward fix; never switch the app back to a SQLite image.
