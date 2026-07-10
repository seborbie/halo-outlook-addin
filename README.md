# Halo Outlook Add-in

The Outlook add-in lives in `Halo Email Integration`. It runs as an Office task pane locally, and can also be built and served by the production Node server for an Azure container deployment.

Current release: `v2026.7.10B` (Beta).

## Required Environment

Use `.env.example` as the template for the required values. The app does not load `.env` files by itself, so either export these values in your shell, use a dotenv loader in your local workflow, or configure them as Azure app/container environment variables.

Required for local development and production:

```text
HALO_TOKEN_ENCRYPTION_KEY=<32-byte base64/base64url key>
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
ADDIN_AUTH_SCOPES=openid profile email User.Read
PUBLIC_BASE_URL=https://your-addin-host.example.com
PORT=3000
```

`HALO_TOKEN_ENCRYPTION_KEY` protects the stored Halo OAuth tokens. Keep the same value across restarts, otherwise existing stored Halo grants cannot be decrypted. In Azure, configure this as an app secret now and later source the same setting from Azure Key Vault.

`HALO_DB_PATH` defaults to `./data/halo.sqlite` relative to the process working directory. When running from `Halo Email Integration`, the default database path is `Halo Email Integration/data/halo.sqlite`.

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
7. The default scope list includes `User.Read`, plus the OpenID profile scopes needed for sign-in. Keep `ADDIN_AUTH_SCOPES` unset unless you have a reason to change the requested Microsoft scopes.

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

The Halo client ID is public OAuth application metadata. The add-in does not ask for or use a Halo client secret. On first use, the task pane asks for the Halo URL, for example `https://your-company.halopsa.com`, and the Halo API application client ID.

## Local Development

Set the required environment variables, then start the Office add-in dev server:

```powershell
$env:HALO_TOKEN_ENCRYPTION_KEY="<generated key>"
$env:ADDIN_CLIENT_ID="<Microsoft Entra application client ID>"
cd ".\Halo Email Integration"
npm run dev-server
```

The local Halo callback URL is:

```text
https://localhost:3000/auth/callback
```

If you sideload the manifest into Outlook, the task pane will silently authenticate the Microsoft user, then prompt for Halo setup if that Microsoft user does not already have a stored Halo grant.

## Production-Style Run

For an Azure-style build, set the public HTTPS origin and run the production server:

```powershell
$env:HALO_TOKEN_ENCRYPTION_KEY="<generated key>"
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

For Azure, mount or persist the SQLite database location if you need state to survive container replacement. Set `HALO_DB_PATH` to that mounted path when it differs from the default.
