# InboxLink for HaloPSA

InboxLink is a multi-tenant SaaS product that connects Microsoft Outlook conversations to HaloPSA tickets. This repository contains:

- a SvelteKit 5 marketing, pricing, signup, and documentation website in `apps/web`;
- an Outlook task-pane and background-send add-in in `Halo Email Integration`;
- an Express API using Microsoft Entra identity, encrypted Halo OAuth grants, and PostgreSQL row-level security.

HaloPSA, the HaloPSA name, and Halo logos are trademarks of Halo Service Solutions. InboxLink is independent software and is not affiliated with or endorsed by Halo.

## Requirements

- Bun 1.3 or newer for the website
- Node.js 22 and npm for the Outlook add-in
- PostgreSQL 15 or newer
- a Microsoft Entra multi-tenant app registration
- one native HaloPSA OAuth application per customer organisation

## Environment

Copy `Halo Email Integration/.env.example` to `.env` and configure at least:

```text
DATABASE_URL=postgresql://inboxlink:password@localhost:5432/inboxlink
HALO_TOKEN_ENCRYPTION_KEY=<32-byte base64url key>
ADDIN_CLIENT_ID=<Microsoft Entra application client ID>
PUBLIC_BASE_URL=https://your-public-origin.example.com
```

Generate the token key with:

```bash
node -e "console.log(require('crypto').randomBytes(32).toString('base64url'))"
```

HaloPSA URLs and public OAuth client IDs are no longer global environment variables. They are registered per Microsoft organisation through `/signup` and stored on that organisation’s isolated PostgreSQL record.

## Local development

Install both dependency sets:

```bash
bun install
npm --prefix "Halo Email Integration" ci
```

Apply the database migration:

```bash
npm --prefix "Halo Email Integration" run migrate
```

Run the Outlook add-in service on HTTPS port 3000:

```bash
npm --prefix "Halo Email Integration" run dev-server
```

Run the SvelteKit/Vite site on port 5173 in another terminal:

```bash
bun run dev
```

Vite proxies `/api` to the local add-in service, so the Microsoft-backed signup form and documentation share one frontend development server.

## Production build

Build both surfaces from the repository root:

```bash
PUBLIC_BASE_URL=https://your-public-origin.example.com bun run build
```

The production container builds the website with Bun, builds the add-in with Node, and serves the website, API, documentation, and add-in assets from one origin:

```bash
docker build \
  --file "Halo Email Integration/Dockerfile" \
  --build-arg PUBLIC_BASE_URL=https://your-public-origin.example.com \
  --tag inboxlink:local \
  .
```

Set `DATABASE_URL`, `HALO_TOKEN_ENCRYPTION_KEY`, and `ADDIN_CLIENT_ID` on the running container. Apply `npm run migrate` as a release step before accepting traffic.

## Multi-tenant security model

The verified Microsoft `tid` claim selects one InboxLink organisation; the verified `oid` claim selects a user inside it. Tenant-owned tables include `organisation_id`, composite foreign keys prevent cross-tenant references, and PostgreSQL row-level-security policies are forced on users, grants, sessions, bug-report sessions, and ticket mappings.

The application establishes `app.current_organisation_id` inside every tenant transaction. Production must use a dedicated non-superuser database role without `BYPASSRLS`.

Opaque session handles include a non-secret organisation UUID prefix. The server uses that prefix only to establish the RLS context before looking up the full hashed token.

To import an existing SQLite deployment, keep its encryption key and run `npm run migrate:sqlite -- --sqlite ./data/halo.sqlite` from `Halo Email Integration`. See the PostgreSQL documentation page for the legacy tenant-mapping requirement and skipped temporary sessions.

## Verification

Website:

```bash
bun run check
bun run build:web
```

Add-in and API:

```bash
npm --prefix "Halo Email Integration" run test:status-route
npm --prefix "Halo Email Integration" run test:send-runtime
npm --prefix "Halo Email Integration" run test:bug-reports
npm --prefix "Halo Email Integration" run test:auth
npm --prefix "Halo Email Integration" run test:tenancy
```

The authentication, bug-report, and tenancy tests require `DATABASE_URL` (or `TEST_DATABASE_URL` for the isolation test) pointing to a disposable PostgreSQL database owned by a non-superuser role.

For the complete administrator setup, see the Markdown guides in `apps/web/src/content/docs` or open `/docs` on the running site.
