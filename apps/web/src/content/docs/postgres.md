# PostgreSQL deployment

SQLite is not supported by the SaaS service. PostgreSQL provides durable concurrency, composite tenant keys, and row-level security.

## Required setting

```text
DATABASE_URL=postgresql://user:password@host:5432/inboxlink?sslmode=require
```

Use a database role dedicated to InboxLink. The service refuses to start with a PostgreSQL superuser or a role carrying `BYPASSRLS`. The migrations force RLS on tenant-owned tables and every application transaction sets `app.current_organisation_id` before reading or writing tenant data.

## Apply migrations

```bash
cd "Halo Email Integration"
npm run migrate
```

Migrations are safe to rerun. Apply them before directing live traffic at a new release. Back up the database before applying a future destructive migration.

## Import an existing SQLite deployment

Keep the existing token-encryption key, set the old deployment’s global `HALO_URL` and `HALO_CLIENT_ID`, then run:

```bash
cd "Halo Email Integration"
npm run migrate:sqlite -- --sqlite ./data/halo.sqlite
```

Users, active encrypted Halo grants, conversations, and message mappings are imported. Opaque login and bug-report sessions cannot be recovered from their hashes and are intentionally skipped. If the old database contains more than one Microsoft tenant, set `LEGACY_DEFAULT_MICROSOFT_TENANT_ID` to the tenant that owns the legacy mapping tables, because those SQLite tables did not record tenancy.

## Connection security

- Require TLS between the application and PostgreSQL.
- Store `DATABASE_URL` in the hosting platform’s secret store.
- Limit network access to the application service and operational tooling.
- Rotate database credentials on a defined schedule.
- Monitor rejected connections and slow queries.
