# Quickstart

Use this checklist for a fresh pilot deployment.

## 1. Prepare PostgreSQL

Create a PostgreSQL 15 or newer database and set `DATABASE_URL` for the add-in service. Start the service once to apply the included migrations, or run the migration command in your deployment pipeline.

```bash
export DATABASE_URL='postgresql://inboxlink:password@localhost:5432/inboxlink'
```

## 2. Register Microsoft Entra

Create a multi-tenant app registration, expose the `access_as_user` delegated permission, and configure the Outlook nested-app redirect origins. Follow [Microsoft Entra setup](/docs/microsoft-entra) for the exact values.

## 3. Build and run the add-in

Set `ADDIN_CLIENT_ID`, `HALO_TOKEN_ENCRYPTION_KEY`, `PUBLIC_BASE_URL`, and `DATABASE_URL`, then build and start the service.

```bash
cd "Halo Email Integration"
npm ci
npm run build
npm run migrate
npm run serve
```

## 4. Register the first company

Open `/signup`, enter the company name and its HaloPSA application details, then continue with a Microsoft administrator from that organisation. InboxLink uses the verified Microsoft tenant ID as the organisation boundary.

## 5. Deploy the manifest

Validate `manifest.xml`, then deploy it to a pilot group from the Microsoft 365 admin centre. Add one received email to a test ticket and reply once to verify both directions.
