# Troubleshooting

## Microsoft sign-in loops or returns 401

Confirm that the requested scope targets the InboxLink API, admin consent is present where required, and the access token contains `tid`, `oid`, and `access_as_user`. A Microsoft Graph token is not valid for the InboxLink API.

## The company is not configured

Open the signup page with a Microsoft administrator from that organisation. Confirm the HaloPSA URL uses HTTPS and contains only the origin, then save the native OAuth client ID.

## Halo login callback fails

The HaloPSA application redirect URI must exactly match `https://your-addin-host.example.com/auth/callback`. Also confirm that `PUBLIC_BASE_URL` contains only the HTTPS origin and matches the host baked into the production manifest.

## PostgreSQL refuses a query

Check that `DATABASE_URL` points to the migrated database and that the application is not running under a superuser. A missing tenant context should produce an access error rather than return unscoped rows.

## Sent replies do not attach

Confirm the background send runtime URL in the built manifest, reconnect the user if their Halo grant expired, and ensure the first message in the conversation was attached manually so a tenant-scoped mapping exists.
