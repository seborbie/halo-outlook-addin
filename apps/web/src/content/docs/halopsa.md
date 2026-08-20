# HaloPSA API setup

Each customer organisation supplies its own HaloPSA origin and OAuth application client ID. These values are stored on that organisation and are never shared with another tenant.

## Create the application

1. Open the API application area in HaloPSA.
2. Create a new application using **Authorisation Code (Native Application)**.
3. Add `https://your-addin-host.example.com/auth/callback` as the redirect URI.
4. Grant ticket read/search access and permission to add ticket actions.
5. Copy the public client ID into the InboxLink signup form.

InboxLink does not ask for a HaloPSA client secret. The user completes OAuth with PKCE, and the resulting token grant is encrypted before it reaches PostgreSQL.

## Changing a connection

Changing the organisation’s Halo URL or client ID invalidates its existing user grants. Ask affected users to connect HaloPSA again so no token issued for the old application is reused with the new configuration.
