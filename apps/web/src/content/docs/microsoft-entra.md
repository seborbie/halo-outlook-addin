# Microsoft Entra setup

InboxLink uses Microsoft nested app authentication (NAA) in Outlook and a delegated API permission. Configure the app as multi-tenant when selling to more than one company.

## App registration

1. Create an app registration in Microsoft Entra ID.
2. Choose **Accounts in any organisational directory**.
3. Add the local SPA redirect URI `brk-multihub://localhost:3000`.
4. Add `brk-multihub://your-addin-host.example.com` for production.
5. Add your website signup callback, such as `https://app.example.com/signup`, as a SPA redirect URI.
6. Under **Expose an API**, accept `api://<ADDIN_CLIENT_ID>`.
7. Create a delegated scope named `access_as_user`.
8. Pre-authorise the app’s own client ID for that scope.
9. Set `api.requestedAccessTokenVersion` to `2` in the manifest.

## Runtime settings

```text
ADDIN_CLIENT_ID=<application-client-id>
ADDIN_API_AUDIENCE=api://<application-client-id>
ADDIN_AUTH_SCOPES=api://<application-client-id>/access_as_user
ADDIN_REQUIRED_SCOPE=access_as_user
```

The API rejects tokens without a stable Microsoft tenant ID (`tid`), user object ID (`oid` or `sub`), and the delegated scope. Never accept a tenant identifier sent in a request body as proof of tenancy.
