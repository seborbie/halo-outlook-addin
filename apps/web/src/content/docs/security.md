# Security and tenant isolation

Tenant separation is the primary security invariant: a request authenticated for one Microsoft organisation must never see, change, or reuse another organisation’s users, grants, sessions, bug-report links, or email-to-ticket mappings.

## Defence in depth

1. Microsoft access tokens are cryptographically verified for audience, issuer, delegated scope, tenant ID, and user ID.
2. Every tenant-owned PostgreSQL row contains `organisation_id`.
3. Composite foreign keys prevent a row from referencing a parent in another organisation.
4. Application queries always include the organisation ID.
5. PostgreSQL row-level security checks the transaction’s `app.current_organisation_id` setting.
6. Opaque session identifiers carry a non-secret organisation prefix so the server can establish RLS context before looking up the hashed token.

## Stored data

InboxLink stores user identity fields, encrypted Halo OAuth tokens, hashed session identifiers, and message-to-ticket mapping identifiers. It does not intentionally persist email bodies or attachments. Bug-report diagnostics are limited to add-in and Outlook version information.

## Encryption

Set `HALO_TOKEN_ENCRYPTION_KEY` to a random 32-byte base64 or base64url value. Keep the key stable across restarts and source it from a managed secret store. Rotating it requires an explicit grant re-encryption plan or user reauthorisation.
