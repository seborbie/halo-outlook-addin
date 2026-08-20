# Tenant onboarding

An InboxLink organisation corresponds to one Microsoft Entra tenant. The Microsoft token—not a form field—provides the authoritative tenant ID.

## Register a company

1. Open the signup page.
2. Enter the company name, HaloPSA HTTPS origin, and public Halo OAuth client ID.
3. Sign in with Microsoft using an account in the company being registered.
4. Confirm that you are authorised to configure the organisation.
5. Deploy the add-in manifest to a small pilot group.

The first verified user becomes the organisation owner. Re-registering the same Microsoft tenant updates its display and Halo connection settings instead of creating a second tenant.

## Add users

Users are created only after a valid Microsoft API token is presented. A user object ID is unique inside its own organisation, so the same identifier in another Microsoft tenant cannot resolve to the first company’s record.

## Offboarding

Disable access in Microsoft Entra first. Then revoke the user’s HaloPSA grant and delete its active InboxLink sessions. Organisation deletion should be handled as a controlled administrative operation with a retention and export decision.
