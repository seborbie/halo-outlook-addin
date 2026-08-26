# Conversation mapping recovery

Replies sent from the event runtime first use Outlook's normal conversation and message identifiers.
When those identifiers do not yet match, the runtime races two bounded metadata lookups: the opaque
`X-Halo-Compose-Id` header and the add-in custom property retained on the exact referenced Sent Items
message. Candidate compose identifiers are revalidated by the Halo add-in server; Outlook metadata is
never trusted as a ticket identifier.

This uses EWS through the add-in's `ReadWriteMailbox` permission. That permission can read or write any
item or folder in the user's mailbox and can send mail, even though this feature restricts itself to
exact Internet Message ID lookups and metadata reads in Sent Items (plus referenced-item header reads).
Deployment therefore requires explicit Microsoft 365 administrator approval, and EWS must be allowed
for the target Exchange mailboxes.

No Microsoft Graph permission is requested. Recovery does not use timestamps, subjects, recipients,
or message-body markers. If recovery is unavailable or produces no validated mapping, Outlook sends
normally and Halo receives nothing. Once a mapping is validated, failures in the required Halo action
or attachment commit use the existing Smart Alerts `PromptUser` warning.

Supported first-release targets are primary Exchange mailboxes in Outlook on the web, new/classic
Outlook for Windows, and Outlook for Mac. Mobile Outlook, shared/delegated mailboxes, Gmail, and other
non-Exchange accounts are not supported acceptance targets.
