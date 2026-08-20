# InboxLink overview

InboxLink is an Outlook add-in that connects a user’s Microsoft identity to that user’s HaloPSA OAuth grant. It lets a service desk attach an email conversation to a ticket, then keep later replies on the same ticket without repeatedly searching for it.

## What it does

- Lists the signed-in agent’s open HaloPSA tickets.
- Searches by ticket number.
- Attaches the full visible chain the first time a conversation is linked.
- Adds only the newest reply on later sends and receives.
- Prevents duplicate message attachments.
- Supports a short-lived bug-report flow without storing mailbox content.

## How the pieces fit

Outlook obtains a delegated Microsoft access token for the InboxLink API. The API validates the token, uses its `tid` claim to select exactly one organisation, and uses its `oid` claim to identify the user inside that organisation. The user then authorises that organisation’s configured HaloPSA OAuth application.

The marketing site, pricing, signup, and these guides are served by the SvelteKit/Vite frontend. During local development, Vite proxies `/api` calls to the add-in service.

> InboxLink is independent software. HaloPSA, the HaloPSA name, and Halo logos are trademarks of Halo Service Solutions.
