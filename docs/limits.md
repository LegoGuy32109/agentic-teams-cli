# Scope and limits

The CLI uses internal Teams web endpoints. It is not a supported Microsoft Graph client.
Endpoint changes and tenant policies can stop an operation. Sign-in may require MFA again.
Work and school accounts are the intended audience. Personal accounts are unverified.

The application paces sends at 20 per minute. A network error or server failure after a write can leave delivery uncertain.
Inspect history and the ledger before deciding what to do next.
Do not delete the ledger to force an immediate retry.

Inline image operations currently support the Amer region. Other regions need verification.
The release includes Windows and macOS builds. Human sign-in, tenant access, and visual rendering still require native checks.

Cards use a separate Power Automate Workflow. Acceptance by the webhook is not confirmation that Teams displayed the card.
Use `--wait` to look for the resulting message, then inspect it in Teams.
Some elements depend on the client version and theme. User mentions, channel mentions, and card mentions have different support.
Do not promise identical rendering on mobile, desktop, and browser clients.

Credentials live on the local machine. File permissions are restricted where supported.
The CLI's Linux and macOS browser profiles use browser flags that avoid operating-system keychain prompts.
These profiles contain sensitive session material. Restrict access to the user account and machine.
