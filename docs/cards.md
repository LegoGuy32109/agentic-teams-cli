# Optional cards

Ordinary chat and channel messages work without a card router.
Cards use a Power Automate Workflow configured for your account and an explicit channel allowlist.
The public flow template contains placeholders, not an account's webhook URL or channel identifiers.

1. Open Workflows in Teams and establish a Teams connection.
2. Sign Azure CLI into the same tenant, or supply `TEAMS_FLOW_TOKEN` and `TEAMS_POWERAPPS_TOKEN` through your environment.
3. Preview setup with `teams-cli card setup --channel 'Team/Channel'`.
4. After approval, repeat with `--allow-write` to create or update the router.

The CLI stores the webhook URL in the protected local profile. Never publish it.
Setup uses the first connected Teams connection. Inspect its preview before applying.
If your tenant blocks Power Automate or the required APIs, ask its administrator about the restriction.
An existing router keeps its allowlist when channels are added.
For a shared operational flow, arrange ownership and recovery. Microsoft documents that user-owned workflows can become orphaned without co-owners.

```sh
teams-cli card templates
teams-cli card example progress > progress.json
teams-cli card send --draft --template progress --data progress.json
# After approval, send to your private Workflows chat:
teams-cli card send --draft --template progress --data progress.json --allow-write --wait
```

Inspect the private result before approving a channel send.
A card draft is a real Teams write. Without `--allow-write`, the CLI only previews.
`--wait` looks for the posted result. A webhook's 202 response only reports acceptance.

The router supports private drafts, allowlisted channel posts, thread replies, and updates.
Cards stay under the CLI's Teams card size limit. Preview checks the serialized size before sending.
External images can load slowly or fail. SVG data URIs do not render in the recorded client checks.
Examples use generic data. Replace it locally and inspect every mention and link.

References: [Microsoft webhook guidance](https://learn.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook),
[card formatting](https://learn.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-format).
