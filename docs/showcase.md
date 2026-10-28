# Screenshots and examples

These examples use fictional projects, people, and results. They contain no user mentions or actual project links.
The screenshots should show messages sent by this CLI and rendered by Teams.

## Styled message

The source is [styled-message.txt](../examples/styled-message.txt).
It demonstrates headings, status badges, colored text, a table, a highlighted scope note, and a labeled link.
Preview it from the repository root:

```sh
teams-cli send me "$(cat examples/styled-message.txt)" --format=rich --no-preview
```

After reviewing the preview, repeat with `--allow-write` to send to Notes to self.
`--no-preview` keeps the link clickable without an extra page preview card.

## Highlighted commands

The source is [code-message.txt](../examples/code-message.txt).
It demonstrates fenced shell code, inline code, and colored text.

```sh
teams-cli send me "$(cat examples/code-message.txt)" --format=rich --no-preview
```

PowerShell can read a whole source file with `Get-Content -Raw`:

```powershell
teams-cli send me (Get-Content -Raw examples\code-message.txt) --format=rich --no-preview
```

## Progress card

[progress.json](../examples/progress.json) supplies data to the built-in progress template.
It demonstrates completion bars and status badges. Configure the [card router](cards.md) first.

```sh
teams-cli card send me --draft --template progress --data examples/progress.json
```

Add `--allow-write --wait` after approval to send to your private Workflows chat and look for the posted result.
`me` is the nominal target. `--draft` routes the card to Workflows.

## Review card

[review.json](../examples/review.json) supplies data to the built-in PR review template.
It demonstrates change counts, review checks, a scoped limitation, and supporting links in an overflow menu.

```sh
teams-cli card send me --draft --template pr-review --data examples/review.json
```

The optional [review-card.json](../examples/review-card.json) starts with that template and adds an expandable file summary.
Send that raw card to demonstrate custom Adaptive Cards:

```sh
teams-cli card send me examples/review-card.json --draft
```

## Capture guide

| Image | Destination | Frame |
| --- | --- | --- |
| `styled-message.png` | Notes to self | Full message from its heading through the example review link |
| `code-message.png` | Notes to self | Heading, both highlighted code blocks, and the inline command |
| `cards-overview.png` | Workflows | Progress and review cards together, including all three progress bars |
| `review-card.png` | Workflows | Entire card with the file summary expanded |

Use one Teams client and theme for the main set. Keep normal zoom and enough width to read the table.
Crop the sidebar, account name, tenant navigation, timestamps, and message actions where practical.
Keep the demo-data labels visible. Capture the card's own content without the Workflow wrapper if the layout allows it.
An optional second capture in a narrow pane helps check wrapping.

Store approved captures in `docs/images/`. The [image checklist](images/README.md) maps each capture to its README position.
The README includes four original captures supplied on October 28, 2026. They show Teams in dark theme.
The captures do not identify whether the client is desktop or browser. Do not infer compatibility with other clients from these images.
