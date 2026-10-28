# Agent instructions

Use Bun. Run `bun run typecheck`, `bun test`, and `bun run build` before proposing a release.
Tests use fake tokens and services. Never authenticate to a real tenant during tests.
Never print, copy, or commit browser profiles, tokens, captured requests, or webhook URLs.

All Teams writes preview by default. `--allow-write` authorizes an approved message and destination.
Preserve the send-once ledger, pacing, and uncertain-delivery behavior.
Do not retry an uncertain write.
Add commands in `src/spec.ts`, then run `bun run docs` to update the command table.

This is a downstream distribution. Read `docs/maintenance.md` and port runtime changes upstream before the next export.
Do not commit, tag, push, publish releases, or send Teams messages without the maintainer's approval.
