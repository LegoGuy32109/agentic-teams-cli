import { commandSpec } from "./spec.ts";

/** Every flag the CLI knows. `true` means the flag takes a value. */
export const FLAG_SPECS = {
  check: false, force: false,
  json: false, content: false, verify: false, "allow-write": false, all: false, system: false, "no-wait": false, refresh: false, unread: false, "all-unread": false,
  profile: true, browser: true, top: true, "after-ts": true, thread: true, format: true, html: true, emoji: true, "thread-ts": true,
  count: true, from: true, after: true, before: true, at: true, image: true, file: true, out: true, access: true, role: true,
  in: true, on: true, type: true, files: false, draft: false, wait: false, channel: true, template: true, data: true, "no-react": false, "no-preview": false,
} as const;
export type FlagName = keyof typeof FLAG_SPECS;

/** The flags a command accepts: its own from `src/spec.ts`, and `--profile` unless the spec says no. */
export function allowedFlags(key: string): FlagName[] {
  const spec = commandSpec(key);
  return spec.profile === false ? spec.flags : [...spec.flags, "profile"];
}
