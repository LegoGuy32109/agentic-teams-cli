import { createInterface } from "node:readline";
import { loadProfile } from "./config.ts";
import { externalTokenSource } from "./tokens.ts";

/**
 * Offers a browser sign-in the first time a command needs tokens and the machine has none.
 *
 * Only a person at a terminal is asked. An agent, a pipe, or `--json` never sees a prompt: the caller gets the
 * ordinary error that points at `auth login`. The prompt and the terminal check are hooks, so tests cover both
 * paths without a browser.
 */

export const SIGN_IN_QUESTION = "No Teams sign-in on this machine. Sign in now in a browser window? [Y/n] ";
export const RUN_LOGIN = "Run: teams-cli auth login";

export type FirstRunHooks = {
  /** True when stdin and stdout are both a terminal. */
  isTerminal(): boolean;
  /** Asks the question and returns the answer line. */
  ask(question: string): Promise<string>;
  /** Runs the same flow as `auth login`. */
  login(options: { browser?: string }): Promise<void>;
};

const defaultHooks: FirstRunHooks = {
  isTerminal: () => !!process.stdin.isTTY && !!process.stdout.isTTY,
  ask: question => new Promise(resolve => {
    const lines = createInterface({ input: process.stdin, output: process.stderr });
    lines.question(question, answer => { lines.close(); resolve(answer); });
  }),
  login: async options => (await import("./browser-auth.ts")).authLogin(options),
};

let hooks: FirstRunHooks = defaultHooks;
let enabled = false;
let browser: string | undefined;
let pending: Promise<boolean> | undefined;

/** The entry point turns the offer on for a run without `--json`, and passes `--browser`. Off by default. */
export function configureFirstRun(options: { enabled: boolean; browser?: string }): void {
  enabled = options.enabled;
  browser = options.browser;
  pending = undefined;
}

/** Tests replace the hooks. Call it with no argument to restore the real ones. */
export function setFirstRunHooks(next?: Partial<FirstRunHooks>): void {
  hooks = { ...defaultHooks, ...next };
  pending = undefined;
}

/** A yes is empty, `y`, or `yes`, in any case. */
export function isYes(answer: string): boolean {
  return /^\s*(y|yes)?\s*$/i.test(answer);
}

async function offer(): Promise<boolean> {
  if (!enabled || externalTokenSource()) return false;
  if ((await loadProfile()).profile) return false;
  if (!hooks.isTerminal()) return false;
  if (!isYes(await hooks.ask(SIGN_IN_QUESTION))) throw new Error(RUN_LOGIN);
  await hooks.login({ browser });
  return true;
}

/**
 * Returns true when the person signed in, so the caller can try again. Returns false when no offer applies.
 * Throws `Run: teams-cli auth login` when the person says no. Parallel callers share one prompt.
 */
export function offerSignIn(): Promise<boolean> {
  pending ??= offer().finally(() => { pending = undefined; });
  return pending;
}
