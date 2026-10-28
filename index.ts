import { authStatus, auth, logout } from "./src/commands/auth.ts";
import { conversations } from "./src/commands/conversations.ts";
import { usersFind } from "./src/commands/users.ts";
import { emojiFind } from "./src/commands/emoji.ts";
import { setPreferredTone } from "./src/emoji.ts";
import { configFileSetting, configSkinTone, preferredSkinTone } from "./src/settings.ts";
import { membersOf } from "./src/commands/members.ts";
import { cacheUsers } from "./src/commands/cache.ts";
import { editMessage } from "./src/commands/edit.ts";
import { history } from "./src/commands/history.ts";
import { ledgerClear, ledgerList, send } from "./src/commands/send.ts";
import { deleteMessages } from "./src/commands/delete.ts";
import { scheduled, scheduledCancel } from "./src/commands/scheduled.ts";
import { RateLimitError } from "./src/pace.ts";
import { splitTerminator } from "./src/args.ts";
import { credentialsPath, selectProfile } from "./src/config.ts";
import { redact } from "./src/tokens.ts";
import { authDoctor } from "./src/doctor.ts";
import { configureFirstRun } from "./src/first-run.ts";
import { VERSION } from "./src/version.ts";
import { update } from "./src/commands/update.ts";
import { setupRefresh, refreshAll } from "./src/refresh.ts";
import { authLogin, authRenew } from "./src/browser-auth.ts";
import { search } from "./src/commands/search.ts";
import { unread } from "./src/commands/unread.ts";
import { context } from "./src/commands/context.ts";
import { api } from "./src/commands/api.ts";
import { reactToMessage } from "./src/commands/react.ts";
import { markRead } from "./src/commands/mark.ts";
import { cardSend, cardUpdate } from "./src/commands/card.ts";
import { cardSetup } from "./src/commands/card-setup.ts";
import { cardExample, cardTemplates, cardVotes } from "./src/commands/card-templates.ts";
import { filesDelete, filesDownload } from "./src/commands/files.ts";
import { allowedFlags, FLAG_SPECS, type FlagName } from "./src/flags.ts";
import { commandHelp, fullHelp, helpJson, topicHelp, topicJson, usageLines } from "./src/help.ts";

const args = process.argv.slice(2);
const command = args[0];
const rest = args.slice(1);
// Everything after a bare `--` is text, even when it starts with `--`.
const { flagArgs, literalArgs } = splitTerminator(rest);
const consumed = new Set<number>();
let parseError: string | undefined;

const flagSpecs = FLAG_SPECS;

function flag(name: string): boolean { return flagArgs.includes(`--${name}`); }

function value(name: string): string | undefined {
  const inline = flagArgs.find(arg => arg.startsWith(`--${name}=`));
  if (inline) return inline.slice(name.length + 3);
  const index = flagArgs.indexOf(`--${name}`);
  const next = index >= 0 ? flagArgs[index + 1] : undefined;
  if (next && !next.startsWith("--")) {
    consumed.add(index + 1);
    return next;
  }
  if (index >= 0) parseError ??= `--${name} requires a value.`;
  return undefined;
}

function invocationError(message: string): never {
  const own = command && commandHelp([command]) ? `Run \`teams-cli ${command} --help\` for this command, or ` : "Run ";
  throw new Error(`${message}\n${own}\`teams-cli --help\` to see valid commands and flags.`);
}

/** Reads every `--image PATH` or `--image=PATH`, or `--file`. The flag can repeat. */
function values(name: string): string[] {
  const found: string[] = [];
  flagArgs.forEach((arg, index) => {
    if (arg.startsWith(`--${name}=`)) found.push(arg.slice(name.length + 3));
    else if (arg === `--${name}`) {
      const next = flagArgs[index + 1];
      if (next && !next.startsWith("--")) { consumed.add(index + 1); found.push(next); } else parseError ??= `--${name} requires a value.`;
    }
  });
  return found;
}

const images = values("image");
const uploadFiles = values("file");
const channels = values("channel");
const access = value("access");
const role = value("role");
const profile = value("profile");
const browser = value("browser");
const topRaw = value("top");
const afterTs = value("after-ts");
const thread = value("thread");
const format = value("format");
const emoji = value("emoji") as "teams" | "plain" | undefined;
const html = value("html");
const threadTs = value("thread-ts");
const count = value("count");
const from = value("from");
const after = value("after");
const before = value("before");
const on = value("on");
const inConversation = value("in");
const fileType = value("type");
const at = value("at");
const out = value("out");
const template = value("template");
const templateData = value("data");
const opts = {
  json: flag("json"), content: flag("content"), verify: flag("verify"), allowWrite: flag("allow-write"),
  all: flag("all"), system: flag("system"), noWait: flag("no-wait"), refresh: flag("refresh"), unread: flag("unread"), allUnread: flag("all-unread"),
  top: topRaw === undefined ? undefined : Number(topRaw), afterTs, thread, format, html, threadTs,
};
const positional = [...flagArgs.filter((arg, index) => !arg.startsWith("--") && !consumed.has(index)), ...literalArgs];

function validateFlags(allowed: FlagName[]) {
  const allowedSet = new Set<string>(allowed);
  for (const arg of flagArgs) {
    if (!arg.startsWith("--")) continue;
    const [name, inlineValue] = arg.slice(2).split("=", 2);
    if (!name || !(name in flagSpecs)) invocationError(`Unknown flag \`${arg}\`.`);
    const expectsValue = flagSpecs[name as FlagName];
    if (!expectsValue && inlineValue !== undefined) invocationError(`Flag \`--${name}\` does not take a value.`);
    if (!allowedSet.has(name)) invocationError(`Flag \`--${name}\` is not valid with \`teams-cli ${command}\`.`);
  }
}

/** Checks one command's flags and its number of arguments against `src/spec.ts`. A usage error prints the spec's usage lines. */
function check(key: string, minimum: number, maximum: number) {
  validateFlags(allowedFlags(key));
  if (positional.length < minimum || positional.length > maximum) invocationError(usageLines(key));
}

/** `card send` and `card update` take a card file, or `--template` with `--data`. Never both. */
function checkCardSource(key: string) {
  validateFlags(allowedFlags(key));
  const withFile = 3;
  if (template === undefined && templateData === undefined) return check(key, withFile, withFile);
  if (template === undefined) invocationError("--data goes with --template. Give --template NAME --data FILE.json.");
  if (positional.length >= withFile) invocationError("Give either a card file or --template, not both. A template builds the card from --data.");
  if (templateData === undefined) invocationError(`--template ${template} needs --data FILE.json. Print example data with: teams-cli card example ${template}`);
  return check(key, withFile - 1, withFile - 1);
}

function usageError(key: string): never {
  return invocationError(usageLines(key));
}

function validateInvocation() {
  const sub = positional[0] ?? "";
  switch (command) {
    case "auth":
      if (["login", "renew", "doctor", "status", "setup-refresh", "refresh"].includes(sub)) return check(`auth ${sub}`, 1, 1);
      return check("auth", 0, 0);
    case "logout":
      return check("logout", 0, 0);
    case "config":
      if (["skin-tone", "file-access", "file-role"].includes(sub)) return check("config", 1, 2);
      if (sub !== "path") { validateFlags(allowedFlags("config")); usageError("config"); }
      return check("config", 1, 1);
    case "conversations":
      return check("conversations", 0, 0);
    case "users":
      if (sub !== "find") usageError("users find");
      return check("users find", 2, Infinity);
    case "emoji":
      if (sub !== "find") usageError("emoji find");
      return check("emoji find", 2, Infinity);
    case "cache":
      if (sub !== "users" || !flag("refresh")) { validateFlags(allowedFlags("cache users")); usageError("cache users"); }
      return check("cache users", 1, 1);
    case "members":
      return check("members", 1, 1);
    case "history":
      return check("history", 1, 1);
    case "search":
      return check("search", 0, Infinity);
    case "unread":
      return check("unread", 0, 0);
    case "context":
      return check("context", 1, 1);
    case "api":
      return check("api", 1, 2);
    case "files":
      if (sub === "delete") return check("files delete", 2, 2);
      if (sub !== "download") { validateFlags(allowedFlags("files download")); invocationError(`${usageLines("files download")}\n${usageLines("files delete")}`); }
      return check("files download", 2, 2);
    case "card":
      if (sub === "setup") return check("card setup", 1, 1);
      if (sub === "templates") return check("card templates", 1, 1);
      if (sub === "example") return check("card example", 2, 2);
      if (sub === "votes") return check("card votes", 2, 2);
      if (sub !== "send" && sub !== "update") {
        validateFlags(allowedFlags("card send"));
        invocationError(["card send", "card update", "card templates", "card example", "card votes", "card setup"].map(usageLines).join("\n"));
      }
      return checkCardSource(`card ${sub}`);
    case "send":
      return check("send", 1, Infinity);
    case "scheduled":
      if (sub === "cancel") return check("scheduled cancel", 2, 2);
      return check("scheduled", 0, 0);
    case "edit":
      return check("edit", 2, Infinity);
    case "delete":
      return check("delete", 2, Infinity);
    case "react":
    case "unreact":
      return check(command, 3, 3);
    case "mark":
      return check("mark", flag("all-unread") ? 0 : 1, flag("all-unread") ? 0 : 1);
    case "ledger":
      if (sub === "list") return check("ledger", 1, 1);
      if (sub === "clear") return check("ledger", 2, 2);
      validateFlags(allowedFlags("ledger"));
      return usageError("ledger");
    case "update":
      return check("update", 0, 0);
    case "version":
      return check("version", 0, 0);
    default:
      invocationError(`Unknown command \`${command}\`.`);
  }
}

function usage() {
  console.log(fullHelp());
}

async function main() {
  if (!command) return usage();
  if (parseError) invocationError(parseError);
  if (command === "help" || command === "--help" || command === "-h") {
    const words = rest.filter(arg => arg !== "--json");
    if (words.length === 1 && topicHelp(words[0]!)) {
      return console.log(rest.includes("--json") ? JSON.stringify(topicJson(words[0]!), null, 2) : topicHelp(words[0]!));
    }
    if (words.length && !commandHelp(words)) invocationError(`No help for \`${words.join(" ")}\`.`);
    if (rest.includes("--json")) return console.log(JSON.stringify(helpJson(words), null, 2));
    return words.length ? console.log(commandHelp(words)) : usage();
  }
  if (command === "--version") {
    if (rest.length) invocationError("Usage: teams-cli --version");
    return console.log(VERSION);
  }
  // `<command> [subcommand] --help` prints that command's help and runs nothing.
  if (flagArgs.includes("--help") || flagArgs.includes("-h")) {
    const words = [command, ...flagArgs.filter(arg => !arg.startsWith("-"))].slice(0, 2);
    if (!commandHelp(words)) invocationError(`Unknown command \`${command}\`.`);
    return console.log(flagArgs.includes("--json") ? JSON.stringify(helpJson(words), null, 2) : commandHelp(words));
  }
  validateInvocation();
  // Tests set this to check that an example parses, without running it.
  if (process.env.TEAMS_CLI_CHECK_ARGS === "1") return console.log("arguments ok");
  if (command === "update") return update({ check: flag("check"), force: flag("force"), json: opts.json });
  selectProfile(profile);
  // A person at a terminal may be offered a sign-in when this machine has none. `--json` never prompts.
  configureFirstRun({ enabled: !opts.json, browser });
  if (command === "version") return console.log(VERSION);
  if (command === "card" && positional[0] === "templates") return cardTemplates(opts.json);
  if (command === "card" && positional[0] === "example") return cardExample(positional[1]!);
  if (command === "config" && positional[0] === "skin-tone") return configSkinTone(positional[1]);
  if (command === "config" && (positional[0] === "file-access" || positional[0] === "file-role")) return configFileSetting(positional[0], positional[1]);
  if (command === "config") return console.log(credentialsPath());
  // Emoji resolve while a message composes, so the saved tone loads once before any command.
  setPreferredTone(await preferredSkinTone());
  if (command === "auth" && positional[0] === "status") return authStatus();
  if (command === "auth" && positional[0] === "doctor") return authDoctor({ json: opts.json, browser });
  if (command === "auth" && positional[0] === "setup-refresh") {
    await setupRefresh(await Bun.stdin.text());
    return console.log("Automatic token renewal configured for this profile.");
  }
  if (command === "auth" && positional[0] === "refresh") return refreshAll();
  if (command === "auth" && positional[0] === "login") return authLogin({ browser });
  if (command === "auth" && positional[0] === "renew") return authRenew({ browser });
  if (command === "auth") return auth();
  if (command === "logout") return logout();
  if (command === "conversations") return conversations(opts.json);
  if (command === "search") return search(positional.join(" "), { json: opts.json, count: count === undefined ? undefined : Number(count), from, after, before, on, in: inConversation, files: flag("files"), type: fileType });
  if (command === "unread") return unread(opts.json, opts.refresh);
  if (command === "context") return context(positional[0]!, { json: opts.json, before: before === undefined ? undefined : Number(before), after: after === undefined ? undefined : Number(after) });
  if (command === "api") return api(positional[0]!, positional[1], opts.json);
  if (command === "emoji") return emojiFind(positional.slice(1).join(" "), opts.json);
  if (command === "users") return usersFind(positional.slice(1).join(" "), opts.json, opts.refresh);
  if (command === "members") return membersOf(positional[0]!, opts.json);
  if (command === "cache") return cacheUsers();
  if (command === "files" && positional[0] === "delete") return filesDelete(positional[1]!, { json: opts.json, allowWrite: opts.allowWrite, noWait: opts.noWait });
  if (command === "files") return filesDownload(positional[1]!, { out, json: opts.json });
  if (command === "card" && positional[0] === "setup") return cardSetup({ json: opts.json, allowWrite: opts.allowWrite, channels });
  const cardSource = template === undefined ? positional[2]! : { template, data: templateData! };
  if (command === "card" && positional[0] === "votes") return cardVotes(positional[1]!, { json: opts.json });
  if (command === "card" && positional[0] === "update") return cardUpdate(positional[1]!, cardSource, { json: opts.json, allowWrite: opts.allowWrite, noWait: opts.noWait });
  if (command === "card") return cardSend(positional[1]!, cardSource, { json: opts.json, allowWrite: opts.allowWrite, draft: flag("draft"), thread: opts.thread, noWait: opts.noWait, wait: flag("wait"), noReact: flag("no-react") });
  if (command === "send") {
    return send(positional[0]!, positional.slice(1), {
      json: opts.json, format, html, verify: opts.verify, allowWrite: opts.allowWrite, threadTs: opts.threadTs, noWait: opts.noWait, noPreview: flag("no-preview"), at, images, files: uploadFiles, emoji, access, role,
    });
  }
  if (command === "scheduled") {
    if (positional[0] === "cancel") return scheduledCancel(positional[1]!, { json: opts.json, allowWrite: opts.allowWrite, noWait: opts.noWait });
    return scheduled(opts.json);
  }
  if (command === "edit") {
    return editMessage(positional[0]!, positional[1]!, positional.slice(2), { json: opts.json, format, html, emoji, allowWrite: opts.allowWrite, noWait: opts.noWait, noPreview: flag("no-preview") });
  }
  if (command === "delete") {
    return deleteMessages(positional[0]!, positional.slice(1), { json: opts.json, allowWrite: opts.allowWrite, noWait: opts.noWait, threadTs: opts.threadTs });
  }
  if (command === "react" || command === "unreact") {
    return reactToMessage(positional[0]!, positional[1]!, positional[2]!, command === "unreact", { json: opts.json, allowWrite: opts.allowWrite, noWait: opts.noWait, verify: opts.verify });
  }
  if (command === "mark") return markRead(positional[0], { json: opts.json, allowWrite: opts.allowWrite, noWait: opts.noWait, allUnread: opts.allUnread });
  if (command === "ledger") {
    if (positional[0] === "list") return ledgerList(opts.json);
    return ledgerClear(positional[1]!);
  }
  if (command === "history") {
    return history(positional[0]!, {
      json: opts.json, content: opts.content, top: topRaw === undefined ? undefined : opts.top, afterTs: opts.afterTs, thread: opts.thread,
      all: opts.all, system: opts.system, unread: opts.unread,
    });
  }
}

main().catch(error => {
  if (error instanceof RateLimitError && opts.json) {
    console.log(JSON.stringify({ ok: false, error: redact(error.message), retry_after_seconds: error.retry_after_seconds }, null, 2));
    process.exitCode = 1;
    return;
  }
  console.error(redact(error instanceof Error ? error.message : String(error)));
  process.exitCode = 1;
});
