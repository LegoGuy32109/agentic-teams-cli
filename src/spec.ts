import type { FlagName } from "./flags.ts";
import {
  DEFAULT_CONTEXT, DEFAULT_SEARCH_COUNT, DEFAULT_TOP, MAX_ALL, MAX_CONTEXT, MAX_DOWNLOAD_BYTES, MAX_FILES, MAX_IMAGE_BYTES, MAX_IMAGES,
  MAX_CARD_BYTES, MAX_TEAMS_CARD_BYTES, MAX_SEARCH_COUNT, MAX_TOP, MAX_UPLOAD_BYTES, mb,
} from "./limits.ts";

/**
 * The one description of every command. The parser takes each command's flags from here, a usage error prints its
 * usage lines, `--help` and `help --json` render it, and `bun run docs` writes the README command table from it.
 * `test/help.test.ts` fails when any of those disagree.
 */

/** One usage form. `text` is the help text. `more` is detail that only the README shows. */
export type UsageEntry = { usage: string; text: string; more?: string };
/** A help example. `args` must pass the real parser (`test/help.test.ts` runs each one). */
export type Example = { args: string[]; note: string };
export type CommandSpec = {
  key: string;
  section: SectionId;
  flags: FlagName[];
  /** False for the few commands that do not take `--profile`. */
  profile?: false;
  entries: UsageEntry[];
  examples?: Example[];
};

export type SectionId = "read" | "write" | "delete" | "react" | "ledger" | "setup";
export const SECTION_TITLES: Record<SectionId, string> = {
  read: "Read (safe by default)",
  write: "Write (preview by default; --allow-write sends)",
  delete: "Delete (preview by default; --allow-write changes Teams)",
  react: "React and mark (preview by default; --allow-write changes Teams)",
  ledger: "Send-once ledger (local only)",
  setup: "Setup",
};

const CHANNEL = "Engineering/General";
const REF = "19:abc123@thread.tacv2#1792428352468";

export const COMMANDS: CommandSpec[] = [
  {
    key: "conversations", section: "read", flags: ["json"],
    entries: [{ usage: "conversations [--json]", text: "List teams, channels, and chats with their ids.", more: "Uses the csa list, and adds `meeting` and `chat` threads from the chat service that the csa list lacks. Names chats of 8 members or fewer." }],
    examples: [{ args: ["conversations"], note: "find a conversation name or id" }],
  },
  {
    key: "history", section: "read", flags: ["json", "content", "top", "after-ts", "thread", "all", "system", "unread"],
    entries: [
      { usage: "history <conversation> [--top=N | --after-ts=T | --all] [--system] [--thread ROOTID] [--content] [--json]", text: `Read messages, oldest first, so the newest is last. Default --top=${DEFAULT_TOP}, at most ${MAX_TOP}. --all reads to the start (capped at ${MAX_ALL.toLocaleString("en-US")}). --system adds calls, recordings, transcripts, and events. --thread ROOTID reads one channel thread. A message with files prints a files: line. Deleted messages are not shown.` },
      { usage: "history mentions --unread", text: "Only mentions newer than the read position of their conversation." },
    ],
    examples: [{ args: ["history", CHANNEL, "--top=5"], note: "each message prints its ref after id:" }],
  },
  {
    key: "context", section: "read", flags: ["json", "before", "after"],
    entries: [{ usage: "context <message-ref> [--before=N] [--after=N] [--json]", text: `Read a message, the messages around it (default ${DEFAULT_CONTEXT} on each side, at most ${MAX_CONTEXT}), and its channel thread.` }],
    examples: [{ args: ["context", REF, "--before=2", "--after=2"], note: "read around one message" }],
  },
  {
    key: "search", section: "read", flags: ["json", "count", "from", "after", "before", "on", "in", "files", "type"],
    entries: [
      { usage: "search <query> [--in CONVERSATION] [--from PERSON] [--after DATE] [--before DATE] [--on DATE] [--count=N] [--json]", text: `Search chat and channel messages, as the Messages tab in Teams does. --in takes any conversation form. --from takes an email, or a name that matches one person. DATE is YYYY-MM-DD, a local day. --after DATE keeps messages sent at or after the start of that day. --before DATE keeps messages sent before the start of that day. --on DATE keeps one day. Default ${DEFAULT_SEARCH_COUNT} hits, at most ${MAX_SEARCH_COUNT}.`, more: "A query of filters only is allowed when --from or --in is given. A name for --from that matches nobody or more than one person fails and lists the people. Each hit prints its time, sender, conversation, text, and ref, then markers: [files: names], [image], [mentions you], and `reply in thread <ref>`. Reads pages of 25 and removes duplicate hits." },
      { usage: "search --files [<query>] [--in CHANNEL] [--from PERSON] [--after DATE] [--before DATE] [--on DATE] [--type EXT] [--count=N] [--json]", text: "Search OneDrive and SharePoint files, as the Files tab in Teams does. --in takes a channel and keeps the files in its files folder; Teams has no file scope for a chat. Dates filter the modified time. --type takes an extension such as pdf, or loop or fluid. Each hit prints the name, the folder or site, the modified time and person, and the link. Save a hit with: files download <link>.", more: "A query of filters only is allowed when --from or --type is given." },
    ],
    examples: [
      { args: ["search", "deploy", "--from", "alex@example.com", "--after", "2026-10-01"], note: "search with filters" },
      { args: ["search", "deploy", "--in", CHANNEL, "--on", "2026-10-18"], note: "search one channel on one day" },
      { args: ["search", "--files", "report", "--type", "pdf"], note: "search for PDF files" },
      { args: ["search", "--files", "screenshot", "--in", CHANNEL], note: "search one channel's files; save a hit with files download <link>" },
    ],
  },
  {
    key: "unread", section: "read", flags: ["json", "refresh"],
    entries: [{ usage: "unread [--json] [--refresh]", text: "List unread chats and channels, with unavailable chats separate. --refresh rechecks access without marking read." }],
  },
  {
    key: "members", section: "read", flags: ["json"],
    entries: [{ usage: "members <conversation> [--json]", text: "List a conversation's members with name, mail, and role." }],
  },
  {
    key: "users find", section: "read", flags: ["json", "refresh"],
    entries: [{ usage: "users find <query> [--json] [--refresh]", text: "Find a user ID, display name, or email. A query of 2 or more characters asks the directory. --refresh skips the cache." }],
  },
  {
    key: "emoji find", section: "read", flags: ["json"], profile: false,
    entries: [{ usage: "emoji find <query> [--json]", text: "Search the Teams emoji catalog by word, emoji, or key. Works offline. Use a key with react, or :key: in message text." }],
    examples: [{ args: ["emoji", "find", "thumbs"], note: "keys and skin tones" }],
  },
  {
    key: "files download", section: "read", flags: ["json", "out"],
    entries: [{ usage: "files download <message-ref|link> [--out DIR] [--json]", text: `Save the files and inline images of a message, or one file by its SharePoint link (as search --files prints it), into DIR (default: the current directory). Works for any file you can open in Teams on a *.sharepoint.com site: chat files, channel files, and SharePoint lists (saved as <name>.json). An inline image (history shows it as [image <id>]) is saved as <message-id>-image-<n>.png, .jpg or .gif. Files are mode 0600. A file that exists is never overwritten. Each file can be ${mb(MAX_DOWNLOAD_BYTES)} at most, and each image ${mb(MAX_IMAGE_BYTES)}.`, more: "Nothing downloads if any target exists. A file on another organization's SharePoint fails with an error that names it. An image is read from the Teams image host by its id, never from the address in the message. Amer region only for images." }],
    examples: [
      { args: ["files", "download", REF, "--out", "./downloads"], note: "save a message's files" },
      { args: ["files", "download", "https://contoso.sharepoint.com/sites/eng/Shared Documents/plan.pdf"], note: "save one file by the link that search --files prints" },
    ],
  },
  {
    key: "api", section: "read", flags: ["json"],
    entries: [{ usage: "api list | api <operation> [argument] [--json]", text: "Inspect or call a verified read operation. Arbitrary URLs and methods are refused." }],
  },
  {
    key: "cache users", section: "read", flags: ["refresh"],
    entries: [{ usage: "cache users --refresh", text: "Clear the user-name cache (no bulk rebuild in Teams)." }],
  },
  {
    key: "send", section: "write", flags: ["json", "allow-write", "no-wait", "format", "html", "emoji", "verify", "thread-ts", "at", "image", "file", "access", "role", "no-preview"],
    entries: [
      { usage: "send <conversation> <text...> [--thread-ts ROOTID] [--format=rich] [--allow-write] [--no-wait] [--json]", text: "Preview or send a message. --thread-ts ROOTID replies in a channel thread; ROOTID is the root message id or its full ref. --format=rich reads Markdown-style marks: *bold*, _italic_, ~strike~, `code`, # headings, lists, > quotes, tables, links, and fenced code. Pauses at the send limit; --no-wait fails at once instead. Use -- before text that starts with --. Links: [label](https://url).", more: "`@{Name}` in the text becomes a mention. A send to `@Name` or an email starts the 1:1 chat when none exists, and only with `--allow-write`." },
      { usage: "send ... \"[label](https://url)\" | \"<https://url|label>\" | https://url", text: "Make a clickable link. Write [label](https://url) to show the label, as in Markdown. <https://url|label> (Slack style) also works. A bare https:// URL becomes a link to itself. All three work in the default format and in --format=rich. The first link gets a preview card by default; --no-preview turns it off. A page that needs a sign-in has no card, and the CLI says so.", more: "Only http and https links are made. Put the URL in the parentheses with no spaces. A label in brackets with no URL after it stays text." },
      { usage: "send ... --no-preview", text: "Send the links without a preview card. Each link stays clickable." },
      { usage: "send <conversation> --html X|@file [--allow-write]", text: "Preview or send HTML. @{Name} in its text becomes a mention (not in tags, attributes, code, or comments). Write @@{ for a literal @{." },
      { usage: "send ... \"@{Name}\" | \"@{everyone}\" | \"@{channel}\"", text: "Mention a person (name or email), everyone in a group or meeting chat, or the whole channel. @{everyone} is refused in a 1:1 chat, Notes to self, and a channel. @{channel} is refused outside a channel. To mention a person named Everyone or Channel, write their email." },
      { usage: "send <conversation> [text...] --image PATH [--image PATH]...", text: `Attach up to ${MAX_IMAGES} PNG, JPEG, or GIF images (${mb(MAX_IMAGE_BYTES)} each), shown inline. A preview uploads nothing. Amer region only. Not valid with --at.`, more: "The text comes first, then one picture per paragraph. `history` and `context` show an image as `[image <id>]`." },
      { usage: "send <conversation> [text...] --file PATH [--file PATH]...", text: `Attach up to ${MAX_FILES} files (${mb(MAX_UPLOAD_BYTES)} each). In a chat, a file goes to your OneDrive folder "Microsoft Teams Chat Files", and by default the members open it through the tenant's default sharing link, as Teams does (see --access). In Notes to self only you can open it. In a channel, it goes into the channel's files folder; a name that is there stops the send. A preview uploads nothing and shows the folder and the people. If an upload fails, no message is sent and the files already uploaded are named. Works with --at, --thread-ts, and --image.`, more: "A send does: pace, send-once check, uploads, invites, message. With `--at`, the files upload and the invites happen now, and the draft carries them. See [Files](#files)." },
      { usage: "send ... --file X --access=link|people|private --role=edit|view", text: "Override who can open a chat file for one send. link: the tenant's default link, nobody added by name. people: each member by name, sign-in required. private: only you. Set the default with config file-access and config file-role." },
      { usage: "send ... --emoji=plain", text: "Keep emoji as typed characters. The default writes the Teams emoticon, as the picker does. :key: works in both." },
      { usage: "send ... --at TIME", text: "Schedule the message instead of sending it now. TIME is ISO 8601 with an offset, HH:MM (the next time that clock time comes, so tomorrow if it has passed today), or +<n>m / +<n>h, 1 minute or more ahead. Works with --thread-ts and --file. Not valid with --verify or --image. Not available in Notes to self." },
      { usage: "send ... --verify", text: "Re-read the sent message and compare it.", more: "The compare ignores whitespace, so list and line-break markup that Teams stores in another form still matches." },
    ],
    examples: [
      { args: ["send", "me", "hello", "--file", "./report.md"], note: "preview a send to Notes to self" },
      { args: ["send", CHANNEL, "Build is *green*", "--format=rich", "--thread-ts", REF], note: "preview a bold reply in a thread" },
      { args: ["send", CHANNEL, "standup in 5 @{channel}", "--at", "09:00"], note: "preview a scheduled channel mention" },
      { args: ["send", CHANNEL, "Ready for review: [PR 1](https://example.com/pr/1)"], note: "preview a message with a labelled link and its preview card" },
    ],
  },
  {
    key: "card send", section: "write", flags: ["json", "allow-write", "no-wait", "draft", "thread", "wait", "template", "data", "no-react"],
    entries: [
      { usage: "card send <channel> <card.json> [--thread ROOTID] [--draft] [--allow-write] [--wait] [--no-wait] [--json]", text: `Preview or send an Adaptive Card to a channel, as you, through the Workflows "teams-cli card router" flow. <card.json> holds an AdaptiveCard (type, version, and a body array) or a {type: "message", attachments: [...]} payload with one card. An Image url or backgroundImage of file:<path> (relative to the card file) is replaced by a PNG, JPEG, or GIF data URI before the send. SVG is refused. The payload can be ${MAX_CARD_BYTES / 1024} KB at most. --thread ROOTID replies in a channel thread; ROOTID is the root message id or its full ref. The preview sends nothing. It shows the target, the card version and size, an outline, and each image.`, more: "The flow URL is the environment variable TEAMS_CARD_ROUTER, or the URL that card setup stored in the profile. It is a secret, and the CLI never prints it. Only an https URL on a host that ends in .powerplatform.com or .logic.azure.com is accepted. Cards do not use Teams image upload (AMS): an AMS image can be read only by the conversation it was uploaded for, so channel members may not see it. A data URI is part of the card, so everyone who sees the card sees the image." },
      { usage: "card send <channel> --template NAME --data FILE.json [--thread ROOTID] [--draft] [--allow-write] [--no-react]", text: `Build the card from a built-in template instead of a card file. --data is a JSON file with the template's fields. The CLI checks the data, names the field that is wrong, and builds the card. A card can be ${MAX_TEAMS_CARD_BYTES / 1024} KB at most, because Teams refuses a larger one; the epic template cuts hidden issues to fit, and the others fail with the size. <card.json> and --template cannot be used together. The preview adds the template name to the outline.`, more: "List the templates with card templates, and print the data of one with card example NAME. The templates are meeting-recap, deadline, epic, pr-review, decision, poll, progress, incident, standup, and qa-checklist. See docs/cards.md." },
      { usage: "card send ... --template decision|poll --allow-write [--no-react]", text: "A template that asks for a vote (decision, poll) adds each option emoji as your own reaction after the post, so people can click it. The CLI first waits up to 20 s for the new message, as --wait does, then reacts through the same path as react, with the same pacing. The preview lists the reactions. --no-react skips them. If the message is not found, the CLI says so and prints the react commands to run by hand. An emoji that has no Teams reaction key fails in the preview, before anything is sent.", more: "Count the votes with card votes REF. A reaction you already have is not added twice. The send-once ledger stops a repeated send, so it stops repeated reactions too. A reaction has no ledger record of its own." },
      { usage: "card send ... --wait", text: "After the 202 answer, read the target conversation (read only) for up to 20 s. Find your newest message, or the Workflows bot's for a draft, that has a card with the same card hash. Print its ref, so card update can use it. The ref is also kept in the ledger record. Without --wait, the command prints how to find the ref." },
      { usage: "card send <channel> <card.json> --draft --allow-write", text: "Send the card to your private Workflows chat instead, so you see the real rendering. Do this first. The card goes there whatever the channel, and a chat is accepted as <channel>. A draft and the final post of the same card are separate sends." },
      { usage: "card send <channel> <card.json> [--thread ROOTID] --allow-write", text: "Post the card in the channel as a new post, or as a reply when --thread is given. Every send passes the send-once ledger. Only a 202 answer counts as accepted, and it means the flow took the card, not that Teams posted it. The flow refuses a channel that is not on its allowlist, and it does this after the 202. A network error, a 408, or a 5xx is an unknown outcome and is never retried." },
    ],
    examples: [
      { args: ["card", "send", CHANNEL, "./card.json", "--thread", REF], note: "preview a card as a thread reply; nothing is sent" },
      { args: ["card", "send", CHANNEL, "./card.json", "--draft"], note: "preview the draft send; add --allow-write to see the real rendering" },
    ],
  },
  {
    key: "card update", section: "write", flags: ["json", "allow-write", "no-wait", "template", "data"],
    entries: [
      { usage: "card update <ref> --template NAME --data FILE.json [--allow-write]", text: "Update a posted card from a template and its data, as card send --template builds it. The incident template is made for this: send it once, then update it with the same ref as the timeline grows. <ref> and the card file rules are the same as below. <card.json> and --template cannot be used together." },
      { usage: "card update <ref> <card.json> [--allow-write] [--no-wait] [--json]", text: "Preview or make an edit of a card that was posted through the card router. <ref> is <channel>#<id> for a channel post, <channel>;messageid=<root>#<reply> for a thread reply, or draft#<id> (or a ref in your Workflows chat) for a draft. card send --wait prints the ref. The preview says what would change and makes no request. Every update passes the send-once ledger, keyed by the target and the card hash. Only a 202 answer counts as accepted.", more: "Router mode `update` takes `channel` and `messageId`. For a reply, `messageId` is <root>/replies/<reply>, exactly; the bare reply id gives 404. A draft uses mode `update-draft` with `messageId`, in your Workflows chat. The channel must be on the flow's allowlist, and the flow refuses it after the 202." },
    ],
    examples: [
      { args: ["card", "update", REF, "./card.json"], note: "preview an edit of a posted card; nothing is sent" },
    ],
  },
  {
    key: "card templates", section: "write", flags: ["json"], profile: false,
    entries: [{ usage: "card templates [--json]", text: "List the built-in card templates. Each has a one-line purpose and its data fields. Offline." }],
    examples: [{ args: ["card", "templates"], note: "list the templates" }],
  },
  {
    key: "card example", section: "write", flags: [], profile: false,
    entries: [{ usage: "card example <template>", text: "Print example data for a template as JSON. Save it, change it, and pass it to card send --template NAME --data FILE. The examples use made-up names. Offline." }],
    examples: [{ args: ["card", "example", "poll"], note: "print the example data of the poll template" }],
  },
  {
    key: "card votes", section: "read", flags: ["json"],
    entries: [{ usage: "card votes <ref> [--json]", text: "Count the votes on a decision or poll card: the reactions of other people under each option emoji. It reads the message and writes nothing. Your own reactions are left out, because card send added them as the seed. <ref> is <conversation>#<id>, <channel>;messageid=<root>#<reply>, or draft#<id>.", more: "The templates mark each option with an id, and the CLI reads the options from the card in the message. A message without that mark lists every reaction. Reactions that are not an option are listed apart. A skin-tone variant of an option emoji counts for the option." }],
    examples: [{ args: ["card", "votes", REF], note: "count the votes on a poll card" }],
  },
  {
    key: "card setup", section: "setup", flags: ["json", "allow-write", "channel"],
    entries: [
      { usage: "card setup [--channel CHANNEL]... [--allow-write] [--json]", text: "Preview or create the \"teams-cli card router\" flow for you from the template that ships with the CLI. --channel adds a channel to the flow's allowlist and can repeat. An existing router is updated, and its allowlist is kept and added to. The preview writes nothing. It shows the environment, the connection, the modes, the allowlist, and whether a router exists. With --allow-write, the CLI creates or updates the flow and keeps its webhook URL in your profile (mode 0600). It never prints the URL.", more: "The Flow API needs a token. Set TEAMS_FLOW_TOKEN (audience https://service.flow.microsoft.com/) and TEAMS_POWERAPPS_TOKEN (audience https://service.powerapps.com/), or sign in with the Azure CLI to the same tenant as your Teams profile (az login). The flow uses your first connected Teams connection. If you have none, open Workflows in Teams once. The update-draft mode needs your Workflows chat, and is left out when the chat does not exist yet." },
    ],
    examples: [
      { args: ["card", "setup", "--channel", CHANNEL], note: "preview the router flow with one channel on the allowlist; nothing is written" },
    ],
  },
  {
    key: "scheduled", section: "write", flags: ["json"],
    entries: [{ usage: "scheduled [--json]", text: "List your scheduled messages with their draft ids, send times, and files." }],
  },
  {
    key: "scheduled cancel", section: "write", flags: ["json", "allow-write", "no-wait"],
    entries: [{ usage: "scheduled cancel <draft-id> [--json] [--allow-write] [--no-wait]", text: "Preview or cancel one scheduled message. Files it carries stay on SharePoint.", more: "An id that is not in the live list is refused." }],
  },
  {
    key: "edit", section: "write", flags: ["json", "allow-write", "no-wait", "format", "html", "emoji", "no-preview"],
    entries: [{ usage: "edit <conversation> <message-id-or-ref> <text...> [--format=rich] [--html X] [--emoji=plain] [--no-preview] [--json] [--allow-write] [--no-wait]", text: "Preview or replace your own message's text. Links work as in send." }],
  },
  {
    key: "delete", section: "delete", flags: ["json", "allow-write", "no-wait", "thread-ts"],
    entries: [{ usage: "delete <conversation> <message-id-or-ref>... [--thread-ts ROOTID] [--json] [--allow-write] [--no-wait]", text: "Preview or soft-delete your own messages. Deleting a message leaves its files on SharePoint. Run \"files delete <ref>\" first." }],
  },
  {
    key: "files delete", section: "delete", flags: ["json", "allow-write", "no-wait"],
    entries: [{ usage: "files delete <message-ref> [--json] [--allow-write] [--no-wait]", text: "Preview, or move to the SharePoint recycle bin, the files of one message that you uploaded. The preview lists each file, its size, and its uploader. A file of another person is refused. A file that is gone is reported as already gone. The Teams message is not changed. Run this before \"delete\", because a deleted message no longer lists its files.", more: "It acts only on `*.sharepoint.com` files whose `Author` email matches your UPN. A SharePoint list is refused. It never hard-deletes, and it prints each recycle bin item id." }],
    examples: [{ args: ["files", "delete", REF], note: "preview recycling a message's files" }],
  },
  {
    key: "react", section: "react", flags: ["json", "allow-write", "no-wait", "verify"],
    entries: [{ usage: "react <conversation> <message-id-or-ref> <emoji> [--verify] [--json] [--allow-write] [--no-wait]", text: "Add your reaction. <emoji> is any emoji (❤️, 🚀, 👍🏽), a catalog key (heart, launch, yes-tone3), or a raw key. Catalog keys are Teams names: 👍 is yes and 🚀 is launch. Says so if you already used it." }],
    examples: [{ args: ["react", CHANNEL, REF, "❤️"], note: "preview a reaction; add --allow-write to apply it" }],
  },
  {
    key: "unreact", section: "react", flags: ["json", "allow-write", "no-wait", "verify"],
    entries: [{ usage: "unreact <conversation> <message-id-or-ref> <emoji> [--verify] [--json] [--allow-write] [--no-wait]", text: "Remove your reaction. Says so if you did not use it." }],
  },
  {
    key: "mark", section: "react", flags: ["json", "allow-write", "no-wait", "all-unread"],
    entries: [
      { usage: "mark <conversation> [--json] [--allow-write] [--no-wait]", text: "Mark a conversation read up to its newest message." },
      { usage: "mark --all-unread [--json] [--allow-write] [--no-wait]", text: "Mark every conversation that \"unread\" lists." },
    ],
  },
  {
    key: "ledger", section: "ledger", flags: ["json"],
    entries: [
      { usage: "ledger list [--json]", text: "List send records." },
      { usage: "ledger clear <key>", text: "Remove one send record." },
    ],
  },
  {
    key: "auth login", section: "setup", flags: ["browser"],
    entries: [{ usage: "auth login [--browser PATH]", text: "Start here. Sign in once in a visible browser (Chrome, Edge, Brave, or Chromium) with its own profile. After that, tokens renew by themselves, headless.", more: "Complete MFA and choose \"Stay signed in\". On a first run at a terminal, any command that needs tokens offers this sign-in. Agents, pipes, and `--json` are never asked. See [Browser sign-in and renewal](#browser-sign-in-and-renewal)." }],
    examples: [{ args: ["auth", "login"], note: "first time on a machine" }],
  },
  {
    key: "auth status", section: "setup", flags: [],
    entries: [{ usage: "auth status", text: "Show each token's expiry, the browser profile, and test the connection." }],
  },
  {
    key: "auth renew", section: "setup", flags: ["browser"],
    entries: [{ usage: "auth renew [--browser PATH]", text: "Renew now from the browser profile, headless, with no prompt." }],
  },
  {
    key: "auth doctor", section: "setup", flags: ["json", "browser"],
    entries: [{ usage: "auth doctor [--browser PATH] [--json]", text: "Check this machine: version, platform, browser, profiles, tokens, SharePoint host, region, and clock. One line per check with ok, warn, or fail, and the fix. Exit code 1 if any check fails. Writes nothing and starts no browser.", more: "`--json` prints `{checks:[{name,status,detail,fix}]}`. Token values are never printed. The clock check sends one unauthenticated request to the Teams host and is skipped when offline." }],
    examples: [{ args: ["auth", "doctor", "--json"], note: "find why a machine cannot sign in" }],
  },
  {
    key: "auth", section: "setup", flags: [],
    entries: [{ usage: "auth < pasted-curl.txt", text: "Fallback without a browser: save tokens from cURL commands copied from the browser's network tab." }],
  },
  {
    key: "auth setup-refresh", section: "setup", flags: [],
    entries: [{ usage: "auth setup-refresh < token-request.txt", text: "Import a captured OAuth refresh request." }],
  },
  {
    key: "auth refresh", section: "setup", flags: [],
    entries: [{ usage: "auth refresh", text: "Renew stored service tokens now." }],
  },
  {
    key: "logout", section: "setup", flags: [],
    entries: [{ usage: "logout", text: "Remove a profile." }],
  },
  {
    key: "config", section: "setup", flags: [], profile: false,
    entries: [
      { usage: "config path", text: "Print the credentials file path." },
      { usage: "config file-access [link|people|private|default]", text: "Show or set who can open a file sent to a chat. Default link." },
      { usage: "config file-role [edit|view|default]", text: "Show or set what members can do with a chat file. Default edit." },
      { usage: "config skin-tone [light|medium-light|medium|medium-dark|dark|none]", text: "Show or set the skin tone for emoji that take one." },
    ],
  },
  {
    key: "update", section: "setup", flags: ["check", "force", "json"], profile: false,
    entries: [{ usage: "update [--check] [--force] [--json]", text: "Check or install a public GitHub release. --check prints versions as JSON and writes nothing. --force permits a downgrade. --json prints the result as JSON. Downloads verify SHA256SUMS before replacement. Work-portal builds remain managed by apps/wp build teams-cli." }],
    examples: [{ args: ["update", "--check"], note: "Inspect the latest public release without installing it." }],
  },
  {
    key: "version", section: "setup", flags: [], profile: false,
    entries: [{ usage: "version | --version", text: "Print the installed version." }],
  },
  {
    key: "help", section: "setup", flags: ["json"], profile: false,
    entries: [{ usage: "help [command] [--json] | <command> --help [--json]", text: "Show this summary, or one command's help. help writing shows advice for messages that need a reply. --json prints the command list for programs." }],
  },
];

export const REFERENCES: UsageEntry[] = [
  { usage: "<conversation>", text: "A conversation id (19:...@thread.v2), \"me\" (Notes to self), \"Team/Channel\", \"#name\" or a bare channel or chat name, \"@Name\" or an email (their 1:1 chat), or \"mentions\" / \"notifications\" (read only). A name that matches more than one conversation fails and lists each id." },
  { usage: "<message-ref>", text: "<conversation-id>#<message-id>, as history, context, search, and unread print it after \"id:\". A channel reply is <channel-id>;messageid=<root-id>#<reply-id>. Commands that also take <conversation> accept a bare message id." },
];

export function commandSpec(key: string): CommandSpec {
  const spec = COMMANDS.find(item => item.key === key);
  if (!spec) throw new Error(`No command spec for "${key}".`);
  return spec;
}

/** A help topic: advice that is not about one command. `teams-cli help <key>` prints it. Each point is one short paragraph. */
export type TopicSection = { heading: string; points: string[] };
export type Topic = { key: string; title: string; sections: TopicSection[] };

export const TOPICS: Topic[] = [
  {
    key: "writing",
    title: "Write a message that someone can answer",
    sections: [
      {
        heading: "Make answering easy",
        points: [
          "Pick one main purpose before you draft: FYI, decision, review, or request.",
          "FYI: say what changed, why it matters, and whether anyone must act. If no one must act, say so. Do not invent a question.",
          "Decision: name the choice, the viable options, the main tradeoffs, and your recommendation with its reason.",
          "Review: name the artifact, what to check, and what feedback completes the review.",
          "Request: state the action or information you need, and what a complete answer contains.",
          "Put the result, why it matters, and the ask before the history of how you got there.",
          "Number the questions, so the reader can answer by number.",
          "Name the person who decides. If you do not know who that is, ask who owns the decision.",
          "Give a deadline only if it is real. Never treat silence as approval.",
          "Keep the context that the reader needs to answer in the message. Link optional queries, logs, and method detail.",
          "Link evidence (a PR, a ticket, a run, a file) with a labelled link, such as [PR 1](https://example.com/pr/1), where it helps the reader check a claim.",
        ],
      },
      {
        heading: "Give the message a readable shape",
        points: [
          "For a long message, use this order: purpose and takeaway, what you need from the reader, the evidence the reader needs, then details and links.",
          "Keep a short message conversational. Do not force a report format on it.",
          "Add headings only if they help the reader find something.",
          "Use a table only to compare the same attributes across options. Use bullets for everything else.",
          "State units, baselines, and time windows. Check totals before the formatting makes them look final.",
        ],
      },
      {
        heading: "Use bold and color with care",
        points: [
          "Use bold for the short key result and the ask.",
          "Red marks a confirmed blocker or a high-consequence risk. Amber marks uncertainty, incomplete validation, or a condition that needs attention. Green marks a specific check that passed, or an issue that is resolved.",
          "For a short status label, such as a check name, a service, or a verdict, a badge is an alternative: [badge:green]PASSED[/] or [badge:red]BLOCKED[/]. Colored text still suits a sentence or a phrase.",
          "These colors are labels that you choose, not risk assessments. Always put a text label and its scope beside the color.",
          "Keep a verified claim as narrow as the check. \"Verified: unit tests passed\" must not become \"Safe to release.\"",
          "Put each limitation next to the claim that it limits, not in a footer. Separate tested behavior from assumptions.",
          "Do not color a whole paragraph.",
        ],
      },
      {
        heading: "Write it with this CLI",
        points: [
          "The default text format keeps each line as one paragraph and escapes HTML. Links work in it. For structure, add --format=rich.",
          "Links: write [label](https://url). <https://url|label> and a bare https:// URL also work, in both formats. The first link gets a preview card unless you add --no-preview.",
          "--format=rich reads *bold*, _italic_, ~strike~, <u>underline</u>, `code`, # headings, - bullets, 1. numbered items, > quotes, --- rules, pipe tables, and links ([label](https://url)). The first link gets a preview card. A fenced block with a language label becomes a code block with colors.",
          "Colors: [red]Blocked: the migration fails on staging[/] is red. [hl:yellow]Unverified: not tested on Windows[/] is amber, as a yellow highlight. [green]Verified: unit tests passed[/] is green. Short names are red, orange, yellow, lime, green, teal, blue, and purple. Write [orange]text[/] for orange text.",
          "Badges: [badge:green]PASSED[/] [badge:red]BLOCKED[/] [badge:yellow]UNVERIFIED[/] make solid labels with bold text. The text color is white or black, chosen for contrast. Do not put a badge inside a badge.",
          "Mentions: @{Name} notifies a person. Plain @Name does not. @{everyone} works only in a group or meeting chat. @{channel} works only in a channel.",
          "To reply in a channel thread, add --thread-ts ROOTID. To schedule a message, add --at TIME. --at does not work with --verify or --image.",
          "To attach files or pictures, add --file PATH or --image PATH. Pictures show inline.",
          "Teams keeps only color, background color, and font size from inline styles. Fonts, borders, padding, SVG, and data-URI images do not show. Send a chart or diagram as a picture with --image, rendered on this machine first (for example, Mermaid with a headless browser, or SVG with rsvg-convert).",
          "--html is an option, not a requirement. Use it only when you want to write the Teams HTML yourself. Do not combine it with text or --format.",
        ],
      },
      {
        heading: "Preview and send without a surprise",
        points: [
          "Read the recent messages with teams-cli history before you write to a conversation.",
          "Run the send without --allow-write first, and read the preview.",
          "A preview may read from Teams and renew tokens. It never sends.",
          "Drafting does not authorize sending. Add --allow-write only when you have approval for this message and this destination.",
          "--verify reads the sent message back and compares its text. It does not check how the message looks.",
          "Never send a test message, for example to see colors, without permission.",
          "If a send ends in an uncertain state, run teams-cli history before you try again.",
        ],
      },
      {
        heading: "Cards",
        points: [
          "A card suits a structured status or a decision. A plain message suits ordinary talk.",
          "Send a card with card send. Cards come from you through Workflows, not from a bot.",
          "Use a template when one fits. card templates lists them, and card send --template NAME --data FILE builds the card from your data.",
          "Draft first, then post. card send --draft --allow-write puts the card in your Workflows chat, so you see the real rendering. Then send it to the channel.",
          "Put the same things in a card that you would put in a message: the result, why it matters, and the ask. Do not use a card to hide a long text.",
        ],
      },
      {
        heading: "Final check",
        points: [
          "For an FYI: can the reader understand the result and why it matters without replying?",
          "For a message that needs a response: Could the recipient write a meaningful reply directly from this message, without first asking what you want?",
        ],
      },
    ],
  },
];

export function topicSpec(key: string): Topic | undefined {
  return TOPICS.find(topic => topic.key === key);
}
