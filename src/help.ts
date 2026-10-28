import { VERSION } from "./version.ts";
import { FLAG_SPECS } from "./flags.ts";
import { COMMANDS, REFERENCES, SECTION_TITLES, TOPICS, topicSpec, type CommandSpec, type SectionId, type UsageEntry } from "./spec.ts";

const INDENT = 2;
const COLUMN = 49;
const WIDTH = 120;
const NOTE = "Every command takes --profile NAME. Writes need --allow-write; without it a command only previews.";

function wrap(text: string, width: number): string[] {
  const lines: string[] = [];
  let line = "";
  for (const word of text.split(/\s+/)) {
    if (line && line.length + 1 + word.length > width) {
      lines.push(line);
      line = word;
    } else line = line ? `${line} ${word}` : word;
  }
  if (line) lines.push(line);
  return lines;
}

function renderEntry(entry: UsageEntry): string {
  const lead = " ".repeat(INDENT);
  const pad = " ".repeat(COLUMN);
  // A long usage wraps with a deeper indent, so the next line does not read as a new command.
  const usageLines = wrap(entry.usage, WIDTH - INDENT - 2).map((line, index) => (index ? `${lead}  ` : lead) + line);
  const textLines = wrap(entry.text, WIDTH - COLUMN);
  const last = usageLines[usageLines.length - 1]!;
  if (usageLines.length === 1 && last.length < COLUMN - 1 && textLines.length) {
    usageLines[usageLines.length - 1] = last.padEnd(COLUMN) + textLines.shift();
  }
  return [...usageLines, ...textLines.map(line => pad + line)].join("\n");
}

function renderSections(commands: CommandSpec[]): string[] {
  const order = Object.keys(SECTION_TITLES) as SectionId[];
  return order
    .map(section => ({ section, entries: commands.filter(spec => spec.section === section).flatMap(spec => spec.entries) }))
    .filter(part => part.entries.length)
    .map(part => `${SECTION_TITLES[part.section]}:\n${part.entries.map(renderEntry).join("\n")}`);
}

/** A shell-ready line for an example. */
export function exampleLine(args: string[]): string {
  return ["teams-cli", ...args.map(arg => (/^[\w./:=@#+-]+$/.test(arg) ? arg : `"${arg.replace(/(["\\$`])/g, "\\$1")}"`))].join(" ");
}

function renderExamples(commands: CommandSpec[]): string | undefined {
  const examples = commands.flatMap(spec => spec.examples ?? []);
  if (!examples.length) return undefined;
  return `Examples:\n${examples.map(example => `  # ${example.note}\n  ${exampleLine(example.args)}`).join("\n")}`;
}

function referencesFor(commands: CommandSpec[]): UsageEntry[] {
  const usages = commands.flatMap(spec => spec.entries.map(entry => entry.usage)).join(" ");
  return REFERENCES.filter(entry => usages.includes(entry.usage) || (entry.usage === "<message-ref>" && usages.includes("message-id-or-ref")));
}

export function fullHelp(): string {
  return [
    `teams-cli ${VERSION}: read and send Microsoft Teams messages with web-client tokens`,
    `Usage: teams-cli <command> [flags]. ${NOTE}`,
    "First time on this machine? Run: teams-cli auth login",
    ...renderSections(COMMANDS),
    `References:\n${REFERENCES.map(renderEntry).join("\n")}`,
    renderExamples(COMMANDS)!,
    "Writing a message someone must act on? Run: teams-cli help writing",
    "Run teams-cli <command> --help for one command.",
  ].join("\n\n");
}

/** The text of one help topic. Each point is a bullet, wrapped to the same width as the rest of the help. */
export function topicHelp(key: string): string | undefined {
  const topic = topicSpec(key);
  if (!topic) return undefined;
  const bullet = (point: string) => wrap(point, WIDTH - 4).map((line, index) => `${index ? "    " : "  - "}${line}`).join("\n");
  return [topic.title, ...topic.sections.map(section => `${section.heading}:\n${section.points.map(bullet).join("\n")}`)].join("\n\n");
}

/** The machine-readable form of one help topic. */
export function topicJson(key: string) {
  const topic = topicSpec(key);
  return topic && { topic: topic.key, title: topic.title, sections: topic.sections };
}

/** The commands that a command line names: `files download` alone, or both `files` commands for `files`. */
export function matchCommands(words: string[]): CommandSpec[] {
  if (!words[0]) return [];
  const pair = `${words[0]} ${words[1] ?? ""}`;
  const exact = COMMANDS.filter(spec => spec.key === pair);
  if (exact.length) return exact;
  return COMMANDS.filter(spec => spec.key === words[0] || spec.key.startsWith(`${words[0]} `));
}

/** Help for one command, or undefined if the words name no command. */
export function commandHelp(words: string[]): string | undefined {
  const commands = matchCommands(words);
  if (!commands.length) return undefined;
  const refs = referencesFor(commands);
  return [
    ...renderSections(commands),
    ...(refs.length ? [`References:\n${refs.map(renderEntry).join("\n")}`] : []),
    ...[renderExamples(commands)].filter((part): part is string => !!part),
    NOTE,
  ].join("\n\n");
}

/** The usage lines a usage error prints for one command. Forms that start with `...` add to another form, so they are left out. */
export function usageLines(key: string): string {
  const spec = COMMANDS.find(item => item.key === key);
  const entries = spec?.entries.filter(entry => !entry.usage.startsWith(`${key.split(" ")[0]} ...`)) ?? [];
  return entries.map(entry => `Usage: teams-cli ${entry.usage}`).join("\n");
}

/** The machine-readable help: every command, its flags, usage forms, and examples. */
export function helpJson(words: string[] = []) {
  const commands = words.length ? matchCommands(words) : COMMANDS;
  return {
    version: VERSION,
    notes: [NOTE, "First time on this machine? Run: teams-cli auth login"],
    commands: commands.map(spec => ({
      command: spec.key,
      section: SECTION_TITLES[spec.section],
      flags: [...spec.flags, ...(spec.profile === false ? [] : ["profile" as const])].map(name => ({ name: `--${name}`, takes_value: FLAG_SPECS[name] })),
      usage: spec.entries.map(entry => ({ usage: `teams-cli ${entry.usage}`, text: entry.text })),
      examples: (spec.examples ?? []).map(example => ({ command: exampleLine(example.args), note: example.note })),
    })),
    ...(words.length ? {} : { topics: TOPICS.map(topic => ({ topic: topic.key, title: topic.title })) }),
    references: (words.length ? referencesFor(commands) : REFERENCES).map(entry => ({ name: entry.usage, text: entry.text })),
  };
}

/** The README command table, between its markers. `bun run docs` writes it. */
export function readmeTable(): string {
  const cell = (text: string) => text.replace(/\|/g, "\\|");
  const rows = COMMANDS.flatMap(spec => spec.entries.map(entry =>
    `| \`${cell(entry.usage)}\` | ${cell(entry.text)}${entry.more ? ` ${cell(entry.more)}` : ""} |`));
  return ["| Command | Purpose |", "| --- | --- |", ...rows].join("\n");
}
