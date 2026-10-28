import { MAX_TEAMS_CARD_BYTES } from "../limits.ts";
import { JIRA_ICONS } from "./icons.ts";
import {
  adaptiveCard, badge, box, bullets, column, header, noted, openUrl, plural, progressColor, row, rowColumns, rows, text, toggleLink, type El,
} from "./kit.ts";
import {
  FieldError, bool, int, link, list, object, oneOf, optInt, optStr, optUrl, reactionEmoji, status, str, strings, style, url,
  type Link, type StatusBadge,
} from "./validate.ts";

/**
 * A template turns validated data into an Adaptive Card. `parse` checks the data and names the field that is wrong.
 * `build` takes parsed data and never reads anything else. `votes` lists the emoji that people react with, for a
 * template that asks for a vote. The sample data in `example` is neutral: no real names, hosts, or tickets.
 */
export type Template<D = any> = {
  name: string;
  purpose: string;
  /** The data fields on one line. A ? after a name marks an optional field. */
  fields: string;
  example: unknown;
  parse(raw: unknown): D;
  build(data: D): El;
  votes?(data: D): string[];
};

const KEYCAPS = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"];
const MAX_OPTIONS = KEYCAPS.length;

/** The Jira type names that have an icon in the CLI. */
export const JIRA_TYPES = Object.keys(JIRA_ICONS).sort();

/** `Production Bug`, `production-bug`, and `production_bug` all name one icon. */
function jiraIcon(value: unknown, path: string): { type: string; icon: string } {
  const name = str(value, path);
  const type = name.toLowerCase().replace(/[\s_]+/g, "-").replace(/^subtask$/, "sub-task");
  const icon = JIRA_ICONS[type];
  if (!icon) throw new FieldError(path, `"${name}" has no icon. Use one of: ${JIRA_TYPES.join(", ")}.`);
  return { type, icon };
}

function optionEmoji(value: unknown, path: string, index: number): string {
  return value === undefined ? KEYCAPS[index]! : reactionEmoji(value, path);
}

/** Two options must not share an emoji, or a vote cannot tell them apart. */
function distinct(emojis: string[], path: string): void {
  const seen = new Set<string>();
  emojis.forEach((emoji, index) => {
    const key = emoji.replaceAll("️", "");
    if (seen.has(key)) throw new FieldError(`${path}[${index}].emoji`, `"${emoji}" is used twice. Each option needs its own emoji.`);
    seen.add(key);
  });
}

// ---------------------------------------------------------------------------------------------------------------------
// meeting-recap

type Recap = {
  title: string; when: string; participants: Array<{ name: string; upn: string }>; decisions: string[];
  actions: Array<{ text: string; owner?: string; status?: StatusBadge }>; notes: string[]; link?: Link;
};

const meetingRecap: Template<Recap> = {
  name: "meeting-recap",
  purpose: "Recap of a meeting: decisions, action items with owners and status, and collapsed participants and notes.",
  fields: "title, when, participants?[{name, upn}], decisions?[], actions?[{text, owner?, status?{text, style}}], notes?[], link?",
  example: {
    title: "Weekly sync: Example project",
    when: "Mon Oct 19 · 2:30 PM",
    participants: [{ name: "Alex", upn: "alex@example.com" }, { name: "Sam", upn: "sam@example.com" }, { name: "Jordan", upn: "jordan@example.com" }],
    decisions: ["Ship the first release on Friday", "Anything we charge for is a feature flag"],
    actions: [
      { text: "Send the draft spec to the team", owner: "Alex", status: { text: "Due soon", style: "Warning" } },
      { text: "Confirm the launch date with the customer", owner: "Sam", status: { text: "Open", style: "Informative" } },
      { text: "Decide on the export format", owner: "Jordan", status: { text: "Blocked", style: "Attention" } },
    ],
    notes: ["Launch date depends on the customer", "The export format needs a second meeting"],
    link: { title: "Open notes", url: "https://notes.example.com/recap/123" },
  },
  parse(raw) {
    const rec = object(raw, "data", ["title", "when", "participants", "decisions", "actions", "notes", "link"]);
    const participants = list(rec.participants, "participants", (item, path) => {
      const person = object(item, path, ["name", "upn"]);
      const name = str(person.name, `${path}.name`);
      if (/[<>&]/.test(name)) throw new FieldError(`${path}.name`, 'must not contain "<", ">", or "&". It goes inside an <at> mention.');
      const upn = str(person.upn, `${path}.upn`);
      if (!upn.includes("@")) throw new FieldError(`${path}.upn`, `must be a user principal name such as alex@example.com, not "${upn}".`);
      return { name, upn };
    });
    const names = new Set<string>();
    participants.forEach((person, index) => {
      if (names.has(person.name)) throw new FieldError(`participants[${index}].name`, `"${person.name}" is used twice. Each mention needs its own name.`);
      names.add(person.name);
    });
    return {
      title: str(rec.title, "title"), when: str(rec.when, "when"), participants,
      decisions: strings(rec.decisions, "decisions"),
      actions: list(rec.actions, "actions", (item, path) => {
        const action = object(item, path, ["text", "owner", "status"]);
        return { text: str(action.text, `${path}.text`), owner: optStr(action.owner, `${path}.owner`), status: action.status === undefined ? undefined : status(action.status, `${path}.status`) };
      }),
      notes: strings(rec.notes, "notes"),
      link: link(rec.link, "link", "Open notes"),
    };
  },
  build(data) {
    const body: El[] = [header({
      title: data.title, icon: "PeopleTeam", badge: { text: "Recap", style: "Accent" },
      subtitle: [data.when, data.participants.length ? plural(data.participants.length, "person", "people") : ""].filter(Boolean).join(" · "),
    })];
    if (data.participants.length) {
      body.push({
        type: "Container", items: [
          toggleLink(`▸ Show participants (${data.participants.length})`, ["participants"]),
          { type: "Container", id: "participants", isVisible: false, items: [text(data.participants.map(person => `<at>${person.name}</at>`).join("  "))] },
        ],
      });
    }
    if (data.decisions.length) body.push(box(bullets(data.decisions), "Decisions", "good"));
    if (data.actions.length) {
      body.push(box(rows(data.actions.map(action => ({
        title: action.text, subline: action.owner, right: action.status ? badge(action.status.text, action.status.style) : undefined,
      }))), "Action items"));
    }
    if (data.notes.length) {
      body.push(
        toggleLink("▸ Show full notes", ["notes"]),
        { type: "Container", id: "notes", isVisible: false, items: bullets(data.notes) },
      );
    }
    return adaptiveCard(body, {
      ...(data.link ? { actions: [openUrl(data.link.title, data.link.url)] } : {}),
      ...(data.participants.length ? {
        msteams: { entities: data.participants.map(person => ({ type: "mention", text: `<at>${person.name}</at>`, mentioned: { id: person.upn, name: person.name } })) },
      } : {}),
    });
  },
};

// ---------------------------------------------------------------------------------------------------------------------
// deadline

const DEADLINE_STATUS = {
  "on-track": { text: "On track", style: "Good" },
  "at-risk": { text: "At risk", style: "Warning" },
  behind: { text: "Behind", style: "Attention" },
} as const;

type Deadline = {
  title: string; subtitle?: string; status: keyof typeof DEADLINE_STATUS; count: string; countLabel: string;
  rows: Array<{ text: string; badge: StatusBadge & { icon?: string } }>; link?: Link;
};

const deadline: Template<Deadline> = {
  name: "deadline",
  purpose: "A deadline: a status badge, a large count with its label, and rows that each have a badge.",
  fields: "title, subtitle?, status (on-track|at-risk|behind), count, countLabel, rows[{text, badge{text, style, icon?}}], link?",
  example: {
    title: "Example launch", subtitle: "Target Fri Oct 30", status: "at-risk", count: "12", countLabel: "days until the target",
    rows: [
      { text: "Draft spec from Alex", badge: { text: "Waiting", style: "Warning" } },
      { text: "Integration check", badge: { text: "Done", style: "Good", icon: "CheckmarkCircle" } },
      { text: "Test plan", badge: { text: "To do", style: "Subtle" } },
      { text: "Ready for testing", badge: { text: "Oct 30", style: "Accent" } },
    ],
    link: { title: "Open the project", url: "https://example.atlassian.net/browse/EX-100" },
  },
  parse(raw) {
    const rec = object(raw, "data", ["title", "subtitle", "status", "count", "countLabel", "rows", "link"]);
    const count = typeof rec.count === "number" && Number.isFinite(rec.count) ? String(rec.count) : str(rec.count, "count");
    return {
      title: str(rec.title, "title"), subtitle: optStr(rec.subtitle, "subtitle"),
      status: oneOf(rec.status, "status", Object.keys(DEADLINE_STATUS) as Array<keyof typeof DEADLINE_STATUS>),
      count, countLabel: str(rec.countLabel, "countLabel"),
      rows: list(rec.rows, "rows", (item, path) => {
        const entry = object(item, path, ["text", "badge"]);
        const badgeRec = object(entry.badge, `${path}.badge`, ["text", "style", "icon"]);
        return {
          text: str(entry.text, `${path}.text`),
          badge: { text: str(badgeRec.text, `${path}.badge.text`), style: style(badgeRec.style, `${path}.badge.style`), icon: optStr(badgeRec.icon, `${path}.badge.icon`) },
        };
      }, { min: 1 }),
      link: link(rec.link, "link", "Open"),
    };
  },
  build(data) {
    const state = DEADLINE_STATUS[data.status];
    return adaptiveCard([
      header({ title: data.title, subtitle: data.subtitle, icon: "CalendarClock", badge: state }),
      {
        type: "ColumnSet", spacing: "Medium", columns: [
          column([text(data.count, { size: "ExtraLarge", weight: "Bolder", color: state.style })]),
          column([text(data.countLabel, { isSubtle: true })], "stretch", { spacing: "Small" }),
        ],
      },
      box(rows(data.rows.map(entry => ({ title: entry.text, right: badge(entry.badge.text, entry.badge.style, "Tint", entry.badge.icon) })))),
    ], data.link ? { actions: [openUrl(data.link.title, data.link.url)] } : {});
  },
};

// ---------------------------------------------------------------------------------------------------------------------
// epic

/** The chart colors of a stacked bar, in status order. More statuses than colors start again at the first. */
export const EPIC_COLORS = ["neutral", "categoricalBlue", "categoricalPurple", "categoricalGreen"] as const;
export const EPIC_VISIBLE_ISSUES = 3;

type Epic = { key: string; title: string; url?: string; statuses: Array<{ name: string; issues: Array<{ key: string; summary: string }> }> };

/** The Jira browse base of an epic URL: its origin and `/browse/`. Without a URL, issue keys are not links. */
function browseBase(epicUrl: string | undefined): string | undefined {
  return epicUrl ? `${new URL(epicUrl).origin}/browse/` : undefined;
}

/**
 * Builds the epic card. `hiddenKept[i]` is how many of status i's hidden issues stay in the card. The issues that
 * do not stay are replaced by one "+N more in Jira" line, which links to the epic.
 */
function epicCard(data: Epic, hiddenKept: number[]): El {
  const base = browseBase(data.url);
  const total = data.statuses.reduce((sum, entry) => sum + entry.issues.length, 0);
  const line = (issue: { key: string; summary: string }, first: boolean): El =>
    text(`${base ? `[${issue.key}](${base}${issue.key})` : `**${issue.key}**`}  ${issue.summary}`, { spacing: "Small", separator: !first });
  const body: El[] = [
    header({ title: `${data.key} ${data.title}`, subtitle: `Epic · ${plural(total, "issue")}`, image: JIRA_ICONS.epic }),
    {
      type: "Chart.HorizontalBar.Stacked", title: "Status",
      data: [{ title: data.key, data: data.statuses.map((entry, index) => ({ legend: entry.name, value: entry.issues.length, color: EPIC_COLORS[index % EPIC_COLORS.length] })) }],
    },
  ];
  data.statuses.forEach((entry, index) => {
    if (!entry.issues.length) return;
    const shown = entry.issues.slice(0, EPIC_VISIBLE_ISSUES);
    const hidden = entry.issues.slice(EPIC_VISIBLE_ISSUES, EPIC_VISIBLE_ISSUES + hiddenKept[index]!);
    const cut = entry.issues.length - EPIC_VISIBLE_ISSUES - hidden.length;
    const items: El[] = [text(`${entry.name} (${entry.issues.length})`, { weight: "Bolder" }), ...shown.map((issue, at) => line(issue, at === 0))];
    const more = cut > 0 ? text(data.url ? `[+${cut} more in Jira](${data.url})` : `+${cut} more in Jira`, { spacing: "Small", separator: true }) : undefined;
    if (hidden.length) {
      const targets = [`rest-${index}`, `more-${index}`, `hide-${index}`];
      items.push(
        { type: "Container", id: targets[0], isVisible: false, items: [...hidden.map(issue => line(issue, false)), ...(more ? [more] : [])] },
        toggleLink(`Show ${hidden.length} more`, targets, { id: targets[1] }),
        toggleLink("Hide", targets, { id: targets[2], isVisible: false }),
      );
    } else if (more) items.push(more);
    body.push(box(items));
  });
  return adaptiveCard(body, data.url ? { actions: [openUrl("Open the epic", data.url)] } : {});
}

const cardBytes = (card: unknown) => Buffer.byteLength(JSON.stringify(card));

/**
 * The full card when it fits. When it does not, issues leave the hidden parts, one at a time from the status that
 * hides the most. The first issues of each status always stay. It fails when even those do not fit.
 */
export function buildEpic(data: Epic, limit = MAX_TEAMS_CARD_BYTES): { card: El; cut: number } {
  const kept = data.statuses.map(entry => Math.max(0, entry.issues.length - EPIC_VISIBLE_ISSUES));
  // No issue line is longer than this, so dropping excess / longest lines can never drop too many.
  const base = browseBase(data.url) ?? "";
  const longest = Math.max(1, ...data.statuses.flatMap(entry => entry.issues.map(issue => cardBytes(text(`[${issue.key}](${base}${issue.key})  ${issue.summary}`, { spacing: "Small", separator: true })))));
  let card = epicCard(data, kept);
  let cut = 0;
  for (;;) {
    const excess = cardBytes(card) - limit;
    if (excess <= 0) return { card, cut };
    for (let drop = Math.max(1, Math.floor(excess / longest)); drop > 0; drop--) {
      let at = -1;
      kept.forEach((count, index) => { if (count > 0 && (at < 0 || count >= kept[at]!)) at = index; });
      if (at < 0) {
        throw new Error(`The epic card is ${cardBytes(card)} bytes even with only the first ${EPIC_VISIBLE_ISSUES} issues of each status. Teams refuses a card over ${limit} bytes (${limit / 1024} KB). Use fewer statuses or shorter summaries.`);
      }
      kept[at]!--;
      cut++;
    }
    card = epicCard(data, kept);
  }
}

const epic: Template<Epic> = {
  name: "epic",
  purpose: "A Jira epic: a stacked status chart, then a box per status with its first 3 issues and a Show more link.",
  fields: "key, title, url?, statuses[{name, issues[{key, summary}]}]",
  example: {
    key: "EX-100", title: "Example epic", url: "https://example.atlassian.net/browse/EX-100",
    statuses: [
      { name: "To Do", issues: [{ key: "EX-111", summary: "Example story one" }, { key: "EX-112", summary: "Example story two" }, { key: "EX-113", summary: "Example story three" }, { key: "EX-114", summary: "Example story four" }] },
      { name: "In Progress", issues: [{ key: "EX-121", summary: "Example story in progress" }, { key: "EX-122", summary: "Another example story" }] },
      { name: "Done", issues: [{ key: "EX-131", summary: "Example story that is done" }] },
    ],
  },
  parse(raw) {
    const rec = object(raw, "data", ["key", "title", "url", "statuses"]);
    return {
      key: str(rec.key, "key"), title: str(rec.title, "title"), url: optUrl(rec.url, "url"),
      statuses: list(rec.statuses, "statuses", (item, path) => {
        const entry = object(item, path, ["name", "issues"]);
        return {
          name: str(entry.name, `${path}.name`),
          issues: list(entry.issues, `${path}.issues`, (issue, issuePath) => {
            const fields = object(issue, issuePath, ["key", "summary"]);
            return { key: str(fields.key, `${issuePath}.key`), summary: str(fields.summary, `${issuePath}.summary`) };
          }),
        };
      }, { min: 1 }),
    };
  },
  build: data => buildEpic(data).card,
};

// ---------------------------------------------------------------------------------------------------------------------
// pr-review

type PrReview = {
  title: string; subtitle?: string; ticket?: { key: string; type: string; icon: string }; adds?: number; dels?: number; files?: number;
  checks: string[]; notTested: string[]; ask: string; links: { pr: string; diff?: string; ticket?: string; pipeline?: string; testPlan?: string };
};

const prReview: Template<PrReview> = {
  name: "pr-review",
  purpose: "A review request: what to check, what is not tested, the size of the change, and links in an overflow menu.",
  fields: "title, subtitle?, ticket?{key, type}, adds?, dels?, files?, checks?[], notTested?[], ask, links{pr, diff?, ticket?, pipeline?, testPlan?}",
  example: {
    title: "Review request: PR #123", subtitle: "Example story title", ticket: { key: "PE-1234", type: "Story" }, adds: 214, dels: 37, files: 6,
    checks: ["The dry run names every row it would change", "Rows that already have a value stay as they are"],
    notTested: ["Production-size data. Tested on staging only."],
    ask: "Reply in this thread with **approve** or your questions.",
    links: {
      pr: "https://example.com/repo/pullrequest/123", diff: "https://example.com/repo/pullrequest/123/files",
      ticket: "https://example.atlassian.net/browse/PE-1234", pipeline: "https://example.com/pipelines/456", testPlan: "https://example.com/test-plans/7",
    },
  },
  parse(raw) {
    const rec = object(raw, "data", ["title", "subtitle", "ticket", "adds", "dels", "files", "checks", "notTested", "ask", "links"]);
    let ticket: PrReview["ticket"];
    if (rec.ticket !== undefined) {
      const entry = object(rec.ticket, "ticket", ["key", "type"]);
      const found = jiraIcon(entry.type, "ticket.type");
      ticket = { key: str(entry.key, "ticket.key"), type: found.type, icon: found.icon };
    }
    const links = object(rec.links, "links", ["pr", "diff", "ticket", "pipeline", "testPlan"]);
    return {
      title: str(rec.title, "title"), subtitle: optStr(rec.subtitle, "subtitle"), ticket,
      adds: optInt(rec.adds, "adds", 0), dels: optInt(rec.dels, "dels", 0), files: optInt(rec.files, "files", 0),
      checks: strings(rec.checks, "checks"), notTested: strings(rec.notTested, "notTested"), ask: str(rec.ask, "ask"),
      links: {
        pr: url(links.pr, "links.pr"), diff: optUrl(links.diff, "links.diff"), ticket: optUrl(links.ticket, "links.ticket"),
        pipeline: optUrl(links.pipeline, "links.pipeline"), testPlan: optUrl(links.testPlan, "links.testPlan"),
      },
    };
  },
  build(data) {
    const body: El[] = [header({
      title: data.title, subtitle: data.subtitle ? [data.ticket?.key, data.subtitle].filter(Boolean).join(" · ") : data.ticket?.key,
      ...(data.ticket ? { image: data.ticket.icon } : { icon: "BranchCompare" }),
    })];
    const size: El[] = [];
    if (data.adds !== undefined) size.push(column([badge(`+${data.adds}`, "Good", "Filled")]));
    if (data.dels !== undefined) size.push(column([badge(`−${data.dels}`, "Attention", "Filled")], "auto", size.length ? { spacing: "Small" } : {}));
    if (data.files !== undefined) size.push(column([badge(plural(data.files, "file"), "Subtle")], "auto", size.length ? { spacing: "Small" } : {}));
    if (size.length) body.push({ type: "ColumnSet", columns: size });
    if (data.checks.length) body.push(box(bullets(data.checks), "What to check"));
    if (data.notTested.length) body.push(box(noted(data.notTested, "Warning"), "Not tested", "warning"));
    body.push(text(data.ask, { spacing: "Medium" }));
    const secondary = { mode: "secondary" };
    // An ActionSet in the body puts its overflow menu next to Open PR. Card-level actions put it in the card header.
    body.push({ type: "ActionSet", spacing: "Medium", actions: [
      openUrl("Open PR", data.links.pr),
      ...(data.links.diff ? [openUrl("View diff", data.links.diff, secondary)] : []),
      ...(data.links.ticket ? [openUrl(data.ticket ? `Open ticket ${data.ticket.key}` : "Open ticket", data.links.ticket, secondary)] : []),
      ...(data.links.pipeline ? [openUrl("Open the pipeline run", data.links.pipeline, secondary)] : []),
      ...(data.links.testPlan ? [openUrl("Open the test plan", data.links.testPlan, secondary)] : []),
    ] });
    return adaptiveCard(body);
  },
};

// ---------------------------------------------------------------------------------------------------------------------
// decision

type Decision = {
  title: string; decider: string; risk?: StatusBadge; recommendation?: string; ask?: string;
  options: Array<{ emoji: string; label: string; effect: string; risk: StatusBadge }>;
};

/** "1️⃣, 2️⃣, or 3️⃣" */
const orList = (items: string[]) => (items.length < 3 ? items.join(" or ") : `${items.slice(0, -1).join(", ")}, or ${items[items.length - 1]}`);

const decision: Template<Decision> = {
  name: "decision",
  purpose: "A decision for one person: an options table with a risk badge per option, a recommendation, and a vote by reaction.",
  fields: "title, decider, risk?{text, style}, options[{emoji?, label, effect, risk{text, style}}], recommendation?, ask?",
  example: {
    title: "Decision needed: example download access", decider: "Alex",
    risk: { text: "Privacy risk", style: "Attention" },
    options: [
      { label: "Admins only", effect: "Downloads limited to admins", risk: { text: "Low", style: "Good" } },
      { label: "Audited", effect: "Everyone, with an audit log", risk: { text: "Medium", style: "Warning" } },
      { label: "As is", effect: "No change", risk: { text: "High", style: "Attention" } },
    ],
    recommendation: "**Option 1**, because it closes the risk now and option 2 can follow.",
    ask: "Please vote before Friday.",
  },
  parse(raw) {
    const rec = object(raw, "data", ["title", "decider", "risk", "options", "recommendation", "ask"]);
    const options = list(rec.options, "options", (item, path) => {
      const option = object(item, path, ["emoji", "label", "effect", "risk"]);
      return { emoji: option.emoji, label: str(option.label, `${path}.label`), effect: str(option.effect, `${path}.effect`), risk: status(option.risk, `${path}.risk`) };
    }, { min: 2, max: MAX_OPTIONS });
    const withEmoji = options.map((option, index) => ({ ...option, emoji: optionEmoji(option.emoji, `options[${index}].emoji`, index) }));
    distinct(withEmoji.map(option => option.emoji), "options");
    return {
      title: str(rec.title, "title"), decider: str(rec.decider, "decider"), risk: rec.risk === undefined ? undefined : status(rec.risk, "risk"),
      options: withEmoji, recommendation: optStr(rec.recommendation, "recommendation"), ask: optStr(rec.ask, "ask"),
    };
  },
  votes: data => data.options.map(option => option.emoji),
  build(data) {
    const cell = (items: El[]): El => ({ type: "TableCell", items });
    const body: El[] = [
      header({ title: data.title, subtitle: `Decider: ${data.decider}`, icon: "QuestionCircle", badge: data.risk }),
      {
        type: "Table", firstRowAsHeaders: true, showGridLines: true, columns: [{ width: 2 }, { width: 3 }, { width: 2 }],
        rows: [
          { type: "TableRow", cells: ["Option", "Effect", "Risk"].map(label => cell([text(label, { weight: "Bolder" })])) },
          ...data.options.map((option, index) => ({
            type: "TableRow",
            cells: [cell([text(`${option.emoji} ${option.label}`, { id: `vote-${index + 1}` })]), cell([text(option.effect)]), cell([badge(option.risk.text, option.risk.style)])],
          })),
        ],
      },
    ];
    if (data.recommendation) body.push(box([text(data.recommendation)], "Recommendation", "good"));
    if (data.ask) body.push(text(data.ask, { spacing: "Medium" }));
    body.push(text(`**React ${orList(data.options.map(option => option.emoji))} to vote.**`, { spacing: "Medium" }));
    return adaptiveCard(body);
  },
};

// ---------------------------------------------------------------------------------------------------------------------
// poll

type Poll = { question: string; context?: string; options: Array<{ emoji: string; label: string; description?: string }> };

const poll: Template<Poll> = {
  name: "poll",
  purpose: "A question with options. People vote by reacting with the option emoji. There are no buttons.",
  fields: "question, context?, options[{emoji, label, description?}]",
  example: {
    question: "Who takes the release this week?", context: "Asked in the example channel",
    options: [
      { emoji: "🙋", label: "I'll take it", description: "You own the deploy calendar this week" },
      { emoji: "🤝", label: "I can back up", description: "Second on call" },
      { emoji: "⛔", label: "I have a conflict", description: "Time off or a meeting clash" },
    ],
  },
  parse(raw) {
    const rec = object(raw, "data", ["question", "context", "options"]);
    const options = list(rec.options, "options", (item, path) => {
      const option = object(item, path, ["emoji", "label", "description"]);
      return { emoji: reactionEmoji(option.emoji, `${path}.emoji`), label: str(option.label, `${path}.label`), description: optStr(option.description, `${path}.description`) };
    }, { min: 2, max: MAX_OPTIONS });
    distinct(options.map(option => option.emoji), "options");
    return { question: str(rec.question, "question"), context: optStr(rec.context, "context"), options };
  },
  votes: data => data.options.map(option => option.emoji),
  build(data) {
    const optionRows = data.options.map((option, index) => {
      const entry = row({ title: `${option.emoji}  ${option.label}`, subline: option.description, separator: index > 0 });
      // The vote reader finds each option by this id.
      entry.items[0].columns[0].items[0].id = `vote-${index + 1}`;
      return entry;
    });
    return adaptiveCard([
      header({ title: data.question, subtitle: data.context, icon: "PeopleQueue" }),
      box([text("**Vote by reacting to this message:**"), ...optionRows], undefined, "default"),
    ]);
  },
};

// ---------------------------------------------------------------------------------------------------------------------
// progress

type Progress = { title: string; subtitle?: string; items: Array<{ name: string; owner?: string; percent: number; status: StatusBadge }> };

const progress: Template<Progress> = {
  name: "progress",
  purpose: "Progress of several items: a status badge and a progress bar in the status color for each.",
  fields: "title, subtitle?, items[{name, owner?, percent, status{text, style}}]",
  example: {
    title: "Progress: example goals", subtitle: "From the weekly check-in",
    items: [
      { name: "First goal", owner: "Alex · part released", percent: 70, status: { text: "On track", style: "Good" } },
      { name: "Second goal", owner: "Sam · missed the first date", percent: 45, status: { text: "Slipped", style: "Attention" } },
      { name: "Third goal", owner: "Jordan · due Nov 27", percent: 25, status: { text: "Needs a meeting", style: "Warning" } },
    ],
  },
  parse(raw) {
    const rec = object(raw, "data", ["title", "subtitle", "items"]);
    return {
      title: str(rec.title, "title"), subtitle: optStr(rec.subtitle, "subtitle"),
      items: list(rec.items, "items", (item, path) => {
        const entry = object(item, path, ["name", "owner", "percent", "status"]);
        return { name: str(entry.name, `${path}.name`), owner: optStr(entry.owner, `${path}.owner`), percent: int(entry.percent, `${path}.percent`, 0, 100), status: status(entry.status, `${path}.status`) };
      }, { min: 1 }),
    };
  },
  build(data) {
    return adaptiveCard([
      header({ title: data.title, subtitle: data.subtitle, icon: "Target" }),
      ...data.items.map((item): El => ({
        type: "Container", separator: true, items: [
          rowColumns({ title: item.name, subline: item.owner, right: badge(item.status.text, item.status.style) }),
          { type: "ProgressBar", value: item.percent, max: 100, color: progressColor(item.status.style) },
        ],
      })),
    ]);
  },
};

// ---------------------------------------------------------------------------------------------------------------------
// incident

const TIMELINE = {
  error: { name: "ErrorCircle", color: "Attention" },
  info: { name: "Info", color: "Accent" },
  investigating: { name: "Search", color: "Warning" },
  resolved: { name: "CheckmarkCircle", color: "Good" },
} as const;

type Incident = {
  title: string; opened: string; severity: string; resolved: boolean; impact: string; owner: string; nextUpdate?: string;
  timeline: Array<{ time: string; text: string; kind: keyof typeof TIMELINE }>;
};

const incident: Template<Incident> = {
  name: "incident",
  purpose: "An incident: severity, impact, owner, next update, and a timeline. Send it, then card update it as the timeline grows.",
  fields: "title, opened, severity, resolved?, impact, owner, nextUpdate?, timeline[{time, text, kind (error|info|investigating|resolved)}]",
  example: {
    title: "Incident: example service is down", opened: "4:10 PM by Alex", severity: "SEV 2", impact: "Example uploads are failing", owner: "Alex", nextUpdate: "5:00 PM",
    timeline: [
      { time: "4:10 PM", text: "Uploads started to fail", kind: "error" },
      { time: "4:24 PM", text: "Sam: no deploy ran today", kind: "info" },
      { time: "4:29 PM", text: "Jordan: looking at the storage settings", kind: "investigating" },
    ],
  },
  parse(raw) {
    const rec = object(raw, "data", ["title", "opened", "severity", "resolved", "impact", "owner", "nextUpdate", "timeline"]);
    return {
      title: str(rec.title, "title"), opened: str(rec.opened, "opened"), severity: str(rec.severity, "severity"), resolved: bool(rec.resolved, "resolved"),
      impact: str(rec.impact, "impact"), owner: str(rec.owner, "owner"), nextUpdate: optStr(rec.nextUpdate, "nextUpdate"),
      timeline: list(rec.timeline, "timeline", (item, path) => {
        const entry = object(item, path, ["time", "text", "kind"]);
        return { time: str(entry.time, `${path}.time`), text: str(entry.text, `${path}.text`), kind: oneOf(entry.kind, `${path}.kind`, Object.keys(TIMELINE) as Array<keyof typeof TIMELINE>) };
      }, { min: 1 }),
    };
  },
  build(data) {
    return adaptiveCard([
      header({ title: data.title, subtitle: `Opened ${data.opened}`, icon: "Alert", badge: data.resolved ? { text: "Resolved", style: "Good" } : { text: data.severity, style: "Attention" } }),
      {
        type: "FactSet", facts: [
          { title: "Impact", value: data.impact },
          { title: "Owner", value: data.owner },
          ...(data.nextUpdate || data.resolved ? [{ title: "Next update", value: data.nextUpdate ?? "Closed" }] : []),
        ],
      },
      box(rows(data.timeline.map(entry => ({ title: entry.time, subline: entry.text, icon: TIMELINE[entry.kind] }))), "Timeline"),
    ]);
  },
};

// ---------------------------------------------------------------------------------------------------------------------
// standup

type Standup = { name: string; date: string; yesterday: string[]; today: string[]; blockers: string[] };

const standup: Template<Standup> = {
  name: "standup",
  purpose: "A stand-up note: yesterday, today, and blockers in a warning box.",
  fields: "name, date, yesterday?[], today?[], blockers?[]",
  example: {
    name: "Alex", date: "Mon Oct 19",
    yesterday: ["Finished the first three example stories", "Reviewed two pull requests"],
    today: ["Deploy the example epic to QA", "Start the next story once the spec arrives"],
    blockers: ["Waiting on the spec from Sam"],
  },
  parse(raw) {
    const rec = object(raw, "data", ["name", "date", "yesterday", "today", "blockers"]);
    return { name: str(rec.name, "name"), date: str(rec.date, "date"), yesterday: strings(rec.yesterday, "yesterday"), today: strings(rec.today, "today"), blockers: strings(rec.blockers, "blockers") };
  },
  build(data) {
    const lines = (items: string[], extra: El = {}) => (items.length ? bullets(items, extra) : [text("None", { isSubtle: true, spacing: "Small" })]);
    return adaptiveCard([
      header({ title: `Stand-up: ${data.name}`, subtitle: data.date, icon: "PersonClock" }),
      box(lines(data.yesterday), "Yesterday", "default"),
      box(lines(data.today), "Today", "default"),
      box(data.blockers.length ? noted(data.blockers, "Warning") : lines([]), "Blockers", data.blockers.length ? "warning" : "default"),
    ]);
  },
};

// ---------------------------------------------------------------------------------------------------------------------
// qa-checklist

type QaChecklist = { title: string; intro?: string; checks: Array<{ title: string; ticket?: string; steps: string[]; expected: string }> };

const qaChecklist: Template<QaChecklist> = {
  name: "qa-checklist",
  purpose: "A QA checklist: a checkbox per check, with numbered steps and the expected result shown under it.",
  fields: "title, intro?, checks[{title, ticket?, steps[], expected}]",
  example: {
    title: "QA pass: example epic", intro: "Run each check on the QA environment.",
    checks: [
      { title: "Totals round to cents", ticket: "PE-1234", steps: ["Open an example claim with three lines", "Save the claim", "Open the claim total"], expected: "The total is the sum of the rounded lines." },
      { title: "Exports match", ticket: "PE-1235", steps: ["Run the first export for one week", "Run the second export for the same week"], expected: "Both exports show the same hours." },
    ],
  },
  parse(raw) {
    const rec = object(raw, "data", ["title", "intro", "checks"]);
    return {
      title: str(rec.title, "title"), intro: optStr(rec.intro, "intro"),
      checks: list(rec.checks, "checks", (item, path) => {
        const check = object(item, path, ["title", "ticket", "steps", "expected"]);
        return { title: str(check.title, `${path}.title`), ticket: optStr(check.ticket, `${path}.ticket`), steps: strings(check.steps, `${path}.steps`, { min: 1 }), expected: str(check.expected, `${path}.expected`) };
      }, { min: 1 }),
    };
  },
  build(data) {
    const body: El[] = [header({ title: data.title, subtitle: "Tick each check as you go. Ticks stay on your screen only.", icon: "TaskListLtr" })];
    if (data.intro) body.push(text(data.intro, { spacing: "Medium" }));
    data.checks.forEach((check, index) => {
      const id = `steps-${index + 1}`;
      body.push({
        type: "Container", separator: true, spacing: "Medium", items: [
          {
            type: "ColumnSet", columns: [
              column([{ type: "Input.Toggle", id: `check-${index + 1}`, title: "", value: "false" }]),
              column([text(check.title, { weight: "Bolder" }), ...(check.ticket ? [text(check.ticket, { isSubtle: true, size: "Small", spacing: "None" })] : [])], "stretch"),
            ],
          },
          toggleLink("▾ How to test", [id]),
          {
            type: "Container", id, isVisible: true, spacing: "Small",
            items: [...check.steps.map((step, at) => text(`${at + 1}. ${step}`, { spacing: "Small" })), text(`**Expected:** ${check.expected}`, { spacing: "Small" })],
          },
        ],
      });
    });
    return adaptiveCard(body);
  },
};

// ---------------------------------------------------------------------------------------------------------------------

export const TEMPLATES: Template[] = [meetingRecap, deadline, epic, prReview, decision, poll, progress, incident, standup, qaChecklist];

export function findTemplate(name: string): Template {
  const found = TEMPLATES.find(template => template.name === name);
  if (!found) throw new Error(`There is no template "${name}". The templates are: ${TEMPLATES.map(template => template.name).join(", ")}. List them with: teams-cli card templates`);
  return found;
}

export type BuiltCard = { template: string; card: El; bytes: number; votes: string[] };

/** Validates the data, builds the card, and checks that Teams accepts its size. */
export function buildTemplateCard(name: string, raw: unknown): BuiltCard {
  const template = findTemplate(name);
  let data: unknown;
  try {
    data = template.parse(raw);
  } catch (error) {
    if (error instanceof FieldError) throw new Error(`The data for template "${name}" is not valid: ${error.message}`);
    throw error;
  }
  const card = template.build(data);
  const bytes = cardBytes(card);
  if (bytes > MAX_TEAMS_CARD_BYTES) {
    throw new Error(`The "${name}" card is ${bytes} bytes. Teams refuses a card over ${MAX_TEAMS_CARD_BYTES} bytes (${MAX_TEAMS_CARD_BYTES / 1024} KB). Shorten the text or remove items.`);
  }
  return { template: name, card, bytes, votes: template.votes?.(data) ?? [] };
}
