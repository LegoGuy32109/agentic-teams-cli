import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { MAX_TEAMS_CARD_BYTES } from "../src/limits.ts";
import { JIRA_ICONS } from "../src/card-templates/icons.ts";
import { EPIC_COLORS, TEMPLATES, buildEpic, buildTemplateCard, findTemplate } from "../src/card-templates/templates.ts";
import { runCli } from "./cli.ts";

type Node = Record<string, any>;

/** Every object below a value that has a type. */
function nodes(value: unknown, found: Node[] = []): Node[] {
  if (Array.isArray(value)) value.forEach(item => nodes(item, found));
  else if (value && typeof value === "object") {
    if (typeof (value as Node).type === "string") found.push(value as Node);
    Object.values(value).forEach(item => nodes(item, found));
  }
  return found;
}
const ofType = (card: unknown, type: string) => nodes(card).filter(node => node.type === type);
const build = (name: string, data?: unknown) => buildTemplateCard(name, data ?? findTemplate(name).example);
const example = (name: string): any => structuredClone(findTemplate(name).example);
const failure = (name: string, data: unknown): string => {
  try { buildTemplateCard(name, data); } catch (error) { return (error as Error).message; }
  throw new Error("The data was accepted.");
};

describe("every template builds a card from its example data", () => {
  test("the registry has the ten templates", () => {
    expect(TEMPLATES.map(template => template.name)).toEqual(
      ["meeting-recap", "deadline", "epic", "pr-review", "decision", "poll", "progress", "incident", "standup", "qa-checklist"]);
  });

  test.each(TEMPLATES.map(template => [template.name]))("%s: a valid card under the Teams size limit", name => {
    const built = build(name);
    expect(built.card.type).toBe("AdaptiveCard");
    expect(built.card.version).toBe("1.5");
    expect(Array.isArray(built.card.body)).toBe(true);
    expect(built.bytes).toBe(Buffer.byteLength(JSON.stringify(built.card)));
    expect(built.bytes).toBeLessThan(MAX_TEAMS_CARD_BYTES);
    for (const block of ofType(built.card, "TextBlock")) expect({ text: block.text, wrap: block.wrap }).toEqual({ text: block.text, wrap: true });
  });

  test.each(TEMPLATES.map(template => [template.name]))("%s: ids are unique and every toggle points at an id", name => {
    const card = build(name).card;
    const ids = nodes(card).map(node => node.id).filter((id): id is string => typeof id === "string");
    expect(new Set(ids).size).toBe(ids.length);
    for (const toggle of ofType(card, "Action.ToggleVisibility")) {
      for (const target of toggle.targetElements) expect({ target, known: ids.includes(target) }).toEqual({ target, known: true });
    }
  });

  test.each(TEMPLATES.map(template => [template.name]))("%s: badges sit in centered auto columns, and no control is a button", name => {
    const card = build(name).card;
    for (const column of ofType(card, "Column").filter(item => item.items.some((child: Node) => child.type === "Badge"))) {
      expect({ width: column.width, align: column.verticalContentAlignment }).toEqual({ width: "auto", align: "Center" });
    }
    // An ActionSet may hold links. Expand and collapse are text links, never buttons.
    for (const set of ofType(card, "ActionSet")) for (const action of set.actions) expect(action.type).toBe("Action.OpenUrl");
  });

  test.each(TEMPLATES.map(template => [template.name]))("%s: only the header holds an image, and the card has no data URI elsewhere", name => {
    const card = build(name).card;
    const images = ofType(card, "Image");
    expect(images.length).toBeLessThanOrEqual(1);
    for (const image of images) expect(card.body[0].columns[0].items[0]).toBe(image);
    expect(JSON.stringify(card).match(/data:image/g)?.length ?? 0).toBe(images.length);
  });

  test("the examples hold no real names, hosts, or addresses", () => {
    const text = JSON.stringify(TEMPLATES.map(template => template.example));
    expect(text).not.toMatch(/josh|hale|suzy|mccreary|seth|foss|mike|broniek|luke|logan|randy|james|daijon|geoh|gogeoh|granola|florida|efax|hha/i);
    for (const address of text.match(/[\w.+-]+@[\w.-]+/g) ?? []) expect(address).toEndWith("@example.com");
    for (const link of text.match(/https?:\/\/[^"\s]+/g) ?? []) expect(new URL(link).hostname).toMatch(/(^|\.)example\.(com|atlassian\.net)$/);
  });

  test("the embedded icons match the PNG files, so a changed icon needs a regenerate", () => {
    const dir = join(import.meta.dir, "..", "src", "assets", "jira-icons");
    const files = readdirSync(dir).filter(name => name.endsWith(".png")).map(name => name.slice(0, -4)).sort();
    expect(Object.keys(JIRA_ICONS).sort()).toEqual(files);
    for (const name of files) expect(JIRA_ICONS[name]).toBe(`data:image/png;base64,${readFileSync(join(dir, `${name}.png`)).toString("base64")}`);
    for (const name of ["story", "bug", "production-bug", "task", "spike", "epic", "feature", "support", "sub-task", "qa-test"]) expect(files).toContain(name);
  });
});

describe("meeting-recap", () => {
  test("participants are a collapsed text link that opens one row of mentions by UPN", () => {
    const card = build("meeting-recap").card;
    const link = card.body[1].items[0];
    expect(link.selectAction).toEqual({ type: "Action.ToggleVisibility", targetElements: ["participants"] });
    expect(link.items[0]).toMatchObject({ text: "▸ Show participants (3)", color: "Accent", size: "Small", weight: "Bolder" });
    const hidden = card.body[1].items[1];
    expect(hidden).toMatchObject({ id: "participants", isVisible: false });
    expect(hidden.items).toHaveLength(1);
    expect(hidden.items[0].text).toBe("<at>Alex</at>  <at>Sam</at>  <at>Jordan</at>");
    expect(card.msteams.entities[1]).toEqual({ type: "mention", text: "<at>Sam</at>", mentioned: { id: "sam@example.com", name: "Sam" } });
    expect(card.msteams.entities).toHaveLength(3);
  });

  test("decisions go in a good box, actions are rows with an owner subline and a status badge, notes are collapsed", () => {
    const card = build("meeting-recap").card;
    const boxes = card.body.filter((node: Node) => node.type === "Container" && node.showBorder);
    expect(boxes[0]).toMatchObject({ style: "good" });
    expect(boxes[0].items[0].text).toBe("Decisions");
    const actions = boxes[1].items.slice(1);
    expect(actions).toHaveLength(3);
    const first = actions[0].items[0].columns;
    expect(first[0].items.map((item: Node) => item.text)).toEqual(["Send the draft spec to the team", "Alex"]);
    expect(first[1].items[0]).toMatchObject({ type: "Badge", text: "Due soon", style: "Warning" });
    expect(card.body.at(-2).items[0].text).toBe("▸ Show full notes");
    expect(card.body.at(-1)).toMatchObject({ id: "notes", isVisible: false });
    expect(card.actions).toEqual([{ type: "Action.OpenUrl", title: "Open notes", url: "https://notes.example.com/recap/123" }]);
  });

  test("with no participants there is no link and no mention entity", () => {
    const data = example("meeting-recap");
    delete data.participants;
    const card = build("meeting-recap", data).card;
    expect(JSON.stringify(card)).not.toContain("Show participants");
    expect(card.msteams).toBeUndefined();
  });
});

describe("deadline", () => {
  test.each([["on-track", "On track", "Good"], ["at-risk", "At risk", "Warning"], ["behind", "Behind", "Attention"]])("%s shows the %s badge in %s", (status, label, style) => {
    const card = build("deadline", { ...example("deadline"), status }).card;
    const badge = ofType(card.body[0], "Badge")[0]!;
    expect(badge).toMatchObject({ text: label, style, appearance: "Filled" });
    expect(card.body[1].columns[0].items[0]).toMatchObject({ text: "12", size: "ExtraLarge", color: style });
  });

  test("the count comes first and the label sits beside it, centered, with the rows in a box that has no title", () => {
    const card = build("deadline").card;
    const [count, label] = card.body[1].columns;
    expect(count.verticalContentAlignment).toBe("Center");
    expect(label).toMatchObject({ width: "stretch", verticalContentAlignment: "Center" });
    expect(label.items[0].text).toBe("days until the target");
    const rowsBox = card.body[2];
    expect(rowsBox.style).toBe("emphasis");
    expect(rowsBox.items[0].type).toBe("Container");
    expect(JSON.stringify(rowsBox)).not.toContain("Checklist");
    expect(ofType(rowsBox, "Badge")[1]).toMatchObject({ text: "Done", icon: "CheckmarkCircle" });
  });
});

describe("epic", () => {
  const issues = (prefix: string, count: number, summary = "A summary of the issue that is about as long as a real one is") =>
    Array.from({ length: count }, (_value, index) => ({ key: `${prefix}-${index + 1}`, summary: `${summary} ${index + 1}` }));
  const big = (counts: number[]) => ({
    key: "EX-1", title: "Big epic", url: "https://example.atlassian.net/browse/EX-1",
    statuses: counts.map((count, index) => ({ name: `Status ${index}`, issues: issues(`EX-${index + 2}`, count) })),
  });

  test("the header has the Epic icon, there are no count tiles, and the chart has the four colors in order", () => {
    const card = build("epic").card;
    expect(card.body[0].columns[0].items[0]).toMatchObject({ type: "Image", url: JIRA_ICONS.epic, width: "32px" });
    expect(card.body[0].columns[1].items[1].text).toBe("Epic · 7 issues");
    expect(card.body[1].type).toBe("Chart.HorizontalBar.Stacked");
    expect(ofType(card, "ColumnSet").filter(set => set.targetWidth)).toHaveLength(0);
    expect(card.body.filter((node: Node) => node.type === "Chart.HorizontalBar.Stacked")).toHaveLength(1);
    expect(card.body[1].data[0].data.map((item: Node) => item.color)).toEqual(["neutral", "categoricalBlue", "categoricalPurple"]);
    expect(EPIC_COLORS).toEqual(["neutral", "categoricalBlue", "categoricalPurple", "categoricalGreen"]);
  });

  test("more statuses than colors start the colors again", () => {
    const card = build("epic", big([1, 1, 1, 1, 1, 1])).card;
    expect(card.body[1].data[0].data.map((item: Node) => item.color)).toEqual([...EPIC_COLORS, EPIC_COLORS[0], EPIC_COLORS[1]]);
  });

  test("a box shows the first 3 issues, then Show N more and Hide text links that toggle the rest", () => {
    const card = build("epic").card;
    const todo = card.body[2];
    expect(todo.items[0].text).toBe("To Do (4)");
    const lines = todo.items.filter((item: Node) => item.type === "TextBlock").slice(1);
    expect(lines).toHaveLength(3);
    expect(lines[0].text).toBe("[EX-111](https://example.atlassian.net/browse/EX-111)  Example story one");
    const [rest, more, hide] = todo.items.slice(4);
    expect(rest).toMatchObject({ id: "rest-0", isVisible: false });
    expect(rest.items).toHaveLength(1);
    expect(more.selectAction.targetElements).toEqual(["rest-0", "more-0", "hide-0"]);
    expect(more.items[0]).toMatchObject({ text: "Show 1 more", color: "Accent", size: "Small", weight: "Bolder" });
    expect(hide).toMatchObject({ id: "hide-0", isVisible: false });
    expect(hide.items[0].text).toBe("Hide");
    // The second box has only 2 issues, so it has nothing to expand.
    expect(card.body[3].items).toHaveLength(3);
    expect(ofType(card.body[3], "Container").filter(item => item.selectAction)).toHaveLength(0);
  });

  test("an issue is one plain TextBlock, with no inline icon", () => {
    const card = build("epic").card;
    expect(ofType(card, "Image")).toHaveLength(1);
    for (const block of ofType(card, "TextBlock").filter(item => /^\[EX-1\d\d\]/.test(item.text))) expect(block.text).toMatch(/^\[EX-\d+\]\(https:\/\/example\.atlassian\.net\/browse\/EX-\d+\)  \S/);
    for (const box of card.body.filter((node: Node) => node.showBorder)) expect(box.items.filter((item: Node) => item.type === "ColumnSet")).toHaveLength(0);
  });

  test("a card that fits is not cut", () => {
    const { card, cut } = buildEpic(big([20, 20]));
    expect(cut).toBe(0);
    expect(JSON.stringify(card)).not.toContain("more in Jira");
  });

  test("a card that is too large loses issues from the hidden parts, with a link to the epic", () => {
    const data = big([80, 80, 80, 80]);
    const { card, cut } = buildEpic(data);
    const bytes = Buffer.byteLength(JSON.stringify(card));
    expect(bytes).toBeLessThanOrEqual(MAX_TEAMS_CARD_BYTES);
    expect(cut).toBeGreaterThan(0);
    // The header still counts every issue, and each status still shows its first 3.
    expect(card.body[0].columns[1].items[1].text).toBe("Epic · 320 issues");
    for (const box of card.body.filter((node: Node) => node.showBorder)) {
      expect(box.items[0].text).toBe("Status " + (card.body.indexOf(box) - 2) + " (80)");
      expect(box.items.slice(1, 4).map((item: Node) => item.text)).toHaveLength(3);
    }
    const links = ofType(card, "TextBlock").filter(item => /more in Jira/.test(item.text));
    expect(links).toHaveLength(4);
    expect(links.map(item => item.text.match(/\[\+(\d+) more in Jira\]\(https:\/\/example\.atlassian\.net\/browse\/EX-1\)/)?.[1]).every(Boolean)).toBe(true);
    // One more issue per box would not have fit: the cut is the smallest that makes it fit.
    const shown = ofType(card, "TextBlock").filter(item => /^\[EX-\d+-?\d*\]/.test(item.text)).length;
    expect(shown + cut).toBe(320);
    expect(buildEpic(data, bytes - 1).cut).toBeGreaterThan(cut);
  });

  test("the cut takes from the status that hides the most", () => {
    const { card } = buildEpic(big([10, 90]), 14_000);
    const hiddenIn = (index: number) => card.body[2 + index].items.find((item: Node) => item.id === `rest-${index}`)?.items.filter((item: Node) => !/more in Jira/.test(item.text)).length ?? 0;
    // The small status keeps all 7 hidden issues. The large one gives up issues first.
    expect(hiddenIn(0)).toBe(7);
    expect(hiddenIn(1)).toBeGreaterThanOrEqual(7);
    expect(hiddenIn(1)).toBeLessThan(87);
    expect(JSON.stringify(card.body[2])).not.toContain("more in Jira");
    expect(JSON.stringify(card.body[3])).toContain("more in Jira");
  });

  test("it fails with a clear message when even the first 3 issues of each status do not fit", () => {
    const statuses = Array.from({ length: 12 }, (_value, index) => ({ name: `Status ${index}`, issues: issues(`EX-${index}`, 5, "x".repeat(900)) }));
    expect(() => buildEpic({ key: "EX-1", title: "Huge", statuses })).toThrow(/even with only the first 3 issues of each status\. Teams refuses a card over 28672 bytes \(28 KB\)/);
    expect(failure("epic", { key: "EX-1", title: "Huge", statuses })).toContain("even with only the first 3 issues");
  });

  test("without a URL, keys are bold text and the cut line is plain", () => {
    const data = { ...big([80, 80, 80, 80]), url: undefined };
    const { card } = buildEpic(data);
    expect(card.actions).toBeUndefined();
    expect(ofType(card, "TextBlock").some(item => item.text.startsWith("**EX-2-1**  "))).toBe(true);
    expect(ofType(card, "TextBlock").some(item => /^\+\d+ more in Jira$/.test(item.text))).toBe(true);
  });
});

describe("pr-review", () => {
  test("the header icon is the Jira icon of the ticket type, and a PR icon without a ticket", () => {
    const withTicket = build("pr-review").card;
    expect(withTicket.body[0].columns[0].items[0]).toMatchObject({ type: "Image", url: JIRA_ICONS.story });
    expect(withTicket.body[0].columns[1].items[1].text).toBe("PE-1234 · Example story title");
    for (const type of ["Production Bug", "production-bug", "Sub-task", "QA Test"]) {
      const card = build("pr-review", { ...example("pr-review"), ticket: { key: "PE-1", type } }).card;
      expect(card.body[0].columns[0].items[0].url).toBe(JIRA_ICONS[type.toLowerCase().replace(/\s+/g, "-")]);
    }
    const data = example("pr-review");
    delete data.ticket;
    delete data.links.ticket;
    delete data.subtitle;
    const plain = build("pr-review", data).card;
    expect(plain.body[0].columns[0].items[0]).toEqual({ type: "Icon", name: "BranchCompare", size: "Medium", color: "Accent" });
  });

  test("the badges are +adds Good, −dels Attention, and files Subtle, with no reviewer or CI tag", () => {
    const card = build("pr-review").card;
    expect(ofType(card.body[1], "Badge")).toEqual([
      { type: "Badge", text: "+214", style: "Good", appearance: "Filled" },
      { type: "Badge", text: "−37", style: "Attention", appearance: "Filled" },
      { type: "Badge", text: "6 files", style: "Subtle", appearance: "Tint" },
    ]);
    expect(ofType(card, "Badge")).toHaveLength(3);
    expect(JSON.stringify(card)).not.toMatch(/reviewer|CI |build/i);
    expect(ofType(build("pr-review", { ...example("pr-review"), files: 1 }).card, "Badge")[2]!.text).toBe("1 file");
  });

  test("What to check is an emphasis box, Not tested is a warning box, and the actions put the extra links in the overflow menu", () => {
    const card = build("pr-review").card;
    const boxes = card.body.filter((node: Node) => node.showBorder);
    expect(boxes.map((node: Node) => [node.style, node.items[0].text])).toEqual([["emphasis", "What to check"], ["warning", "Not tested"]]);
    expect(boxes[1].items[1]).toMatchObject({ text: "Production-size data. Tested on staging only.", color: "Warning" });
    // The links sit in a body ActionSet, so the overflow menu is next to Open PR, not in the card header.
    expect(card.actions).toBeUndefined();
    const set = card.body[card.body.length - 1];
    expect(set.type).toBe("ActionSet");
    expect(set.actions.map((action: Node) => [action.title, action.mode])).toEqual([
      ["Open PR", undefined], ["View diff", "secondary"], ["Open ticket PE-1234", "secondary"], ["Open the pipeline run", "secondary"], ["Open the test plan", "secondary"],
    ]);
  });
});

describe("decision", () => {
  test("the table has a header row in the default style, and each option has an emoji label and a risk badge", () => {
    const card = build("decision").card;
    const table = card.body[1];
    expect(table).toMatchObject({ type: "Table", firstRowAsHeaders: true });
    expect(table.rows[0].style).toBeUndefined();
    expect(JSON.stringify(table.rows[0])).not.toContain("accent");
    expect(table.rows[0].cells.map((cell: Node) => cell.items[0].text)).toEqual(["Option", "Effect", "Risk"]);
    expect(table.rows.slice(1).map((row: Node) => row.cells[0].items[0].text)).toEqual(["1️⃣ Admins only", "2️⃣ Audited", "3️⃣ As is"]);
    expect(table.rows[3].cells[2].items[0]).toMatchObject({ type: "Badge", text: "High", style: "Attention" });
  });

  test("options without an emoji get 1️⃣ 2️⃣ 3️⃣ in order, and a given emoji is kept", () => {
    expect(build("decision").votes).toEqual(["1️⃣", "2️⃣", "3️⃣"]);
    const data = example("decision");
    data.options[1].emoji = "🚀";
    const built = build("decision", data);
    expect(built.votes).toEqual(["1️⃣", "🚀", "3️⃣"]);
    expect(JSON.stringify(built.card)).toContain("🚀 Audited");
    const ten = { ...data, options: Array.from({ length: 10 }, (_v, index) => ({ label: `Option ${index}`, effect: "x", risk: { text: "Low", style: "good" } })) };
    expect(build("decision", ten).votes).toEqual(["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"]);
  });

  test("it says to react to vote, and has the recommendation in a good box", () => {
    const card = build("decision").card;
    const texts = ofType(card, "TextBlock").map(item => item.text);
    expect(texts).toContain("**React 1️⃣, 2️⃣, or 3️⃣ to vote.**");
    expect(card.body.find((node: Node) => node.style === "good").items[0].text).toBe("Recommendation");
    expect(card.body[0].columns.at(-1).items[0]).toMatchObject({ text: "Privacy risk", style: "Attention", appearance: "Filled" });
    expect(ofType(card, "Action.Submit")).toHaveLength(0);
  });

  test("two options give 'or' with no comma", () => {
    const data = example("decision");
    data.options = data.options.slice(0, 2);
    expect(ofType(build("decision", data).card, "TextBlock").map(item => item.text)).toContain("**React 1️⃣ or 2️⃣ to vote.**");
  });
});

describe("poll", () => {
  test("a Vote by reacting box lists the options, and there are no buttons", () => {
    const built = build("poll");
    const card = built.card;
    expect(card.body[1].style).toBe("default");
    expect(card.body[1].items[0].text).toBe("**Vote by reacting to this message:**");
    expect(card.body[1].items.slice(1).map((item: Node) => item.items[0].columns[0].items.map((child: Node) => child.text))).toEqual([
      ["🙋  I'll take it", "You own the deploy calendar this week"], ["🤝  I can back up", "Second on call"], ["⛔  I have a conflict", "Time off or a meeting clash"],
    ]);
    expect(built.votes).toEqual(["🙋", "🤝", "⛔"]);
    expect(card.actions).toBeUndefined();
    expect(nodes(card).filter(node => node.type.startsWith("Action.") || node.type.startsWith("Input."))).toHaveLength(0);
  });
});

describe("progress", () => {
  test("each item has a status badge and a ProgressBar in the matching color", () => {
    const data = example("progress");
    data.items.push({ name: "Fourth", percent: 10, status: { text: "Planned", style: "Informative" } });
    const card = build("progress", data).card;
    const items = card.body.slice(1);
    expect(items.map((item: Node) => [item.items[0].columns[1].items[0].style, item.items[1].color, item.items[1].value, item.items[1].max])).toEqual([
      ["Good", "Good", 70, 100], ["Attention", "Attention", 45, 100], ["Warning", "Warning", 25, 100], ["Informative", "Accent", 10, 100],
    ]);
    expect(items[3].items[0].columns[0].items).toHaveLength(1);
  });
});

describe("incident", () => {
  const timeline = (card: Node) => card.body[2].items.slice(1).map((item: Node) => ofType(item, "Icon")[0]);

  test("the timeline icons are ErrorCircle/Attention, Info/Accent, Search/Warning, and CheckmarkCircle/Good", () => {
    const data = example("incident");
    data.timeline.push({ time: "4:55 PM", text: "Fixed", kind: "resolved" });
    const icons = timeline(build("incident", data).card).map((icon: Node) => [icon.name, icon.color]);
    expect(icons).toEqual([["ErrorCircle", "Attention"], ["Info", "Accent"], ["Search", "Warning"], ["CheckmarkCircle", "Good"]]);
  });

  test("open shows the severity badge, and resolved turns the badge green while the icon stays Alert", () => {
    const open = build("incident").card;
    expect(open.body[0].columns[0].items[0].name).toBe("Alert");
    expect(open.body[0].columns[2].items[0]).toMatchObject({ text: "SEV 2", style: "Attention", appearance: "Filled" });
    expect(open.body[1].facts).toEqual([{ title: "Impact", value: "Example uploads are failing" }, { title: "Owner", value: "Alex" }, { title: "Next update", value: "5:00 PM" }]);
    const done = build("incident", { ...example("incident"), resolved: true, nextUpdate: undefined }).card;
    expect(done.body[0].columns[0].items[0].name).toBe("Alert");
    expect(done.body[0].columns[2].items[0]).toMatchObject({ text: "Resolved", style: "Good", appearance: "Filled" });
    expect(done.body[1].facts[2]).toEqual({ title: "Next update", value: "Closed" });
  });

  test("a card update with a longer timeline changes only the timeline and the badge", () => {
    const before = build("incident").card;
    const data = example("incident");
    data.timeline.push({ time: "4:55 PM", text: "Fixed", kind: "resolved" });
    const after = build("incident", { ...data, resolved: true }).card;
    expect(after.body[2].items).toHaveLength(before.body[2].items.length + 1);
    expect(after.body[1]).toEqual({ ...before.body[1], facts: [before.body[1].facts[0], before.body[1].facts[1], before.body[1].facts[2]] });
  });
});

describe("standup", () => {
  test("yesterday and today are default boxes, and blockers use the warning style", () => {
    const card = build("standup").card;
    expect(card.body.slice(1).map((node: Node) => [node.style, node.items[0].text])).toEqual([["default", "Yesterday"], ["default", "Today"], ["warning", "Blockers"]]);
    expect(card.body[3].items[1]).toMatchObject({ text: "Waiting on the spec from Sam", color: "Warning" });
    expect(card.body[1].items[1].text).toBe("• Finished the first three example stories");
  });

  test("no blockers shows None in a default box", () => {
    const card = build("standup", { ...example("standup"), blockers: [] }).card;
    expect(card.body[3]).toMatchObject({ style: "default" });
    expect(card.body[3].items[1].text).toBe("None");
  });
});

describe("qa-checklist", () => {
  test("each check has a checkbox, a How to test text link, visible numbered steps, and the expected result", () => {
    const card = build("qa-checklist").card;
    expect(card.body[0].columns[1].items[1].text).toBe("Tick each check as you go. Ticks stay on your screen only.");
    const checks = card.body.filter((node: Node) => node.separator);
    expect(checks).toHaveLength(2);
    const [row, link, steps] = checks[0].items;
    expect(row.columns[0].items[0]).toMatchObject({ type: "Input.Toggle", id: "check-1", title: "", value: "false" });
    expect(row.columns[1].items.map((item: Node) => item.text)).toEqual(["Totals round to cents", "PE-1234"]);
    expect(link.selectAction).toEqual({ type: "Action.ToggleVisibility", targetElements: ["steps-1"] });
    expect(link.items[0]).toEqual({ type: "TextBlock", text: "▾ How to test", wrap: true, color: "Accent", size: "Small", weight: "Bolder" });
    expect(steps).toMatchObject({ id: "steps-1", isVisible: true });
    expect(steps.items.map((item: Node) => item.text)).toEqual([
      "1. Open an example claim with three lines", "2. Save the claim", "3. Open the claim total", "**Expected:** The total is the sum of the rounded lines.",
    ]);
  });

  test("there are no outside links, and no submit", () => {
    const card = build("qa-checklist").card;
    expect(card.actions).toBeUndefined();
    expect(nodes(card).filter(node => node.type === "Action.OpenUrl" || node.type === "Action.Submit")).toHaveLength(0);
    expect(JSON.stringify(card)).not.toMatch(/"url"|\]\(http/);
    expect(ofType(card, "Input.Toggle")).toHaveLength(2);
  });
});

describe("data validation", () => {
  const cases: Array<[string, string, (data: any) => unknown, string]> = [
    ["a missing required field", "poll", data => { delete data.question; return data; }, "question is required."],
    ["a field of the wrong type", "poll", data => ({ ...data, question: 7 }), "question must be a string, not number."],
    ["an empty string", "standup", data => ({ ...data, name: "  " }), "name must not be empty."],
    ["a field that is not in the template", "poll", data => ({ ...data, titel: "x" }), '"titel" is not a field here. The fields are: question, context, options.'],
    ["a nested field that is not in the template", "decision", data => { data.options[0].colour = "red"; return data; }, "options[0].colour is not a field here"],
    ["a list that is not a list", "standup", data => ({ ...data, today: "x" }), "today must be a list, not string."],
    ["too few options", "poll", data => ({ ...data, options: data.options.slice(0, 1) }), "options needs at least 2 items."],
    ["too many options", "decision", data => ({ ...data, options: Array.from({ length: 11 }, () => data.options[0]) }), "options has 11 items. The most is 10."],
    ["an emoji that is not in the catalog", "poll", data => { data.options[1].emoji = "ab"; return data; }, 'options[1].emoji "ab" is not an emoji in the Teams catalog'],
    ["a word instead of an emoji", "poll", data => { data.options[1].emoji = "launch"; return data; }, "is not an emoji in the Teams catalog"],
    ["a repeated emoji", "poll", data => { data.options[1].emoji = data.options[0].emoji; return data; }, `options[1].emoji "🙋" is used twice.`],
    ["a repeated default and given emoji", "decision", data => { data.options[2].emoji = "1️⃣"; return data; }, 'options[2].emoji "1️⃣" is used twice.'],
    ["a badge style that does not exist", "progress", data => { data.items[0].status.style = "Purple"; return data; }, 'items[0].status.style "Purple" is not a badge style. Use one of: Default, Accent, Attention, Good, Informative, Subtle, Warning.'],
    ["a percent over 100", "progress", data => { data.items[2].percent = 120; return data; }, "items[2].percent must be from 0 to 100."],
    ["a percent with decimals", "progress", data => { data.items[2].percent = 12.5; return data; }, "items[2].percent must be a whole number."],
    ["a status that is not allowed", "deadline", data => ({ ...data, status: "late" }), 'status "late" is not allowed. Use one of: on-track, at-risk, behind.'],
    ["a timeline kind that is not allowed", "incident", data => { data.timeline[0].kind = "boom"; return data; }, 'timeline[0].kind "boom" is not allowed. Use one of: error, info, investigating, resolved.'],
    ["a Jira type with no icon", "pr-review", data => { data.ticket.type = "Saga"; return data; }, 'ticket.type "Saga" has no icon. Use one of: bug, epic,'],
    ["a link that is not a URL", "pr-review", data => { data.links.pr = "not a url"; return data; }, 'links.pr must be an http or https URL, not "not a url".'],
    ["a link with a script scheme", "deadline", data => ({ ...data, link: "javascript:alert(1)" }), "link must be an http or https URL"],
    ["a UPN that is not an address", "meeting-recap", data => { data.participants[1].upn = "sam"; return data; }, 'participants[1].upn must be a user principal name such as alex@example.com, not "sam".'],
    ["a participant name with markup", "meeting-recap", data => { data.participants[0].name = "A<b>"; return data; }, "participants[0].name must not contain"],
    ["a repeated participant name", "meeting-recap", data => { data.participants[1].name = "Alex"; return data; }, 'participants[1].name "Alex" is used twice.'],
    ["a check with no steps", "qa-checklist", data => { data.checks[0].steps = []; return data; }, "checks[0].steps needs at least 1 item."],
    ["a negative count", "pr-review", data => ({ ...data, adds: -1 }), "adds must be from 0 to"],
    ["an epic with no statuses", "epic", data => ({ ...data, statuses: [] }), "statuses needs at least 1 item."],
  ];

  test.each(cases)("%s names the field", (_name, template, mutate, expected) => {
    const message = failure(template, mutate(example(template)));
    expect(message).toStartWith(`The data for template "${template}" is not valid: `);
    expect(message).toContain(expected);
  });

  test("data that is not an object fails", () => {
    expect(failure("poll", [])).toContain("data must be an object, not a list.");
    expect(failure("poll", null)).toContain("data must be an object, not null.");
  });

  test("an unknown template lists the known ones", () => {
    expect(() => findTemplate("nope")).toThrow('There is no template "nope". The templates are: meeting-recap, deadline, epic,');
  });

  test("a card over the Teams size limit fails with its size, for a template that does not cut", () => {
    const data = example("standup");
    data.today = Array.from({ length: 700 }, (_value, index) => `Item number ${index} with some words to make the line long enough`);
    expect(failure("standup", data)).toMatch(/The "standup" card is \d+ bytes\. Teams refuses a card over 28672 bytes \(28 KB\)\. Shorten the text or remove items\./);
  });
});

describe("the card commands", () => {
  const env = { TEAMS_CLI_CHECK_ARGS: "1" };

  test("a card file and --template are mutually exclusive, for send and update", async () => {
    for (const args of [
      ["card", "send", "Eng/General", "card.json", "--template", "poll", "--data", "d.json"],
      ["card", "update", "19:abc@thread.tacv2#1792428352468", "card.json", "--template", "poll", "--data", "d.json"],
    ]) {
      const result = await runCli(args, { env });
      expect(result.code).toBe(1);
      expect(result.stderr).toContain("Give either a card file or --template, not both.");
    }
  });

  test("--template needs --data, and --data needs --template", async () => {
    const noData = await runCli(["card", "send", "Eng/General", "--template", "poll"], { env });
    expect(noData.stderr).toContain("--template poll needs --data FILE.json.");
    const noTemplate = await runCli(["card", "send", "Eng/General", "--data", "d.json"], { env });
    expect(noTemplate.stderr).toContain("--data goes with --template.");
    const ok = await runCli(["card", "send", "Eng/General", "--template", "poll", "--data", "d.json", "--no-react"], { env });
    expect(ok.stdout).toContain("arguments ok");
  });

  test("a card still needs its file and a conversation without --template", async () => {
    const result = await runCli(["card", "send", "Eng/General"], { env });
    expect(result.code).toBe(1);
    expect(result.stderr).toContain("card send <channel> <card.json>");
  });

  test("card templates lists the ten templates, in text and in JSON", async () => {
    const text = await runCli(["card", "templates"]);
    expect(text.code).toBe(0);
    for (const template of TEMPLATES) {
      expect(text.stdout).toContain(template.name);
      expect(text.stdout).toContain(template.fields);
    }
    const json = JSON.parse((await runCli(["card", "templates", "--json"])).stdout);
    expect(json.map((item: Node) => item.name)).toEqual(TEMPLATES.map(template => template.name));
    expect(json.filter((item: Node) => item.votes).map((item: Node) => item.name)).toEqual(["decision", "poll"]);
    expect(Object.keys(json[0])).toEqual(["name", "purpose", "fields", "votes"]);
  });

  test("card example prints data that the template accepts, offline", async () => {
    for (const template of TEMPLATES) {
      const result = await runCli(["card", "example", template.name]);
      expect(result.code).toBe(0);
      expect(JSON.parse(result.stdout)).toEqual(template.example);
      expect(() => buildTemplateCard(template.name, JSON.parse(result.stdout))).not.toThrow();
    }
    const missing = await runCli(["card", "example", "nope"]);
    expect(missing.code).toBe(1);
    expect(missing.stderr).toContain('There is no template "nope".');
  });
});
