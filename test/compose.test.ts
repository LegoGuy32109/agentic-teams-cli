import { describe, expect, test } from "bun:test";
import { buildSendBody, checkBodySize, composePlain, composeRich, escapeHtml, htmlTextSpans, unresolvedMentionWarning } from "../src/compose.ts";
import { PLACEHOLDER_END, PLACEHOLDER_START, type User } from "../src/users.ts";

const ALEX: User = { mri: "8:orgid:33333333-3333-3333-3333-333333333333", displayName: "Alex Doe", fetchedAt: "" };
const MENTION_TYPE = "http://schema.skype.com/Mention";
const ONE = (itemid: string) => `<readonly class="skipProofing" itemtype="${MENTION_TYPE}" spellcheck="false"><span itemtype="${MENTION_TYPE}" itemscope itemid="${itemid}">Alex Doe</span></readonly>`;

describe("plain composition", () => {
  test("one mention matches the browser shape byte for byte", () => {
    const composed = composePlain("0 we", [ALEX]);
    expect(composed.html).toBe(`<p>${ONE("0")} we</p>`);
    expect(composed.entries).toEqual([{ "@type": MENTION_TYPE, itemid: "0", mri: ALEX.mri, mentionType: "person", displayName: "Alex Doe" }]);
  });

  test("two mentions of one person count itemid 0 and 1", () => {
    const composed = composePlain("0 and 1", [ALEX, ALEX]);
    expect(composed.html).toBe(`<p>${ONE("0")} and ${ONE("1")}</p>`);
    expect(composed.entries.map(entry => entry.itemid)).toEqual(["0", "1"]);
  });

  test("the mentions property is a JSON string with the expected key order", () => {
    const body = buildSendBody({
      conversationId: "48:notes",
      content: "<p>x</p>",
      mentions: composePlain("0", [ALEX]).entries,
      mri: "8:orgid:11111111-1111-1111-1111-111111111111",
      name: "Test User",
      now: "2026-10-14T14:41:23.194Z",
      clientMessageId: "1234567890123456789",
    });
    expect(body.properties.mentions).toBe(`[{"@type":"${MENTION_TYPE}","itemid":"0","mri":"${ALEX.mri}","mentionType":"person","displayName":"Alex Doe"}]`);
    expect(buildSendBody({ ...baseBody(), mentions: [] }).properties.mentions).toBe("[]");
  });

  test("text is escaped and each line is a paragraph", () => {
    expect(composePlain("a < b & \"c\" 'd'\nsecond", []).html).toBe("<p>a &lt; b &amp; &quot;c&quot; &#39;d&#39;</p><p>second</p>");
    expect(escapeHtml("<&>")).toBe("&lt;&amp;&gt;");
  });

  test("an empty line becomes a non-breaking space paragraph", () => {
    expect(composePlain("one\n\nthree", []).html).toBe("<p>one</p><p>&nbsp;</p><p>three</p>");
  });

  test("the body keeps the browser field shape", () => {
    const body = buildSendBody({ ...baseBody(), mentions: [] });
    expect(Object.keys(body)).toEqual(["id", "type", "conversationid", "conversationLink", "from", "fromUserId", "composetime", "originalarrivaltime", "content", "messagetype", "contenttype", "imdisplayname", "clientmessageid", "callId", "state", "version", "amsreferences", "properties", "crossPostChannels"]);
    expect(body.messagetype).toBe("RichText/Html");
    expect(body.properties.formatVariant).toBe("TEAMS");
    expect(body.conversationLink).toBe("blah/48:notes");
  });
});

describe("rich composition", () => {
  test("consecutive bullets become one list and other lines are paragraphs", () => {
    expect(composeRich("Title\n- one\n* two\n\nlast", []).html).toBe("<p>Title</p><ul><li>one</li><li>two</li></ul><p>last</p>");
  });

  test("a mention inside a bullet keeps its item id", () => {
    expect(composeRich("- 0 here", [ALEX]).html).toBe(`<ul><li>${ONE("0")} here</li></ul>`);
  });
});

describe("rich composition with mentions", () => {
  const at = (index: number) => `${PLACEHOLDER_START}${index}${PLACEHOLDER_END}`;

  test("a mention in a table cell, a quote, a color span and a list item gets item ids in document order", () => {
    const composed = composeRich(`| who |\n|---|\n| ${at(0)} |\n> ${at(1)} said\n[red]${at(2)}[/]\n- ${at(3)}`, [ALEX, ALEX, ALEX, ALEX]);
    expect(composed.html).toContain(`<td><p>${ONE("0")}</p></td>`);
    expect(composed.html).toContain(`<blockquote><p>${ONE("1")} said</p></blockquote>`);
    expect(composed.html).toContain(`<span style="color:#B6424C">${ONE("2")}</span>`);
    expect(composed.html).toContain(`<ul><li>${ONE("3")}</li></ul>`);
    expect(composed.entries.map(entry => entry.itemid)).toEqual(["0", "1", "2", "3"]);
  });
});

describe("rich text: bold, numbered lists and links", () => {
  const URL = "https://claude.ai/artifact/6rw2pZMRMdscZdVXSrU8p1";
  const anchor = (url: string, label: string, id: string) =>
    `<a href="${url}" itemtype="http://schema.skype.com/HyperLink" rel="noreferrer noopener" title="${url}" target="_blank" itemid="${id}">${label}</a>`;

  test("*bold* and **bold** become strong, inside a line or alone", () => {
    expect(composeRich("*How I ran it*", []).html).toBe("<p><strong>How I ran it</strong></p>");
    expect(composeRich("*On rules:* no change, **really**.", []).html).toBe("<p><strong>On rules:</strong> no change, <strong>really</strong>.</p>");
    expect(composeRich("1. *C-36 (Ohio Table 4)*: is it right?", []).html).toBe("<ol><li><strong>C-36 (Ohio Table 4)</strong>: is it right?</li></ol>");
  });

  test("a star that is not a bold pair stays text, and bold text is escaped", () => {
    expect(composeRich("2 * 3 * 4 = 24", []).html).toBe("<p>2 * 3 * 4 = 24</p>");
    expect(composeRich("a*b*c and * alone and *", []).html).toBe("<p>a*b*c and * alone and *</p>");
    expect(composeRich("a*b* tail", []).html).toBe("<p>a*b* tail</p>");
    expect(composeRich("*b*c tail", []).html).toBe("<p>*b*c tail</p>");
    expect(composeRich("*a < b*", []).html).toBe("<p><strong>a &lt; b</strong></p>");
  });

  test("numbered lines become one ol, and a bullet after them starts a ul", () => {
    expect(composeRich("Intro\n1. one\n2) two\n\n- bullet\nend", []).html).toBe("<p>Intro</p><ol><li>one</li><li>two</li></ol><ul><li>bullet</li></ul><p>end</p>");
  });

  test("a Slack link, a labelled link and a bare link become Teams anchors with a links property", () => {
    const composed = composeRich(`see <${URL}> and <https://x.test/a?b=1&c=2|the *page*>, or https://y.test/z.`, []);
    expect(composed.html).toBe(
      `<p>see ${anchor(URL, URL, composed.links.length ? composed.html.match(/itemid="([^"]+)"/)![1]! : "")} and ` +
      `${anchor("https://x.test/a?b=1&amp;c=2", "the *page*", composed.html.match(/title="https:\/\/x.test[^"]*" target="_blank" itemid="([^"]+)"/)![1]!)}, or ` +
      `${anchor("https://y.test/z", "https://y.test/z", composed.html.match(/title="https:\/\/y.test\/z" target="_blank" itemid="([^"]+)"/)![1]!)}.</p>`,
    );
    expect(composed.links).toEqual([
      { "@type": "http://schema.skype.com/HyperLink", itemid: "0", url: URL, previewenabled: true },
      { "@type": "http://schema.skype.com/HyperLink", itemid: "1", url: "https://x.test/a?b=1&c=2", previewenabled: true },
      { "@type": "http://schema.skype.com/HyperLink", itemid: "2", url: "https://y.test/z", previewenabled: true },
    ]);
  });

  test("the anchor id is a UUID that depends only on the URL, so a repeat message has the same content", () => {
    const first = composeRich(`<${URL}>`, []);
    expect(first.html).toMatch(/itemid="[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}"/);
    expect(composeRich(`<${URL}>`, []).html).toBe(first.html);
    expect(composeRich("<https://other.test/>", []).html).not.toBe(first.html);
  });

  test("a trailing period or bracket is not part of a bare URL", () => {
    expect(composeRich("(see https://x.test/a).", []).links[0]!.url).toBe("https://x.test/a");
    expect(composeRich("go to https://x.test/a, then", []).links[0]!.url).toBe("https://x.test/a");
  });

  test("a mention and a link in one line both work", () => {
    const composed = composeRich(`${PLACEHOLDER_START}0${PLACEHOLDER_END} read <${URL}>`, [ALEX]);
    expect(composed.html).toContain(ONE("0"));
    expect(composed.html).toContain(`<a href="${URL}"`);
    expect(composed.entries).toHaveLength(1);
  });

  test("a Markdown link shows its label and links to its URL, in both formats", () => {
    for (const compose of [composeRich, composePlain]) {
      const composed = compose(`see [the recert](${URL}).`, []);
      expect(composed.html).toContain(`<a href="${URL}"`);
      expect(composed.html).toContain(`>the recert</a>.</p>`);
      expect(composed.html).not.toContain("[the recert]");
      expect(composed.links).toEqual([{ "@type": "http://schema.skype.com/HyperLink", itemid: "0", url: URL, previewenabled: true }]);
    }
  });

  test("a Markdown link label is not read as a color tag, and brackets with no URL stay text", () => {
    expect(composeRich(`[red](${URL})`, []).html).toContain(`>red</a>`);
    expect(composeRich("[red]x[/]", []).html).toContain(`<span style="color:`);
    expect(composePlain("[note] (later)", []).html).toBe("<p>[note] (later)</p>");
    expect(composePlain("[note] (later)", []).links).toEqual([]);
  });

  test("the plain format links Slack and bare URLs, and keeps other marks as text", () => {
    const composed = composePlain(`<${URL}|page> *x* https://y.test/z.`, []);
    expect(composed.html).toContain(`>page</a> *x* <a href="https://y.test/z"`);
    expect(composed.html.endsWith(`>https://y.test/z</a>.</p>`)).toBe(true);
    expect(composed.links.map(link => link.url)).toEqual([URL, "https://y.test/z"]);
    expect(composePlain("x", []).links).toEqual([]);
  });

  test("the body carries links as a JSON string, or [] when there are none", () => {
    const links = composeRich(`<${URL}>`, []).links;
    expect(JSON.parse(buildSendBody({ ...baseBody(), links }).properties.links)).toEqual(links);
    expect(buildSendBody({ ...baseBody(), links: [] }).properties.links).toBe("[]");
    expect(buildSendBody(baseBody()).properties.links).toBe("[]");
  });
});

describe("guards", () => {
  test("a plain @Word warns and an email does not", () => {
    expect(unresolvedMentionWarning("ping @alex now")).toContain("Message contains @alex as plain text, not a Teams mention.");
    expect(unresolvedMentionWarning("mail adoe@example.test")).toBeUndefined();
    expect(unresolvedMentionWarning(`done 0`)).toBeUndefined();
  });

  test("the body limit is 100,000 UTF-8 bytes of the final JSON", () => {
    const body = (text: string) => buildSendBody({ ...baseBody(), content: `<p>${text}</p>` });
    const size = (text: string) => Buffer.byteLength(JSON.stringify(body(text)), "utf8");
    const base = size("");
    // ASCII: one byte per character. Each character adds one byte.
    expect(size("x".repeat(99_000 - base))).toBe(99_000);
    expect(() => checkBodySize(body("x".repeat(99_000 - base)))).not.toThrow();
    expect(() => checkBodySize(body("x".repeat(101_000 - base)))).toThrow(
      "Message is 101,000 bytes, above the 100,000 byte limit. Teams allows 102,400 bytes in total.",
    );
  });

  test("multi-byte text counts its UTF-8 bytes, not its characters", () => {
    // "é" is two bytes in UTF-8. Half as many of them fill the same byte budget.
    const body = (text: string) => buildSendBody({ ...baseBody(), content: `<p>${text}</p>` });
    const base = Buffer.byteLength(JSON.stringify(body("")), "utf8");
    const bytesFor = (count: number) => Buffer.byteLength(JSON.stringify(body("é".repeat(count))), "utf8");
    const at99k = Math.floor((99_000 - base) / 2);
    expect(bytesFor(at99k)).toBe(99_000 - ((99_000 - base) % 2));
    expect(() => checkBodySize(body("é".repeat(at99k)))).not.toThrow();
    expect(() => checkBodySize(body("é".repeat(Math.ceil((101_000 - base) / 2))))).toThrow("above the 100,000 byte limit");
  });
});

function baseBody() {
  return {
    conversationId: "48:notes",
    content: "<p>x</p>",
    mentions: [],
    mri: "8:orgid:11111111-1111-1111-1111-111111111111",
    name: "Test User",
    now: "2026-10-14T14:41:23.194Z",
    clientMessageId: "1234567890123456789",
  };
}

describe("html text spans", () => {
  const texts = (html: string) => htmlTextSpans(html).map(span => html.slice(span.start, span.end));

  test("tags and attribute values are not text", () => {
    expect(texts(`<p class="a>b" title='x'>one <b>two</b></p>`)).toEqual(["one ", "two"]);
  });

  test("code, pre, script, style, and comments are not text, nested or not", () => {
    expect(texts("a<code>b<code>c</code>d</code>e<pre>f</pre>g<script>h</script><style>i</style>j<!-- k -->l")).toEqual(["a", "e", "g", "j", "l"]);
  });

  test("a < that cannot start a tag is text, and an unclosed tag hides the rest", () => {
    expect(texts("1 < 2 and 3 <4")).toEqual(["1 < 2 and 3 <4"]);
    expect(texts("a<b title='x")).toEqual(["a"]);
  });
});

describe("badges", () => {
  const badge = (bg: string, fg: string, text: string) => `<span style="background-color:${bg};color:${fg}"><b>&nbsp;${text}&nbsp;</b></span>`;
  const at = (index: number) => `${PLACEHOLDER_START}${index}${PLACEHOLDER_END}`;

  test("a badge is the exact Teams span, with bold and a non-breaking space on each side", () => {
    expect(composeRich("[badge:green]PASSED[/]", []).html).toBe(`<p>${badge("#2B9B62", "white", "PASSED")}</p>`);
  });

  test("text color follows luminance: black on light, white on dark", () => {
    const fg = (name: string) => composeRich(`[badge:${name}]x[/]`, []).html.match(/;color:(\w+)"/)![1];
    expect(fg("yellow")).toBe("black");
    expect(fg("lime")).toBe("black");
    for (const name of ["red", "green", "blue", "purple", "teal", "orange"]) expect(fg(name)).toBe("white");
    expect(composeRich("[badge:faded-jade]x[/]", []).html).toContain("#37797B");
  });

  test("a mention and an emoji work inside a badge", () => {
    const html = composeRich(`[badge:blue]${at(0)} :thumbsup:[/]`, [ALEX]).html;
    expect(html).toStartWith('<p><span style="background-color:#1E53A3;color:white"><b>&nbsp;');
    expect(html).toContain(ONE("0"));
    expect(html).not.toContain(":thumbsup:");
    expect(html).toEndWith("&nbsp;</b></span></p>");
  });

  test("an escaped badge tag stays literal", () => {
    expect(composeRich("\\[badge:green]x", []).html).toBe("<p>[badge:green]x</p>");
  });

  test("an unknown color stays literal, like an unknown text color", () => {
    expect(composeRich("[badge:nope]x[/]", []).html).toBe("<p>[badge:nope]x[/]</p>");
    expect(composeRich("[nope]x[/]", []).html).toBe("<p>[nope]x[/]</p>");
  });

  test("a badge inside a badge is flattened to the outer one", () => {
    expect(composeRich("[badge:green]a [badge:red]b[/] c[/] d", []).html)
      .toBe(`<p>${badge("#2B9B62", "white", "a b c")} d</p>`);
  });

  test("a badge works in a table cell, a list item, and a heading", () => {
    const green = badge("#2B9B62", "white", "OK");
    expect(composeRich("| a |\n|---|\n| [badge:green]OK[/] |", []).html).toContain(`<td><p>${green}</p></td>`);
    expect(composeRich("- [badge:green]OK[/] done", []).html).toBe(`<ul><li>${green} done</li></ul>`);
    expect(composeRich("# [badge:green]OK[/] title", []).html).toBe(`<h1>${green} title</h1>`);
  });

  test("colored text and highlights are unchanged", () => {
    expect(composeRich("[red]a[/] [hl:yellow]b[/]", []).html).toBe('<p><span style="color:#B6424C">a</span> <span style="background-color:#FDD472">b</span></p>');
  });
});
