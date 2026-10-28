import { describe, expect, test } from "bun:test";
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { codeBlockHtml, highlightCode, highlightTokens, resolveLanguage } from "../src/highlight.ts";
import { composeRich, extractFences } from "../src/compose.ts";
import { HIGHLIGHTS, TEXT_COLORS, highlightColor, textColor } from "../src/palette.ts";

const FIXTURES = join(import.meta.dir, "fixtures", "teams-code");
const LANGUAGE: Record<string, string> = {
  javascript: "javascript", typescript: "typescript", typescript2: "typescript", jsx: "javascript", python: "python", rust: "rust",
  bash: "bash", powershell: "powershell", yaml: "yaml", json: "json", markdown: "markdown", xml: "xml", http: "http",
  sql: "sql", "sql-kql-text": "sql", kql: "kql",
};

const unescape = (text: string) => text.replace(/&nbsp;| /g, " ").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, "&");

/** Teams output as `{class:text}` runs. Spaces and line breaks are normalized, because they do not change the look. */
function normalize(html: string): string {
  let out = "";
  for (const match of html.replace(/<br>/g, "\n").matchAll(/<span class="hljs-(\w+)">([^<]*)<\/span>|([^<]+)/g)) {
    out += match[1] ? `{${match[1]}:${unescape(match[2]!)}}` : unescape(match[3]!);
  }
  return out;
}

describe("syntax colors match what the Teams web client wrote for the same source", () => {
  const files = readdirSync(FIXTURES).filter(name => name.endsWith(".html"));
  test("every language has a fixture", () => {
    expect(files.map(name => name.replace(".html", "")).sort()).toEqual(Object.keys(LANGUAGE).sort());
  });
  for (const file of files) {
    const name = file.replace(".html", "");
    test(name, () => {
      const fixture = readFileSync(join(FIXTURES, file), "utf8");
      const source = unescape(fixture.replace(/<br>/g, "\n").replace(/<[^>]+>/g, ""));
      const mine = highlightTokens(LANGUAGE[name]!, source).map(([cls, text]) => cls ? `{${cls}:${text}}` : text).join("");
      expect(mine).toBe(normalize(fixture));
    });
  }
});

describe("language names", () => {
  test("common labels map to the Teams language names", () => {
    const expected: Record<string, string> = {
      js: "javascript", jsx: "javascript", ts: "typescript", py: "python", rs: "rust", sh: "bash", shell: "bash", ps1: "powershell",
      pwsh: "powershell", yml: "yaml", md: "markdown", html: "xml", kusto: "kql", JSON: "json", " Sql ": "sql", http: "http",
    };
    for (const [label, language] of Object.entries(expected)) expect(resolveLanguage(label)).toBe(language);
    expect(resolveLanguage("cobol")).toBeUndefined();
    expect(resolveLanguage("")).toBeUndefined();
  });
});

describe("code block HTML", () => {
  test("a known language has the editor marker, the label classes, an id without x_, and spans", () => {
    const html = codeBlockHtml("py", "print(1)");
    const id = html.match(/id="(codeBlockEditor-[0-9a-f-]{36})"/)![1]!;
    expect(html).toBe(
      `<p itemtype="http://schema.skype.com/CodeBlockEditor" id="${id}">&nbsp;</p>` +
      `<pre class="language-python language-was-manually-selected" itemid="${id}"><code><span class="hljs-function">print</span>(1)</code></pre>`,
    );
    expect(id).not.toContain("x_");
  });

  test("the id depends only on the language and the code", () => {
    expect(codeBlockHtml("py", "a")).toBe(codeBlockHtml("python", "a"));
    expect(codeBlockHtml("py", "a")).not.toBe(codeBlockHtml("py", "b"));
  });

  test("text is escaped, spaces are no-break spaces and newlines are br", () => {
    expect(highlightCode(undefined, "a < b &\n  c")).toBe("a&nbsp;&lt;&nbsp;b&nbsp;&amp;<br>&nbsp;&nbsp;c");
  });

  test("an unknown label keeps its own class with no spans, and an odd label gives a bare block", () => {
    expect(codeBlockHtml("scala", "val x = 1")).toContain(`<pre class="language-scala language-was-manually-selected"`);
    expect(codeBlockHtml("scala", "val x = 1")).toContain("<code>val&nbsp;x&nbsp;=&nbsp;1</code>");
    expect(codeBlockHtml("", "x")).toBe("<pre><code>x</code></pre>");
    expect(codeBlockHtml("a b<c>", "x")).toBe("<pre><code>x</code></pre>");
  });
});

describe("fences in rich text", () => {
  test("a fenced block between paragraphs becomes a code block", () => {
    const html = composeRich("before\n```js\nconst a = 1;\n```\nafter", []).html;
    expect(html).toStartWith("<p>before</p><p itemtype=");
    expect(html).toContain(`<pre class="language-javascript language-was-manually-selected"`);
    expect(html).toContain('<span class="hljs-keyword">const</span>');
    expect(html).toEndWith("</code></pre><p>after</p>");
  });

  test("a fence closes an open list, and the block keeps its text as written", () => {
    const html = composeRich("- one\n```\n*not bold* <b>\n```", []).html;
    expect(html).toBe("<ul><li>one</li></ul><pre><code>*not&nbsp;bold*&nbsp;&lt;b&gt;</code></pre>");
  });

  test("a fence that never closes runs to the end", () => {
    expect(composeRich("```json\n{\"a\": 1}", []).html).toContain('<span class="hljs-function">&quot;a&quot;</span>');
  });

  test("extractFences sets blocks aside so mention syntax inside code is not touched", () => {
    const fenced = extractFences("run:\n```powershell\n$h = @{a=1}\n```\ndone");
    expect(fenced.text).not.toContain("@{");
    expect(fenced.blocks).toEqual([{ label: "powershell", code: "$h = @{a=1}" }]);
    expect(composeRich(fenced.text, [], fenced.blocks).html).toContain("@{a=1}");
  });
});

describe("headings", () => {
  test("one to three hashes make h1 to h3, deeper levels stay h3 as in Teams, and a hash without a space stays text", () => {
    expect(composeRich("# One\n## Two\n### Three\n#### Four\n###### Six\n#123 not a heading\n####### seven", []).html)
      .toBe("<h1>One</h1><h2>Two</h2><h3>Three</h3><h3>Four</h3><h3>Six</h3><p>#123 not a heading</p><p>####### seven</p>");
  });

  test("a heading can hold bold and a link", () => {
    expect(composeRich("# *Big* news", []).html).toBe("<h1><strong>Big</strong> news</h1>");
  });
});

describe("colors", () => {
  test("the palette holds the 8 text and 8 highlight colors from the Teams picker", () => {
    expect(TEXT_COLORS.map(c => c.hex)).toEqual(["#B6424C", "#CD5937", "#FDC030", "#BDCB4C", "#2B9B62", "#37797B", "#1E53A3", "#A5397A"]);
    expect(HIGHLIGHTS.map(c => c.hex)).toEqual(["#DF9299", "#F4A593", "#FDD472", "#E5F18F", "#82CDA8", "#9DD9DB", "#C7D4E8", "#EBD3E1"]);
    expect(TEXT_COLORS.map(c => c.name)[1]).toBe("Tuscan Orange");
    expect(HIGHLIGHTS.map(c => c.name)[5]).toBe("Regent St. Blue");
  });

  test("a color is found by alias or by its Teams name", () => {
    expect(textColor("red")).toBe("#B6424C");
    expect(textColor("faded-jade")).toBe("#37797B");
    expect(textColor("purple")).toBe("#A5397A");
    expect(highlightColor("yellow")).toBe("#FDD472");
    expect(highlightColor("regent-st-blue")).toBe("#9DD9DB");
    expect(highlightColor("goldenrod")).toBe("#FDD472");
    expect(textColor("goldenrod")).toBeUndefined();
    expect(textColor("chartreuse")).toBeUndefined();
  });

  test("[color] and [hl:color] wrap text in the Teams spans", () => {
    expect(composeRich("a [red]hot[/] and [hl:yellow]marked[/] end", []).html)
      .toBe('<p>a <span style="color:#B6424C">hot</span> and <span style="background-color:#FDD472">marked</span> end</p>');
  });

  test("tags nest, bold works inside, and a tag that is never closed ends with its line", () => {
    expect(composeRich("[hl:green][blue]x[/] y[/] *b*", []).html)
      .toBe('<p><span style="background-color:#82CDA8"><span style="color:#1E53A3">x</span> y</span> <strong>b</strong></p>');
    expect(composeRich("[red]open\nnext", []).html).toBe('<p><span style="color:#B6424C">open</span></p><p>next</p>');
  });

  test("an unknown name, a stray close tag and a number in brackets stay text", () => {
    expect(composeRich("[chartreuse]x[/] [1] [/]", []).html).toBe("<p>[chartreuse]x[/] [1] [/]</p>");
  });

  test("a color tag inside a code block is code", () => {
    expect(composeRich("```\n[red]x[/]\n```", []).html).toBe("<pre><code>[red]x[/]</code></pre>");
  });
});

describe("italics, strikethrough, inline code and underline", () => {
  test("_italic_ ~strike~ and `code` become the shapes that Teams stores", () => {
    expect(composeRich("an _italic_ word, ~struck~ out, and `inline code` here", []).html)
      .toBe("<p>an <i>italic</i> word, <s>struck</s> out, and <code>inline code</code> here</p>");
  });

  test("inline code keeps its text as written, including stars, tags and quotes", () => {
    expect(composeRich("`*x* [red] <b> it's`", []).html).toBe("<p><code>*x* [red] &lt;b&gt; it&#39;s</code></p>");
  });

  test("underscores in words and URLs, tildes in paths, and lone marks stay text", () => {
    expect(composeRich("snake_case_name and my_var_ ok", []).html).toBe("<p>snake_case_name and my_var_ ok</p>");
    expect(composeRich("cd ~/a and ~/b", []).html).toBe("<p>cd ~/a and ~/b</p>");
    expect(composeRich("1 ~ 2 and _ alone and ~", []).html).toBe("<p>1 ~ 2 and _ alone and ~</p>");
    expect(composeRich("_a_b and cd ~/a and~ b and ~a~/b", []).html).toBe("<p>_a_b and cd ~/a and~ b and ~a~/b</p>");
    const html = composeRich("see https://x.test/a_b_c and https://x.test/~u/~v", []).html;
    expect(html).not.toContain("<i>");
    expect(html).not.toContain("<s>");
  });

  test("[u] underlines, nests with colors, and closes in order", () => {
    expect(composeRich("[u]under[/] and [red][u]both[/] red[/]", []).html)
      .toBe('<p><u>under</u> and <span style="color:#B6424C"><u>both</u> red</span></p>');
    expect(composeRich("[u]never closed", []).html).toBe("<p><u>never closed</u></p>");
  });

  test("<u>…</u> underlines like [u], nests with other marks, and a stray close tag stays text", () => {
    expect(composeRich("a <u>word</u> and <u>*bold* under</u> and [red]<u>both</u>[/]", []).html)
      .toBe('<p>a <u>word</u> and <u><strong>bold</strong> under</u> and <span style="color:#B6424C"><u>both</u></span></p>');
    expect(composeRich("stray </u> here and <u>open", []).html).toBe("<p>stray &lt;/u&gt; here and <u>open</u></p>");
    expect(composeRich("[red]x</u>y[/]", []).html).toBe('<p><span style="color:#B6424C">x&lt;/u&gt;y</span></p>');
    expect(composeRich("`<u>code</u>`", []).html).toBe("<p><code>&lt;u&gt;code&lt;/u&gt;</code></p>");
  });

  test("a backslash makes a color tag literal", () => {
    expect(composeRich("this is \\[red] in prose, and \\[hl:blue] too", []).html).toBe("<p>this is [red] in prose, and [hl:blue] too</p>");
  });

  test("none of it applies inside a code block", () => {
    expect(composeRich("```\n_a_ ~b~ `c`\n```", []).html).toBe("<pre><code>_a_&nbsp;~b~&nbsp;`c`</code></pre>");
  });
});

describe("horizontal rules", () => {
  test("three or more dashes, stars or underscores alone on a line give a bare hr", () => {
    expect(composeRich("above\n---\n***\n_____\nbelow", []).html).toBe("<p>above</p><hr><hr><hr><p>below</p>");
  });

  test("a rule ends an open list, and two dashes, a bullet and a rule inside code stay as they are", () => {
    expect(composeRich("- one\n---\n--\n- two", []).html).toBe("<ul><li>one</li></ul><hr><p>--</p><ul><li>two</li></ul>");
    expect(composeRich("```\n---\n```", []).html).toBe("<pre><code>---</code></pre>");
  });
});

describe("pipe tables", () => {
  const table = "| Name | Qty | Note |\n|:-----|:---:|-----:|\n| a | 1 | left |\n| b | 2 | right |";

  test("the separator row sets the align attribute, on header and body cells, with a thead and a p in every cell", () => {
    expect(composeRich(table, []).html).toBe(
      "<table><thead><tr><th align=\"left\"><p>Name</p></th><th align=\"center\"><p>Qty</p></th><th align=\"right\"><p>Note</p></th></tr></thead>" +
      "<tbody><tr><td align=\"left\"><p>a</p></td><td align=\"center\"><p>1</p></td><td align=\"right\"><p>left</p></td></tr>" +
      "<tr><td align=\"left\"><p>b</p></td><td align=\"center\"><p>2</p></td><td align=\"right\"><p>right</p></td></tr></tbody></table>",
    );
  });

  test("a plain separator gives no align attribute, and a header with no body rows is a table", () => {
    expect(composeRich("a | b\n--- | ---", []).html).toBe("<table><thead><tr><th><p>a</p></th><th><p>b</p></th></tr></thead><tbody></tbody></table>");
  });

  test("cells take bold, links, colors and an escaped pipe, and short rows are padded", () => {
    const html = composeRich("| x | y |\n|---|---|\n| **b** | [red]r[/] \\| p |\n| only |", []).html;
    expect(html).toContain("<td><p><strong>b</strong></p></td><td><p><span style=\"color:#B6424C\">r</span> | p</p></td>");
    expect(html).toContain("<tr><td><p>only</p></td><td><p>&nbsp;</p></td></tr>");
  });

  test("text before and after a table, and a rule right after it, keep their own blocks", () => {
    const html = composeRich("before\n| a |\n|---|\n| 1 |\n---\nafter", []).html;
    expect(html).toBe("<p>before</p><table><thead><tr><th><p>a</p></th></tr></thead><tbody><tr><td><p>1</p></td></tr></tbody></table><hr><p>after</p>");
  });

  test("a pipe in prose, a separator with a different cell count, or a separator that is not one is not a table", () => {
    expect(composeRich("a | b", []).html).toBe("<p>a | b</p>");
    expect(composeRich("| a | b |\n|---|", []).html).toBe("<p>| a | b |</p><p>|---|</p>");
    expect(composeRich("| a | b |\n|---| |", []).html).toBe("<p>| a | b |</p><p>|---| |</p>");
    expect(composeRich("| a |\n|:|", []).html).toBe("<p>| a |</p><p>|:|</p>");
    expect(composeRich("| a |\n| x |", []).html).toBe("<p>| a |</p><p>| x |</p>");
  });

  test("a table inside a code block is code", () => {
    expect(composeRich("```\n| a |\n|---|\n```", []).html).toBe("<pre><code>|&nbsp;a&nbsp;|<br>|---|</code></pre>");
  });
});

describe("quotes", () => {
  test("consecutive > lines are one blockquote with a br between lines, and marks work inside", () => {
    expect(composeRich("before\n> first *line*\n>second\n> \nafter", []).html)
      .toBe("<p>before</p><blockquote><p>first <strong>line</strong><br>second<br></p></blockquote><p>after</p>");
  });

  test("two quotes apart are two blocks, and a quote ends a list and is code inside a fence", () => {
    expect(composeRich("> a\ntext\n> b", []).html).toBe("<blockquote><p>a</p></blockquote><p>text</p><blockquote><p>b</p></blockquote>");
    expect(composeRich("- x\n> y", []).html).toBe("<ul><li>x</li></ul><blockquote><p>y</p></blockquote>");
    expect(composeRich("```\n> y\n```", []).html).toBe("<pre><code>&gt;&nbsp;y</code></pre>");
  });
});
