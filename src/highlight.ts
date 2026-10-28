import { createHash } from "node:crypto";
import { escapeHtml } from "./html.ts";

/**
 * Code blocks for `--format=rich`. Teams stores syntax colors as `hljs-*` spans inside the message, and the
 * sending client writes them. A viewer does not highlight again. So the CLI writes the spans itself.
 * Teams uses five classes only: comment, keyword, string, function, default. The token rules below follow the
 * code blocks that the Teams web client produced for the same source (test/fixtures/teams-code). They are close to
 * that output, not a copy of highlight.js.
 */

type Cls = "comment" | "keyword" | "string" | "function" | "default";
export type Tok = [Cls | undefined, string];
type Rule = (code: string, at: number) => [Tok[], number] | undefined;

const sticky = (source: string, flags = "") => new RegExp(source, `${flags}y`);

/** A rule for one regex that gives one class. */
const re = (source: string, cls: Cls, flags = ""): Rule => {
  const pattern = sticky(source, flags);
  return (code, at) => {
    pattern.lastIndex = at;
    const match = pattern.exec(code);
    return match && match[0] ? [[[cls, match[0]]], at + match[0].length] : undefined;
  };
};

/** A rule for a word list. */
const words = (list: string, cls: Cls, flags = ""): Rule => re(`\\b(?:${list})\\b`, cls, flags);

/** A rule that gives the first group a class and keeps the text before and after it plain. */
const group = (source: string, cls: Cls, flags = ""): Rule => {
  const pattern = sticky(source, flags);
  return (code, at) => {
    pattern.lastIndex = at;
    const match = pattern.exec(code);
    if (!match || !match[0]) return undefined;
    const start = match[0].indexOf(match[1]!);
    const toks: Tok[] = [];
    if (start > 0) toks.push([undefined, match[0].slice(0, start)]);
    toks.push([cls, match[1]!]);
    const tail = match[0].slice(start + match[1]!.length);
    if (tail) toks.push([undefined, tail]);
    return [toks, at + match[0].length];
  };
};

function scan(code: string, rules: Rule[]): Tok[] {
  const out: Tok[] = [];
  let plain = "";
  const flush = () => { if (plain) { out.push([undefined, plain]); plain = ""; } };
  let at = 0;
  while (at < code.length) {
    let hit: [Tok[], number] | undefined;
    for (const rule of rules) { hit = rule(code, at); if (hit) break; }
    if (hit) {
      flush();
      out.push(...hit[0]);
      at = hit[1];
    } else {
      // Take a whole word as plain text, so a rule never starts inside a word.
      const word = /^[\w$]+/.exec(code.slice(at))?.[0] ?? code[at]!;
      plain += word;
      at += word.length;
    }
  }
  flush();
  return out;
}

/** A string with embedded `${…}` or `{…}` parts. The parts get the default class. */
const interpolated = (opener: string, quote: string, open: "${" | "{"): Rule => {
  const start = sticky(opener);
  return (code, at) => {
    start.lastIndex = at;
    const head = start.exec(code);
    if (!head) return undefined;
    const toks: Tok[] = [];
    let text = head[0];
    let i = at + head[0].length;
    for (;;) {
      const ch = code[i];
      if (ch === undefined || (ch === "\n" && quote !== "`")) return undefined;
      if (ch === "\\") { text += code.slice(i, i + 2); i += 2; continue; }
      if (code.startsWith(open, i) && !(open === "{" && code[i + 1] === "{")) {
        const end = code.indexOf("}", i);
        if (end < 0) return undefined;
        if (text) toks.push(["string", text]);
        toks.push(["default", code.slice(i, end + 1)]);
        text = "";
        i = end + 1;
        continue;
      }
      text += ch;
      i++;
      if (ch === quote) break;
    }
    if (text) toks.push(["string", text]);
    return [toks, i];
  };
};

const lineComment = (marker: string) => re(`${marker}[^\\n]*`, "comment");
const dq = re(`"(?:[^"\\\\\\n]|\\\\.)*"`, "string");
const sq = re(`'(?:[^'\\\\\\n]|\\\\.)*'`, "string");

const JS_KEYWORDS = "const|let|var|function|async|await|return|if|else|for|while|do|switch|case|break|continue|new|class|extends|import|export|from|default|throw|try|catch|finally|typeof|instanceof|in|of|yield|delete|this|super|static|get|set";
const TS_KEYWORDS = `${JS_KEYWORDS}|interface|type|enum|implements|as|public|private|protected|readonly|abstract|declare|namespace|keyof|satisfies`;

/** A JSX opening tag: the tag name is a keyword, each attribute name a function, each value a string. */
const jsxTag: Rule = (code, at) => {
  if (code[at] !== "<" || /[\w)\]]/.test(code[at - 1] ?? " ")) return undefined;
  const name = /^<(\/?)([A-Za-z][\w.]*)/.exec(code.slice(at));
  if (!name) return undefined;
  const toks: Tok[] = [[undefined, `<${name[1]}`], ["keyword", name[2]!]];
  let i = at + name[0].length;
  if (name[1]) return [toks, i];
  for (;;) {
    const rest = code.slice(i);
    const space = /^\s+/.exec(rest);
    if (space) { toks.push([undefined, space[0]]); i += space[0].length; continue; }
    if (rest.startsWith("=>")) { toks.push([undefined, "=>"]); i += 2; continue; }
    const close = /^\/?>/.exec(rest);
    if (close) { toks.push([undefined, close[0]]); return [toks, i + close[0].length]; }
    const attr = /^[\w:-]+/.exec(rest);
    if (!attr) { toks.push([undefined, rest[0]!]); i++; if (i >= code.length) return [toks, i]; continue; }
    toks.push(["function", attr[0]]);
    i += attr[0].length;
    if (code[i] !== "=") continue;
    toks.push([undefined, "="]);
    i++;
    const value = /^(?:"[^"\n]*"|'[^'\n]*'|\{[^})\n]*[)}])/.exec(code.slice(i));
    if (value) {
      toks.push(["string", value[0]]);
      i += value[0].length;
      // After `{(` the rest of the braces is plain text.
      if (value[0].startsWith("{") && value[0].endsWith(")")) {
        const end = code.indexOf("}", i);
        if (end >= 0) { toks.push([undefined, code.slice(i, end + 1)]); i = end + 1; }
      }
    }
  }
};

function javascript(tsx: boolean): Rule[] {
  return [
    re("\\/\\*[\\s\\S]*?\\*\\/", "comment"), lineComment("\\/\\/"),
    interpolated("`", "`", "${"), dq, sq,
    ...(tsx ? [] : [jsxTag]),
    words(JS_KEYWORDS, "keyword"),
    group("(?<=\\b(?:function|class)\\s+)([A-Za-z_$][\\w$]*)", "function"),
    re("[A-Za-z_$][\\w$]*(?=\\s*\\()", "function"),
  ];
}

function typescript(): Rule[] {
  return [
    re("\\/\\*[\\s\\S]*?\\*\\/", "comment"), lineComment("\\/\\/"),
    interpolated("`", "`", "${"), dq, sq,
    words(TS_KEYWORDS, "keyword"),
    // Names after a keyword, types after `:` or `as`, generic arguments, and parameter or property names.
    re("(?<=\\b(?:function|class|interface|type|enum|as)\\s+)[A-Za-z_$][\\w$]*", "function"),
    re("(?<=:\\s+)(?!void\\b)[A-Za-z_$][\\w$]*(?=[\\s;,)<>=}{]|$)", "function"),
    re("(?<=<)[A-Za-z_$][\\w$]*(?=[>,])", "function"),
    re("[A-Za-z_$][\\w$]*(?=\\??:\\s+[A-Za-z_$])", "function"),
    re("[A-Za-z_$][\\w$]*(?=\\s*\\()", "function"),
  ];
}

const PY_BUILTINS = "str|int|float|bool|list|dict|set|tuple|bytes|print|len|range|open|type|isinstance|sum|min|max|sorted|enumerate|zip|map|filter|abs|round|input|super";

function python(): Rule[] {
  return [
    lineComment("#"),
    re(`[rRbB]*"""[\\s\\S]*?"""`, "string"), re(`[rRbB]*'''[\\s\\S]*?'''`, "string"),
    interpolated(`[fF][rR]?"`, `"`, "{"), interpolated(`[fF][rR]?'`, `'`, "{"),
    dq, sq,
    re("@[\\w.]+", "keyword"),
    words("import|from|def|class|return|if|elif|else|for|while|in|is|not|and|or|raise|try|except|finally|with|as|lambda|pass|break|continue|yield|assert|del|global|nonlocal|async|await", "keyword"),
    group("(?<=\\b(?:def|class)\\s+)([A-Za-z_]\\w*)", "function"),
    words(PY_BUILTINS, "function"),
  ];
}

const RUST_TYPES = "u8|u16|u32|u64|u128|usize|i8|i16|i32|i64|i128|isize|f32|f64|bool|char|str|String|Option|Result|Vec|Box|HashMap|Self";

function rust(): Rule[] {
  return [
    lineComment("\\/\\/"), re("\\/\\*[\\s\\S]*?\\*\\/", "comment"),
    dq,
    re("#!?\\[[^\\]\\n]*\\]", "keyword"),
    words(`use|struct|fn|let|mut|impl|enum|pub|mod|match|if|else|for|while|loop|return|trait|where|as|in|const|static|type|unsafe|async|await|move|ref|crate|self|super|dyn|${RUST_TYPES}`, "keyword"),
    re("(?<=\\b(?:fn|struct|enum|trait)\\s+)[A-Za-z_]\\w*", "function"),
    re("[A-Za-z_]\\w*!?(?=\\()", "function"),
  ];
}

/** A shebang line is plain text in Teams. */
const shebang: Rule = (code, at) => {
  if (at !== 0 || !code.startsWith("#!")) return undefined;
  const line = code.split("\n")[0]!;
  return [[[undefined, line]], line.length];
};

function bash(): Rule[] {
  return [
    shebang,
    re("#[^\\n]*", "comment"),
    dq, sq,
    words("for|in|do|done|if|then|else|elif|fi|while|until|case|esac|function|select", "keyword"),
    words("echo|printf|cd|pwd|export|unset|source|alias|read|exit|eval|exec|wc|sort|uniq|head|tail|cat", "function"),
  ];
}

/** `function Name` is a keyword, then a function-class space, then the name in the default class. */
const powershellFunction: Rule = (code, at) => {
  const pattern = sticky("function(\\s+)([A-Za-z][\\w-]*)", "i");
  pattern.lastIndex = at;
  const found = pattern.exec(code);
  if (!found) return undefined;
  return [[["keyword", "function"], ["function", found[1]!], ["default", found[2]!]], at + found[0].length];
};

function powershell(): Rule[] {
  return [
    lineComment("#"), dq, sq,
    powershellFunction,
    words("param|function|if|else|elseif|foreach|for|while|return|switch|try|catch|finally|throw|begin|process|end", "keyword", "i"),
    group("(?<=\\[)([A-Za-z][\\w.]*)(?=\\])", "function"),
    re("\\b[A-Z][a-z]+-[A-Z][A-Za-z]+\\b", "function"),
  ];
}

const jsonRules = (): Rule[] => [
  re(`"(?:[^"\\\\\\n]|\\\\.)*"(?=\\s*:)`, "function"), dq,
  words("true|false|null", "keyword"),
];

/** YAML is handled by line, because Teams colors a key by whether it holds a value. */
function yaml(code: string): Tok[] {
  const out: Tok[] = [];
  const lines = code.split("\n");
  const value = (text: string) => {
    for (const part of text.split(/("[^"]*"|'[^']*'|[^\s,[\]{}]+)/)) {
      if (!part) continue;
      if (/^("[^"]*"|'[^']*'|[^\s,[\]{}]+)$/.test(part) && !/^(?:-?\d+(?:\.\d+)?|true|false|null|~),?$/.test(part)) out.push(["string", part]);
      else out.push([undefined, part]);
    }
  };
  lines.forEach((line, index) => {
    if (index) out.push([undefined, "\n"]);
    const comment = /^(\s*)(#.*)$/.exec(line);
    if (comment) {
      if (comment[1]) out.push([undefined, comment[1]]);
      comment[2]!.split(/(\s+)/).forEach(part => { if (part) out.push([/^\s+$/.test(part) ? undefined : "string", part]); });
      return;
    }
    const key = /^(\s*)(- )?([\w.-]+:)(?=\s|$)(.*)$/.exec(line);
    if (key) {
      if (key[1]) out.push([undefined, key[1]]);
      if (key[2]) { out.push(["string", "-"], [undefined, " "]); }
      out.push([key[4]!.trim() ? "string" : "function", key[3]!]);
      value(key[4]!);
      return;
    }
    const item = /^(\s*)- (.*)$/.exec(line);
    if (item) {
      if (item[1]) out.push([undefined, item[1]]);
      out.push(["string", "-"], [undefined, " "]);
      value(item[2]!);
      return;
    }
    value(line);
  });
  return out;
}

function markdown(): Rule[] {
  return [group("\\[([^\\]\\n]+)\\](?=\\()", "string")];
}

function xml(): Rule[] {
  const tag: Rule = (code, at) => {
    const name = /^<(\/?)([A-Za-z][\w:-]*)/.exec(code.slice(at));
    if (!name) return undefined;
    const toks: Tok[] = [[undefined, `<${name[1]}`], ["keyword", name[2]!]];
    let i = at + name[0].length;
    for (;;) {
      const rest = code.slice(i);
      const space = /^\s+/.exec(rest);
      if (space) { toks.push([undefined, space[0]]); i += space[0].length; continue; }
      const close = /^\/?>/.exec(rest);
      if (close) {
        toks.push([undefined, close[0]]);
        i += close[0].length;
        if (!name[1] && name[2] === "script") {
          const end = code.indexOf("</script", i);
          const body = code.slice(i, end < 0 ? code.length : end);
          toks.push(...scan(body, [words(JS_KEYWORDS, "keyword")]));
          i += body.length;
        }
        return [toks, i];
      }
      const attr = /^[\w:-]+/.exec(rest);
      if (!attr) return [toks, i];
      toks.push(["function", attr[0]]);
      i += attr[0].length;
      if (code[i] === "=") {
        toks.push([undefined, "="]);
        i++;
        const quoted = /^(?:"[^"]*"|'[^']*')/.exec(code.slice(i));
        if (quoted) { toks.push(["string", quoted[0]]); i += quoted[0].length; }
      }
    }
  };
  return [re("<!--[\\s\\S]*?-->", "comment"), tag];
}

const SQL_KEYWORDS = "select|from|left|right|inner|outer|full|cross|join|on|where|and|or|not|in|is|like|as|group|by|order|desc|asc|having|union|all|distinct|insert|into|values|update|set|delete|create|table|alter|drop|index|primary|key|null|timestamp|last";

function sql(): Rule[] {
  return [
    lineComment("--"), re("\\/\\*[\\s\\S]*?\\*\\/", "comment"),
    sq,
    words(SQL_KEYWORDS, "keyword", "i"),
    re("\\b(?:count|max|min|sum|avg|coalesce|lower|upper|length|concat|substring|round|abs|now)\\b(?=\\s*\\()", "function", "i"),
  ];
}

function kql(): Rule[] {
  return [
    lineComment("\\/\\/"), dq, sq,
    words("where|and|or|summarize|count|by|order|desc|asc|take|limit|extend|project|join|on|distinct|top|sort|union|let|print|render|parse|mv-expand|bin|between|contains|has|startswith|endswith", "keyword"),
  ];
}

const LANGUAGES: Record<string, { cls: string; tokens: (code: string) => Tok[] }> = {
  javascript: { cls: "javascript", tokens: code => scan(code, javascript(false)) },
  typescript: { cls: "typescript", tokens: code => scan(code, typescript()) },
  python: { cls: "python", tokens: code => scan(code, python()) },
  rust: { cls: "rust", tokens: code => scan(code, rust()) },
  bash: { cls: "bash", tokens: code => scan(code, bash()) },
  powershell: { cls: "powershell", tokens: code => scan(code, powershell()) },
  yaml: { cls: "yaml", tokens: yaml },
  json: { cls: "json", tokens: code => scan(code, jsonRules()) },
  markdown: { cls: "markdown", tokens: code => scan(code, markdown()) },
  xml: { cls: "xml", tokens: code => scan(code, xml()) },
  http: { cls: "http", tokens: code => [[undefined, code]] },
  sql: { cls: "sql", tokens: code => scan(code, sql()) },
  kql: { cls: "kql", tokens: code => scan(code, kql()) },
};

/** Names that people write after a code fence, mapped to the language that Teams names. */
const ALIASES: Record<string, string> = {
  js: "javascript", jsx: "javascript", mjs: "javascript", cjs: "javascript", node: "javascript",
  ts: "typescript", tsx: "javascript",
  py: "python", python3: "python",
  rs: "rust",
  sh: "bash", shell: "bash", zsh: "bash", console: "bash",
  ps: "powershell", ps1: "powershell", pwsh: "powershell",
  yml: "yaml",
  md: "markdown",
  html: "xml", htm: "xml", svg: "xml",
  kusto: "kql",
};

/** The Teams language name for a fence label, or undefined when the CLI has no rules for it. */
export function resolveLanguage(label: string): string | undefined {
  const name = label.trim().toLowerCase();
  const canonical = ALIASES[name] ?? name;
  return LANGUAGES[canonical] ? canonical : undefined;
}

/** Text in a code block. Teams writes a space as a no-break space and a newline as `<br>`. */
function codeText(text: string): string {
  return escapeHtml(text).replace(/ /g, "&nbsp;").replace(/\n/g, "<br>");
}

export function highlightTokens(language: string, code: string): Tok[] {
  return LANGUAGES[language]?.tokens(code) ?? [[undefined, code]];
}

/** The code HTML inside `<code>`, with highlight spans for a language the CLI knows. */
export function highlightCode(language: string | undefined, code: string): string {
  const tokens: Tok[] = language ? highlightTokens(language, code) : [[undefined, code]];
  return tokens.map(([cls, text]) => cls ? `<span class="hljs-${cls}">${codeText(text)}</span>` : codeText(text)).join("");
}

function blockId(language: string, code: string): string {
  const hex = createHash("sha256").update(`${language}\n${code}`).digest("hex");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}`;
}

/**
 * One fenced block in the Teams shape. A known language gets the editor marker, the language label, and spans.
 * An unknown label keeps its own name as the language class, so Teams can show a label for names it knows.
 * No label gives a bare code block. The server adds an `x_` prefix to the marker id, so the id has none.
 */
export function codeBlockHtml(label: string, code: string): string {
  const language = resolveLanguage(label);
  const name = language ?? (/^[a-z][a-z0-9+#.-]{0,30}$/.test(label.trim().toLowerCase()) ? label.trim().toLowerCase() : undefined);
  if (!name) return `<pre><code>${highlightCode(undefined, code)}</code></pre>`;
  const id = `codeBlockEditor-${blockId(name, code)}`;
  return `<p itemtype="http://schema.skype.com/CodeBlockEditor" id="${id}">&nbsp;</p>` +
    `<pre class="language-${name} language-was-manually-selected" itemid="${id}"><code>${highlightCode(language, code)}</code></pre>`;
}
