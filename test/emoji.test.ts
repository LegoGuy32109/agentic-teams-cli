import { describe, expect, test } from "bun:test";
import { composeHtml, composePlain, composeRich, extractFences } from "../src/compose.ts";
import { emojiByName, emojiByUnicode, emojiHtml, emojifyText, findEmoji, searchEmoji, withEmojiStyle } from "../src/emoji.ts";
import { composeMessage } from "../src/commands/send.ts";
import { htmlToText } from "../src/message.ts";

// The two emoji the web client sent in one message on 2026-10-19 (capture 08, request 3652).
const LAUNCH = `<span contenteditable="false" title="Rocket launch" type="(launch)" class="animated-emoticon-20-launch" itemscope><img itemscope itemtype="http://schema.skype.com/Emoji" itemid="launch" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/launch/default/20_f.png" title="Rocket launch" alt="🚀" style="width:20px;height:20px;"></span>`;
const YES_TONE3 = `<span contenteditable="false" title="Yes with medium skin tone" type="(yes-tone3)" class="animated-emoticon-20-yes-tone3" itemscope><img itemscope itemtype="http://schema.skype.com/Emoji" itemid="yes-tone3" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/yes/default/20_f_s4.png" title="Yes with medium skin tone" alt="👍🏽" style="width:20px;height:20px;"></span>`;

describe("emoji catalog", () => {
  test("finds an emoji by character, with and without the variation selector", () => {
    expect(emojiByUnicode("🚀")?.id).toBe("launch");
    expect(emojiByUnicode("👍")?.id).toBe("like");
    expect(emojiByUnicode("❤️")?.id).toBe(emojiByUnicode("❤")?.id);
  });

  test("a skin tone takes the first entry with tones, as the web client did", () => {
    const toned = emojiByUnicode("👍🏽")!;
    expect(toned.id).toBe("yes-tone3");
    expect(toned.unicode).toBe("👍🏽");
    expect(toned.title).toBe("Yes with medium skin tone");
  });

  test("finds an emoji by key or shortcut in each written form", () => {
    for (const form of ["launch", "(launch)", ":launch:", "LAUNCH"]) expect(emojiByName(form)?.id).toBe("launch");
    expect(emojiByName("yes-tone3")?.unicode).toBe("👍🏽");
    expect(emojiByName("nosuchemoji")).toBeUndefined();
  });

  test("an emoji with no skin tones refuses a tone", () => {
    expect(emojiByName("launch-tone2")).toBeUndefined();
    expect(emojiByUnicode("🚀🏽")).toBeUndefined();
  });

  test("search matches names and keywords", () => {
    expect(searchEmoji("rocket").map(entry => entry.id)).toContain("launch");
    expect(searchEmoji("rocket")[0]!.id).toBe("launch");
    expect(searchEmoji("")).toEqual([]);
    expect(findEmoji(" 🚀 ")?.id).toBe("launch");
  });
});

describe("emoji markup", () => {
  test("is byte-identical to the captured markup", () => {
    expect(emojiHtml(emojiByUnicode("🚀")!)).toBe(LAUNCH);
    expect(emojiHtml(emojiByUnicode("👍🏽")!)).toBe(YES_TONE3);
  });

  test("plain text turns emoji and :key: into markup and escapes the rest", () => {
    expect(emojifyText("a < b 🚀👍🏽")).toBe(`a &lt; b ${LAUNCH}${YES_TONE3}`);
    expect(emojifyText("go :launch:")).toBe(`go ${LAUNCH}`);
  });

  test("leaves a :name: that is not in the catalog, times, and URLs as text", () => {
    expect(emojifyText(":nosuchemoji:")).toBe(":nosuchemoji:");
    expect(emojifyText("at 10:30:45")).toBe("at 10:30:45");
    expect(emojifyText("a:launch:b")).toBe("a:launch:b");
    expect(emojifyText("x:launch:")).toBe("x:launch:");
  });

  test("the plain and rich formats use the markup, and code stays as written", () => {
    expect(composePlain("emoji test 🚀", []).html).toBe(`<p>emoji test ${LAUNCH}</p>`);
    const fenced = extractFences("**big 🚀** and `🚀`\n```\n🚀\n```");
    const rich = composeRich(fenced.text, [], fenced.blocks).html;
    expect(rich).toContain(`<strong>big ${LAUNCH}</strong>`);
    expect(rich).toContain("<code>🚀</code>");
    expect(rich.split("itemtype=\"http://schema.skype.com/Emoji\"").length - 1).toBe(1);
  });

  test("--html text gets the markup, and code and attributes do not", async () => {
    const { content } = await composeHtml(`<p title="🚀">go 🚀</p><pre><code>🚀</code></pre>`);
    expect(content).toBe(`<p title="🚀">go ${LAUNCH}</p><pre><code>🚀</code></pre>`);
  });

  test("history and --verify read the markup back as the character", () => {
    expect(htmlToText(`<p>emoji test ${LAUNCH}${YES_TONE3}</p>`)).toBe("emoji test 🚀👍🏽");
  });

  test("the plain style keeps characters, turns :key: into its character, and restores the default", async () => {
    const plain = await composeMessage("hi ❤️ :launch: 👍🏽", { emoji: "plain" });
    expect(plain.content).toBe("<p>hi ❤️ 🚀 👍🏽</p>");
    expect(await withEmojiStyle("plain", async () => emojifyText("a&b 🚀"))).toBe("a&amp;b 🚀");
    // A character keeps its exact bytes: no variation selector is added.
    expect(await withEmojiStyle("plain", async () => emojifyText("\u2764"))).toBe("\u2764");
    expect(emojifyText("🚀")).toBe(LAUNCH);
    const teams = await composeMessage("hi 🚀", {});
    expect(teams.content).toBe(`<p>hi ${LAUNCH}</p>`);
  });
});

describe("preferred skin tone", () => {
  test("applies to emoji that take tones, and an explicit tone or a toneless key still wins", async () => {
    const { setPreferredTone } = await import("../src/emoji.ts");
    const { resolveReactionKey } = await import("../src/reactions.ts");
    setPreferredTone(1);
    try {
      expect(emojiByUnicode("👍")?.id).toBe("yes-tone1");
      expect(emojiByUnicode("👍🏽")?.id).toBe("yes-tone3");
      expect(emojiByUnicode("🚀")?.id).toBe("launch");
      expect(emojiByName("1f44b_wavinghand")?.unicode).toBe("👋🏻");
      expect(resolveReactionKey("👍")).toBe("yes-tone1");
      expect(resolveReactionKey("like")).toBe("like");
      expect(resolveReactionKey("❤️")).toBe("heart");
      expect(emojifyText("👋")).toContain('itemid="1f44b_wavinghand-tone1"');
    } finally {
      setPreferredTone(undefined);
    }
    expect(emojiByUnicode("👍")?.id).toBe("like");
  });

  test("tone words parse, including white, and other words are refused", async () => {
    const { parseSkinTone } = await import("../src/settings.ts");
    expect(["white", "light", "1", "🏻", "tone1"].map(parseSkinTone)).toEqual([1, 1, 1, 1, 1]);
    expect(parseSkinTone("Medium-Dark")).toBe(4);
    expect(parseSkinTone("none")).toBe(0);
    expect(() => parseSkinTone("purple")).toThrow("is not a skin tone");
  });
});
