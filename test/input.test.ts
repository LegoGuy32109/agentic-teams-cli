import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { splitTerminator } from "../src/args.ts";
import { send } from "../src/commands/send.ts";
import { ledgerKey, LedgerRefusal, reserve } from "../src/ledger.ts";
import { mentionTokens } from "../src/users.ts";
import { installFakeTeams, resetLocalState, seedProfile } from "./fake-teams.ts";
import { TID } from "./jwt.ts";

const plain = { json: false, verify: false, allowWrite: false } as const;
let fake: ReturnType<typeof installFakeTeams> | undefined;
const originalLog = console.log;

beforeEach(async () => {
  await resetLocalState();
  await seedProfile();
  console.log = () => {};
});

afterEach(() => {
  console.log = originalLog;
  fake?.uninstall();
  fake = undefined;
});

describe("the -- terminator", () => {
  test("text after -- is literal, even when it starts with --", () => {
    expect(splitTerminator(["me", "--allow-write", "--", "--allow-write", "--json"])).toEqual({
      flagArgs: ["me", "--allow-write"],
      literalArgs: ["--allow-write", "--json"],
    });
  });

  test("without --, every argument is a flag or positional", () => {
    expect(splitTerminator(["me", "hi", "--json"])).toEqual({ flagArgs: ["me", "hi", "--json"], literalArgs: [] });
  });
});

describe("empty messages", () => {
  test("spaces, tabs, newlines, and no-break spaces are empty", async () => {
    for (const text of ["   ", "\t", "\n", "  "]) {
      await expect(send("me", [text], plain)).rejects.toThrow("The message is empty.");
    }
  });

  test("an emoji alone is a message", async () => {
    await send("me", ["🙂"], plain);
  });

  test("empty --html is refused", async () => {
    await expect(send("me", [], { ...plain, html: "  " })).rejects.toThrow("The message is empty.");
  });
});

describe("mention syntax", () => {
  test("an empty mention is refused before any lookup", async () => {
    fake = installFakeTeams({});
    await expect(send("19:dev@thread.tacv2", ["hi @{}"], plain)).rejects.toThrow("Empty mention. Write @{Name}.");
    await expect(send("19:dev@thread.tacv2", ["hi @{  }"], plain)).rejects.toThrow("Empty mention. Write @{Name}.");
    expect(fake.requests).toHaveLength(0);
  });

  test("an unclosed mention names the first 20 characters", () => {
    expect(() => mentionTokens("ping @{Alexandra Doe and more words here")).toThrow(
      "Mention is not closed: \"@{Alexandra Doe and mo\". Write @{Name}.",
    );
  });

  test("nested braces are refused", () => {
    expect(() => mentionTokens("ping @{Al{ex}}")).toThrow("Mention names cannot contain braces.");
  });

  test("plain braces and a closed mention parse", () => {
    expect(mentionTokens("a {b} @{ Alex Doe } c")).toEqual([{ start: 6, end: 19, query: "Alex Doe" }]);
  });
});

describe("ledger wording", () => {
  const key = ledgerKey({ conversationId: "48:notes", content: "<p>x</p>", mentionMris: [], contentType: "Text" });
  const context = { label: "Notes to self", conversationId: "48:notes", input: "me" };
  const refusalOf = (work: Promise<void>) => work.then(() => undefined, (error: LedgerRefusal) => error);

  test("a reservation younger than 30 seconds says a send is in progress", async () => {
    const now = Date.now();
    await reserve(TID, key, context, now);
    const refusal = await refusalOf(reserve(TID, key, context, now + 5_000));
    expect(refusal?.message).toBe("Another identical send is in progress. Not sending again.");
  });

  test("an older reservation keeps the outcome-unknown wording", async () => {
    const now = Date.now();
    await reserve(TID, key, context, now);
    const refusal = await refusalOf(reserve(TID, key, context, now + 60_000));
    expect(refusal?.message).toContain("A previous identical send may have posted (outcome unknown).");
  });
});
