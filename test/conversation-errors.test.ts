import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { history } from "../src/commands/history.ts";
import { send } from "../src/commands/send.ts";
import { editMessage } from "../src/commands/edit.ts";
import { deleteMessages } from "../src/commands/delete.ts";
import { membersOf } from "../src/commands/members.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Reply } from "./fake-teams.ts";

const CONV = "19:bogus@thread.v2";
const BASE = `/api/chatsvc/amer/v1/users/ME/conversations/${CONV}`;
const MSG = "1791990424088";
const DELETE_ID = "1791990424001";
const plain = { json: false, allowWrite: false } as const;
const write = { json: false, verify: false, allowWrite: true } as const;
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
let errors: string[] = [];
const originalLog = console.log;
const originalError = console.error;

/** The two shapes the service uses: the code at the top, or nested as a JSON string inside `message`. */
const BODIES: Array<[string, string, string]> = [
  ["LocationLookupFailed", JSON.stringify({ subCode: "LocationLookupFailed", details: "Location lookup failed for thread 19:bogus" }), "Conversation not found (404)."],
  ["SoftDeleted", JSON.stringify({ subCode: "SoftDeleted" }), "This conversation was deleted (404)."],
  ["nested LocationLookupFailed", JSON.stringify({ message: JSON.stringify({ subCode: "LocationLookupFailed" }) }), "Conversation not found (404)."],
  ["nested SoftDeleted", JSON.stringify({ details: JSON.stringify({ subCode: "SoftDeleted", message: "gone" }) }), "This conversation was deleted (404)."],
];

function serve(routes: Record<string, Reply>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
}

beforeEach(async () => {
  await resetLocalState();
  await seedProfile();
  logs = [];
  errors = [];
  console.log = (...args: unknown[]) => { logs.push(args.map(String).join(" ")); };
  console.error = (...args: unknown[]) => { errors.push(args.map(String).join(" ")); };
});

afterEach(() => {
  console.log = originalLog;
  console.error = originalError;
  fake?.uninstall();
  fake = undefined;
  process.exitCode = 0;
});

const notFound = (text: string): Reply => ({ status: 404, text });

/** The error a command throws. It fails the test if the command succeeds. */
async function failureOf(run: () => Promise<unknown>): Promise<Error> {
  try {
    await run();
  } catch (error) {
    return error as Error;
  }
  throw new Error("The command was expected to fail.");
}

describe("conversation errors in plain words", () => {
  for (const [name, body, sentence] of BODIES) {
    test(`history shows "${sentence}" for ${name}`, async () => {
      serve({ [`GET ${BASE}/messages`]: notFound(body) });
      const error = await failureOf(() => history(CONV, { json: false, content: false, top: 5 }));
      expect(error.message).toBe(sentence);
    });

    test(`send shows "${sentence}" for ${name}`, async () => {
      serve({ [`POST ${BASE}/messages`]: notFound(body) });
      const error = await failureOf(() => send(CONV, ["hi"], write));
      expect(error.message).toBe(sentence);
    });

    test(`edit shows "${sentence}" for ${name}`, async () => {
      serve({ [`GET ${BASE}/messages/${MSG}`]: notFound(body) });
      const error = await failureOf(() => editMessage(CONV, `${CONV}#${MSG}`, ["new"], write));
      expect(error.message).toBe(sentence);
    });

    test(`delete preview shows "${sentence}" for ${name}`, async () => {
      serve({ [`GET ${BASE}/messages/${DELETE_ID}`]: notFound(body) });
      await deleteMessages(CONV, [DELETE_ID], plain);
      expect(logs.join("\n")).toContain(`unreadable: ${sentence}`);
      expect(logs.join("\n")).not.toContain("{");
    });

    test(`delete shows "${sentence}" for ${name} when it writes`, async () => {
      serve({ [`GET ${BASE}/messages/${DELETE_ID}`]: notFound(body) });
      await deleteMessages(CONV, [DELETE_ID], write);
      expect(logs.join("\n")).toContain(`failed ${DELETE_ID}: ${sentence}`);
      expect(logs.join("\n")).not.toContain("{");
    });

    test(`members shows "${sentence}" for ${name}`, async () => {
      serve({ [`GET /api/chatsvc/amer/v1/threads/${CONV}`]: notFound(body) });
      const error = await failureOf(() => membersOf(CONV, false));
      expect(error.message).toBe(sentence);
    });
  }
});
