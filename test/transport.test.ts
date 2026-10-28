import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { networkCause, request, requestHeaders, serviceFailure, setSleepForTests, TeamsError, utcOffset, mapConcurrent } from "../src/transport.ts";
import { fakeJwt, AUD } from "./jwt.ts";
import { installFakeTeams, seedProfile, type Route } from "./fake-teams.ts";

const CHAT_PATH = "/api/chatsvc/amer/v1/users/ME/properties";
const CSA_PATH = "/api/csa/amer/api/v3/teams/users/me";
let fake: ReturnType<typeof installFakeTeams> | undefined;

function serve(routes: Record<string, Route>) {
  fake = installFakeTeams(routes);
  return fake;
}

beforeEach(async () => {
  setSleepForTests(async () => undefined);
  await seedProfile();
});

afterEach(() => {
  fake?.uninstall();
  fake = undefined;
  setSleepForTests(ms => Bun.sleep(ms));
});

describe("headers", () => {
  const now = new Date("2026-10-14T12:00:00Z");

  test("every service sends the common headers", () => {
    for (const service of ["chat", "csa", "spaces"] as const) {
      const headers = requestHeaders(service, "tok", { region: "amer", hasBody: false }, now);
      expect(headers.authorization).toBe("Bearer tok");
      expect(headers.behavioroverride).toBe("redirectAs404");
      expect(headers["x-ms-migration"]).toBe("True");
      expect(headers["x-ms-test-user"]).toBe("False");
      expect(headers.clientinfo).toContain("clientName=skypeteams");
      expect(headers.clientinfo).toContain("clientVer=1415/26091712222");
      expect(headers["user-agent"]).toContain("Chrome/154.0.0.0");
    }
  });

  test("csa adds its own headers and chat does not", () => {
    const csa = requestHeaders("csa", "tok", { region: "amer", hasBody: false }, now);
    expect(csa["x-ms-region"]).toBe("amer");
    expect(csa["x-ringoverride"]).toBe("general");
    expect(csa["x-ms-user-type"]).toBe("real-user");
    const chat = requestHeaders("chat", "tok", { region: "amer", hasBody: false }, now);
    expect(chat["x-ms-region"]).toBeUndefined();
  });

  test("spaces adds its own headers", () => {
    const spaces = requestHeaders("spaces", "tok", { region: "amer", hasBody: false }, now);
    expect(spaces["x-ms-client-type"]).toBe("cdlworker");
    expect(spaces.accept).toBe("application/json");
  });

  test("content-type is set only when there is a body", () => {
    expect(requestHeaders("chat", "tok", { region: "amer", hasBody: false }, now)["content-type"]).toBeUndefined();
    expect(requestHeaders("chat", "tok", { region: "amer", hasBody: true }, now)["content-type"]).toBe("application/json");
  });

  test("TEAMS_CLIENT_VER overrides the client version", () => {
    process.env.TEAMS_CLIENT_VER = "9999/1";
    try {
      expect(requestHeaders("csa", "tok", { region: "amer", hasBody: false }, now)["x-ms-client-version"]).toBe("9999/1");
    } finally {
      delete process.env.TEAMS_CLIENT_VER;
    }
  });

  test("the UTC offset has a sign and two-digit parts", () => {
    expect(utcOffset(new Date())).toMatch(/^[+-]\d{2}:\d{2}$/);
  });
});

describe("request", () => {
  test("returns parsed JSON and sends the chat token to the chat service", async () => {
    serve({ [`GET ${CHAT_PATH}`]: { status: 200, json: { online: true } } });
    expect(await request("chat", "GET", "/users/ME/properties")).toEqual({ online: true });
    const sent = fake!.requests[0]!;
    expect(sent.path).toBe(CHAT_PATH);
    expect(sent.headers.authorization?.startsWith("Bearer ")).toBe(true);
  });

  test("an empty body returns undefined", async () => {
    serve({ [`DELETE ${CHAT_PATH}`]: { status: 204 } });
    expect(await request("chat", "DELETE", "/users/ME/properties")).toBeUndefined();
  });

  test("a 429 is retried after Retry-After and then succeeds", async () => {
    const waits: number[] = [];
    setSleepForTests(async ms => { waits.push(ms); });
    serve({
      [`GET ${CHAT_PATH}`]: [
        { status: 429, headers: { "retry-after": "2" }, text: "slow down" },
        { status: 200, json: { ok: 1 } },
      ],
    });
    expect(await request("chat", "GET", "/users/ME/properties")).toEqual({ ok: 1 });
    expect(fake!.requests.length).toBe(2);
    expect(waits).toEqual([2000]);
  });

  test("the wait line says request for a change, and send only for a send", async () => {
    const lines: string[] = [];
    const originalError = console.error;
    console.error = (...args: unknown[]) => { lines.push(args.map(String).join(" ")); };
    try {
      setSleepForTests(async () => undefined);
      serve({
        [`DELETE ${CHAT_PATH}`]: [{ status: 429, headers: { "retry-after": "2" }, text: "slow" }, { status: 200, text: "null" }],
        [`POST ${CHAT_PATH}`]: [{ status: 429, headers: { "retry-after": "3" }, text: "slow" }, { status: 201, json: {} }],
      });
      await request("chat", "DELETE", "/users/ME/properties");
      await request("chat", "POST", "/users/ME/properties", { body: {}, send: true });
    } finally {
      console.error = originalError;
    }
    expect(lines).toEqual(["Rate limit: waiting 2s before the next request.", "Rate limit: waiting 3s before the next send."]);
  });

  test("a 429 is retried three times, then fails and is not uncertain", async () => {
    serve({ [`GET ${CHAT_PATH}`]: { status: 429, headers: { "retry-after": "1" }, text: "slow" } });
    const error = await request("chat", "GET", "/users/ME/properties").catch(e => e as TeamsError);
    expect(error).toBeInstanceOf(TeamsError);
    expect(error.status).toBe(429);
    expect(error.uncertain).toBe(false);
    expect(error.message).toBe("Teams refused the request: 20 per minute quota. Try again in 1s.");
    expect(fake!.requests.length).toBe(4);
  });

  test("a Retry-After above 120 seconds is refused at once", async () => {
    serve({ [`GET ${CHAT_PATH}`]: { status: 429, headers: { "retry-after": "300" }, text: "slow" } });
    const error = await request("chat", "GET", "/users/ME/properties").catch(e => e as TeamsError);
    expect(error.message).toBe("Teams refused the request: 20 per minute quota. Try again in 300s.");
    expect(fake!.requests.length).toBe(1);
  });

  test("a 5xx is not retried and is uncertain", async () => {
    serve({ [`GET ${CHAT_PATH}`]: { status: 503, text: "busy" } });
    const error = await request("chat", "GET", "/users/ME/properties").catch(e => e as TeamsError);
    expect(error.status).toBe(503);
    expect(error.uncertain).toBe(true);
    // The raw body is never shown, so a text body adds nothing.
    expect(error.message).toBe("GET /users/ME/properties failed: 503");
    expect(fake!.requests.length).toBe(1);
  });

  test("a 408 is uncertain, a 400 is not", async () => {
    serve({ [`GET ${CHAT_PATH}`]: { status: 408 } });
    expect((await request("chat", "GET", "/users/ME/properties").catch(e => e as TeamsError)).uncertain).toBe(true);
    fake!.uninstall();
    serve({ [`GET ${CHAT_PATH}`]: { status: 400, json: { errorCode: 400, message: "bad input" } } });
    const error = await request("chat", "GET", "/users/ME/properties").catch(e => e as TeamsError);
    expect(error.uncertain).toBe(false);
    expect(error.message).toBe("GET /users/ME/properties failed: 400 bad input");
  });

  test("a network error is uncertain and names no host", async () => {
    serve({ [`POST ${CHAT_PATH}`]: "network" });
    const error = await request("chat", "POST", "/users/ME/properties", { body: {} }).catch(e => e as TeamsError);
    expect(error.uncertain).toBe(true);
    expect(error.status).toBeUndefined();
    expect(error.message).toContain("network error");
  });

  test("a 401 with a future exp says the token was rejected, not expired", async () => {
    serve({ [`GET ${CHAT_PATH}`]: { status: 401, json: { errorCode: 401, message: "Unauthorized" } } });
    await seedProfile(); // fresh token
    const error = await request("chat", "GET", "/users/ME/properties").catch(e => e as Error);
    expect(error.message).toBe("Teams rejected the chat token (401). It may be revoked, signed out, or for another service. Supply a new one: teams-cli auth");
  });

  test("a 401 with an exp in the past says it expired", async () => {
    serve({ [`GET ${CHAT_PATH}`]: { status: 401, json: { errorCode: 401, message: "Unauthorized" } } });
    await seedProfile(["chat"], Math.floor(Date.now() / 1000) - 3600);
    const error = await request("chat", "GET", "/users/ME/properties").catch(e => e as Error);
    expect(error.message).toMatch(/^The chat token expired at .+\. Supply a new one: teams-cli auth$/);
  });

  test("a subCode maps to a plain sentence, with the status", () => {
    expect(serviceFailure("GET", "/x", 404, JSON.stringify({ subCode: "LocationLookupFailed", details: "Location lookup failed for thread 19:x" }), "amer"))
      .toBe("Conversation not found (404).");
    expect(serviceFailure("GET", "/x", 404, JSON.stringify({ subCode: "SoftDeleted" }), "amer")).toBe("This conversation was deleted (404).");
    expect(serviceFailure("POST", "/x", 403, JSON.stringify({ subCode: "AclCheckFailed" }), "amer"))
      .toBe("You are not a member of this conversation, or the chat token lacks access (403).");
    expect(serviceFailure("POST", "/x", 400, JSON.stringify({ subCode: "MessageSizeExceeded" }), "amer"))
      .toBe("Teams refused the message: it is too large (400).");
  });

  test("an unknown error shows the message field cut to 200 characters, never the raw body", () => {
    const long = "x".repeat(300);
    expect(serviceFailure("GET", "/x", 500, JSON.stringify({ message: long }), "amer")).toBe(`GET /x failed: 500 ${"x".repeat(200)}`);
    expect(serviceFailure("GET", "/x", 500, "<html>secret body</html>", "amer")).toBe("GET /x failed: 500");
  });

  test("a 400 with no detail names TEAMS_REGION when it is set", () => {
    process.env.TEAMS_REGION = "zzzz";
    try {
      expect(serviceFailure("GET", "/x", 400, "", "zzzz")).toBe("Teams rejected the request (400). Check TEAMS_REGION: zzzz.");
    } finally {
      delete process.env.TEAMS_REGION;
    }
    expect(serviceFailure("GET", "/x", 400, "", "amer")).toBe("GET /x failed: 400");
  });

  test("a network error names its cause", () => {
    expect(networkCause({ code: "ECONNREFUSED" })).toBe("ECONNREFUSED");
    expect(networkCause({ cause: { code: "ENOTFOUND" } })).toBe("ENOTFOUND");
    expect(networkCause({ name: "TimeoutError" })).toBe("timeout");
    expect(networkCause({ code: "ConnectionRefused" })).toBe("ECONNREFUSED");
    expect(networkCause(new TypeError("fetch failed"))).toBeUndefined();
  });

  test("a token of the wrong kind gets 401 from the fake", async () => {
    await seedProfile(["chat"]);
    fake = installFakeTeams({ [`GET ${CSA_PATH}`]: { status: 200, json: {} } });
    // The stored csa token is missing, so the request never leaves the process.
    await expect(request("csa", "GET", "/v3/teams/users/me")).rejects.toThrow("No csa token. Run: teams-cli auth");
    expect(fake.requests.length).toBe(0);
  });

  test("a token value in an error body is redacted", async () => {
    const leaked = fakeJwt({ aud: AUD.chat });
    serve({ [`GET ${CHAT_PATH}`]: { status: 500, json: { errorCode: 500, message: `oops ${leaked}` } } });
    const error = await request("chat", "GET", "/users/ME/properties").catch(e => e as Error);
    expect(error.message).not.toContain(leaked);
    expect(error.message).toContain("<token>");
  });

  test("a request to an unexpected host is blocked", async () => {
    serve({ [`GET ${CHAT_PATH}`]: { status: 200, json: {} } });
    process.env.TEAMS_HOST = "http://127.0.0.1:9";
    try {
      await expect(request("chat", "GET", "/users/ME/properties")).rejects.toThrow("network error");
    } finally {
      delete process.env.TEAMS_HOST;
    }
  });

  test("TEAMS_HOST accepts only localhost or 127.0.0.1", async () => {
    process.env.TEAMS_HOST = "https://example.test";
    try {
      await expect(request("chat", "GET", "/users/ME/properties")).rejects.toThrow("TEAMS_HOST may only be");
    } finally {
      delete process.env.TEAMS_HOST;
    }
  });
});

describe("mapConcurrent", () => {
  test("keeps input order and respects the limit", async () => {
    let active = 0;
    let peak = 0;
    const results = await mapConcurrent([1, 2, 3, 4, 5, 6], 4, async n => {
      active++;
      peak = Math.max(peak, active);
      await Bun.sleep(1);
      active--;
      return n * 10;
    });
    expect(results).toEqual([10, 20, 30, 40, 50, 60]);
    expect(peak).toBeLessThanOrEqual(4);
  });
});
