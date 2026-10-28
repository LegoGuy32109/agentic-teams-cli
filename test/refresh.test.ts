import { afterEach, beforeEach, expect, test } from "bun:test";
import { stat, readFile } from "node:fs/promises";
import { join } from "node:path";
import { loadProfile, saveProfile, credentialsPath } from "../src/config.ts";
import { requireToken, resetTokenCache } from "../src/tokens.ts";
import { parseRefreshSetup, refreshToken, setupRefresh } from "../src/refresh.ts";
import { request } from "../src/transport.ts";
import { resetLocalState, seedProfile } from "./fake-teams.ts";
import { AUD, TID, fakeJwt } from "./jwt.ts";

let original: typeof fetch;
beforeEach(async () => {
  original = globalThis.fetch;
  await resetLocalState();
  await seedProfile(["chat"], Math.floor(Date.now() / 1000) - 60);
  await setupRefresh(JSON.stringify({ clientId: "fake-client", tenantId: TID, refreshToken: "fake-refresh-secret", origin: "https://teams.cloud.microsoft" }));
});
afterEach(() => { globalThis.fetch = original; resetTokenCache(); });

function fakeRenew(work: (form: URLSearchParams) => Response | Promise<Response>) {
  globalThis.fetch = (async (input: any, init: RequestInit) => {
    expect(new URL(String(input)).host).toBe("login.microsoftonline.com");
    expect(init.redirect).toBe("error");
    return work(new URLSearchParams(String(init.body)));
  }) as typeof fetch;
}

test("expired stored access automatically renews and persists rotation with private permissions", async () => {
  fakeRenew(form => {
    expect(form.get("refresh_token")).toBe("fake-refresh-secret");
    expect(form.get("scope")).toContain("ic3.teams.office.com");
    return Response.json({ access_token: fakeJwt(), refresh_token: "fake-rotated-secret" });
  });
  expect((await requireToken("chat")).exp * 1000).toBeGreaterThan(Date.now());
  expect((await loadProfile()).profile?.refresh?.refreshToken).toBe("fake-rotated-secret");
  expect((await stat(credentialsPath())).mode & 0o777).toBe(0o600);
});

test("simultaneous renewals share the rotated profile and only issue one request", async () => {
  let calls = 0;
  fakeRenew(async () => { calls++; await Bun.sleep(20); return Response.json({ access_token: fakeJwt(), refresh_token: "fake-rotated" }); });
  await Promise.all([refreshToken("chat"), refreshToken("chat")]);
  expect(calls).toBe(1);
});

test("separate CLI processes serialize renewal and preserve rotation", async () => {
  const calls = `${credentialsPath()}.fake-calls`;
  const workers = [0, 1].map(() => Bun.spawn([process.execPath, join(import.meta.dir, "refresh-worker.ts"), calls], {
    env: { ...process.env }, stdout: "pipe", stderr: "pipe",
  }));
  expect(await Promise.all(workers.map(worker => worker.exited))).toEqual([0, 0]);
  expect((await readFile(calls, "utf8")).trim().split("\n")).toHaveLength(1);
  expect((await loadProfile()).profile?.refresh?.refreshToken).toBe("fake-rotated-by-worker");
});

test("wrong resource token is refused but refresh rotation survives", async () => {
  fakeRenew(() => Response.json({ access_token: fakeJwt({ aud: AUD.spaces }), refresh_token: "fake-rotated" }));
  await expect(refreshToken("chat")).rejects.toThrow("did not return a usable chat token");
  expect((await loadProfile()).profile?.refresh?.refreshToken).toBe("fake-rotated");
});

test("another account is refused", async () => {
  fakeRenew(() => Response.json({ access_token: fakeJwt({ oid: "another-user" }) }));
  await expect(refreshToken("chat")).rejects.toThrow("another account");
});

test("sign-in-required errors do not echo Microsoft descriptions or credentials", async () => {
  fakeRenew(() => Response.json({ error: "invalid_grant", error_description: "fake-refresh-secret" }, { status: 400 }));
  const error = await refreshToken("chat").then(() => { throw new Error("Expected renewal failure"); }, e => e as Error);
  expect(error.message).toContain("new sign-in");
  expect(error.message).not.toContain("fake-refresh-secret");
});

test("pasted token cURL imports form data without shell execution", () => {
  const config = parseRefreshSetup(`curl --url 'https://login.microsoftonline.com/${TID}/oauth2/v2.0/token' --data-raw 'client_id=fake-client&grant_type=refresh_token&refresh_token=fake%2Bsecret'`);
  expect(config.refreshToken).toBe("fake+secret");
  expect(config.origin).toBe("https://teams.cloud.microsoft");
  expect(() => parseRefreshSetup('{"clientId":"x","tenantId":"../../evil","refreshToken":"secret"}')).toThrow();
});

test("a service 401 renews once, while a second 401 stops", async () => {
  const {name,profile} = await loadProfile();
  profile!.tokens.chat = { token: fakeJwt(), exp: Math.floor(Date.now()/1000)+3600 };
  await saveProfile(name,profile!);
  let renewals=0, reads=0;
  globalThis.fetch = (async (input: any) => {
    if(new URL(String(input)).host === "login.microsoftonline.com") {
      renewals++; return Response.json({access_token:fakeJwt()});
    }
    reads++; return Response.json({}, {status:401});
  }) as typeof fetch;
  await expect(request("chat","GET","/users/ME/properties")).rejects.toThrow("rejected");
  expect(renewals).toBe(1);
  expect(reads).toBe(2);
});
