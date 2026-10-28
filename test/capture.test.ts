import { expect, test } from "bun:test";
import { isSignInUrl, SignInWatch, TokenCapture, type CaptureIO } from "../src/capture.ts";
import type { CdpEvent } from "../src/cdp.ts";
import { AUD, fakeJwt, OID, TID } from "./jwt.ts";
import { TENANT_URL, tokenExchange } from "./fake-cdp.ts";

const FORM = { client_id: "fake-client", grant_type: "refresh_token", refresh_token: "fake-old", scope: "x" };

function harness(bodies: Record<string, unknown>, extra: Partial<CaptureIO> = {}) {
  const reads: string[] = [];
  const io: CaptureIO = {
    getResponseBody: async id => { reads.push(id); const body = bodies[id]; return typeof body === "string" ? { body, base64Encoded: false } : { body: JSON.stringify(body), base64Encoded: false }; },
    getPostData: async () => undefined,
    ...extra,
  };
  return { capture: new TokenCapture(io), reads };
}
const feed = (capture: TokenCapture, events: CdpEvent[]) => events.forEach(event => capture.onEvent(event));

test("each access token is classified by its audience into chat, csa, spaces and substrate", async () => {
  const bodies = Object.fromEntries(Object.entries({ chat: AUD.chat, csa: AUD.csa, spaces: AUD.spaces, substrate: "https://substrate.office.com" })
    .map(([id, aud]) => [id, { access_token: fakeJwt({ aud }) }]));
  const { capture } = harness(bodies);
  for (const id of Object.keys(bodies)) feed(capture, tokenExchange(id, TID, FORM));
  await capture.settled();
  expect(Object.keys(capture.captured.tokens).sort()).toEqual(["chat", "csa", "spaces", "substrate"]);
  expect(capture.captured.claims?.oid).toBe(OID);
  expect(capture.missing(["chat", "csa", "spaces", "substrate"])).toEqual([]);
});

test("an unknown audience, an expiring token, and a non-200 response are not kept", async () => {
  const { capture, reads } = harness({
    graph: { access_token: fakeJwt({ aud: "https://graph.microsoft.com" }) },
    soon: { access_token: fakeJwt({ aud: AUD.chat, exp: Math.floor(Date.now() / 1000) + 20 }) },
    denied: { access_token: fakeJwt({ aud: AUD.csa }) },
  });
  feed(capture, [...tokenExchange("graph", TID, FORM), ...tokenExchange("soon", TID, FORM), ...tokenExchange("denied", TID, FORM, 400)]);
  await capture.settled();
  expect(capture.captured.tokens).toEqual({});
  expect(reads).toEqual(["graph", "soon"]);
});

test("a request to another host or a GET is never read", async () => {
  const { capture, reads } = harness({ a: { access_token: fakeJwt() }, b: { access_token: fakeJwt() } });
  feed(capture, [
    { method: "Network.requestWillBeSent", params: { requestId: "a", request: { url: `https://evil.test/${TID}/oauth2/v2.0/token`, method: "POST" } } },
    { method: "Network.responseReceived", params: { requestId: "a", response: { status: 200 } } },
    { method: "Network.loadingFinished", params: { requestId: "a" } },
    { method: "Network.requestWillBeSent", params: { requestId: "b", request: { url: TENANT_URL(TID), method: "GET" } } },
    { method: "Network.responseReceived", params: { requestId: "b", response: { status: 200 } } },
    { method: "Network.loadingFinished", params: { requestId: "b" } },
  ]);
  await capture.settled();
  expect(reads).toEqual([]);
});

test("the newest refresh token and its client and tenant become the refresh config", async () => {
  const { capture } = harness({
    one: { access_token: fakeJwt({ aud: AUD.chat }), refresh_token: "fake-refresh-1" },
    two: { access_token: fakeJwt({ aud: AUD.csa }), refresh_token: "fake-refresh-2" },
    other: { access_token: fakeJwt({ aud: "https://graph.microsoft.com" }), refresh_token: "fake-refresh-graph" },
  });
  feed(capture, [...tokenExchange("one", TID, FORM), ...tokenExchange("two", TID, { ...FORM, client_id: "fake-client-2" }), ...tokenExchange("other", TID, FORM)]);
  await capture.settled();
  expect(capture.captured.refresh).toEqual({ clientId: "fake-client-2", tenantId: TID, refreshToken: "fake-refresh-2", origin: "https://teams.cloud.microsoft" });
});

test("response bodies are read in the order the loads finished, even when an early read is slow", async () => {
  const bodies: Record<string, unknown> = {
    slow: { access_token: fakeJwt({ aud: AUD.chat }), refresh_token: "fake-refresh-first" },
    fast: { access_token: fakeJwt({ aud: AUD.csa }), refresh_token: "fake-refresh-last" },
  };
  const { capture } = harness(bodies, {
    getResponseBody: async id => { if (id === "slow") await Bun.sleep(40); return { body: JSON.stringify(bodies[id]), base64Encoded: false }; },
  });
  feed(capture, [...tokenExchange("slow", TID, FORM), ...tokenExchange("fast", TID, FORM)]);
  await capture.settled();
  expect(capture.captured.refresh?.refreshToken).toBe("fake-refresh-last");
});

test("post data that the event left out is read separately, and a base64 body is decoded", async () => {
  const { capture } = harness({}, {
    getResponseBody: async () => ({ body: Buffer.from(JSON.stringify({ access_token: fakeJwt(), refresh_token: "fake-refresh-b64" })).toString("base64"), base64Encoded: true }),
    getPostData: async () => new URLSearchParams(FORM).toString(),
  });
  const [sent, received, finished] = tokenExchange("big", TID, FORM);
  sent!.params.request.postData = undefined;
  feed(capture, [sent!, received!, finished!]);
  await capture.settled();
  expect(capture.captured.refresh?.refreshToken).toBe("fake-refresh-b64");
  expect(capture.captured.refresh?.clientId).toBe("fake-client");
});

test("the API region comes from a chat service URL", () => {
  const { capture } = harness({});
  capture.onEvent({ method: "Network.requestWillBeSent", params: { requestId: "r", request: { url: "https://teams.cloud.microsoft/api/chatsvc/emea/v1/x", method: "GET" } } });
  expect(capture.captured.region).toBe("emea");
});

test("a token for another account is ignored and counted, and a body error never carries text", async () => {
  const { capture } = harness({
    mine: { access_token: fakeJwt({ aud: AUD.chat }) },
    theirs: { access_token: fakeJwt({ aud: AUD.csa, oid: "another-user" }) },
    broken: "not json fake-refresh-secret",
  });
  feed(capture, [...tokenExchange("mine", TID, FORM), ...tokenExchange("theirs", TID, FORM), ...tokenExchange("broken", TID, FORM)]);
  await capture.settled();
  expect(Object.keys(capture.captured.tokens)).toEqual(["chat"]);
  expect(capture.otherAccount).toBe(1);
  expect(capture.errors).toBe(1);
});

test("an unreadable response body is counted and later responses still read", async () => {
  const { capture } = harness({ ok: { access_token: fakeJwt() } }, {
    getResponseBody: async id => { if (id === "gone") throw new Error("No resource"); return { body: JSON.stringify({ access_token: fakeJwt() }), base64Encoded: false }; },
  });
  feed(capture, [...tokenExchange("gone", TID, FORM), ...tokenExchange("ok", TID, FORM)]);
  await capture.settled();
  expect(capture.errors).toBe(1);
  expect(capture.captured.tokens.chat).toBeDefined();
});

test("sign-in pages are recognised by host", () => {
  expect(isSignInUrl("https://login.microsoftonline.com/common/oauth2/v2.0/authorize?x=1")).toBe(true);
  expect(isSignInUrl("https://login.live.com/")).toBe(true);
  expect(isSignInUrl("https://teams.cloud.microsoft/v2/")).toBe(false);
  expect(isSignInUrl("about:blank")).toBe(false);
  expect(isSignInUrl("not a url")).toBe(false);
});

test("a pass through the login host is silent, and a stay past the grace period needs a person", () => {
  const watch = new SignInWatch(8_000);
  watch.observe("https://teams.cloud.microsoft/", 0);
  expect(watch.needsSignIn(60_000)).toBe(false);
  watch.observe("https://login.microsoftonline.com/x", 1_000);
  expect(watch.needsSignIn(5_000)).toBe(false);
  expect(watch.onSignInPage).toBe(true);
  watch.observe("https://login.microsoftonline.com/x?step=2", 6_000);
  expect(watch.needsSignIn(9_000)).toBe(true);
  watch.observe("https://teams.cloud.microsoft/v2/", 10_000);
  expect(watch.needsSignIn(60_000)).toBe(false);
  expect(watch.onSignInPage).toBe(false);
});
