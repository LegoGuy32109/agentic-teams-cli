import { afterEach, expect, test } from "bun:test";
import { CdpClient, parseCdpMessage, type CdpEvent } from "../src/cdp.ts";
import { startFakeCdp } from "./fake-cdp.ts";

let fake: ReturnType<typeof startFakeCdp> | undefined;
afterEach(() => { fake?.stop(); fake = undefined; });

test("a frame is read as a response or an event, and junk is ignored", () => {
  expect(parseCdpMessage('{"id":3,"result":{"a":1}}')).toEqual({ kind: "response", id: 3, result: { a: 1 }, error: undefined });
  expect(parseCdpMessage('{"id":4,"error":{"code":-1,"message":"no"}}')).toMatchObject({ kind: "response", id: 4, error: { message: "no" } });
  expect(parseCdpMessage('{"method":"Network.loadingFinished","params":{"requestId":"r"},"sessionId":"S"}')).toEqual({
    kind: "event", event: { method: "Network.loadingFinished", params: { requestId: "r" }, sessionId: "S" },
  });
  expect(parseCdpMessage('{"method":"X"}')).toEqual({ kind: "event", event: { method: "X", params: {}, sessionId: undefined } });
  for (const junk of ["not json", "[]", "null", '{"nothing":1}', new ArrayBuffer(2)]) expect(parseCdpMessage(junk)).toBeUndefined();
});

test("a command resolves with its own response, and the session id travels with it", async () => {
  fake = startFakeCdp();
  const cdp = await CdpClient.connect(fake.wsUrl);
  const [targets, attached] = await Promise.all([cdp.send("Target.getTargets"), cdp.send("Target.attachToTarget", { targetId: "T1", flatten: true })]);
  expect(targets.targetInfos).toHaveLength(1);
  expect(attached.sessionId).toBe("S1");
  await cdp.send("Network.enable", { maxPostDataSize: 5 }, "S1");
  expect(fake.commands.at(-1)).toEqual({ method: "Network.enable", params: { maxPostDataSize: 5 }, sessionId: "S1" });
  cdp.close();
});

test("an error response rejects with the redacted message", async () => {
  fake = startFakeCdp();
  const cdp = await CdpClient.connect(fake.wsUrl);
  const error = await cdp.send("Network.getResponseBody", { requestId: "none" }).then(() => undefined, e => e as Error);
  expect(error?.message).toContain("Network.getResponseBody failed");
  cdp.close();
});

test("events reach a listener until it is removed", async () => {
  fake = startFakeCdp({ onNavigate: () => [{ method: "Page.loadEventFired", params: { n: 1 } }, { method: "Page.loadEventFired", params: { n: 2 }, delayMs: 60 }] });
  const cdp = await CdpClient.connect(fake.wsUrl);
  const seen: CdpEvent[] = [];
  const off = cdp.on(event => { seen.push(event); off(); });
  await cdp.send("Page.navigate", { url: "https://x.test/" }, "S1");
  await Bun.sleep(150);
  expect(seen.map(event => event.params.n)).toEqual([1]);
  cdp.close();
});

test("a closed connection rejects pending commands and resolves `closed`", async () => {
  fake = startFakeCdp();
  const cdp = await CdpClient.connect(fake.wsUrl);
  fake.dropConnections();
  await cdp.closed;
  await expect(cdp.send("Target.getTargets")).rejects.toThrow();
});

test("connecting to a dead port fails with a plain message", async () => {
  fake = startFakeCdp();
  const url = fake.wsUrl;
  fake.stop();
  fake = undefined;
  await expect(CdpClient.connect(url, 500)).rejects.toThrow(/Could not connect|timed out/);
});
