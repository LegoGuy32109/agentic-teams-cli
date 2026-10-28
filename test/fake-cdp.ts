import { AUD, fakeJwt } from "./jwt.ts";

/** A fake Chrome DevTools server on a local port. It speaks only the commands that teams-cli sends. */

export type FakeEvent = { method: string; params: any; sessionId?: string; delayMs?: number };

export type FakeCdpOptions = {
  /** Events to send after `Page.navigate` to the Teams URL, in order. */
  onNavigate?: (url: string) => FakeEvent[];
  /** Response bodies by request id, for `Network.getResponseBody`. */
  bodies?: Record<string, { body: string; base64Encoded?: boolean }>;
  /** Post data by request id, for `Network.getRequestPostData`. */
  postData?: Record<string, string>;
  /** The value that a `Runtime.evaluate` of the storage scan returns. */
  storage?: string[];
  /** Pages that `Target.getTargets` reports. Default one page. */
  pages?: number;
  userAgent?: string;
};

export const SESSION = "S1";

export function startFakeCdp(options: FakeCdpOptions = {}) {
  const commands: Array<{ method: string; params: any; sessionId?: string }> = [];
  const sockets = new Set<any>();
  let closedByCommand = false;
  const handle = (socket: { send(data: string): unknown; close(): void }, raw: unknown): unknown => {
    const message = JSON.parse(String(raw));
    commands.push({ method: message.method, params: message.params, sessionId: message.sessionId });
    const reply = (result: unknown = {}) => socket.send(JSON.stringify({ id: message.id, result }));
    switch (message.method) {
      case "Target.getTargets":
        return reply({ targetInfos: Array.from({ length: options.pages ?? 1 }, (_, index) => ({ targetId: `T${index + 1}`, type: "page" })) });
      case "Target.createTarget": return reply({ targetId: "T1" });
      case "Target.attachToTarget": return reply({ sessionId: SESSION });
      case "Browser.getVersion": return reply({ userAgent: options.userAgent ?? "Mozilla/5.0 (X11; Linux x86_64) HeadlessChrome/154.0.0.0 Safari/537.36" });
      case "Network.getResponseBody": {
        const body = options.bodies?.[message.params.requestId];
        return body ? reply({ body: body.body, base64Encoded: body.base64Encoded === true })
          : socket.send(JSON.stringify({ id: message.id, error: { code: -32000, message: "No resource with given identifier found" } }));
      }
      case "Network.getRequestPostData": return reply({ postData: options.postData?.[message.params.requestId] });
      case "Runtime.evaluate":
        return reply({ result: { value: String(message.params.expression).includes("localStorage") ? options.storage ?? [] : false } });
      case "Browser.close":
        closedByCommand = true;
        reply();
        return void setTimeout(() => socket.close(), 5);
      case "Page.navigate": {
        reply({ frameId: "F1" });
        let at = 0;
        for (const event of options.onNavigate?.(String(message.params.url)) ?? []) {
          at += event.delayMs ?? 0;
          setTimeout(() => {
            try { socket.send(JSON.stringify({ method: event.method, params: event.params, sessionId: event.sessionId ?? SESSION })); } catch { /* Closed. */ }
          }, at);
        }
        return;
      }
      default: return reply();
    }
  };
  const server = Bun.serve({
    port: 0,
    hostname: "127.0.0.1",
    fetch(request, srv) {
      return srv.upgrade(request) ? undefined : new Response("fake cdp", { status: 426 });
    },
    websocket: {
      open(socket) { sockets.add(socket); },
      close(socket) { sockets.delete(socket); },
      message(socket, raw): void { handle(socket, raw); },
    },
  });
  return {
    wsUrl: `ws://127.0.0.1:${server.port}/devtools/browser/fake`,
    commands,
    get closedByCommand() { return closedByCommand; },
    dropConnections: () => { for (const socket of sockets) socket.close(); },
    stop: () => server.stop(true),
  };
}

export const TENANT_URL = (tid: string) => `https://login.microsoftonline.com/${tid}/oauth2/v2.0/token`;

/**
 * The three Network events of one POST to the token endpoint. The body goes in the server's `bodies`,
 * keyed by `id`. Pass the form as the request body.
 */
export function tokenExchange(id: string, tid: string, form: Record<string, string>, status = 200): FakeEvent[] {
  return [
    { method: "Network.requestWillBeSent", params: { requestId: id, request: { url: TENANT_URL(tid), method: "POST", postData: new URLSearchParams(form).toString(), hasPostData: true } } },
    { method: "Network.responseReceived", params: { requestId: id, response: { url: TENANT_URL(tid), status } } },
    { method: "Network.loadingFinished", params: { requestId: id } },
  ];
}

export const body = (value: unknown) => ({ body: JSON.stringify(value) });

/** A token endpoint success body for one audience. */
export const tokenBody = (aud: string, extra: Record<string, unknown> = {}, claims: Record<string, unknown> = {}) =>
  body({ access_token: fakeJwt({ aud, ...claims }), refresh_token: `fake-refresh-for-${aud.length}`, ...extra });

export { AUD };
