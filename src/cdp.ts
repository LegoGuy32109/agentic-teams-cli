import { redact } from "./tokens.ts";

/** A minimal Chrome DevTools Protocol client over Bun's built-in WebSocket. It has no dependency. */

export type CdpEvent = { method: string; params: any; sessionId?: string };

export type CdpMessage =
  | { kind: "response"; id: number; result?: any; error?: { code?: number; message?: string } }
  | { kind: "event"; event: CdpEvent };

/** Reads one DevTools frame. Anything that is not a response or an event returns undefined. */
export function parseCdpMessage(raw: unknown): CdpMessage | undefined {
  if (typeof raw !== "string") return undefined;
  let message: any;
  try { message = JSON.parse(raw); } catch { return undefined; }
  if (!message || typeof message !== "object") return undefined;
  if (typeof message.id === "number") return { kind: "response", id: message.id, result: message.result, error: message.error };
  if (typeof message.method === "string") {
    return { kind: "event", event: { method: message.method, params: message.params ?? {}, sessionId: message.sessionId } };
  }
  return undefined;
}

type Pending = { method: string; resolve: (value: any) => void; reject: (error: Error) => void; timer: ReturnType<typeof setTimeout> };

export class CdpClient {
  private nextId = 1;
  private readonly pending = new Map<number, Pending>();
  private readonly listeners = new Set<(event: CdpEvent) => void>();
  /** Resolves when the socket closes, for any reason. */
  readonly closed: Promise<void>;

  private constructor(private readonly socket: WebSocket) {
    this.closed = new Promise<void>(resolve => {
      socket.addEventListener("close", () => {
        for (const [id, item] of this.pending) {
          clearTimeout(item.timer);
          item.reject(new Error(`The browser closed the connection during ${item.method}.`));
          this.pending.delete(id);
        }
        resolve();
      });
    });
    socket.addEventListener("message", message => {
      const parsed = parseCdpMessage(message.data);
      if (!parsed) return;
      if (parsed.kind === "event") {
        for (const listener of this.listeners) listener(parsed.event);
        return;
      }
      const item = this.pending.get(parsed.id);
      if (!item) return;
      this.pending.delete(parsed.id);
      clearTimeout(item.timer);
      if (parsed.error) item.reject(new Error(`Browser command ${item.method} failed: ${redact(String(parsed.error.message ?? "unknown error"))}`));
      else item.resolve(parsed.result ?? {});
    });
  }

  static connect(url: string, timeoutMs = 10_000): Promise<CdpClient> {
    return new Promise((resolve, reject) => {
      let socket: WebSocket;
      try { socket = new WebSocket(url); } catch { return reject(new Error("Could not open the browser debugging connection.")); }
      const timer = setTimeout(() => { socket.close(); reject(new Error("The browser debugging connection timed out.")); }, timeoutMs);
      socket.addEventListener("open", () => { clearTimeout(timer); resolve(new CdpClient(socket)); });
      socket.addEventListener("error", () => { clearTimeout(timer); reject(new Error("Could not connect to the browser debugging port.")); });
    });
  }

  send<T = any>(method: string, params: Record<string, unknown> = {}, sessionId?: string, timeoutMs = 30_000): Promise<T> {
    const id = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      if (this.socket.readyState !== WebSocket.OPEN) return reject(new Error(`The browser connection is closed (${method}).`));
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Browser command ${method} timed out.`));
      }, timeoutMs);
      this.pending.set(id, { method, resolve, reject, timer });
      try {
        this.socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
      } catch {
        this.pending.delete(id);
        clearTimeout(timer);
        reject(new Error(`Could not send ${method} to the browser.`));
      }
    });
  }

  /** Adds an event listener. Returns a function that removes it. */
  on(listener: (event: CdpEvent) => void): () => void {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }

  close(): void {
    try { this.socket.close(); } catch { /* Already closed. */ }
  }
}
