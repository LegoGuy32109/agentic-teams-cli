import { identity, authFailureMessage, redact, requireToken, type ResolvedToken } from "./tokens.ts";
import { loadProfile, type TokenKind } from "./config.ts";
import { clearDenial, rememberDenial } from "./access.ts";

/** The three services. Each one uses the token of the same kind. */
export type Service = TokenKind;

const PREFIX: Record<Service, (region: string) => string> = {
  chat: region => `/api/chatsvc/${region}/v1`,
  csa: region => `/api/csa/${region}/api`,
  spaces: region => `/api/mt/${region}/beta`,
  substrate: () => "",
  sharepoint: () => "",
};

const USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36";
const DEFAULT_CLIENT_VERSION = "1415/26091712222";
const MAX_IN_FLIGHT = 4;
const MAX_RETRIES = 3;
/** Retry-After above this is refused at once. Shorter waits are honored. */
const MAX_RETRY_AFTER_S = 120;
const TIMEOUT_MS = 30_000;
const SUCCESS = new Set([200, 201, 204, 207]);

/** A failed request. `uncertain` means a write may still have posted: a network error, 408, or 5xx. */
export class TeamsError extends Error {
  constructor(message: string, readonly status?: number, readonly uncertain = false) {
    super(message);
    this.name = "TeamsError";
  }
}

let sleep = (ms: number) => Bun.sleep(ms);

/** Replaces the wait used on 429 responses. Tests call it to skip real waits. */
export function setSleepForTests(fn: (ms: number) => Promise<unknown>): void {
  sleep = fn as (ms: number) => Promise<void>;
}

export function teamsHost(): string {
  const override = process.env.TEAMS_HOST;
  if (!override) return "https://teams.cloud.microsoft";
  if (/^http:\/\/(127\.0\.0\.1|localhost)(:\d+)?$/.test(override)) return override;
  throw new Error("TEAMS_HOST may only be http://127.0.0.1 or http://localhost. It exists for tests.");
}

/** The two image hosts. Only the amer pair was captured, so any other region is refused. */
export function amsHosts(region: string): { upload: string; view: string } {
  if (region !== "amer") throw new Error(`Inline images are verified only for the amer region. This profile uses ${region}.`);
  return { upload: "https://us-prod.asyncgw.teams.microsoft.com", view: "https://us-api.asm.skype.com" };
}

/** Checks a paging link from the service. It must be on the Teams host, or the request is refused. */
function linkUrl(link: string, host: string): URL {
  if (!link.startsWith(`${host}/`)) throw new Error("Refused a paging link that does not point at the Teams host.");
  return new URL(link);
}

export function clientVersion(): string {
  return process.env.TEAMS_CLIENT_VER || DEFAULT_CLIENT_VERSION;
}

/** Formats the local UTC offset as `-04:00`. */
export function utcOffset(date = new Date()): string {
  const minutes = -date.getTimezoneOffset();
  const sign = minutes >= 0 ? "+" : "-";
  const abs = Math.abs(minutes);
  return `${sign}${String(Math.floor(abs / 60)).padStart(2, "0")}:${String(abs % 60).padStart(2, "0")}`;
}

export function localTimeZone(): string {
  return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
}

/**
 * Builds the headers for one request. Every header is sent to every service
 * that the web client sends it to. Trim this table once Josh knows which are needed.
 *
 * All: authorization, behavioroverride, x-ms-migration, x-ms-test-user, clientinfo, user-agent.
 * csa adds: x-ms-client-type, x-ms-region, x-ms-client-version, x-ringoverride, x-ms-user-type.
 * spaces adds: x-ms-client-type, x-ms-client-version, accept.
 * sharepoint sends only: authorization, accept, application, user-agent.
 * A body adds: content-type.
 */
export function requestHeaders(
  service: Service,
  token: string,
  options: { region: string; hasBody: boolean; contentType?: string },
  now = new Date(),
): Record<string, string> {
  if (service === "sharepoint") {
    // The web client sends SharePoint no Teams client headers, so none go here.
    const headers: Record<string, string> = { authorization: `Bearer ${token}`, accept: "application/json", application: "Teams_Web", "user-agent": USER_AGENT };
    if (options.hasBody) headers["content-type"] = options.contentType ?? "application/json";
    return headers;
  }
  const version = clientVersion();
  const headers: Record<string, string> = {
    authorization: `Bearer ${token}`,
    behavioroverride: "redirectAs404",
    "x-ms-migration": "True",
    "x-ms-test-user": "False",
    clientinfo: `os=linux; osVer=undefined; proc=x86; lcid=en-us; deviceType=1; country=us; clientName=skypeteams; clientVer=${version}; utcOffset=${utcOffset(now)}; timezone=${localTimeZone()}`,
    "user-agent": USER_AGENT,
  };
  if (service === "csa") {
    Object.assign(headers, {
      "x-ms-client-type": "cdlworker",
      "x-ms-region": options.region,
      "x-ms-client-version": version,
      "x-ringoverride": "general",
      "x-ms-user-type": "real-user",
    });
  }
  if (service === "spaces") {
    Object.assign(headers, { "x-ms-client-type": "cdlworker", "x-ms-client-version": version, accept: "application/json" });
  }
  if (options.hasBody) headers["content-type"] = options.contentType ?? "application/json";
  return headers;
}

export type RequestOptions = {
  /** A full URL from a paging link. It must start with the Teams host. Replaces the path and query. */
  fullUrl?: string;
  /** The SharePoint host for this request. The caller checks that it is a host of the user's tenant. */
  sharepointHost?: string;
  query?: Record<string, string>;
  body?: unknown;
  rawBody?: string;
  /** Raw bytes for the body, as in an image upload. A `Blob` (such as `Bun.file()`) streams a file from disk. */
  bytes?: Uint8Array | Blob;
  /** Replaces the 30 second limit for one request. File uploads set it from the size. */
  timeoutMs?: number;
  /** Calls an image host instead of a Teams service path: `upload` to create an image, `view` to read one. The path then starts at `/v1/objects/`. */
  ams?: "upload" | "view";
  /** Extra headers for this request. */
  headers?: Record<string, string>;
  contentType?: string;
  /** Return the body as a Buffer. A response over `maxBytes` is refused before it is read. */
  binary?: boolean;
  maxBytes?: number;
  /** True for a message send. It only changes the rate-limit wait line, which says `send` or `request`. */
  send?: boolean;
  /** Internal guard against repeated renewal when the service rejects a new token. */
  renewed?: boolean;
};

let inFlight = 0;
const waiting: Array<() => void> = [];

async function slot<T>(work: () => Promise<T>): Promise<T> {
  while (inFlight >= MAX_IN_FLIGHT) await new Promise<void>(resolve => waiting.push(resolve));
  inFlight++;
  try {
    return await work();
  } finally {
    inFlight--;
    waiting.shift()?.();
  }
}

/** The cause of a failed request, in plain words: a code such as ECONNREFUSED, or `timeout`. */
export function networkCause(error: unknown): string | undefined {
  const record = (error ?? {}) as { code?: unknown; name?: unknown; cause?: { code?: unknown } };
  if (record.name === "TimeoutError" || record.name === "AbortError") return "timeout";
  const code = typeof record.code === "string" ? record.code : typeof record.cause?.code === "string" ? record.cause.code : undefined;
  // Bun names a refused connection in its own style. Show the usual system code.
  return code === "ConnectionRefused" ? "ECONNREFUSED" : code;
}

async function fetchOnce(url: URL, init: RequestInit & { method: string }, timeoutMs = TIMEOUT_MS): Promise<Response> {
  try {
    return await fetch(url, { ...init, redirect: "error", signal: AbortSignal.timeout(timeoutMs) });
  } catch (error) {
    const cause = networkCause(error);
    throw new TeamsError(`${init.method} ${url.pathname} failed: network error${cause ? ` (${cause})` : ""}.`, undefined, true);
  }
}

/** Service errors that have a plain sentence. Matched on the `subCode` the service sends. */
const PLAIN_ERRORS: Record<string, string> = {
  LocationLookupFailed: "Conversation not found (404).",
  SoftDeleted: "This conversation was deleted (404).",
  AclCheckFailed: "You are not a member of this conversation, or the chat token lacks access (403).",
  MessageSizeExceeded: "Teams refused the message: it is too large (400).",
};

/** Reads a `subCode` from an error object that the service wrote as a JSON string. */
function nestedError(text: string): string | undefined {
  if (!text.trimStart().startsWith("{")) return undefined;
  try {
    const code = JSON.parse(text)?.subCode;
    return typeof code === "string" ? code : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Turns a failed service response into plain words with the status. The raw body is never shown.
 * For any other error, only the service's `message` or `details` field is used, cut to 200 characters.
 */
export function serviceFailure(method: string, path: string, status: number, text: string, region: string): string {
  let parsed: any;
  try { parsed = JSON.parse(text); } catch { parsed = undefined; }
  const code = typeof parsed?.subCode === "string" ? parsed.subCode : undefined;
  if (code && PLAIN_ERRORS[code]) return PLAIN_ERRORS[code];
  const detail = [parsed?.message, parsed?.details].find((value): value is string => typeof value === "string" && value.trim() !== "");
  // The service can nest its error object as a JSON string inside `message` or `details`. Read the code from there too.
  const nested = detail ? nestedError(detail) : undefined;
  if (nested && PLAIN_ERRORS[nested]) return PLAIN_ERRORS[nested];
  if (status === 400 && !detail && process.env.TEAMS_REGION) {
    return `Teams rejected the request (400). Check TEAMS_REGION: ${region}.`;
  }
  // A detail that is itself a JSON object is never shown. It would print the raw body.
  const short = detail && !detail.trimStart().startsWith("{") ? redact(detail).replace(/\s+/g, " ").trim().slice(0, 200) : "";
  return `${method} ${path} failed: ${status}${short ? ` ${short}` : ""}`;
}

/** The message for a 429 after the retries. It never includes the service body. */
function limitMessage(method: string, seconds: number): string {
  const verb = method === "GET" ? "request" : method === "POST" ? "send" : "change";
  return `Teams refused the ${verb}: 20 per minute quota. Try again in ${seconds}s.`;
}

/**
 * Calls one Teams service. Returns parsed JSON, or undefined for an empty body.
 * Retries only 429, up to three times, honoring Retry-After up to 120 seconds. Any other failure is thrown at once.
 */
export async function request(service: Service, method: string, path: string, options: RequestOptions = {}): Promise<any> {
  const who = await identity();
  const token: ResolvedToken = await requireToken(service);
  const host = options.ams ? amsHosts(who.region)[options.ams]
    : service === "substrate" ? "https://substrate.office.com"
    : service === "sharepoint" ? `https://${options.sharepointHost ?? await (await import("./sharepoint.ts")).sharepointHost()}`
    : teamsHost();
  const prefix = options.ams ? "" : PREFIX[service](who.region);
  const url = options.fullUrl ? linkUrl(options.fullUrl, host) : new URL(`${host}${prefix}${path}`);
  if (url.origin !== host) throw new Error("Refused a request outside the service host.");
  if (!options.fullUrl) for (const [key, value] of Object.entries(options.query ?? {})) url.searchParams.set(key, value);
  const hasBody = options.body !== undefined || options.rawBody !== undefined || options.bytes !== undefined;
  const body = options.bytes ?? options.rawBody ?? (options.body !== undefined ? JSON.stringify(options.body) : undefined);
  const init = {
    method,
    headers: {
      ...requestHeaders(service, token.token, { region: who.region, hasBody, contentType: options.contentType }),
      // The upload host refuses a request without it, with a 400 that names the User-Agent header instead.
      ...(options.ams ? { "x-ms-client-version": clientVersion() } : {}),
      ...options.headers,
    },
    body: body as RequestInit["body"],
  };
  const historyMatch = service === "chat" && method === "GET" ? url.pathname.match(/\/users\/ME\/conversations\/([^/]+)\/messages$/) : undefined;
  const historyId = historyMatch ? decodeURIComponent(historyMatch[1]!) : undefined;
  if (service === "substrate") Object.assign(init.headers, { Referer: "https://teams.cloud.microsoft/", "x-client-version": "T2.1" });

  let response: Response | undefined;
  for (let retry = 0; ; retry++) {
    response = await slot(() => fetchOnce(url, init, options.timeoutMs));
    if (response.status !== 429) break;
    const seconds = Math.max(1, Number(response.headers.get("retry-after")) || 1);
    if (retry === MAX_RETRIES || seconds > MAX_RETRY_AFTER_S) throw new TeamsError(limitMessage(method, seconds), 429);
    console.error(`Rate limit: waiting ${seconds}s before the next ${options.send ? "send" : "request"}.`);
    await sleep(seconds * 1000);
  }
  if (!response) throw new TeamsError(`${method} ${path} failed: no response.`, undefined, true);

  if (response.status === 401) {
    if (!options.renewed && token.source === "store" && (await loadProfile()).profile?.refresh) {
      const { refreshToken } = await import("./refresh.ts");
      await refreshToken(service, true);
      return request(service, method, path, { ...options, renewed: true });
    }
    throw new TeamsError(authFailureMessage(service, token.exp), 401);
  }
  if (SUCCESS.has(response.status)) {
    if (options.binary) {
      if (options.maxBytes !== undefined && Number(response.headers.get("content-length")) > options.maxBytes) {
        throw new TeamsError(`The file is larger than ${options.maxBytes} bytes.`, response.status);
      }
      const bytes = Buffer.from(await response.arrayBuffer());
      if (options.maxBytes !== undefined && bytes.length > options.maxBytes) throw new TeamsError(`The file is larger than ${options.maxBytes} bytes.`, response.status);
      return bytes;
    }
    if (historyId) await clearDenial(who, historyId).catch(() => {});
    const text = await response.text();
    if (!text) return undefined;
    try {
      return JSON.parse(text);
    } catch {
      return text;
    }
  }
  const uncertain = response.status === 408 || response.status >= 500;
  if (historyId && response.status === 403) await rememberDenial(who, historyId).catch(() => {});
  const text = await response.text().catch(() => "");
  throw new TeamsError(serviceFailure(method, path, response.status, text, who.region), response.status, uncertain);
}

/** Runs async work over items with a fixed limit. Results keep the input order. */
export async function mapConcurrent<T, R>(items: T[], limit: number, work: (item: T) => Promise<R>): Promise<R[]> {
  const results = new Array<R>(items.length);
  let next = 0;
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const index = next++;
      results[index] = await work(items[index]!);
    }
  }));
  return results;
}
