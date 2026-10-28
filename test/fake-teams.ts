import { decodeJwt, KIND_BY_AUD } from "../src/tokens.ts";
import { saveProfile } from "../src/config.ts";
import { AUD, OID, TID, fakeJwt } from "./jwt.ts";

export type Reply = { status: number; headers?: Record<string, string>; json?: unknown; text?: string };
/** A single reply, a sequence (the last one repeats), "network" to throw a network error, or a function that answers from the request. */
export type Route = Reply | Reply[] | "network" | ((request: FakeRequest) => Reply);

export type FakeRequest = {
  host: string;
  method: string;
  path: string;
  query: Record<string, string>;
  /** The path as sent, not decoded, and the query string as sent. */
  rawPath: string;
  search: string;
  body: unknown;
  headers: Record<string, string>;
};

const AUD_BY_PREFIX: Array<[string, string]> = [
  ["/api/chatsvc/", AUD.chat],
  ["/api/csa/", AUD.csa],
  ["/api/mt/", AUD.spaces],
  ["/searchservice/", AUD.substrate],
  ["/v1/objects/", AUD.chat],
];
/** The image hosts. A request to one is routed by path only, and needs the chat token. */
const AMS_HOSTS = new Set(["us-prod.asyncgw.teams.microsoft.com", "us-api.asm.skype.com"]);
/** SharePoint hosts are routed by path only. They need the SharePoint token. */
export const isSharepointHost = (host: string) => host.endsWith(".sharepoint.com");

/** The Workflows card router host. A request to one is routed by path only, and needs no token. */
export const isRouterHost = (host: string) => host.endsWith(".logic.azure.com") || host.endsWith(".powerplatform.com");

/** The Flow and Power Apps API hosts that `card setup` reads and writes. A request to one is routed by path only. */
export const isFlowApiHost = (host: string) => host === "api.flow.microsoft.com" || host === "api.powerapps.com";

function matches(pattern: string, path: string): boolean {
  const want = pattern.split("/");
  const got = path.split("/");
  if (want.length !== got.length) return false;
  return want.every((part, index) => (part.startsWith("{") && part.endsWith("}")) || part === got[index]);
}

/**
 * Replaces globalThis.fetch with a fake Teams server. Routes use `METHOD /path`
 * keys with `{param}` templates. The query string is ignored for matching.
 * A request must carry a token with the audience for its service, or it gets 401.
 */
export function installFakeTeams(routes: Record<string, Route>) {
  const original = globalThis.fetch;
  const requests: FakeRequest[] = [];
  const served = new Map<string, number>();
  const host = new URL(process.env.TEAMS_HOST || "https://teams.cloud.microsoft").host;

  globalThis.fetch = (async (input: string | URL, init: RequestInit = {}) => {
    const url = new URL(String(input));
    if (url.host !== host && url.host !== "substrate.office.com" && !AMS_HOSTS.has(url.host) && !isSharepointHost(url.host) && !isRouterHost(url.host) && !isFlowApiHost(url.host)) throw new Error(`Test request went to an unexpected host: ${url.host}`);
    const method = (init.method || "GET").toUpperCase();
    const path = decodeURIComponent(url.pathname);
    const headers: Record<string, string> = {};
    new Headers(init.headers).forEach((value, key) => { headers[key.toLowerCase()] = value; });
    const rawBody = typeof init.body === "string" ? init.body : undefined;
    let body: unknown = rawBody;
    try { body = rawBody === undefined ? undefined : JSON.parse(rawBody); } catch { /* Raw text body. */ }
    if (init.body instanceof Uint8Array) body = init.body;
    if (init.body instanceof Blob) body = new Uint8Array(await init.body.arrayBuffer());
    const query = Object.fromEntries(url.searchParams.entries());
    const request: FakeRequest = { host: url.host, method, path, query, rawPath: url.pathname, search: url.search, body, headers };
    requests.push(request);

    const key = Object.keys(routes).find(route => {
      const [routeMethod, ...rest] = route.split(" ");
      return routeMethod === method && matches(rest.join(" "), path);
    });
    const expectedAud = isSharepointHost(url.host) ? AUD.sharepoint : AUD_BY_PREFIX.find(([prefix]) => path.startsWith(prefix))?.[1];
    const aud = decodeJwt((headers.authorization || "").replace(/^Bearer /, ""))?.aud;
    if (expectedAud && (typeof aud !== "string" || KIND_BY_AUD[aud] === undefined || aud !== expectedAud)) {
      return new Response(JSON.stringify({ errorCode: 401, message: "Unauthorized" }), { status: 401, headers: { "content-type": "application/json" } });
    }
    // The real upload host answers this way when x-ms-client-version is missing (2026-10-19).
    if (url.host === "us-prod.asyncgw.teams.microsoft.com" && !headers["x-ms-client-version"]) {
      return new Response("The 'User-Agent' header is missing or invalid.", { status: 400 });
    }
    if (!key) {
      return new Response(JSON.stringify({ errorCode: 404, message: "no fake route" }), { status: 404, headers: { "content-type": "application/json" } });
    }
    const route = routes[key]!;
    if (route === "network") throw new TypeError("fetch failed");
    const sequence = typeof route === "function" ? [route(request)] : Array.isArray(route) ? route : [route];
    const index = served.get(key) ?? 0;
    served.set(key, index + 1);
    const reply = sequence[Math.min(index, sequence.length - 1)]!;
    const responseHeaders = { ...(reply.headers ?? {}) };
    if (reply.json !== undefined) responseHeaders["content-type"] = "application/json";
    const payload = reply.json !== undefined ? JSON.stringify(reply.json) : reply.text;
    return new Response(reply.status === 204 ? null : payload ?? null, { status: reply.status, headers: responseHeaders });
  }) as typeof fetch;

  return {
    requests,
    /** Returns requests that match a method and path. */
    find: (method: string, path: string) => requests.filter(item => item.method === method && item.path === path),
    uninstall: () => { globalThis.fetch = original; },
  };
}

/** Stores a profile with fake tokens. Pass the kinds to include. */
export async function seedProfile(kinds: Array<"chat" | "csa" | "spaces" | "sharepoint" | "substrate"> = ["chat", "csa", "spaces"], exp?: number, profile = "default") {
  const expires = exp ?? Math.floor(Date.now() / 1000) + 3600;
  const tokens = Object.fromEntries(kinds.map(kind => [kind, { token: fakeJwt({ aud: AUD[kind], exp: expires }), exp: expires }]));
  await saveProfile(profile, {
    tokens,
    userId: OID,
    tenantId: TID,
    upn: "tester@example.test",
    name: "Test User",
    ...(kinds.includes("sharepoint") ? { sharepointHost: "fake-my.sharepoint.com" } : {}),
    region: "amer",
    updatedAt: new Date().toISOString(),
  });
  return tokens;
}

/** Removes the config and cache directories between tests and clears the token cache. */
export async function resetLocalState(): Promise<void> {
  const { rm } = await import("node:fs/promises");
  const { configDir, cacheDir } = await import("../src/config.ts");
  const { resetTokenCache } = await import("../src/tokens.ts");
  await rm(configDir(), { recursive: true, force: true });
  await rm(cacheDir(), { recursive: true, force: true });
  resetTokenCache();
}
