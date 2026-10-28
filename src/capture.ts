import type { RefreshConfig, StoredToken, TokenKind } from "./config.ts";
import type { CdpEvent } from "./cdp.ts";
import { classifyToken, type Claims } from "./tokens.ts";

const TOKEN_URL = /^https:\/\/login\.microsoftonline\.com\/([\w.-]+)\/oauth2\/v2\.0\/token(?:[?#]|$)/;
const REGION = /\/api\/(?:chatsvc|csa|mt)\/([A-Za-z0-9-]+)\//;
const TEAMS_ORIGIN = "https://teams.cloud.microsoft";

export type Captured = {
  tokens: Partial<Record<TokenKind, StoredToken>>;
  claims?: Claims;
  region?: string;
  /** The newest refresh token that the page's token requests returned, with the request that owns it. */
  refresh?: RefreshConfig;
};

/** The two DevTools reads that the capture needs. A test supplies a fake. */
export type CaptureIO = {
  getResponseBody(requestId: string, sessionId?: string): Promise<{ body: string; base64Encoded: boolean }>;
  getPostData(requestId: string, sessionId?: string): Promise<string | undefined>;
};

type Seen = { sessionId?: string; tenant: string; postData?: string; hasPostData: boolean; status?: number };

/**
 * Watches Network events for token responses from the Microsoft token endpoint. Each access token is
 * classified by its audience. Response bodies are read one at a time, in the order the loads finished,
 * so the newest refresh token wins.
 */
export class TokenCapture {
  private readonly seen = new Map<string, Seen>();
  private queue: Promise<void> = Promise.resolve();
  private readonly tokens: Partial<Record<TokenKind, StoredToken>> = {};
  private claims?: Claims;
  private region?: string;
  private refresh?: RefreshConfig;
  /** Counts of problems, never their text. A response body can hold a credential. */
  errors = 0;
  otherAccount = 0;

  constructor(private readonly io: CaptureIO, private readonly now: () => number = Date.now) {}

  get captured(): Captured {
    return { tokens: { ...this.tokens }, claims: this.claims, region: this.region, refresh: this.refresh };
  }

  /** Waits for every response body read that has started. */
  async settled(): Promise<void> {
    let queue: Promise<void>;
    do { queue = this.queue; await queue; } while (queue !== this.queue);
  }

  missing(kinds: TokenKind[]): TokenKind[] {
    return kinds.filter(kind => !this.tokens[kind]);
  }

  onEvent(event: CdpEvent): void {
    const params = event.params ?? {};
    if (event.method === "Network.requestWillBeSent") {
      const request = params.request ?? {};
      const url = String(request.url ?? "");
      const region = url.match(REGION)?.[1];
      if (region) this.region = region;
      const tenant = url.match(TOKEN_URL)?.[1];
      if (tenant && String(request.method).toUpperCase() === "POST" && typeof params.requestId === "string") {
        this.seen.set(params.requestId, {
          sessionId: event.sessionId,
          tenant,
          postData: typeof request.postData === "string" ? request.postData : undefined,
          hasPostData: request.hasPostData === true,
        });
      }
    } else if (event.method === "Network.responseReceived") {
      const item = this.seen.get(params.requestId);
      if (item) item.status = Number(params.response?.status);
    } else if (event.method === "Network.loadingFinished") {
      const item = this.seen.get(params.requestId);
      if (!item || item.status !== 200) return;
      const requestId = String(params.requestId);
      this.queue = this.queue.then(() => this.read(requestId, item)).catch(() => { this.errors++; });
    }
  }

  /** Adds an access token from any source. Returns its kind, or undefined when it is unusable. */
  addAccessToken(token: string): TokenKind | undefined {
    const { kind, claims } = classifyToken(token);
    if (!kind || !claims || typeof claims.exp !== "number" || claims.exp * 1000 - this.now() < 60_000) return undefined;
    if (this.claims && (claims.tid !== this.claims.tid || claims.oid !== this.claims.oid)) {
      this.otherAccount++;
      return undefined;
    }
    this.claims ??= claims;
    const existing = this.tokens[kind];
    if (!existing || claims.exp > existing.exp) this.tokens[kind] = { token, exp: claims.exp };
    return kind;
  }

  private async read(requestId: string, item: Seen): Promise<void> {
    const { body, base64Encoded } = await this.io.getResponseBody(requestId, item.sessionId);
    const text = base64Encoded ? Buffer.from(body, "base64").toString("utf8") : body;
    let data: any;
    try { data = JSON.parse(text); } catch { this.errors++; return; }
    if (typeof data?.access_token !== "string") return;
    const kind = this.addAccessToken(data.access_token);
    if (!kind || typeof data.refresh_token !== "string" || !data.refresh_token) return;
    // Only a response that returned a Teams service token owns a refresh token that is worth keeping.
    let postData = item.postData;
    if (postData === undefined && item.hasPostData) postData = await this.io.getPostData(requestId, item.sessionId).catch(() => undefined);
    const clientId = postData ? new URLSearchParams(postData).get("client_id") : null;
    if (!clientId) return;
    this.refresh = { clientId, tenantId: item.tenant, refreshToken: data.refresh_token, origin: TEAMS_ORIGIN };
  }
}

const SIGN_IN_HOSTS = new Set(["login.microsoftonline.com", "login.live.com", "login.microsoft.com"]);

export function isSignInUrl(url: string): boolean {
  try { return SIGN_IN_HOSTS.has(new URL(url).host); } catch { return false; }
}

/**
 * Tells a silent sign-in from a sign-in page that waits for a person. A silent sign-in passes through
 * the Microsoft login host for a moment. A page that stays there longer than the grace period needs a person.
 */
export class SignInWatch {
  private since?: number;
  constructor(private readonly graceMs = 8_000) {}

  observe(url: string, now: number): void {
    if (!isSignInUrl(url)) this.since = undefined;
    else this.since ??= now;
  }

  get onSignInPage(): boolean {
    return this.since !== undefined;
  }

  needsSignIn(now: number): boolean {
    return this.since !== undefined && now - this.since >= this.graceMs;
  }
}
