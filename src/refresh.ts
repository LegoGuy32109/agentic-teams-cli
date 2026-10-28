import { loadProfile, saveProfile, type RefreshConfig, type TokenKind } from "./config.ts";
import { SIGN_IN_AGAIN, tryBrowserRenew } from "./browser-auth.ts";
import { locked } from "./lock.ts";
import { classifyToken, KINDS, redact } from "./tokens.ts";

export const DEFAULT_SCOPES: Record<Exclude<TokenKind, "sharepoint">, string> = {
  chat: "https://ic3.teams.office.com/.default openid profile offline_access",
  csa: "https://chatsvcagg.teams.microsoft.com/.default openid profile offline_access",
  spaces: "https://api.spaces.skype.com/.default openid profile offline_access",
  substrate: "https://substrate.office.com/.default openid profile offline_access",
};

/** The scope for a kind. The sharepoint scope names the personal site host, so it needs the profile. */
export function scopeFor(kind: TokenKind, config: RefreshConfig, host?: string): string {
  const set = config.scopes?.[kind];
  if (set) return set;
  if (kind !== "sharepoint") return DEFAULT_SCOPES[kind];
  if (!host) throw new Error("The SharePoint host is unknown. Run: teams-cli auth refresh");
  return `https://${host}/.default openid profile offline_access`;
}

function validate(config: RefreshConfig): void {
  if (!/^[\w.-]+$/.test(config.tenantId) || !/^[\w-]+$/.test(config.clientId) || !config.refreshToken) {
    throw new Error("Refresh setup needs tenantId, clientId, and refreshToken.");
  }
  if (config.origin && config.origin !== "https://teams.cloud.microsoft") {
    throw new Error("The refresh Origin must be https://teams.cloud.microsoft.");
  }
}

/** Imports a captured token request without executing any part of the pasted command. */
export function parseRefreshSetup(text: string): RefreshConfig {
  if (text.trimStart().startsWith("{")) {
    let config: RefreshConfig;
    try { config = JSON.parse(text); } catch { throw new Error("Refresh setup is not valid JSON."); }
    if (!config || typeof config !== "object" || typeof config.clientId !== "string" ||
        typeof config.tenantId !== "string" || typeof config.refreshToken !== "string") {
      throw new Error("Refresh setup needs tenantId, clientId, and refreshToken.");
    }
    validate(config);
    return config;
  }
  const tenant = text.match(/https:\/\/login\.microsoftonline\.com\/([\w.-]+)\/oauth2\/v2\.0\/token/)?.[1];
  const body = text.match(/--data(?:-raw|-binary)?\s+'([^']*)'/s)?.[1] ?? text.match(/--data(?:-raw|-binary)?\s+"([^"]*)"/s)?.[1];
  if (!tenant || !body) throw new Error("Paste the Microsoft token cURL with its --data-raw form body, or refresh setup JSON.");
  const form = new URLSearchParams(body);
  if (form.get("grant_type") !== "refresh_token") throw new Error("The captured request must use grant_type=refresh_token.");
  const config: RefreshConfig = {
    clientId: form.get("client_id") ?? "",
    tenantId: tenant,
    refreshToken: form.get("refresh_token") ?? "",
    origin: "https://teams.cloud.microsoft",
  };
  validate(config);
  return config;
}

export async function setupRefresh(text: string): Promise<void> {
  const config = parseRefreshSetup(text);
  const { name, profile } = await loadProfile();
  if (!profile) throw new Error("Import access tokens first with teams-cli auth.");
  if (profile.tenantId !== config.tenantId) throw new Error("The refresh request belongs to another tenant. Select the matching profile.");
  await saveProfile(name, { ...profile, refresh: config });
}

/** What the user reads when Microsoft refuses the stored refresh token and the browser cannot fix it. */
function signInMessage(outcome: { reason?: string; signInRequired?: boolean }): string {
  if (outcome.signInRequired) return SIGN_IN_AGAIN;
  if (outcome.reason) return `Microsoft requires a new sign-in. The browser renewal failed: ${outcome.reason} Run: teams-cli auth login`;
  return "Microsoft requires a new sign-in. Run: teams-cli auth login";
}

/**
 * Renews a stored token and saves the rotated refresh token before releasing the lock. When Microsoft
 * refuses the refresh token, a browser profile gets one chance to sign in silently, then the renewal runs again.
 */
export async function refreshToken(kind: TokenKind, force = false): Promise<void> {
  await locked(async () => {
    let browserTried = false;
    for (;;) {
      const { name, profile } = await loadProfile();
      const config = profile?.refresh;
      if (!profile || !config) throw new Error("No refresh setup. Run: teams-cli auth setup-refresh < token-request.txt");
      validate(config);
      if (!force && (profile.tokens[kind]?.exp ?? 0) * 1000 - Date.now() >= 60_000) return;
      const form = new URLSearchParams({
        client_id: config.clientId,
        grant_type: "refresh_token",
        refresh_token: config.refreshToken,
        scope: scopeFor(kind, config, profile.sharepointHost),
      });
      let response: Response;
      try {
        response = await fetch(`https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`, {
          method: "POST", redirect: "error", signal: AbortSignal.timeout(30_000),
          headers: { "content-type": "application/x-www-form-urlencoded", ...(config.origin ? { Origin: config.origin } : {}) },
          body: form.toString(),
        });
      } catch { throw new Error(`Could not renew the ${kind} token. Check the network and try again.`); }
      let data: any;
      try { data = await response.json(); } catch { throw new Error(`Token renewal returned an invalid response (${response.status}).`); }
      if (!response.ok) {
        if (data?.error === "invalid_grant" || data?.error === "interaction_required") {
          if (browserTried) throw new Error(signInMessage({ reason: "Microsoft still refused the renewed session." }));
          browserTried = true;
          const outcome = await tryBrowserRenew();
          if (!outcome.ok) throw new Error(signInMessage(outcome));
          continue;
        }
        throw new Error(`Microsoft refused ${kind} token renewal (${response.status}). The requested service may need a different scope.`);
      }
      // Rotation must survive even if the access token does not match the requested resource.
      if (typeof data?.refresh_token === "string" && data.refresh_token) {
        config.refreshToken = data.refresh_token;
        await saveProfile(name, { ...profile, refresh: config });
      }
      if (typeof data?.access_token !== "string") throw new Error("Token renewal returned no access token.");
      const parsed = classifyToken(data.access_token);
      if (parsed.kind !== kind || typeof parsed.claims?.exp !== "number" || parsed.claims.exp * 1000 - Date.now() < 60_000) {
        throw new Error(`Token renewal did not return a usable ${kind} token. Check its configured scope.`);
      }
      if (parsed.claims.tid !== profile.tenantId || parsed.claims.oid !== profile.userId) {
        throw new Error("Token renewal returned another account. Import the matching profile before continuing.");
      }
      profile.tokens[kind] = { token: data.access_token, exp: parsed.claims.exp };
      await saveProfile(name, { ...profile, refresh: config, updatedAt: new Date().toISOString() });
      return;
    }
  });
}

export async function refreshAll(): Promise<void> {
  for (const kind of KINDS) {
    if (kind === "sharepoint") {
      // The personal site host comes from a chat read. A failure here must not hide the other renewals.
      try {
        await (await import("./sharepoint.ts")).sharepointHost();
      } catch (error) {
        console.log(`${kind}: skipped (${redact(error instanceof Error ? error.message : String(error))})`);
        process.exitCode = 1;
        continue;
      }
    }
    await refreshToken(kind, true);
    console.log(`${kind}: renewed`);
  }
}
