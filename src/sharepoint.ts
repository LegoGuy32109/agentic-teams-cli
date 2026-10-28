import { loadProfile, saveProfile } from "./config.ts";

const HOST_PATTERN = /^[a-z0-9][a-z0-9-]*\.sharepoint\.com$/;

/** Returns a lowercase SharePoint host from a host or a URL, or undefined when it is not one. */
export function parseSharepointHost(value: unknown): string | undefined {
  if (typeof value !== "string" || !value.trim()) return undefined;
  let host: string;
  try {
    host = new URL(value.includes("://") ? value : `https://${value}`).host.toLowerCase();
  } catch {
    return undefined;
  }
  return HOST_PATTERN.test(host) ? host : undefined;
}

/**
 * The path of a personal site for a user name: `.` and `@` become `_`. The web client sent
 * `/personal/alex_example_com` for `alex@example.com`. Other characters are unproven.
 */
export function personalSitePath(upn: string): string {
  return `/personal/${upn.toLowerCase().replace(/[.@]/g, "_")}`;
}

export function personalSiteUrl(host: string, upn: string): string {
  return `https://${host}${personalSitePath(upn)}`;
}

/**
 * Finds the host of the personal SharePoint site. `TEAMS_SHAREPOINT_HOST` wins, then the host saved in the
 * profile, then the `personalFileSite` property of the chat service. A host found that way is saved.
 */
export async function sharepointHost(): Promise<string> {
  const fromEnv = process.env.TEAMS_SHAREPOINT_HOST;
  if (fromEnv) {
    const host = parseSharepointHost(fromEnv);
    if (!host) throw new Error("TEAMS_SHAREPOINT_HOST must be a host such as contoso-my.sharepoint.com.");
    return host;
  }
  const { name, profile } = await loadProfile();
  const saved = parseSharepointHost(profile?.sharepointHost);
  if (saved) return saved;
  const { request } = await import("./transport.ts");
  const properties = await request("chat", "GET", "/users/ME/properties");
  const host = parseSharepointHost(siteUrlFromProperty(properties?.personalFileSite));
  if (!host) throw new Error("Teams did not name your SharePoint site. Set TEAMS_SHAREPOINT_HOST, for example contoso-my.sharepoint.com.");
  const latest = (await loadProfile()).profile ?? profile;
  if (latest) await saveProfile(name, { ...latest, sharepointHost: host });
  return host;
}

/**
 * The personal site that holds `Microsoft Teams Chat Files`: its host and its path, such as `/personal/jo_contoso_com`.
 * Both come from the `personalFileSite` property, so no name rule is guessed. When the property has no site, the host
 * comes from `TEAMS_SHAREPOINT_HOST` or the profile, and the path from `personalSitePath()`.
 */
export async function personalSite(): Promise<{ host: string; path: string }> {
  const { request } = await import("./transport.ts");
  const properties = await request("chat", "GET", "/users/ME/properties");
  const url = siteUrlFromProperty(properties?.personalFileSite);
  let path: string | undefined;
  try {
    const found = typeof url === "string" ? new URL(url).pathname.replace(/\/+$/, "") : "";
    if (/^\/personal\/[^/]+$/.test(found)) path = found;
  } catch { /* The property holds no usable address. */ }
  const host = process.env.TEAMS_SHAREPOINT_HOST ? await sharepointHost() : (parseSharepointHost(url) ?? await sharepointHost());
  if (path) return { host, path };
  const { identity } = await import("./tokens.ts");
  return { host, path: personalSitePath((await identity()).upn) };
}

/** `personalFileSite` is a JSON string, such as `{"personalSiteUrl":"https://contoso-my.sharepoint.com/personal/a_b/",...}`. */
function siteUrlFromProperty(value: unknown): unknown {
  if (typeof value !== "string" || !value.trimStart().startsWith("{")) return value;
  try { return JSON.parse(value)?.personalSiteUrl; } catch { return undefined; }
}

/**
 * The SharePoint hosts of the tenant: the team-site host and the personal-site host, such as
 * `contoso.sharepoint.com` and `contoso-my.sharepoint.com`. One SharePoint token reads both (proven 2026-10-19).
 */
export function tenantSharepointHosts(personalHost: string): string[] {
  const root = personalHost.replace(/-my\.sharepoint\.com$/, ".sharepoint.com");
  return [...new Set([root, root.replace(/\.sharepoint\.com$/, "-my.sharepoint.com")])];
}

/** The share id of a file URL, as the web client builds it: `u!` and the unpadded base64url of the URL. */
export function shareId(fileUrl: string): string {
  return `u!${Buffer.from(fileUrl, "utf8").toString("base64url")}`;
}
