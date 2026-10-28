import type { LinkEntry } from "./compose.ts";
import { requireToken } from "./tokens.ts";

/** The link preview service. The web client calls `url/info` and attaches the answer to the message. */
const HOSTS: Record<string, string> = { amer: "us-prod.asyncgw.teams.microsoft.com" };

export type LinkPreview = { previewurl?: string; previewmeta?: { width: number; height: number }; title: string; description: string };

/** Maps a `url/info` answer to the `preview` object that the web client stores on the link. */
export function toPreview(info: any): LinkPreview | undefined {
  if (!info || typeof info.title !== "string" || !info.title) return undefined;
  const meta = info.thumbnail_meta;
  return {
    ...(typeof info.thumbnail === "string" && info.thumbnail ? { previewurl: info.thumbnail } : {}),
    ...(meta && typeof meta.width === "number" && typeof meta.height === "number" ? { previewmeta: { width: meta.width, height: meta.height } } : {}),
    title: info.title,
    description: typeof info.description === "string" ? info.description : "",
  };
}

/** Reads the preview for one URL. It returns undefined for any failure, so a send never fails because of a card. */
export async function fetchPreview(url: string, region: string): Promise<LinkPreview | undefined> {
  const host = HOSTS[region];
  if (!host) return undefined;
  try {
    const token = (await requireToken("chat")).token;
    const response = await fetch(`https://${host}/urlp/v1/url/info?url=${encodeURIComponent(url)}`, {
      headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
      redirect: "error",
      signal: AbortSignal.timeout(10_000),
    });
    return response.ok ? toPreview(await response.json()) : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Adds a preview card to the first link, as the web client does. Other links, and a link whose
 * preview cannot be read, get no card. With `enabled` false, no link gets a card and nothing is read.
 * Returns true when a card was attached, there was no link, or cards are off.
 */
export async function attachPreviews(links: LinkEntry[], region: string, enabled = true): Promise<boolean> {
  links.forEach((link, index) => { if (index > 0 || !enabled) link.previewenabled = false; });
  if (!enabled) return true;
  const first = links[0];
  if (!first) return true;
  const preview = await fetchPreview(first.url, region);
  if (!preview) { first.previewenabled = false; return false; }
  first.preview = preview;
  return true;
}
