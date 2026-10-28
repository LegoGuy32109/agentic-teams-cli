import { readFile } from "node:fs/promises";
import { dirname, isAbsolute, resolve } from "node:path";
import { parseImageHeader, type ImageKind } from "./image.ts";

export const SVG_REFUSAL = "Teams does not show SVG in cards. Render it to PNG first (rsvg-convert, or a headless browser for Mermaid)";
const FILE_PREFIX = "file:";

/** One image that was read from disk and put into the card as a data URI. */
export type CardImage = { path: string; type: ImageKind; bytes: number };

function isObject(value: unknown): value is Record<string, any> {
  return !!value && typeof value === "object" && !Array.isArray(value);
}

function looksLikeSvg(path: string, bytes: Uint8Array): boolean {
  if (/\.svgz?$/i.test(path)) return true;
  const head = Buffer.from(bytes.subarray(0, 512)).toString("utf8").trimStart().toLowerCase();
  return head.startsWith("<svg") || (head.startsWith("<?xml") && head.includes("<svg"));
}

async function toDataUri(reference: string, baseDir: string): Promise<{ uri: string; image: CardImage }> {
  const raw = reference.slice(FILE_PREFIX.length).trim();
  if (!raw) throw new Error(`An image url "${reference}" names no file. Use file:<path>.`);
  const path = isAbsolute(raw) ? raw : resolve(baseDir, raw);
  let bytes: Uint8Array;
  try {
    bytes = new Uint8Array(await readFile(path));
  } catch {
    throw new Error(`Cannot read the card image file: ${path}`);
  }
  if (looksLikeSvg(path, bytes)) throw new Error(`${raw}: ${SVG_REFUSAL}.`);
  const header = parseImageHeader(bytes);
  if (!header) throw new Error(`${raw} is not a PNG, JPEG, or GIF image. Cards take only those types.`);
  const mime = header.kind === "jpeg" ? "image/jpeg" : `image/${header.kind}`;
  return { uri: `data:${mime};base64,${Buffer.from(bytes).toString("base64")}`, image: { path: raw, type: header.kind, bytes: bytes.length } };
}

/**
 * Replaces each `file:<path>` in an Image `url` or a `backgroundImage` with a data URI. The path is relative to the
 * card file. The card is not changed in place. The images come back in document order.
 */
export async function inlineCardImages(card: Record<string, any>, baseDir: string): Promise<{ card: Record<string, any>; images: CardImage[] }> {
  const images: CardImage[] = [];
  const isFile = (value: unknown): value is string => typeof value === "string" && value.startsWith(FILE_PREFIX);

  async function replace(reference: string): Promise<string> {
    const done = await toDataUri(reference, baseDir);
    images.push(done.image);
    return done.uri;
  }

  async function visit(value: any): Promise<any> {
    if (Array.isArray(value)) {
      const out: any[] = [];
      for (const item of value) out.push(await visit(item));
      return out;
    }
    if (!isObject(value)) return value;
    const out: Record<string, any> = {};
    for (const [key, child] of Object.entries(value)) {
      if (key === "url" && value.type === "Image" && isFile(child)) out[key] = await replace(child);
      else if (key === "backgroundImage" && isFile(child)) out[key] = await replace(child);
      else if (key === "backgroundImage" && isObject(child) && isFile(child.url)) out[key] = { ...child, url: await replace(child.url) };
      else out[key] = await visit(child);
    }
    return out;
  }

  return { card: await visit(card), images };
}

export function cardBaseDir(cardFile: string): string {
  return dirname(resolve(cardFile));
}
