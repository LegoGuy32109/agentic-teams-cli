import { MAX_IMAGE_BYTES, mb } from "./limits.ts";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { basename } from "node:path";
import { amsHosts, request, TeamsError } from "./transport.ts";

export { MAX_IMAGES, MAX_IMAGE_BYTES } from "./limits.ts";
/** The web client scales an inline image to this display height. */
export const MAX_DISPLAY_HEIGHT = 250;
export const IMAGE_TYPE = "http://schema.skype.com/AMSImage";

export type ImageKind = "png" | "jpeg" | "gif";

export type LoadedImage = {
  path: string;
  name: string;
  kind: ImageKind;
  bytes: Uint8Array;
  size: number;
  width: number;
  height: number;
  digest: string;
};

/** Reads the type and pixel size from the file header. Returns undefined for any other file. */
export function parseImageHeader(bytes: Uint8Array): { kind: ImageKind; width: number; height: number } | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const at = (offset: number) => (offset < bytes.length ? bytes[offset]! : -1);
  if (bytes.length >= 24 && at(0) === 0x89 && at(1) === 0x50 && at(2) === 0x4e && at(3) === 0x47 && at(4) === 0x0d && at(5) === 0x0a && at(6) === 0x1a && at(7) === 0x0a) {
    return { kind: "png", width: view.getUint32(16), height: view.getUint32(20) };
  }
  const gif = String.fromCharCode(...bytes.subarray(0, 6));
  if (bytes.length >= 10 && (gif === "GIF87a" || gif === "GIF89a")) {
    return { kind: "gif", width: view.getUint16(6, true), height: view.getUint16(8, true) };
  }
  if (at(0) === 0xff && at(1) === 0xd8) {
    let offset = 2;
    while (offset + 4 <= bytes.length) {
      if (bytes[offset] !== 0xff) return undefined;
      const marker = bytes[offset + 1]!;
      // Fill bytes and markers without a length.
      if (marker === 0xff) { offset++; continue; }
      if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { offset += 2; continue; }
      const length = view.getUint16(offset + 2);
      const isFrame = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
      if (isFrame) {
        if (offset + 9 > bytes.length) return undefined;
        return { kind: "jpeg", height: view.getUint16(offset + 5), width: view.getUint16(offset + 7) };
      }
      offset += 2 + length;
    }
  }
  return undefined;
}

/** Reads one image file. Refuses a missing file, another file type, an empty picture, and a file over 25 MB. */
export async function loadImage(path: string): Promise<LoadedImage> {
  let bytes: Uint8Array;
  try {
    bytes = new Uint8Array(await readFile(path));
  } catch {
    throw new Error(`Cannot read the image file: ${path}`);
  }
  if (bytes.length > MAX_IMAGE_BYTES) throw new Error(`${path} is ${(bytes.length / 1024 / 1024).toFixed(1)} MB. The limit is ${mb(MAX_IMAGE_BYTES)}.`);
  const header = parseImageHeader(bytes);
  if (!header) throw new Error(`${path} is not a PNG, JPEG, or GIF image.`);
  if (!header.width || !header.height) throw new Error(`${path} has no readable width and height.`);
  return {
    path,
    name: basename(path),
    kind: header.kind,
    bytes,
    size: bytes.length,
    width: header.width,
    height: header.height,
    digest: createHash("sha256").update(bytes).digest("hex"),
  };
}

/** The display size: the picture size, or scaled down to 250 pixels high with the same ratio. */
export function displaySize(width: number, height: number): { width: number; height: number } {
  if (height <= MAX_DISPLAY_HEIGHT) return { width, height };
  return { width: Math.max(1, Math.round((width * MAX_DISPLAY_HEIGHT) / height)), height: MAX_DISPLAY_HEIGHT };
}

export function viewUrl(region: string, id: string): string {
  return `${amsHosts(region).view}/v1/objects/${id}/views/imgo`;
}

/** One paragraph with the browser shape for an inline image. */
export function imageHtml(region: string, id: string, image: LoadedImage): string {
  const url = viewUrl(region, id);
  const size = displaySize(image.width, image.height);
  return `<p><img src="${url}" itemtype="${IMAGE_TYPE}" itemscope="${image.kind}" itemid="${id}" id="${id}" width="${size.width}" height="${size.height}" alt="image" href="${url}" target-src="${url}"></p>`;
}

/** Stands in for an object id in a preview and in the send-once key. */
export function placeholderId(index: number): string {
  return `0-cus-d5-<image-${index + 1}>`;
}

/** Creates an object and uploads the bytes. Returns the object id. Nothing is sent to the conversation. */
export async function uploadImage(conversationId: string, image: LoadedImage): Promise<string> {
  try {
    const created = await request("chat", "POST", "/v1/objects/", {
      ams: "upload",
      headers: { "x-ams-post-sharing-mode": "Inline" },
      body: { type: "pish/image", permissions: { [conversationId]: ["read"] }, sharingMode: "Inline", filename: image.name },
    });
    const id = created?.id;
    if (typeof id !== "string" || !/^[\w-]+$/.test(id)) throw new TeamsError("Teams did not return an object id for the image.");
    await request("chat", "PUT", `/v1/objects/${id}/content/imgpsh`, { ams: "upload", bytes: image.bytes, contentType: "application/octet-stream" });
    return id;
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    throw new Error(`Upload of ${image.name} failed. No message was sent. An uploaded image with no message is harmless. ${reason}`);
  }
}

/** An inline image of a received message: its object id and the type its tag names. */
export type InlineImage = { id: string; scope?: string };

/**
 * Reads the inline images of a message from its HTML content, in order. Only the object id is kept: the download
 * address is built from it, never taken from `src`, because anyone in the conversation writes the content.
 */
export function inlineImagesOf(message: any): InlineImage[] {
  const content = typeof message?.content === "string" ? message.content : "";
  const images: InlineImage[] = [];
  const seen = new Set<string>();
  for (const [tag] of content.matchAll(/<img\b[^>]*>/gi)) {
    if (!/itemtype="http:\/\/schema\.skype\.com\/AMSImage"/i.test(tag)) continue;
    const id = /\bitemid="([^"]+)"/i.exec(tag)?.[1];
    if (!id || !/^[\w-]+$/.test(id) || seen.has(id)) continue;
    seen.add(id);
    const scope = /\bitemscope="([a-z]+)"/i.exec(tag)?.[1]?.toLowerCase();
    images.push({ id, ...(scope ? { scope } : {}) });
  }
  return images;
}

/** Reads the full-size picture of one inline image. */
export async function downloadImage(id: string): Promise<Buffer> {
  return request("chat", "GET", `/v1/objects/${id}/views/imgo`, { ams: "view", binary: true, maxBytes: MAX_IMAGE_BYTES });
}

/** The file extension for downloaded image bytes: from the header, then from the tag, then `bin`. */
export function imageExtension(bytes: Uint8Array, scope?: string): string {
  const kind = parseImageHeader(bytes)?.kind ?? scope;
  return kind === "jpeg" ? "jpg" : kind === "png" || kind === "gif" ? kind : "bin";
}
