import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { existsSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { send } from "../src/commands/send.ts";
import { cacheDir } from "../src/config.ts";
import { displaySize, loadImage, parseImageHeader } from "../src/image.ts";
import { ledgerKey, LedgerRefusal } from "../src/ledger.ts";
import { htmlToText } from "../src/message.ts";
import { installFakeTeams, resetLocalState, seedProfile, type Route } from "./fake-teams.ts";
import { runCli } from "./cli.ts";
import { TID } from "./jwt.ts";

const SEND = "POST /api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages";
const CREATE = "POST /v1/objects/";
const UPLOAD = "PUT /v1/objects/{id}/content/imgpsh";
const OBJECT_ID = "0-cus-d5-00000000000000000000000000000001";
const OK_SEND = { status: 201, json: { OriginalArrivalTime: 1791990424088 } };
const OK_CREATE = { status: 201, json: { id: OBJECT_ID } };
const OK_UPLOAD = { status: 201 };

const dir = mkdtempSync(join(tmpdir(), "teams-image-"));

function png(width: number, height: number, filler = 0): Uint8Array {
  const bytes = new Uint8Array(40 + filler);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const view = new DataView(bytes.buffer);
  view.setUint32(8, 13);
  bytes.set([0x49, 0x48, 0x44, 0x52], 12);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return bytes;
}

function gif(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(20);
  bytes.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]);
  new DataView(bytes.buffer).setUint16(6, width, true);
  new DataView(bytes.buffer).setUint16(8, height, true);
  return bytes;
}

/** A JPEG with an APP0 segment before the frame header, as a camera file has. */
function jpeg(width: number, height: number, marker = 0xc0): Uint8Array {
  const bytes = new Uint8Array(2 + 2 + 16 + 2 + 2 + 15);
  bytes.set([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10]);
  const frame = 2 + 2 + 16;
  bytes.set([0xff, marker, 0x00, 0x0b, 0x08], frame);
  const view = new DataView(bytes.buffer);
  view.setUint16(frame + 5, height);
  view.setUint16(frame + 7, width);
  return bytes;
}

function file(name: string, bytes: Uint8Array | string): string {
  const path = join(dir, name);
  writeFileSync(path, bytes);
  return path;
}

describe("image header parsing", () => {
  test("reads PNG, GIF, and JPEG sizes from the header", () => {
    expect(parseImageHeader(png(800, 600))).toEqual({ kind: "png", width: 800, height: 600 });
    expect(parseImageHeader(gif(120, 90))).toEqual({ kind: "gif", width: 120, height: 90 });
    expect(parseImageHeader(jpeg(1024, 768))).toEqual({ kind: "jpeg", width: 1024, height: 768 });
    expect(parseImageHeader(jpeg(640, 480, 0xc2))).toEqual({ kind: "jpeg", width: 640, height: 480 });
  });

  test("returns nothing for other files and cut-off headers", () => {
    expect(parseImageHeader(new TextEncoder().encode("hello world, this is text"))).toBeUndefined();
    expect(parseImageHeader(png(1, 1).subarray(0, 12))).toBeUndefined();
    expect(parseImageHeader(jpeg(5, 5).subarray(0, 12))).toBeUndefined();
  });

  test("the display height is at most 250 and the ratio holds", () => {
    expect(displaySize(1000, 500)).toEqual({ width: 500, height: 250 });
    expect(displaySize(300, 250)).toEqual({ width: 300, height: 250 });
    expect(displaySize(40, 30)).toEqual({ width: 40, height: 30 });
  });

  test("loadImage refuses text, an empty picture, and a file over 25 MB", async () => {
    await expect(loadImage(file("note.txt", "not an image at all, just text"))).rejects.toThrow("is not a PNG, JPEG, or GIF image.");
    await expect(loadImage(file("empty.png", png(0, 10)))).rejects.toThrow("no readable width and height");
    await expect(loadImage(file("huge.png", png(10, 10, 25 * 1024 * 1024)))).rejects.toThrow("The limit is 25 MB.");
    await expect(loadImage(join(dir, "missing.png"))).rejects.toThrow("Cannot read the image file");
  });
});

describe("history text", () => {
  test("an AMS image shows its object id and another image stays [image]", () => {
    const tag = `<p><img src="https://x.test/v1/objects/${OBJECT_ID}/views/imgo" itemtype="http://schema.skype.com/AMSImage" itemid="${OBJECT_ID}"></p>`;
    expect(htmlToText(tag)).toBe(`[image ${OBJECT_ID}]`);
    expect(htmlToText("<img src=\"x.png\">")).toBe("[image]");
  });
});

let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
const originalLog = console.log;
const originalError = console.error;
const plain = { json: false, verify: false, allowWrite: false } as const;

function serve(routes: Record<string, Route>) {
  fake?.uninstall();
  fake = installFakeTeams(routes);
}

beforeEach(async () => {
  await resetLocalState();
  await seedProfile();
  logs = [];
  console.log = (...args: unknown[]) => { logs.push(args.map(String).join(" ")); };
  console.error = () => {};
});

afterEach(() => {
  console.log = originalLog;
  console.error = originalError;
  fake?.uninstall();
  fake = undefined;
  delete process.env.TEAMS_REGION;
  process.exitCode = 0;
});

describe("preview", () => {
  test("a preview uploads nothing, sends nothing, and shows placeholder ids", async () => {
    serve({ [CREATE]: OK_CREATE, [UPLOAD]: OK_UPLOAD, [SEND]: OK_SEND });
    const path = file("small.png", png(100, 50));
    await send("me", ["image test"], { ...plain, json: true, images: [path] });
    expect(fake!.requests).toHaveLength(0);
    expect(existsSync(join(cacheDir(), TID, "ledger"))).toBe(false);
    const preview = JSON.parse(logs.join("\n"));
    expect(preview.images).toEqual([{ file: "small.png", bytes: 40, type: "png", width: 100, height: 50, display: { width: 100, height: 50 } }]);
    expect(preview.body.amsreferences).toEqual(["0-cus-d5-<image-1>"]);
    expect(preview.body.content).toStartWith("<p>image test</p><p><img src=\"https://us-api.asm.skype.com/v1/objects/0-cus-d5-<image-1>/views/imgo\"");
  });

  test("the text line names the files, sizes, and dimensions", async () => {
    serve({});
    await send("me", ["hi"], { ...plain, images: [file("a.png", png(10, 20)), file("b.gif", gif(5, 6))] });
    expect(logs[0]).toBe("Dry run (effect: message): would send to Notes to self: hi [2 images: a.png 40 bytes 10x20, b.gif 20 bytes 5x6; no upload in a preview]");
  });

  test("an image with no text is a valid message", async () => {
    serve({});
    await send("me", [], { ...plain, json: true, images: [file("only.png", png(10, 10))] });
    expect(JSON.parse(logs.join("\n")).body.content).toStartWith("<p><img ");
  });

  test("a mention and rich text still compose before the images", async () => {
    serve({});
    await send("me", ["**bold** text"], { ...plain, json: true, format: "rich", images: [file("r.png", png(10, 10))] });
    expect(JSON.parse(logs.join("\n")).body.content).toStartWith("<p><strong>bold</strong> text</p><p><img ");
  });

  test("a region other than amer is refused with a clear error", async () => {
    process.env.TEAMS_REGION = "emea";
    serve({});
    await expect(send("me", ["x"], { ...plain, images: [file("e.png", png(1, 1))] })).rejects.toThrow("Inline images are verified only for the amer region. This profile uses emea.");
  });

  test("more than 10 images is refused", async () => {
    serve({});
    const one = file("many.png", png(1, 1));
    await expect(send("me", ["x"], { ...plain, images: Array(11).fill(one) })).rejects.toThrow("10 images at most");
  });
});

describe("send with images", () => {
  test("creates the object, uploads the bytes, then posts the message with the object id", async () => {
    serve({ [CREATE]: OK_CREATE, [UPLOAD]: OK_UPLOAD, [SEND]: OK_SEND });
    const bytes = png(1000, 500, 8);
    await send("me", ["image test"], { ...plain, allowWrite: true, images: [file("big.png", bytes)] });
    expect(fake!.requests.map(item => `${item.method} ${item.host}${item.path}`)).toEqual([
      "POST us-prod.asyncgw.teams.microsoft.com/v1/objects/",
      `PUT us-prod.asyncgw.teams.microsoft.com/v1/objects/${OBJECT_ID}/content/imgpsh`,
      "POST teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages",
    ]);
    const [create, upload, message] = fake!.requests;
    expect(create!.body).toEqual({ type: "pish/image", permissions: { "48:notes": ["read"] }, sharingMode: "Inline", filename: "big.png" });
    expect(create!.headers["x-ams-post-sharing-mode"]).toBe("Inline");
    expect(upload!.headers["content-type"]).toBe("application/octet-stream");
    expect(upload!.body).toEqual(bytes);
    const body = message!.body as any;
    expect(body.amsreferences).toEqual([OBJECT_ID]);
    const url = `https://us-api.asm.skype.com/v1/objects/${OBJECT_ID}/views/imgo`;
    expect(body.content).toBe(
      `<p>image test</p><p><img src="${url}" itemtype="http://schema.skype.com/AMSImage" itemscope="png" itemid="${OBJECT_ID}" id="${OBJECT_ID}" width="500" height="250" alt="image" href="${url}" target-src="${url}"></p>`,
    );
    expect(logs[0]).toBe("Sent to Notes to self as 48:notes#1791990424088");
  });

  test("a failed upload sends no message, says so, and leaves no send-once record", async () => {
    serve({ [CREATE]: OK_CREATE, [UPLOAD]: { status: 400, json: { message: "bad image" } }, [SEND]: OK_SEND });
    const path = file("fail.png", png(10, 10));
    const error = await send("me", ["x"], { ...plain, allowWrite: true, images: [path] }).then(() => undefined, e => e as Error);
    expect(error!.message).toContain("Upload of fail.png failed. No message was sent. An uploaded image with no message is harmless.");
    expect(fake!.requests.some(item => item.path.endsWith("/messages"))).toBe(false);
    // The same send can run again at once.
    serve({ [CREATE]: OK_CREATE, [UPLOAD]: OK_UPLOAD, [SEND]: OK_SEND });
    await send("me", ["x"], { ...plain, allowWrite: true, images: [path] });
    expect(fake!.requests.filter(item => item.path.endsWith("/messages"))).toHaveLength(1);
  });

  test("a failed object create sends no message", async () => {
    serve({ [CREATE]: { status: 500, text: "boom" }, [SEND]: OK_SEND });
    await expect(send("me", ["x"], { ...plain, allowWrite: true, images: [file("c.png", png(10, 10))] })).rejects.toThrow("No message was sent.");
    expect(fake!.requests.some(item => item.path.endsWith("/messages"))).toBe(false);
  });

  test("the ledger key includes the image digest", async () => {
    serve({ [CREATE]: OK_CREATE, [UPLOAD]: OK_UPLOAD, [SEND]: OK_SEND });
    const first = file("one.png", png(10, 10, 1));
    const second = file("two.png", png(10, 10, 2));
    await send("me", ["same text"], { ...plain, allowWrite: true, images: [first] });
    // Same text, different picture: allowed.
    await send("me", ["same text"], { ...plain, allowWrite: true, images: [second] });
    // Same text and the same picture: refused.
    const refusal = await send("me", ["same text"], { ...plain, allowWrite: true, images: [first] }).then(() => undefined, e => e as Error);
    expect(refusal).toBeInstanceOf(LedgerRefusal);
    expect(fake!.requests.filter(item => item.path.endsWith("/messages"))).toHaveLength(2);
  });

  test("a key with no images matches the key from before this feature", () => {
    const parts = { conversationId: "48:notes", content: "<p>x</p>", mentionMris: [], contentType: "Text" };
    expect(ledgerKey({ ...parts, images: [] })).toBe(ledgerKey(parts));
    expect(ledgerKey({ ...parts, images: ["abc"] })).not.toBe(ledgerKey(parts));
  });
});

describe("command line", () => {
  test("the flag parser reports an unknown flag and a missing --image value", async () => {
    const result = await runCli(["send", "me", "--image", file("l1.png", png(1, 1)), `--image=${file("l2.png", png(1, 1))}`, "--bogus"]);
    expect(result.stderr).toContain("Unknown flag `--bogus`");
    const missing = await runCli(["send", "me", "hi", "--image"]);
    expect(missing.stderr).toContain("--image requires a value.");
  });
});
