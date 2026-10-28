import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { attachPreviews, fetchPreview, toPreview } from "../src/preview.ts";
import type { LinkEntry } from "../src/compose.ts";
import { resetLocalState, seedProfile } from "./fake-teams.ts";
import { AUD, fakeJwt } from "./jwt.ts";

const INFO = {
  url: "https://x.test/a", title: "Claude Artifact", description: "Try out Artifacts created by Claude users",
  thumbnail: "https://us-prod.asyncgw.teams.microsoft.com/urlp/v1/url/image/Thumbnail?url=https%3a%2f%2fx.test%2fi.png",
  thumbnail_meta: { width: 320, height: 180 }, largethumbnail: "ignored", user_pic: "",
};
const link = (url: string, index = 0): LinkEntry => ({ "@type": "http://schema.skype.com/HyperLink", itemid: String(index), url, previewenabled: true });

let original: typeof fetch;
let calls: Array<{ url: string; authorization: string | null }>;
beforeEach(async () => {
  original = globalThis.fetch;
  calls = [];
  await resetLocalState();
  await seedProfile(["chat"]);
});
afterEach(() => { globalThis.fetch = original; });

function serve(reply: (url: URL) => Response | Promise<Response>) {
  globalThis.fetch = (async (input: any, init: RequestInit = {}) => {
    const url = new URL(String(input));
    calls.push({ url: String(input), authorization: new Headers(init.headers).get("authorization") });
    return reply(url);
  }) as typeof fetch;
}

describe("preview data", () => {
  test("url/info becomes the preview object that the web client stores", () => {
    expect(toPreview(INFO)).toEqual({ previewurl: INFO.thumbnail, previewmeta: { width: 320, height: 180 }, title: "Claude Artifact", description: "Try out Artifacts created by Claude users" });
  });

  test("a page with no title has no card, and a page with no thumbnail has a text card", () => {
    expect(toPreview({ url: "x" })).toBeUndefined();
    expect(toPreview({ title: "" })).toBeUndefined();
    expect(toPreview(null)).toBeUndefined();
    expect(toPreview({ title: "T" })).toEqual({ title: "T", description: "" });
  });
});

describe("fetching", () => {
  test("it asks the amer preview host with the chat token and the encoded URL", async () => {
    serve(() => Response.json(INFO));
    expect((await fetchPreview("https://x.test/a?b=1&c=2", "amer"))?.title).toBe("Claude Artifact");
    expect(calls).toHaveLength(1);
    expect(calls[0]!.url).toBe("https://us-prod.asyncgw.teams.microsoft.com/urlp/v1/url/info?url=https%3A%2F%2Fx.test%2Fa%3Fb%3D1%26c%3D2");
    expect(calls[0]!.authorization).toMatch(/^Bearer eyJ/);
  });

  test("an unknown region, an error status, a network error and bad JSON all give no preview", async () => {
    serve(() => Response.json(INFO));
    expect(await fetchPreview("https://x.test/a", "emea")).toBeUndefined();
    expect(calls).toHaveLength(0);
    serve(() => Response.json(INFO, { status: 401 }));
    expect(await fetchPreview("https://x.test/a", "amer")).toBeUndefined();
    serve(() => { throw new TypeError("fetch failed"); });
    expect(await fetchPreview("https://x.test/a", "amer")).toBeUndefined();
    serve(() => new Response("<html>", { status: 200 }));
    expect(await fetchPreview("https://x.test/a", "amer")).toBeUndefined();
  });

  test("a missing chat token gives no preview instead of an error", async () => {
    await resetLocalState();
    serve(() => Response.json(INFO));
    expect(await fetchPreview("https://x.test/a", "amer")).toBeUndefined();
    expect(calls).toHaveLength(0);
  });
});

describe("attaching", () => {
  test("the first link gets the card and the others get none", async () => {
    serve(() => Response.json(INFO));
    const links = [link("https://x.test/a"), link("https://x.test/b", 1)];
    expect(await attachPreviews(links, "amer")).toBe(true);
    expect(links[0]).toMatchObject({ previewenabled: true, preview: { title: "Claude Artifact" } });
    expect(links[1]!.previewenabled).toBe(false);
    expect(links[1]!.preview).toBeUndefined();
    expect(calls).toHaveLength(1);
  });

  test("a failed preview turns the flag off and reports it", async () => {
    serve(() => new Response("no", { status: 500 }));
    const links = [link("https://x.test/a")];
    expect(await attachPreviews(links, "amer")).toBe(false);
    expect(links[0]).toMatchObject({ previewenabled: false });
    expect(links[0]!.preview).toBeUndefined();
  });

  test("no links means no request", async () => {
    serve(() => Response.json(INFO));
    expect(await attachPreviews([], "amer")).toBe(true);
    expect(calls).toHaveLength(0);
  });
});

describe("turning cards off", () => {
  test("with cards off, no link gets a card and nothing is read", async () => {
    serve(() => Response.json(INFO));
    const links = [link("https://x.test/a"), link("https://x.test/b", 1)];
    expect(await attachPreviews(links, "amer", false)).toBe(true);
    expect(calls).toHaveLength(0);
    expect(links.map(item => item.previewenabled)).toEqual([false, false]);
    expect(links[0]!.preview).toBeUndefined();
  });
});
