// This helper replaces fetch completely. It cannot make a real network request.
import { appendFile } from "node:fs/promises";
import { fakeJwt } from "./jwt.ts";
import { refreshToken } from "../src/refresh.ts";

globalThis.fetch = (async () => {
  await appendFile(process.argv[2]!, "renewal\n");
  await Bun.sleep(50);
  return Response.json({ access_token: fakeJwt(), refresh_token: "fake-rotated-by-worker" });
}) as unknown as typeof fetch;
await refreshToken("chat");
