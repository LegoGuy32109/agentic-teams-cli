import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { readFileSync, statSync } from "node:fs";
import { credentialsPath, loadProfile } from "../src/config.ts";
import { saveConversations } from "../src/conversations.ts";
import { buildDefinition, callbackUrl, cardSetup, setCommandRunnerForTests, type CommandResult } from "../src/commands/card-setup.ts";
import { OID, TID, fakeJwt } from "./jwt.ts";
import { installFakeTeams, resetLocalState, seedProfile, type FakeRequest, type Route } from "./fake-teams.ts";

const CHANNEL = "19:fake@thread.tacv2";
const OLD_CHANNEL = "19:old@thread.tacv2";
const GROUP = "33333333-3333-3333-3333-333333333333";
const BOT_CHAT = "19:aaaa_bbbb@unq.gbl.spaces";
const ENV = "Default-env-1";
const CONNECTION = "shared-teams-abc";
const FLOW_ID = "flow-1";
const SECRET = "SECRETSIG-0123456789";
const HOOK = `https://prod-00.westus.logic.azure.com:443/workflows/abc/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sig=${SECRET}`;
const AZ_TOKEN = fakeJwt({ aud: "https://service.flow.microsoft.com/", tid: TID });
const AZ_TOKEN_APPS = fakeJwt({ aud: "https://service.powerapps.com/", tid: TID });

const P = "/providers/Microsoft.ProcessSimple/environments";
const ENVIRONMENTS = `GET ${P}`;
const CONNECTIONS = "GET /providers/Microsoft.PowerApps/apis/shared_teams/connections";
const FLOWS = `GET ${P}/{env}/flows`;
const FLOW = `GET ${P}/{env}/flows/{id}`;
const CREATE = `POST ${P}/{env}/flows`;
const PATCH = `PATCH ${P}/{env}/flows/{id}`;
const CALLBACK = `POST ${P}/{env}/flows/{id}/triggers/manual/listCallbackUrl`;
const MESSAGES = "GET /api/chatsvc/amer/v1/users/ME/conversations/{id}/messages";

const environments: Route = { status: 200, json: { value: [
  { name: "other", properties: { displayName: "Other", isDefault: false } },
  { name: ENV, properties: { displayName: "Contoso (default)", isDefault: true } },
] } };
const connected = (name = CONNECTION): Route => ({ status: 200, json: { value: [
  { name: "shared-teams-broken", properties: { statuses: [{ status: "Error" }] } },
  { name, properties: { statuses: [{ status: "Connected" }] } },
] } });
const noFlows: Route = { status: 200, json: { value: [{ name: "x", properties: { displayName: "Send webhook alerts" } }] } };
const botChat: Route = request => request.path.includes(BOT_CHAT)
  ? { status: 200, json: { messages: [{ id: "1", imdisplayname: "Workflows" }] } }
  : { status: 200, json: { messages: [] } };

let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[] = [];
let azCalls: string[][] = [];
const originalLog = console.log;
const originalError = console.error;

function az(overrides: Partial<{ tenant: string; code: number; missing: boolean }> = {}) {
  setCommandRunnerForTests(async (command, args) => {
    azCalls.push([command, ...args]);
    if (overrides.missing) throw Object.assign(new Error("spawn az ENOENT"), { code: "ENOENT" });
    const resource = args[args.indexOf("--resource") + 1];
    const result: CommandResult = overrides.code
      ? { code: overrides.code, stdout: "", stderr: "Please run az login" }
      : { code: 0, stdout: JSON.stringify({ accessToken: resource === "https://service.powerapps.com/" ? AZ_TOKEN_APPS : AZ_TOKEN, tenant: overrides.tenant ?? TID }), stderr: "" };
    return result;
  });
}

function serve(routes: Record<string, Route> = {}) {
  fake?.uninstall();
  fake = installFakeTeams({
    [ENVIRONMENTS]: environments, [CONNECTIONS]: connected(), [FLOWS]: noFlows, [MESSAGES]: botChat,
    [CREATE]: { status: 200, json: { name: FLOW_ID } },
    [CALLBACK]: { status: 200, json: { basePath: HOOK.split("?")[0], method: "POST", queries: { "api-version": "2016-06-01", sp: "/triggers/manual/run", sig: SECRET } } },
    ...routes,
  });
  return fake;
}
const output = () => logs.join("\n");
const apiRequests = () => fake!.requests.filter(request => request.host.endsWith(".flow.microsoft.com") || request.host === "api.powerapps.com");
const writes = () => apiRequests().filter(request => request.method !== "GET");
const byMethod = (method: string, suffix = "") => fake!.requests.filter((request: FakeRequest) => request.method === method && request.path.endsWith(suffix));
async function failure(call: Promise<unknown>): Promise<Error> {
  try { await call; } catch (error) { return error as Error; }
  throw new Error("The call did not fail.");
}
const preview = { json: false, allowWrite: false, channels: [] as string[] };

beforeEach(async () => {
  logs = [];
  azCalls = [];
  console.log = (value?: unknown) => { logs.push(String(value)); };
  console.error = () => {};
  await resetLocalState();
  await seedProfile(["chat", "csa", "spaces"]);
  await saveConversations(TID, {
    version: 1, updatedAt: "", chats: [{ id: BOT_CHAT, title: "", type: "chat", isOneOnOne: true, members: [`8:orgid:${OID}`, "28:bot-guid"] }],
    teams: [{ id: "t1", name: "Eng", groupId: GROUP, channels: [{ id: CHANNEL, name: "General" }, { id: OLD_CHANNEL, name: "Old" }] }],
  });
  az();
});

afterEach(() => {
  console.log = originalLog;
  console.error = originalError;
  setCommandRunnerForTests(undefined);
  delete process.env.TEAMS_FLOW_TOKEN;
  delete process.env.TEAMS_POWERAPPS_TOKEN;
  fake?.uninstall();
  fake = undefined;
  process.exitCode = 0;
});

describe("card setup preview", () => {
  test("makes no write, and shows the environment, connection, flow, modes, and allowlist with labels", async () => {
    serve();
    await cardSetup({ ...preview, channels: ["Eng/General"] });
    expect(writes()).toHaveLength(0);
    expect(byMethod("POST")).toHaveLength(0);
    const text = output();
    expect(text).toContain("environment: Contoso (default) (Default-env-1)");
    expect(text).toContain(`connection: ${CONNECTION}`);
    expect(text).toContain("flow: teams-cli card router");
    expect(text).toContain("existing router: none");
    expect(text).toContain("modes: draft, post, reply, update, update-draft, live-demo");
    expect(text).toContain("allowlist: 1 channel");
    expect(text).toContain(`Eng/General: ${CHANNEL}`);
    expect(text).not.toContain(AZ_TOKEN);
  });

  test("the JSON preview carries the same facts", async () => {
    serve();
    await cardSetup({ json: true, allowWrite: false, channels: ["Eng/General"] });
    const preview = JSON.parse(logs[0]!);
    expect(preview).toMatchObject({ dry_run: true, connection: CONNECTION, flow_name: "teams-cli card router", action: "create", existing_router: null, update_draft: true });
    expect(preview.allowlist).toEqual([{ channel_id: CHANNEL, label: "Eng/General", group_id: GROUP }]);
  });

  test("a router with the same name is found and the preview says it will be updated", async () => {
    serve({ [FLOWS]: { status: 200, json: { value: [{ name: FLOW_ID, properties: { displayName: "teams-cli card router (draft, post, reply)" } }] } },
      [FLOW]: { status: 200, json: { properties: { definition: { actions: { Init_allow: { inputs: { variables: [{ value: { [OLD_CHANNEL]: GROUP } }] } } } } } } } });
    await cardSetup({ ...preview, channels: ["Eng/General"] });
    const text = output();
    expect(text).toContain(`found (${FLOW_ID})`);
    expect(text).toContain("flow: teams-cli card router (draft, post, reply)");
    expect(text).toContain("allowlist: 2 channels");
    expect(text).toContain(`Eng/Old (kept): ${OLD_CHANNEL}`);
    expect(writes()).toHaveLength(0);
  });

  test("with no Workflows chat the update-draft case is left out, and the preview says so", async () => {
    await saveConversations(TID, { version: 1, updatedAt: "", chats: [], teams: [{ id: "t1", name: "Eng", groupId: GROUP, channels: [{ id: CHANNEL, name: "General" }] }] });
    serve({ "GET /api/csa/amer/api/v3/teams/users/me": { status: 200, json: { teams: [{ id: "t1", displayName: "Eng", teamSiteInformation: { groupId: GROUP }, channels: [{ id: CHANNEL, displayName: "General" }] }], chats: [] } } });
    await cardSetup({ ...preview, channels: [CHANNEL] });
    expect(output()).toContain("modes: draft, post, reply, update, live-demo");
    expect(output()).toContain("update-draft: left out");
  });

  test("a channel with no known group id is refused, and a chat is not a channel", async () => {
    await saveConversations(TID, { version: 1, updatedAt: "", chats: [], teams: [{ id: "t1", name: "Eng", channels: [{ id: CHANNEL, name: "General" }] }] });
    serve({ "GET /api/csa/amer/api/v3/teams/users/me": { status: 200, json: { teams: [{ id: "t1", displayName: "Eng", channels: [{ id: CHANNEL, displayName: "General" }] }], chats: [] } } });
    expect((await failure(cardSetup({ ...preview, channels: ["Eng/General"] }))).message).toContain("does not list the group id");
    expect((await failure(cardSetup({ ...preview, channels: [BOT_CHAT] }))).message).toContain("--channel takes a channel");
  });

  test("an empty allowlist is called out", async () => {
    serve();
    await cardSetup(preview);
    expect(output()).toContain("allowlist: empty");
  });
});

describe("card setup connection and tokens", () => {
  test("no connected Teams connection stops with advice to open Workflows", async () => {
    serve({ [CONNECTIONS]: { status: 200, json: { value: [{ name: "c", properties: { statuses: [{ status: "Error" }] } }] } } });
    const error = await failure(cardSetup(preview));
    expect(error.message).toContain("Open Workflows in Teams once");
    expect(writes()).toHaveLength(0);
  });

  test("the connection list is read for the default environment with the powerapps token", async () => {
    serve();
    await cardSetup(preview);
    const request = byMethod("GET", "/connections")[0]!;
    expect(request.query.$filter).toBe(`environment eq '${ENV}'`);
    expect(request.headers.authorization).toBe(`Bearer ${AZ_TOKEN_APPS}`);
    expect(byMethod("GET", "/environments")[0]!.headers.authorization).toBe(`Bearer ${AZ_TOKEN}`);
  });

  test("az signed in to another tenant is refused, with no API request", async () => {
    az({ tenant: "99999999-9999-9999-9999-999999999999" });
    serve();
    const error = await failure(cardSetup(preview));
    expect(error.message).toContain("signed in to tenant 99999999-9999-9999-9999-999999999999");
    expect(error.message).toContain(TID);
    expect(apiRequests()).toHaveLength(0);
  });

  test("a missing az names both options", async () => {
    az({ missing: true });
    serve();
    const error = await failure(cardSetup(preview));
    expect(error.message).toContain("TEAMS_FLOW_TOKEN");
    expect(error.message).toContain("Azure CLI");
    expect(apiRequests()).toHaveLength(0);
  });

  test("az that is not signed in names both options", async () => {
    az({ code: 1 });
    serve();
    const error = await failure(cardSetup(preview));
    expect(error.message).toContain("az is not signed in");
    expect(error.message).toContain("TEAMS_FLOW_TOKEN");
  });

  test("the token variables win over az, and a token for another tenant is refused", async () => {
    process.env.TEAMS_FLOW_TOKEN = fakeJwt({ tid: TID });
    process.env.TEAMS_POWERAPPS_TOKEN = fakeJwt({ tid: TID });
    serve();
    await cardSetup(preview);
    expect(azCalls).toHaveLength(0);
    process.env.TEAMS_FLOW_TOKEN = fakeJwt({ tid: "99999999-9999-9999-9999-999999999999" });
    expect((await failure(cardSetup(preview))).message).toContain("another tenant");
  });
});

describe("card setup --allow-write", () => {
  const write = { json: false, allowWrite: true, channels: ["Eng/General"] };

  test("a new router: POST the flow with the template filled, read the callback, store the URL, and never print it", async () => {
    serve();
    await cardSetup(write);
    const created = byMethod("POST", `/${ENV}/flows`);
    expect(created).toHaveLength(1);
    expect(byMethod("PATCH")).toHaveLength(0);
    const body = created[0]!.body as any;
    expect(body.properties.displayName).toBe("teams-cli card router");
    expect(body.properties.connectionReferences.shared_teams.connectionName).toBe(CONNECTION);
    const text = JSON.stringify(body.properties.definition);
    expect(text).not.toContain("{{");
    expect(text).toContain("tester@example.test");
    expect(text).toContain(BOT_CHAT);
    expect(body.properties.definition.actions.Init_allow.inputs.variables[0].value).toEqual({ [CHANNEL]: GROUP });
    const callback = byMethod("POST", "/listCallbackUrl");
    expect(callback).toHaveLength(1);
    expect(callback[0]!.body).toBeUndefined();
    expect(callback[0]!.path).toContain(`/flows/${FLOW_ID}/`);

    const stored = (await loadProfile()).profile!.cardRouterUrl!;
    // The default port is dropped when the URL is parsed.
    expect(stored).toBe(HOOK.replace(":443", ""));
    expect(statSync(credentialsPath()).mode & 0o777).toBe(0o600);
    expect(output()).not.toContain(SECRET);
    expect(output()).not.toContain("logic.azure.com");
    expect(output()).not.toContain(AZ_TOKEN);
    expect(output()).toContain("Card router created");
    expect(readFileSync(credentialsPath(), "utf8")).toContain(SECRET);
  });

  test("an existing router is PATCHed with its allowlist kept and added to, and no new flow is created", async () => {
    serve({
      [FLOWS]: { status: 200, json: { value: [{ name: FLOW_ID, properties: { displayName: "teams-cli card router" } }] } },
      [FLOW]: { status: 200, json: { properties: { definition: { actions: { Init_allow: { inputs: { variables: [{ value: { [OLD_CHANNEL]: GROUP } }] } } } } } } },
      [PATCH]: { status: 200, json: {} },
    });
    await cardSetup(write);
    expect(byMethod("PATCH")).toHaveLength(1);
    expect(byMethod("POST", `/${ENV}/flows`)).toHaveLength(0);
    const patch = byMethod("PATCH")[0]!;
    expect(patch.path).toBe(`${P}/${ENV}/flows/${FLOW_ID}`);
    expect((patch.body as any).properties.definition.actions.Init_allow.inputs.variables[0].value).toEqual({ [OLD_CHANNEL]: GROUP, [CHANNEL]: GROUP });
    expect((patch.body as any).properties.connectionReferences.shared_teams.connectionName).toBe(CONNECTION);
    expect(output()).toContain("Card router updated");
    expect((await loadProfile()).profile!.cardRouterUrl).toContain(SECRET);
  });

  test("a failure of the callback call shows the status only, and nothing is stored", async () => {
    serve({ [CALLBACK]: { status: 500, text: `boom ${HOOK}` } });
    const error = await failure(cardSetup(write));
    expect(error.message).toContain("answered 500");
    expect(error.message).not.toContain(SECRET);
    expect((await loadProfile()).profile!.cardRouterUrl).toBeUndefined();
  });

  test("a callback answer with no URL, or a URL on another host, stores nothing and prints no URL", async () => {
    serve({ [CALLBACK]: { status: 200, json: {} } });
    expect((await failure(cardSetup(write))).message).toContain("gave no webhook URL");
    serve({ [CALLBACK]: { status: 200, json: { response: { value: `https://evil.example.com/x?sig=${SECRET}` } } } });
    const error = await failure(cardSetup(write));
    expect(error.message).toContain("is not accepted");
    expect(error.message).not.toContain(SECRET);
    expect((await loadProfile()).profile!.cardRouterUrl).toBeUndefined();
  });

});

describe("the flow definition", () => {
  test("fills every placeholder, and drops update-draft without a chat", () => {
    const full = buildDefinition({ upn: "me@example.test", chat: BOT_CHAT, allowlist: { [CHANNEL]: GROUP } });
    expect(JSON.stringify(full)).not.toContain("{{");
    expect(Object.keys(full.actions.Route.cases)).toContain("update-draft");
    expect(full.actions.Route.cases["update-draft"].actions.Update_draft.inputs.parameters["body/recipient"]).toBe(BOT_CHAT);
    expect(full.actions.Route.cases.draft.actions.Post_draft_to_Josh.inputs.parameters["body/recipient"]).toBe("me@example.test");
    const none = buildDefinition({ upn: "me@example.test", allowlist: {} });
    expect(Object.keys(none.actions.Route.cases)).not.toContain("update-draft");
    expect(none.actions.Init_allow.inputs.variables[0].value).toEqual({});
  });

  test("callbackUrl reads basePath with queries, and the plain forms", () => {
    expect(callbackUrl({ basePath: "https://h.logic.azure.com/p", queries: { a: "1", sig: "s" } })).toBe("https://h.logic.azure.com/p?a=1&sig=s");
    expect(callbackUrl({ response: { value: "https://x" } })).toBe("https://x");
    expect(callbackUrl({ value: "https://y" })).toBe("https://y");
    expect(callbackUrl({})).toBeUndefined();
  });
});

test("the router template posts to a channel only as the user, never as the Workflows bot", async () => {
  // teams-cli is for one person: every channel message comes from that person. The bot posts only in their own Workflows chat.
  const template = JSON.parse(await Bun.file(new URL("../docs/card-router-flow.template.json", import.meta.url)).text());
  const steps: Array<{ poster: string; location: string }> = [];
  const walk = (node: unknown) => {
    if (Array.isArray(node)) node.forEach(walk);
    else if (node && typeof node === "object") {
      const params = (node as any).inputs?.parameters;
      if (params?.poster) steps.push({ poster: params.poster, location: params.location });
      Object.values(node).forEach(walk);
    }
  };
  walk(template);
  expect(steps.length).toBeGreaterThan(0);
  for (const step of steps.filter(item => item.location === "Channel")) expect(step.poster).toBe("User");
  for (const step of steps.filter(item => item.poster === "Flow bot")) expect(["Chat with Flow bot", "Group chat"]).toContain(step.location);
});
