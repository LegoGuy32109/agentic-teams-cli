import template from "../../docs/card-router-flow.template.json";
import { checkRouterUrl } from "../card.ts";
import { findWorkflowsChat } from "../card-ref.ts";
import { loadProfile, saveCardRouterUrl } from "../config.ts";
import { findChannel, labelFor, loadConversations, resolveConversation } from "../conversations.ts";
import { isChannelId } from "../refs.ts";
import { decodeJwt, identity, redact } from "../tokens.ts";

export const FLOW_NAME = "teams-cli card router";
export const FLOW_AUDIENCE = "https://service.flow.microsoft.com/";
export const POWERAPPS_AUDIENCE = "https://service.powerapps.com/";
export const FLOW_TOKEN_VARIABLE = "TEAMS_FLOW_TOKEN";
export const POWERAPPS_TOKEN_VARIABLE = "TEAMS_POWERAPPS_TOKEN";
const FLOW_API = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple";
const CONNECTIONS_API = "https://api.powerapps.com/providers/Microsoft.PowerApps/apis/shared_teams/connections";
const API_VERSION = "2016-11-01";
const TEAMS_API_ID = "/providers/Microsoft.PowerApps/apis/shared_teams";
const TIMEOUT_MS = 30_000;

export type CommandResult = { code: number | null; stdout: string; stderr: string };
/** Runs a program. It throws an error with `code: "ENOENT"` when the program is not installed. */
export type CommandRunner = (command: string, args: string[]) => Promise<CommandResult>;

const defaultRunner: CommandRunner = async (command, args) => {
  const child = Bun.spawn([command, ...args], { stdout: "pipe", stderr: "pipe", stdin: "ignore" });
  const [stdout, stderr] = await Promise.all([new Response(child.stdout).text(), new Response(child.stderr).text()]);
  return { code: await child.exited, stdout, stderr };
};
let runner: CommandRunner = defaultRunner;

/** Replaces the program runner. Tests call it, so no test starts a real `az`. */
export function setCommandRunnerForTests(fn: CommandRunner | undefined): void {
  runner = fn ?? defaultRunner;
}

export type SetupOptions = { json: boolean; allowWrite: boolean; channels: string[] };

function noToken(variable: string, audience: string, reason: string): Error {
  return new Error(`${reason} The card router setup needs a token for ${audience}. Do one of these: set ${variable} to such a token, or sign in with the Azure CLI (az login) to your work tenant.`);
}

/**
 * Finds a token for the Flow or Power Apps API, in this order: the environment variable, then `az account get-access-token`.
 * The token must belong to the tenant of the Teams profile. Nothing here prints a token.
 */
export async function apiToken(audience: string, variable: string, tenantId: string): Promise<string> {
  const fromEnv = process.env[variable]?.trim();
  if (fromEnv) {
    const claims = decodeJwt(fromEnv);
    if (!claims) throw new Error(`${variable} is not a JWT.`);
    if (claims.tid && claims.tid !== tenantId) throw new Error(`${variable} is for another tenant than your Teams profile (${tenantId}).`);
    return fromEnv;
  }
  let result: CommandResult;
  try {
    result = await runner("az", ["account", "get-access-token", "--resource", audience, "--output", "json"]);
  } catch (error) {
    if ((error as NodeJS.ErrnoException)?.code === "ENOENT") throw noToken(variable, audience, `${variable} is not set, and the Azure CLI (az) is not on PATH.`);
    throw new Error(`Could not run az: ${redact(error instanceof Error ? error.message : String(error))}`);
  }
  if (result.code !== 0) throw noToken(variable, audience, `${variable} is not set, and az is not signed in (az account get-access-token failed).`);
  let parsed: any;
  try { parsed = JSON.parse(result.stdout); } catch { parsed = undefined; }
  if (typeof parsed?.accessToken !== "string" || !parsed.accessToken) throw new Error("az answered with no accessToken.");
  if (parsed.tenant !== tenantId) {
    throw new Error(`az is signed in to tenant ${parsed.tenant ?? "(unknown)"}, but your Teams profile is in ${tenantId}. Run az login --tenant ${tenantId}, or set ${variable}.`);
  }
  return parsed.accessToken;
}

async function apiRequest(method: string, url: string, token: string, body?: unknown, quiet = false): Promise<any> {
  let response: Response;
  try {
    response = await fetch(url, {
      method,
      headers: { authorization: `Bearer ${token}`, accept: "application/json", ...(body === undefined ? {} : { "content-type": "application/json" }) },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      redirect: "error",
      signal: AbortSignal.timeout(TIMEOUT_MS),
    });
  } catch {
    throw new Error(`${method} ${new URL(url).pathname} failed with a network error.`);
  }
  const text = await response.text().catch(() => "");
  if (!response.ok) {
    // The callback URL answer holds a secret, so a failure of that call shows only the status.
    const detail = quiet ? "" : ` ${redact(text).replace(/\s+/g, " ").trim().slice(0, 200)}`;
    throw new Error(`${method} ${new URL(url).pathname} answered ${response.status}.${detail}`.trimEnd());
  }
  try { return text ? JSON.parse(text) : undefined; } catch { return undefined; }
}

export type Allowlist = Record<string, string>;
export type PlannedChannel = { id: string; label: string; groupId: string; existing: boolean };

/** Turns each `--channel` into a channel id and the group id of its team. */
export async function resolveChannels(inputs: string[]): Promise<Array<{ id: string; label: string; groupId: string }>> {
  const found: Array<{ id: string; label: string; groupId: string }> = [];
  for (const input of inputs) {
    const conversation = await resolveConversation(input);
    if (!isChannelId(conversation.id)) throw new Error(`--channel takes a channel. ${conversation.label} is not one.`);
    let { team } = await findChannel(conversation.id);
    if (!team.groupId) ({ team } = await findChannel(conversation.id, { fresh: true }));
    if (!team.groupId) throw new Error(`Teams does not list the group id of the team ${team.name}, so ${conversation.label} cannot go on the allowlist.`);
    found.push({ id: conversation.id, label: conversation.label, groupId: team.groupId });
  }
  return found;
}

function fill(value: any, vars: { upn: string; chat?: string; allowlist: Allowlist }): any {
  if (Array.isArray(value)) return value.map(item => fill(item, vars));
  if (value && typeof value === "object") return Object.fromEntries(Object.entries(value).map(([key, child]) => [key, fill(child, vars)]));
  if (typeof value !== "string") return value;
  if (value.startsWith("{{ALLOWLIST")) return vars.allowlist;
  return value.replaceAll("{{USER_UPN}}", vars.upn).replaceAll("{{FLOW_BOT_CHAT_ID}}", vars.chat ?? "");
}

/**
 * The flow definition from the shipped template. `{{USER_UPN}}`, `{{FLOW_BOT_CHAT_ID}}` and `{{ALLOWLIST}}` are filled.
 * With no Workflows chat, the `update-draft` case is left out.
 */
export function buildDefinition(vars: { upn: string; chat?: string; allowlist: Allowlist }): any {
  const definition = fill(structuredClone(template), vars);
  if (!vars.chat) delete definition.actions.Route.cases["update-draft"];
  return definition;
}

export function modesOf(definition: any): string[] {
  return [...Object.keys(definition.actions.Route.cases)];
}

async function defaultEnvironment(token: string): Promise<{ name: string; displayName: string }> {
  const json = await apiRequest("GET", `${FLOW_API}/environments?api-version=${API_VERSION}`, token);
  const list: any[] = Array.isArray(json?.value) ? json.value : [];
  const found = list.find(item => item?.properties?.isDefault === true);
  if (!found || typeof found.name !== "string") throw new Error("The Flow API lists no default environment for you.");
  return { name: found.name, displayName: String(found.properties?.displayName ?? found.name) };
}

async function teamsConnection(token: string, environment: string): Promise<string> {
  const filter = encodeURIComponent(`environment eq '${environment}'`);
  const json = await apiRequest("GET", `${CONNECTIONS_API}?api-version=${API_VERSION}&$filter=${filter}`, token);
  const list: any[] = Array.isArray(json?.value) ? json.value : [];
  const connected = list.find(item => typeof item?.name === "string" && (item.properties?.statuses ?? []).some((status: any) => status?.status === "Connected"));
  if (!connected) throw new Error("You have no connected Microsoft Teams connection in Power Automate. Open Workflows in Teams once, then run this command again.");
  return connected.name;
}

type ExistingRouter = { id: string; displayName: string; allowlist: Allowlist };

async function existingRouter(token: string, environment: string): Promise<ExistingRouter | undefined> {
  const json = await apiRequest("GET", `${FLOW_API}/environments/${environment}/flows?api-version=${API_VERSION}`, token);
  const list: any[] = Array.isArray(json?.value) ? json.value : [];
  const found = list.find(item => typeof item?.properties?.displayName === "string" && item.properties.displayName.startsWith(FLOW_NAME));
  if (!found || typeof found.name !== "string") return undefined;
  const detail = await apiRequest("GET", `${FLOW_API}/environments/${environment}/flows/${found.name}?api-version=${API_VERSION}`, token);
  const value = detail?.properties?.definition?.actions?.Init_allow?.inputs?.variables?.[0]?.value;
  const allowlist: Allowlist = value && typeof value === "object" ? Object.fromEntries(Object.entries(value).filter(([, group]) => typeof group === "string")) as Allowlist : {};
  return { id: found.name, displayName: found.properties.displayName, allowlist };
}

/** The webhook URL out of a `listCallbackUrl` answer: `{basePath, queries}`, or a plain `value`/`response.value`/`url`. */
export function callbackUrl(json: any): string | undefined {
  if (typeof json?.basePath === "string") {
    const url = new URL(json.basePath);
    for (const [key, value] of Object.entries(json.queries ?? {})) url.searchParams.set(key, String(value));
    return url.toString();
  }
  for (const candidate of [json?.response?.value, json?.value, json?.url]) if (typeof candidate === "string") return candidate;
  return undefined;
}

/** Previews the card router flow, or with `--allow-write` creates or updates it and keeps its webhook URL in the profile. */
export async function cardSetup(options: SetupOptions): Promise<void> {
  const who = await identity();
  const flowToken = await apiToken(FLOW_AUDIENCE, FLOW_TOKEN_VARIABLE, who.tenantId);
  const appsToken = await apiToken(POWERAPPS_AUDIENCE, POWERAPPS_TOKEN_VARIABLE, who.tenantId);
  const environment = await defaultEnvironment(flowToken);
  const connection = await teamsConnection(appsToken, environment.name);
  const existing = await existingRouter(flowToken, environment.name);
  const chat = await findWorkflowsChat();
  const requested = await resolveChannels(options.channels);

  // The existing allowlist stays. A channel that is named now is added to it.
  const cache = await loadConversations(who.tenantId);
  const allowlist: Allowlist = { ...(existing?.allowlist ?? {}) };
  for (const channel of requested) allowlist[channel.id] = channel.groupId;
  const entries: PlannedChannel[] = Object.entries(allowlist).map(([id, groupId]) => ({
    id, groupId, label: requested.find(channel => channel.id === id)?.label ?? labelFor(cache, id) ?? id, existing: !requested.some(channel => channel.id === id),
  }));
  const definition = buildDefinition({ upn: who.upn, chat, allowlist });
  const modes = modesOf(definition);
  const flowName = existing?.displayName ?? FLOW_NAME;

  if (!options.allowWrite) {
    if (options.json) {
      console.log(JSON.stringify({
        ok: true, dry_run: true, effect: "card-router-setup", environment: { name: environment.name, display_name: environment.displayName },
        connection, flow_name: flowName, action: existing ? "update" : "create", existing_router: existing ? { id: existing.id } : null,
        modes, update_draft: chat !== undefined, allowlist: entries.map(entry => ({ channel_id: entry.id, label: entry.label, group_id: entry.groupId })),
      }, null, 2));
      return;
    }
    console.log(`Dry run (effect: card router setup): would ${existing ? "update" : "create"} the flow.`);
    console.log(`  environment: ${environment.displayName} (${environment.name})`);
    console.log(`  connection: ${connection} (Connected)`);
    console.log(`  flow: ${flowName}`);
    console.log(`  existing router: ${existing ? `found (${existing.id}); its definition is replaced and its allowlist kept` : "none; a new flow is created"}`);
    console.log(`  modes: ${modes.join(", ")}`);
    if (!chat) console.log("  update-draft: left out. Your Workflows chat was not found. Open Workflows in Teams once and run setup again.");
    console.log(`  allowlist: ${entries.length ? `${entries.length} channel${entries.length === 1 ? "" : "s"}` : "empty; post, reply, and update are refused until a channel is added with --channel"}`);
    for (const entry of entries) console.log(`    ${entry.label}${entry.existing ? " (kept)" : ""}: ${entry.id}`);
    console.log("Run again with --allow-write to apply. The webhook URL is stored in your profile and is never shown.");
    return;
  }

  if (!(await loadProfile()).profile) throw new Error("No stored profile exists, so the router URL could not be kept. Run: teams-cli auth");
  const connectionReferences = { shared_teams: { connectionName: connection, source: "Embedded", id: TEAMS_API_ID } };
  let flowId: string;
  if (existing) {
    await apiRequest("PATCH", `${FLOW_API}/environments/${environment.name}/flows/${existing.id}?api-version=${API_VERSION}`, flowToken, { properties: { definition, connectionReferences } });
    flowId = existing.id;
  } else {
    const created = await apiRequest("POST", `${FLOW_API}/environments/${environment.name}/flows?api-version=${API_VERSION}`, flowToken, {
      properties: { displayName: FLOW_NAME, definition, connectionReferences, state: "Started" },
    });
    if (typeof created?.name !== "string" || !created.name) throw new Error("The Flow API did not return the id of the new flow.");
    flowId = created.name;
  }
  const answer = await apiRequest("POST", `${FLOW_API}/environments/${environment.name}/flows/${flowId}/triggers/manual/listCallbackUrl?api-version=${API_VERSION}`, flowToken, undefined, true);
  const raw = callbackUrl(answer);
  if (!raw) throw new Error(`The flow ${flowId} was ${existing ? "updated" : "created"}, but the Flow API gave no webhook URL. Open the flow in Power Automate and copy it to TEAMS_CARD_ROUTER.`);
  checkRouterUrl(raw, "The webhook URL of the new flow");
  await saveCardRouterUrl(raw);

  const message = `Card router ${existing ? "updated" : "created"}. Its webhook URL is stored in your profile (mode 0600) and is not shown. TEAMS_CARD_ROUTER overrides it.`;
  if (options.json) console.log(JSON.stringify({ ok: true, effect: "card-router-setup", action: existing ? "update" : "create", flow_id: flowId, modes, message }, null, 2));
  else console.log(message);
}
