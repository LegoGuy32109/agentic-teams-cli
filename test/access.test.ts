import { afterEach, beforeEach, expect, test } from "bun:test";
import { knownDenial, ACCESS_TTL_MS } from "../src/access.ts";
import { identity } from "../src/tokens.ts";
import { unread } from "../src/commands/unread.ts";
import { request } from "../src/transport.ts";
import { seedProfile, resetLocalState, installFakeTeams } from "./fake-teams.ts";

const ID="19:denied@thread.v2";
const HISTORY=`GET /api/chatsvc/amer/v1/users/ME/conversations/${ID}/messages`;
const AGGREGATE="GET /api/csa/amer/api/v3/teams/users/me";
let fake: ReturnType<typeof installFakeTeams> | undefined;
let originalLog: typeof console.log;
beforeEach(async()=>{originalLog=console.log;await resetLocalState();await seedProfile(["chat","csa"]);});
afterEach(()=>{console.log=originalLog;fake?.uninstall();fake=undefined;});

test("an observed history denial is separate from unread conversations without repeating the request",async()=>{
  fake=installFakeTeams({[HISTORY]:{status:403,json:{subCode:"AclCheckFailed"}},[AGGREGATE]:{status:200,json:{teams:[],chats:[{id:ID,title:"Denied",isRead:false}]}}});
  await expect(request("chat","GET",`/users/ME/conversations/${encodeURIComponent(ID)}/messages`)).rejects.toThrow("(403)");
  let output="";console.log=v=>{output=String(v);};await unread(true);
  const d=JSON.parse(output);expect(d.conversations).toHaveLength(0);expect(d.unavailable).toHaveLength(1);
  expect(d.unavailable[0].status).toBe(403);
  expect(fake.requests.filter(r=>r.path.endsWith("/messages"))).toHaveLength(1);
});

test("a refresh restores the normal unread list when access returns",async()=>{
  fake=installFakeTeams({[HISTORY]:[{status:403,json:{subCode:"AclCheckFailed"}},{status:200,json:{messages:[]}}],[AGGREGATE]:{status:200,json:{teams:[],chats:[{id:ID,title:"Denied",isRead:false}]}}});
  await request("chat","GET",`/users/ME/conversations/${encodeURIComponent(ID)}/messages`).catch(()=>{});
  let output="";console.log=v=>{output=String(v);};await unread(true,true);
  const d=JSON.parse(output);expect(d.conversations).toHaveLength(1);expect(d.unavailable).toHaveLength(0);
  expect(fake.requests.every(r=>r.method==="GET")).toBe(true);
});

test("denials expire and are isolated by account and profile",async()=>{
  fake=installFakeTeams({[HISTORY]:{status:403,json:{subCode:"AclCheckFailed"}}});
  await request("chat","GET",`/users/ME/conversations/${encodeURIComponent(ID)}/messages`).catch(()=>{});
  const who=await identity();const record=await knownDenial(who,ID);expect(record).toBeDefined();
  expect(await knownDenial({...who,profile:"another-profile"},ID)).toBeUndefined();
  expect(await knownDenial({...who,mri:"another-account"},ID)).toBeUndefined();
  expect(await knownDenial(who,ID,Date.parse(record!.checked_at)+ACCESS_TTL_MS+1)).toBeUndefined();
});

test("a single-message denial does not mark the whole conversation unavailable",async()=>{
  fake=installFakeTeams({[`${HISTORY}/1791990000000`]:{status:403,json:{subCode:"AclCheckFailed"}}});
  await request("chat","GET",`/users/ME/conversations/${encodeURIComponent(ID)}/messages/1791990000000`).catch(()=>{});
  expect(await knownDenial(await identity(),ID)).toBeUndefined();
});
