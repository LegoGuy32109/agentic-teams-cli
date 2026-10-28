import { beforeEach, afterEach, expect, test } from "bun:test";
import { installFakeTeams, resetLocalState, seedProfile } from "./fake-teams.ts";
import { unread } from "../src/commands/unread.ts";
import { context } from "../src/commands/context.ts";
import { api } from "../src/commands/api.ts";
import { loadHistory } from "../src/commands/history.ts";

let originalLog: typeof console.log;
let fake: ReturnType<typeof installFakeTeams> | undefined;
beforeEach(async () => {originalLog=console.log;await resetLocalState();await seedProfile(["chat","csa"]);});
afterEach(() => {console.log=originalLog;fake?.uninstall();fake=undefined;});
const message=(id:string)=>({id,messagetype:"Text",content:`fake ${id}`,imdisplayname:"Test",composetime:new Date(Number(id)).toISOString()});

test("unread excludes read conversations and unknown state",async()=>{
  fake=installFakeTeams({"GET /api/csa/amer/api/v3/teams/users/me":{status:200,json:{chats:[{id:"19:unread@thread.v2",isRead:false},{id:"19:read@thread.v2",isRead:true},{id:"19:unknown@thread.v2"}],teams:[]}}});
  let output="";console.log=v=>{output=String(v);};await unread(true);
  expect(JSON.parse(output).conversations.map((c:any)=>c.id)).toEqual(["19:unread@thread.v2"]);
});

test("unread reads the channel flag and object horizons, excluding former memberships",async()=>{
  fake=installFakeTeams({"GET /api/csa/amer/api/v3/teams/users/me":{status:200,json:{chats:[],teams:[{id:"19:team@thread.tacv2",displayName:"Test",channels:[
    {id:"19:unread@thread.tacv2",displayName:"Inbox",isMessageRead:false,isMember:true,consumptionHorizon:{originalArrivalTime:1791990000000}},
    {id:"19:left@thread.tacv2",isMessageRead:false,isMember:false},
    {id:"19:deleted@thread.tacv2",isMessageRead:false,isMember:true,isDeleted:true},
  ]}]}}});
  let output="";console.log=v=>{output=String(v);};await unread(true);
  expect(JSON.parse(output).conversations).toEqual([{id:"19:unread@thread.tacv2",label:"Test/Inbox",kind:"channel",consumptionHorizon:"1791990000000"}]);
});

test("context reads an old anchor through pagination and handles zero surroundings",async()=>{
  const anchor=message("1791990000001");
  fake=installFakeTeams({
    "GET /api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages/1791990000001":{status:200,json:anchor},
    "GET /api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages":{status:200,json:{messages:[message("1791990000004"),message("1791990000003")],_metadata:{backwardLink:"https://teams.cloud.microsoft/page2"}}},
    "GET /page2":{status:200,json:{messages:[message("1791990000003"),message("1791990000002"),anchor,message("1791990000000")]}}
  });
  let output="";console.log=v=>{output=String(v);};
  await context("48:notes#1791990000001",{json:true,before:1,after:1});
  expect(JSON.parse(output).messages.map((m:any)=>m.ts)).toEqual(["1791990000000","1791990000001","1791990000002"]);
  await context("48:notes#1791990000001",{json:true,before:0,after:0});
  expect(JSON.parse(output).messages.map((m:any)=>m.ts)).toEqual(["1791990000001"]);
  expect(fake.requests.every(r=>r.method==="GET")).toBe(true);
});

test("paging overlap does not produce repeated history messages",async()=>{
  fake=installFakeTeams({
    "GET /api/chatsvc/amer/v1/users/ME/conversations/48:notes/messages":{status:200,json:{messages:[message("1791990000002")],_metadata:{backwardLink:"https://teams.cloud.microsoft/page2"}}},
    "GET /page2":{status:200,json:{messages:[message("1791990000002"),message("1791990000001")]}}
  });
  const result=await loadHistory("me",{json:true,content:false,all:true});
  expect(result.messages).toHaveLength(2);
});

test("API catalog rejects arbitrary URLs and calls only the named read",async()=>{
  fake=installFakeTeams({"GET /api/chatsvc/amer/v1/users/ME/properties":{status:200,json:{ok:true}}});
  console.log=()=>{};
  await expect(api("https://evil.test",undefined,true)).rejects.toThrow("Unknown API operation");
  await api("self",undefined,true);
  expect(fake.requests).toHaveLength(1);
  expect(fake.requests[0]?.method).toBe("GET");
});
