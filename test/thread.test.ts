import { beforeEach, afterEach, expect, test } from "bun:test";
import { send } from "../src/commands/send.ts";
import { deleteMessages } from "../src/commands/delete.ts";
import { loadHistory } from "../src/commands/history.ts";
import { seedProfile, resetLocalState, installFakeTeams } from "./fake-teams.ts";

const CHANNEL="19:fake@thread.tacv2";
const ROOT="1791990000000";
const REPLY="1791990000001";
const CHILD=`${CHANNEL};messageid=${ROOT}`;
const PREFIX="/api/chatsvc/amer/v1/users/ME/conversations/";
let fake: ReturnType<typeof installFakeTeams> | undefined;
let logs: string[]=[];
let original: typeof console.log;
beforeEach(async()=>{original=console.log;logs=[];console.log=v=>{logs.push(String(v));};await resetLocalState();await seedProfile(["chat"]);});
afterEach(()=>{console.log=original;fake?.uninstall();fake=undefined;process.exitCode=0;});
const options={json:true,allowWrite:false,verify:false};

test("a channel post uses the captured Standard post type",async()=>{
  fake=installFakeTeams({});await send(CHANNEL,["fake"],options);
  expect(JSON.parse(logs[0]!).body.postType).toBe("Standard");
  expect(fake.requests).toHaveLength(0);
});

test("a reply preview retains the base channel in the body and addresses the child thread",async()=>{
  fake=installFakeTeams({});await send(CHANNEL,["fake"],{...options,threadTs:ROOT});
  const preview=JSON.parse(logs[0]!);
  expect(preview.path).toContain(encodeURIComponent(CHILD));
  expect(preview.body.conversationid).toBe(CHANNEL);
  expect(preview.body.conversationLink).toBe(`blah/${CHILD}`);
  expect(preview.body.replyChainId).toBe(ROOT);
  expect(preview.body.postType).toBeUndefined();
  expect(fake.requests).toHaveLength(0);
});

test("the reply send verifies through the thread path and returns a reusable child ref",async()=>{
  fake=installFakeTeams({
    [`POST ${PREFIX}${CHILD}/messages`]:{status:201,json:{OriginalArrivalTime:Number(REPLY)}},
    [`GET ${PREFIX}${CHILD}/messages/${REPLY}`]:{status:200,json:{content:"<p>fake</p>"}},
  });
  await send(CHANNEL,["fake"],{...options,threadTs:ROOT,allowWrite:true,verify:true});
  const receipt=JSON.parse(logs[0]!);
  expect(receipt.id).toBe(`${CHILD}#${REPLY}`);
  expect(receipt.reply_to).toBe(ROOT);
  expect(receipt.verification.ok).toBe(true);
  await expect(send(CHANNEL,["fake"],{...options,threadTs:ROOT,allowWrite:true})).rejects.toThrow("Identical message already sent");
  expect(fake.requests.filter(r=>r.method==="POST")).toHaveLength(1);
});

test("the same text in another thread has a different ledger identity",async()=>{
  fake=installFakeTeams({[`POST ${PREFIX}{thread}/messages`]:{status:201,json:{OriginalArrivalTime:Number(REPLY)}}});
  for(const root of [ROOT,"1791990000002"]) await send(CHANNEL,["fake"],{...options,threadTs:root,allowWrite:true});
  expect(fake.requests.filter(r=>r.method==="POST")).toHaveLength(2);
});

test("invalid roots and non-channel replies fail before a write",async()=>{
  fake=installFakeTeams({});
  await expect(send(CHANNEL,["fake"],{...options,threadTs:"bad"})).rejects.toThrow("message id");
  await expect(send("me",["fake"],{...options,threadTs:ROOT})).rejects.toThrow("only with channels");
  expect(fake.requests).toHaveLength(0);
});

test("reply delete preview reads the child but deletes nothing, then the real delete uses that path",async()=>{
  fake=installFakeTeams({
    [`GET ${PREFIX}${CHILD}/messages/${REPLY}`]:{status:200,json:{id:REPLY,content:"<p>fake</p>"}},
    [`DELETE ${PREFIX}${CHILD}/messages/${REPLY}`]:{status:200,json:null},
  });
  await deleteMessages(CHANNEL,[`${CHILD}#${REPLY}`],{json:true,allowWrite:false,threadTs:ROOT});
  expect(fake.requests.filter(r=>r.method==="DELETE")).toHaveLength(0);
  await deleteMessages(CHANNEL,[REPLY],{json:true,allowWrite:true,threadTs:ROOT});
  const deletion=fake.requests.find(r=>r.method==="DELETE");
  expect(deletion?.path).toBe(`${PREFIX}${CHILD}/messages/${REPLY}`);
  expect(deletion?.query.behavior).toBe("softDelete");
});

test("thread history returns child refs suitable for readback and deletion",async()=>{
  fake=installFakeTeams({[`GET ${PREFIX}${CHILD}/messages`]:{status:200,json:{messages:[{id:REPLY,content:"fake",messagetype:"Text",rootMessageId:ROOT}]}}});
  const result=await loadHistory(CHANNEL,{json:true,content:false,thread:ROOT});
  expect(result.messages[0]?.id).toBe(`${CHILD}#${REPLY}`);
  expect(result.messages[0]?.reply_to).toBe(ROOT);
});
