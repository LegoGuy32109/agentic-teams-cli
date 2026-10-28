import { describe, expect, test } from "bun:test";
import { encodeConversation, formatMessageRef, isChannelId, isConversationId, parseMessageRef, threadRootId } from "../src/refs.ts";

describe("references", () => {
  test("accepts every conversation id form", () => {
    for (const id of ["19:abc@thread.v2", "19:abc@unq.gbl.spaces", "19:meeting_abc@thread.v2", "19:xyz@thread.tacv2", "19:xyz@thread.skype", "48:notes"]) {
      expect(isConversationId(id)).toBe(true);
    }
    expect(isConversationId("general")).toBe(false);
    expect(isConversationId("#general")).toBe(false);
  });

  test("channel ids are the tacv2 and skype threads", () => {
    expect(isChannelId("19:xyz@thread.tacv2")).toBe(true);
    expect(isChannelId("19:abc@thread.v2")).toBe(false);
  });

  test("a message reference round-trips", () => {
    const ref = formatMessageRef("19:abc@thread.v2", "1791990424088");
    expect(ref).toBe("19:abc@thread.v2#1791990424088");
    expect(parseMessageRef(ref)).toEqual({ conversationId: "19:abc@thread.v2", messageId: "1791990424088" });
  });

  test("a message reference needs a 10 digit or longer id", () => {
    expect(parseMessageRef("19:abc@thread.v2#12345")).toBeUndefined();
    expect(parseMessageRef("19:abc@thread.v2")).toBeUndefined();
    expect(parseMessageRef("#1791990424088")).toBeUndefined();
  });

  test("notes to self parses as a message reference", () => {
    expect(parseMessageRef("48:notes#1791988883194")?.conversationId).toBe("48:notes");
  });

  test("a conversation id encodes colon and at-sign for a path", () => {
    expect(encodeConversation("19:abc@thread.v2")).toBe("19%3Aabc%40thread.v2");
  });
});

test("a thread root is a message id or a full message ref", () => {
  expect(threadRootId("1792428352468", "--thread-ts")).toBe("1792428352468");
  expect(threadRootId("19:abc@thread.tacv2#1792428352468", "--thread-ts")).toBe("1792428352468");
  expect(() => threadRootId("1", "--thread-ts")).toThrow("--thread-ts needs the root message id (10 or more digits) or its ref");
  expect(() => threadRootId("19:abc@thread.tacv2#12", "--thread")).toThrow("--thread needs the root message id");
});
