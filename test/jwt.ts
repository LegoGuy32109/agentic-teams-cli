// Fake JWTs for tests. The signature is a placeholder, and nothing verifies it.
export const OID = "11111111-1111-1111-1111-111111111111";
export const TID = "22222222-2222-2222-2222-222222222222";

export const AUD = {
  chat: "https://ic3.teams.office.com",
  csa: "https://chatsvcagg.teams.microsoft.com",
  spaces: "https://api.spaces.skype.com",
  sharepoint: "00000003-0000-0ff1-ce00-000000000000",
  substrate: "https://substrate.office.com",
} as const;

function base64url(text: string): string {
  return Buffer.from(text, "utf8").toString("base64url");
}

/** Builds an unsigned JWT. Claims override the defaults. Default audience is the chat kind. */
export function fakeJwt(claims: Record<string, unknown> = {}): string {
  const payload = {
    aud: AUD.chat,
    exp: Math.floor(Date.now() / 1000) + 3600,
    oid: OID,
    tid: TID,
    upn: "tester@example.test",
    name: "Test User",
    ...claims,
  };
  return `${base64url(JSON.stringify({ alg: "none", typ: "JWT" }))}.${base64url(JSON.stringify(payload))}.sig`;
}
