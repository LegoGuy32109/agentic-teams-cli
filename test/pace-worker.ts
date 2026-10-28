// Runs in a child process. It tries `count` sends with --no-wait and prints how many got through.
import { paceWrite, RateLimitError } from "../src/pace.ts";

const [tenant, countRaw] = process.argv.slice(2);
let accepted = 0;
let refused = 0;
for (let index = 0; index < Number(countRaw); index++) {
  try {
    await paceWrite(tenant!, { noWait: true });
    accepted++;
  } catch (error) {
    if (error instanceof RateLimitError) refused++;
    else throw error;
  }
}
console.log(JSON.stringify({ accepted, refused }));
