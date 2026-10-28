/** Splits arguments at the first bare `--`. Flags are read only before it. Everything after it is text. */
export function splitTerminator(args: string[]): { flagArgs: string[]; literalArgs: string[] } {
  const index = args.indexOf("--");
  if (index < 0) return { flagArgs: args, literalArgs: [] };
  return { flagArgs: args.slice(0, index), literalArgs: args.slice(index + 1) };
}
