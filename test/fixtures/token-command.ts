// Stands in for TEAMS_TOKEN_COMMAND in tests. It prints the JSON file named by an env variable.
process.stdout.write(await Bun.file(process.env.FAKE_TOKEN_JSON_PATH!).text());
