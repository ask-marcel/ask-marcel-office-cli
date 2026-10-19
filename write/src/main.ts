import pkg from '../package.json' with { type: 'json' };
import { formatError } from '@ask-marcel/office-core/domain/utilities/format-error.ts';

/*
 * The bin of @ask-marcel/office-write (dist/cli.js). Two paths, chosen from the
 * first argument before anything else loads:
 *
 * - `mcp` serves the write MCP server (list-commands, get-command-docs,
 *   run-write-command) over stdio. stdout is the JSON-RPC frame channel, so the
 *   update notifier never runs. `mcp` is not a commander command, so a plain
 *   command never loads the MCP SDK.
 * - anything else is the CLI (the 7 write commands, docs and help-json), after
 *   the update notice of this package.
 *
 * The write bin signs through the tier variables, else the token helper of
 * @ask-marcel/office-auth. It runs no token path and writes no locator: only
 * the auth bin writes it (LESSONS 2026-10-18). Every import that a path alone
 * needs is dynamic, so the other path never evaluates it.
 */

const ONE_WEEK_MS = 1000 * 60 * 60 * 24 * 7;

const isMcpInvocation = (argv: ReadonlyArray<string>): boolean => argv[2] === 'mcp';

const serveMcp = async (): Promise<void> => {
  const [{ buildWriteDeps }, { buildWriteMcpServer }, { StdioServerTransport }] = await Promise.all([
    import('./composition/build-write-deps.ts'),
    import('./composition/write-server.ts'),
    import('@modelcontextprotocol/sdk/server/stdio.js'),
  ]);
  const deps = buildWriteDeps();
  const server = buildWriteMcpServer({ graph: deps.graph, fs: deps.fs, version: pkg.version });
  await server.connect(new StdioServerTransport());
  // Deliberately no resolve: the transport owns the process lifetime now and
  // exits when the client closes stdin.
};

const runCli = async (): Promise<void> => {
  const [{ default: updateNotifier }, { buildWriteDeps }, { buildWriteCli }] = await Promise.all([
    import('update-notifier'),
    import('./composition/build-write-deps.ts'),
    import('./composition/write-program.ts'),
  ]);
  updateNotifier({ pkg: { name: pkg.name, version: pkg.version }, updateCheckInterval: ONE_WEEK_MS }).notify({ defer: false });
  const deps = buildWriteDeps();
  const cli = buildWriteCli({
    graph: deps.graph,
    logger: deps.logger,
    fs: deps.fs,
    version: pkg.version,
    onCommandError: () => {
      process.exitCode = 1;
    },
  });
  await cli.parseAsync();
};

const main = async (): Promise<void> => (isMcpInvocation(process.argv) ? serveMcp() : runCli());

// Commander parser errors (unknown option, missing required flag) are already
// rendered as a JSON envelope on stdout by the exitOverride of the CLI builder:
// the catch writes no second line for them.
const isCommanderError = (e: unknown): boolean =>
  e !== null && typeof e === 'object' && 'code' in e && typeof e.code === 'string' && (e as { code: string }).code.startsWith('commander.');

// The one crash catch of this entry (rule 17; .claude/LESSONS.md, 2026-10-04).
try {
  await main();
} catch (e) {
  if (isCommanderError(e)) process.exit(1);
  process.stderr.write(`[crash] ${formatError(e)}\n`);
  process.exit(1);
}
