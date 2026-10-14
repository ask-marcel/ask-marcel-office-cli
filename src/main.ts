import { homedir } from 'node:os';
import { dirname, extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import updateNotifier from 'update-notifier';
import pkg from '../package.json' with { type: 'json' };
import { buildDeps } from './composition/build-deps.ts';
import type { BuiltDeps } from './composition/build-deps.ts';
import { buildCli } from './composition/cli.ts';
import { recordTokenHelperLocation, runTokenHelper, tokenHelperLocatorPath } from './composition/token-helper.ts';
import { formatError } from './domain/utilities/format-error.ts';
import { resolveAuthPaths } from './infra/auth-paths.ts';

const ONE_WEEK_MS = 1000 * 60 * 60 * 24 * 7;

/*
 * `ask-marcel-office mcp` serves the MCP gateway over stdio instead of running a
 * command. It is intercepted HERE, before anything else in main, for two
 * reasons:
 *
 * 1. stdout is the JSON-RPC frame channel in this mode. update-notifier must
 *    never run — a "new version available" banner would corrupt the stream.
 * 2. It is deliberately NOT a commander subcommand. Registering it would pull
 *    the MCP SDK into the module graph for all 183 commands, and would add a
 *    lifecycle entry that `docs.test.ts` pins exactly.
 *
 * The imports are dynamic so a normal command never evaluates the SDK.
 */
const isMcpInvocation = (argv: ReadonlyArray<string>): boolean => argv[2] === 'mcp';

/*
 * `ask-marcel-office token --tier <tier> ...` is the token helper, intercepted
 * here like `mcp` and for the same stdout reason: its one JSON line is all it
 * prints, so no update banner, no commander and no `help-json` entry. The fast
 * path is its own entry, dist/token.js (src/token.ts), which sits next to this
 * file in both the bundle and the source tree.
 */
const isTokenInvocation = (argv: ReadonlyArray<string>): boolean => argv[2] === 'token';

const selfPath = fileURLToPath(import.meta.url);
const TOKEN_HELPER = { execPath: process.execPath, entry: join(dirname(selfPath), `token${extname(selfPath)}`), version: pkg.version };

// Every run records where the token helper lives, so read and write find it
// after a plain `login` without PATH. Best effort: a failed write is ignored.
const recordTokenHelper = (deps: BuiltDeps): Promise<void> => recordTokenHelperLocation(deps.fs, tokenHelperLocatorPath(resolveAuthPaths(homedir(), process.env)), TOKEN_HELPER);

const serveMcp = async (): Promise<void> => {
  const [{ buildMcpServer }, { StdioServerTransport }] = await Promise.all([import('./composition/mcp.ts'), import('@modelcontextprotocol/sdk/server/stdio.js')]);
  const deps = buildDeps();
  await recordTokenHelper(deps);
  const server = buildMcpServer({
    auth: deps.auth,
    graph: deps.graph,
    fs: deps.fs,
    version: pkg.version,
    makeLoginAuth: deps.makeLoginAuth,
  });
  await server.connect(new StdioServerTransport());
  // Deliberately no resolve: the transport owns the process lifetime now and
  // exits when the client closes stdin.
};

const main = async (): Promise<void> => {
  if (isTokenInvocation(process.argv)) {
    process.exitCode = await runTokenHelper({ argv: process.argv.slice(3), location: TOKEN_HELPER });
    return;
  }
  if (isMcpInvocation(process.argv)) {
    await serveMcp();
    return;
  }
  updateNotifier({ pkg: { name: pkg.name, version: pkg.version }, updateCheckInterval: ONE_WEEK_MS }).notify({ defer: false });
  const deps = buildDeps();
  await recordTokenHelper(deps);
  const cli = buildCli({
    auth: deps.auth,
    graph: deps.graph,
    logger: deps.logger,
    fs: deps.fs,
    makeLoginAuth: deps.makeLoginAuth,
    version: pkg.version,
    onCommandError: () => {
      process.exitCode = 1;
    },
  });
  await cli.parseAsync();
};

const isCommanderError = (e: unknown): boolean =>
  e !== null && typeof e === 'object' && 'code' in e && typeof e.code === 'string' && (e as { code: string }).code.startsWith('commander.');

try {
  await main();
} catch (e) {
  // Commander parser errors (unknown option, missing required, etc.) are
  // already rendered as a JSON envelope on stdout by exitOverride in
  // buildCli. Don't double-print them to stderr — that would re-introduce
  // the two-channel error contract the audit flagged. Truly uncaught errors
  // (assertion failures, OOM, etc.) still hit the [crash] line.
  if (isCommanderError(e)) {
    process.exit(1);
  }
  process.stderr.write(`[crash] ${formatError(e)}\n`);
  process.exit(1);
}
