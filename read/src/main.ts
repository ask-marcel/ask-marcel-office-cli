import updateNotifier from 'update-notifier';
import pkg from '../package.json' with { type: 'json' };
import { buildDeps } from './composition/build-deps.ts';
import { buildCli } from './composition/cli.ts';
import { formatError } from '@ask-marcel/office-core/domain/utilities/format-error.ts';

const ONE_WEEK_MS = 1000 * 60 * 60 * 24 * 7;

/*
 * The bin of @ask-marcel/office-read (dist/cli.js).
 * `ask-marcel-office-read mcp` serves the read MCP gateway over stdio instead
 * of running a command. It is intercepted HERE, before anything else in main,
 * for two reasons:
 *
 * 1. stdout is the JSON-RPC frame channel in this mode. update-notifier must
 *    never run — a "new version available" banner would corrupt the stream.
 * 2. It is deliberately NOT a commander subcommand. Registering it would pull
 *    the MCP SDK into the module graph of every read command.
 *
 * The imports are dynamic so a normal command never evaluates the SDK.
 */
const isMcpInvocation = (argv: ReadonlyArray<string>): boolean => argv[2] === 'mcp';

const serveMcp = async (): Promise<void> => {
  const [{ buildMcpServer }, { StdioServerTransport }] = await Promise.all([import('./composition/mcp.ts'), import('@modelcontextprotocol/sdk/server/stdio.js')]);
  const deps = buildDeps();
  const server = buildMcpServer({
    graph: deps.graph,
    fs: deps.fs,
    version: pkg.version,
  });
  await server.connect(new StdioServerTransport());
  // Deliberately no resolve: the transport owns the process lifetime now and
  // exits when the client closes stdin.
};

const main = async (): Promise<void> => {
  if (isMcpInvocation(process.argv)) {
    await serveMcp();
    return;
  }
  updateNotifier({ pkg: { name: pkg.name, version: pkg.version }, updateCheckInterval: ONE_WEEK_MS }).notify({ defer: false });
  const deps = buildDeps();
  const cli = buildCli({
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
