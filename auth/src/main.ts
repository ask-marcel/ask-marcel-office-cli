import { homedir } from 'node:os';
import { dirname, extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import pkg from '../package.json' with { type: 'json' };
import type { BuiltAuthDeps } from './composition/build-auth-deps.ts';
import { crashText, isCommanderError, isTokenInvocation } from './composition/crash-text.ts';
import { resolveAuthPaths } from '@ask-marcel/office-core/infra/auth-paths.ts';
import { tokenHelperLocatorPath } from '@ask-marcel/office-core/infra/token-helper-locator.ts';

/*
 * The bin of @ask-marcel/office-auth (dist/cli.js). Three paths, chosen from
 * the first argument before anything else loads:
 *
 * - `token --tier <tier> ...` is the token helper. Its one JSON line is all it
 *   prints, so it loads no commander, no winston and no update notifier: the
 *   helper is a dynamic import, as in its own entry, dist/token.js
 *   (src/token.ts), which sits next to this file in the bundle and the source.
 * - `mcp` serves the auth MCP server (`login`, `status`) over stdio. stdout is
 *   the JSON-RPC frame channel, so the update notifier never runs.
 * - anything else is the CLI, after the update notice of this package.
 *
 * The CLI and MCP paths record where the token helper lives, so read and write
 * find it after a plain `login` without PATH. Every import that a path alone
 * needs is dynamic, so the other paths never evaluate it.
 */

const ONE_WEEK_MS = 1000 * 60 * 60 * 24 * 7;

const isMcpInvocation = (argv: ReadonlyArray<string>): boolean => argv[2] === 'mcp';

const selfPath = fileURLToPath(import.meta.url);
const TOKEN_HELPER = { execPath: process.execPath, entry: join(dirname(selfPath), `token${extname(selfPath)}`), version: pkg.version };

// Best effort: a failed write is ignored.
const recordTokenHelper = async (deps: BuiltAuthDeps): Promise<void> => {
  const { recordTokenHelperLocation } = await import('./composition/token-helper.ts');
  await recordTokenHelperLocation(deps.fs, tokenHelperLocatorPath(resolveAuthPaths(homedir(), process.env)), TOKEN_HELPER);
};

const runToken = async (): Promise<void> => {
  const { runTokenHelper } = await import('./composition/token-helper.ts');
  process.exitCode = await runTokenHelper({ argv: process.argv.slice(3), location: TOKEN_HELPER });
};

const serveMcp = async (): Promise<void> => {
  const [{ buildAuthDeps }, { buildAuthMcpServer }, { StdioServerTransport }] = await Promise.all([
    import('./composition/build-auth-deps.ts'),
    import('./composition/auth-server.ts'),
    import('@modelcontextprotocol/sdk/server/stdio.js'),
  ]);
  const deps = buildAuthDeps();
  await recordTokenHelper(deps);
  const server = buildAuthMcpServer({ auth: deps.auth, makeLoginAuth: deps.makeLoginAuth, version: pkg.version });
  await server.connect(new StdioServerTransport());
  // Deliberately no resolve: the transport owns the process lifetime now and
  // exits when the client closes stdin.
};

const runCli = async (): Promise<void> => {
  const [{ default: updateNotifier }, { buildAuthDeps }, { buildAuthCli }] = await Promise.all([
    import('update-notifier'),
    import('./composition/build-auth-deps.ts'),
    import('./composition/auth-program.ts'),
  ]);
  updateNotifier({ pkg: { name: pkg.name, version: pkg.version }, updateCheckInterval: ONE_WEEK_MS }).notify({ defer: false });
  const deps = buildAuthDeps();
  await recordTokenHelper(deps);
  const cli = buildAuthCli({
    auth: deps.auth,
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

const main = async (): Promise<void> => {
  if (isTokenInvocation(process.argv)) return runToken();
  if (isMcpInvocation(process.argv)) return serveMcp();
  return runCli();
};

// The one crash catch of this entry (rule 17; .claude/LESSONS.md, 2026-10-04).
try {
  await main();
} catch (e) {
  if (isCommanderError(e)) process.exit(1);
  // On the token helper path, only the error's kind (crash-text.ts).
  process.stderr.write(`[crash] ${crashText(process.argv, e)}\n`);
  process.exit(1);
}
