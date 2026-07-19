import { join } from 'node:path';
import type { AuthManager, SecondaryTokenCommands } from '../infra/auth.ts';
import { createAuthManager } from '../infra/auth.ts';
import { commands } from '../use-cases/commands/index.ts';
import type { CommandMeta } from '../use-cases/commands/command-types.ts';
import { createBunFileSystem } from '../infra/filesystem-bun.ts';
import { createNodeFileSystem } from '../infra/filesystem-node.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import { createGraphClient } from '../infra/graph-client.ts';
import { createWinstonLogger } from '../infra/logger.ts';
import { createBunProcessRunner } from '../infra/process-runner-bun.ts';
import { createNodeProcessRunner } from '../infra/process-runner-node.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { ProcessRunner } from '../use-cases/ports/process-runner.ts';

export type BuildDepsConfig = {
  readonly logLevel?: string;
  readonly cachePath?: string;
  readonly home?: string;
  readonly fs?: FileSystem;
  readonly processRunner?: ProcessRunner;
  readonly createAuth?: typeof createAuthManager;
  /**
   * Whether this is an interactive session (a real TTY with a human present).
   * Defaults to `process.stdin.isTTY`. When true, the command-path auth keeps
   * the first-run auto-launched sign-in browser for the basic token; when false
   * (a headless / piped agent run) that rung fails fast instead of hanging on
   * the interactive-login poll. Injectable so tests pin the behavior.
   */
  readonly interactive?: boolean;
};

/**
 * Builds an AuthManager configured for an interactive `login` run — unlike
 * the command-path manager it may recapture the secondary (elevated /
 * chatsvcagg / ic3) tokens via the browser. The composition root hands the
 * CLI a factory rather than a single pre-built manager — keeping the concrete
 * `createAuthManager` wiring (cache path, env) out of `cli.ts` and the action
 * testable with an injected fake.
 */
export type LoginAuthFactory = () => AuthManager;

export type BuiltDeps = Readonly<{ logger: Logger; auth: AuthManager; graph: GraphClient; processRunner: ProcessRunner; fs: FileSystem; makeLoginAuth: LoginAuthFactory }>;

const defaultCachePath = (home: string): string => join(home, '.ask-marcel', 'token-cache.json');

// The secondary-token error messages name the commands that need each token.
// Deriving the lists from the registry flags here (instead of hardcoding them
// inside infra/auth.ts, which cannot import the registry) keeps the messages
// truthful as commands gain or lose a flag — the old hardcoded elevated list
// omitted `get-user`.
const commandNamesWhere = (matches: (meta: CommandMeta) => boolean): ReadonlyArray<string> =>
  Object.entries(commands)
    .filter(([, cmd]) => matches(cmd.meta))
    .map(([name]) => name)
    .toSorted((a, b) => a.localeCompare(b));

const secondaryTokenCommands: SecondaryTokenCommands = {
  elevated: commandNamesWhere((meta) => meta.needsElevatedToken === true),
  chatsvcagg: commandNamesWhere((meta) => meta.needsSubstrateToken === 'chatsvcagg'),
  ic3: commandNamesWhere((meta) => meta.needsSubstrateToken === 'ic3'),
};

const defaultFileSystem = (): FileSystem => (typeof globalThis.Bun !== 'undefined' ? createBunFileSystem() : createNodeFileSystem());

const defaultProcessRunner = (): ProcessRunner => (typeof globalThis.Bun !== 'undefined' ? createBunProcessRunner() : createNodeProcessRunner());

export const buildDeps = (config: BuildDepsConfig = {}): BuiltDeps => {
  const home = config.home ?? process.env.HOME ?? process.env.USERPROFILE ?? '';
  const cachePath = config.cachePath ?? defaultCachePath(home);
  const logLevel = config.logLevel ?? process.env.ASKMARCEL_LOG_LEVEL ?? 'error';
  const fs = config.fs ?? defaultFileSystem();
  const processRunner = config.processRunner ?? defaultProcessRunner();
  const logger = createWinstonLogger({ logLevel });
  const makeAuth = config.createAuth ?? createAuthManager;
  // Secondary tokens (elevated / chatsvcagg / ic3) NEVER pop a per-command
  // browser: `recaptureSecondaryViaBrowser: false` fails them fast ("run
  // `ask-marcel-office login`") whether or not the session is interactive.
  //
  // The BASIC token is gated on the SESSION instead. A real terminal keeps the
  // first-run convenience of an auto-launched sign-in browser; a headless / piped
  // run (an agent, CI) has no TTY and no human to complete that sign-in, so its
  // basic-token browser rung fails fast rather than blocking on the 5-minute
  // interactive-login poll that hung a headless agent. Full browser capture stays
  // available through the explicit `login` command (`makeLoginAuth` leaves
  // `acquireBasicViaBrowser` at its `true` default).
  const interactive = config.interactive ?? process.stdin.isTTY === true;
  const auth = makeAuth({ cachePath, logger, fs, recaptureSecondaryViaBrowser: false, acquireBasicViaBrowser: interactive, secondaryTokenCommands });
  const graph = createGraphClient(auth);
  const makeLoginAuth: LoginAuthFactory = () => makeAuth({ cachePath, logger, fs, secondaryTokenCommands });
  return { logger, auth, graph, processRunner, fs, makeLoginAuth };
};
