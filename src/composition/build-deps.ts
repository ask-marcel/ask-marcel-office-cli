import { homedir } from 'node:os';
import { commands } from '../use-cases/commands/index.ts';
import type { CommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { createBunFileSystem } from '@ask-marcel/office-core/infra/filesystem-bun.ts';
import { createNodeFileSystem } from '@ask-marcel/office-core/infra/filesystem-node.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import { createTokenSourceGraphClient } from '../infra/graph-client.ts';
import { createWinstonLogger } from '@ask-marcel/office-core/infra/logger.ts';
import { createBunProcessRunner } from '@ask-marcel/office-core/infra/process-runner-bun.ts';
import { createNodeProcessRunner } from '@ask-marcel/office-core/infra/process-runner-node.ts';
import type { SecondaryTokenCommands } from '@ask-marcel/office-core/infra/token-source-needed-by.ts';
import type { AtomicFileWrites, FileExistence, FileSystem } from '@ask-marcel/office-core/use-cases/ports/filesystem.ts';
import type { Logger } from '@ask-marcel/office-core/use-cases/ports/logger.ts';
import type { ProcessRunner } from '@ask-marcel/office-core/use-cases/ports/process-runner.ts';
import { createEnvThenHelperTokenSource } from '@ask-marcel/office-core/composition/token-source.ts';

export type BuildDepsConfig = {
  readonly logLevel?: string;
  readonly home?: string;
  readonly env?: Readonly<Record<string, string | undefined>>;
  readonly fs?: FileSystem & AtomicFileWrites & FileExistence;
  readonly processRunner?: ProcessRunner;
  /**
   * Whether this is an interactive session (a real TTY with a human present).
   * Defaults to `process.stdin.isTTY`. When true, the token helper gets this
   * terminal's stdin; when false (a headless / piped agent run) it gets none,
   * so it fails fast instead of waiting on a sign-in that nobody can complete.
   * Injectable so tests pin the behavior.
   */
  readonly interactive?: boolean;
};

export type BuiltDeps = Readonly<{
  logger: Logger;
  graph: GraphClient;
  processRunner: ProcessRunner;
  fs: FileSystem & AtomicFileWrites;
}>;

// The secondary-token error messages name the commands that need each token.
// Deriving the lists from the registry flags here keeps the messages truthful
// as commands gain or lose a flag. The token source adds them after the
// remedy of the auth package.
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

const defaultFileSystem = (): FileSystem & AtomicFileWrites & FileExistence => (typeof globalThis.Bun !== 'undefined' ? createBunFileSystem() : createNodeFileSystem());

const defaultProcessRunner = (): ProcessRunner => (typeof globalThis.Bun !== 'undefined' ? createBunProcessRunner() : createNodeProcessRunner());

// The root signs only through the package split's source (D1 of phase 4): a
// tier's environment variable when it is set, otherwise the token helper of
// the auth package. The root holds no token cache and opens no browser.
export const buildDeps = (config: BuildDepsConfig = {}): BuiltDeps => {
  const home = config.home ?? homedir();
  const env = config.env ?? process.env;
  const logLevel = config.logLevel ?? process.env.ASKMARCEL_LOG_LEVEL ?? 'error';
  const fs = config.fs ?? defaultFileSystem();
  const processRunner = config.processRunner ?? defaultProcessRunner();
  const logger = createWinstonLogger({ logLevel });
  const interactive = config.interactive ?? process.stdin.isTTY === true;
  const tokens = createEnvThenHelperTokenSource({ env, home, fs, runner: processRunner, interactive, neededBy: secondaryTokenCommands });
  return { logger, graph: createTokenSourceGraphClient(tokens), processRunner, fs };
};
