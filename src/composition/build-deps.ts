import { join } from 'node:path';
import type { AuthManager } from '../infra/auth.ts';
import { createAuthManager } from '../infra/auth.ts';
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
  // The command-path auth must never pop a browser per command:
  // `recaptureSecondaryViaBrowser: false` makes the elevated/chatsvcagg/ic3
  // getters FAIL-FAST ("run `ask-marcel-office login`") instead of opening a visible
  // window per command when a secondary token lapses (~hourly for the elevated
  // token). Secondary browser capture is reserved for the explicit `login`
  // command, whose manager comes from `makeLoginAuth`.
  const auth = makeAuth({ cachePath, logger, fs, recaptureSecondaryViaBrowser: false });
  const graph = createGraphClient(auth);
  const makeLoginAuth: LoginAuthFactory = () => makeAuth({ cachePath, logger, fs });
  return { logger, auth, graph, processRunner, fs, makeLoginAuth };
};
