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
};

export type BuiltDeps = Readonly<{
  logger: Logger;
  auth: AuthManager;
  graph: GraphClient;
  processRunner: ProcessRunner;
  fs: FileSystem;
  // Deferred AuthManager construction for the `login` command: the
  // `--use-extension` flag is parsed at action time (after deps are built),
  // and it flips the system-browser / Playwright-fallback config, so the
  // login manager can't be the eagerly-built `auth`. The composition root
  // owns this wiring; cli.ts only invokes the factory (or falls back to the
  // injected `auth` in tests).
  makeLoginAuth: (opts: { useExtension: boolean }) => AuthManager;
}>;

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
  // The shared auth every command uses (via the graph client). It must NEVER
  // open the user's default browser mid-command: `skipSystemBrowser: true`
  // confines it to cache → refresh → Playwright fallback (the pre-extension
  // behaviour). The interactive system-browser / extension flow is reached
  // only by `login --use-extension` through `makeLoginAuth` below.
  const auth = createAuthManager({ cachePath, logger, fs, skipSystemBrowser: true });
  const graph = createGraphClient(auth);
  const makeLoginAuth = (opts: { useExtension: boolean }): AuthManager =>
    createAuthManager({ cachePath, logger, fs, skipSystemBrowser: !opts.useExtension, usePlaywrightFallback: !opts.useExtension });
  return { logger, auth, graph, processRunner, fs, makeLoginAuth };
};
