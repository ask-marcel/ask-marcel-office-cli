import { homedir } from 'node:os';
import type { AuthLadder, LoginAuthFactory } from '../infra/auth.ts';
import { createAuthManager } from '../infra/auth-browser.ts';
import { resolveAuthPaths } from '@ask-marcel/office-core/infra/auth-paths.ts';
import { createBunFileSystem } from '@ask-marcel/office-core/infra/filesystem-bun.ts';
import { createNodeFileSystem } from '@ask-marcel/office-core/infra/filesystem-node.ts';
import { createWinstonLogger } from '@ask-marcel/office-core/infra/logger.ts';
import type { AtomicFileWrites, FileSystem } from '@ask-marcel/office-core/use-cases/ports/filesystem.ts';
import type { Logger } from '@ask-marcel/office-core/use-cases/ports/logger.ts';

/*
 * The composition root of the auth package: the logger, the file system, the
 * command-path auth ladder with its interactive gates, and the factory of the
 * login ladder. It builds no graph and no process runner: the auth package runs
 * no Graph command. It names no secondary-token commands either, so the ladder
 * keeps its default (none): the read and write packages add their own commands
 * to a token failure (`neededBy`).
 */
export type BuildAuthDepsConfig = {
  readonly logLevel?: string;
  readonly home?: string;
  readonly env?: Readonly<Record<string, string | undefined>>;
  readonly fs?: FileSystem & AtomicFileWrites;
  // Production builds the real ladder; a test records the options it was given.
  readonly createAuth?: typeof createAuthManager;
  // A person at the terminal. Defaults to `process.stdin.isTTY`.
  readonly interactive?: boolean;
};

export type BuiltAuthDeps = Readonly<{
  logger: Logger;
  auth: AuthLadder;
  fs: FileSystem & AtomicFileWrites;
  makeLoginAuth: LoginAuthFactory;
}>;

const defaultFileSystem = (): FileSystem & AtomicFileWrites => (typeof globalThis.Bun === 'undefined' ? createNodeFileSystem() : createBunFileSystem());

export const buildAuthDeps = (config: BuildAuthDepsConfig = {}): BuiltAuthDeps => {
  const paths = resolveAuthPaths(config.home ?? homedir(), config.env ?? process.env);
  const location = { cachePath: paths.tokenCache, browserProfileDir: paths.browserProfile };
  const logger = createWinstonLogger({ logLevel: config.logLevel ?? process.env.ASKMARCEL_LOG_LEVEL ?? 'error' });
  const fs = config.fs ?? defaultFileSystem();
  const makeAuth = config.createAuth ?? createAuthManager;
  // A secondary token never pops a browser on the command path. The basic and
  // the elevated tokens may, but only for a person at the terminal: an agent
  // gets a fast failure that names `login` (as in the single package).
  const interactive = config.interactive ?? process.stdin.isTTY === true;
  const auth = makeAuth({
    ...location,
    logger,
    fs,
    interactive,
    recaptureSecondaryViaBrowser: false,
    recaptureElevatedViaBrowser: interactive,
    acquireBasicViaBrowser: interactive,
  });
  const makeLoginAuth: LoginAuthFactory = () => makeAuth({ ...location, logger, fs });
  return { logger, auth, fs, makeLoginAuth };
};
