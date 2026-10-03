import { homedir } from 'node:os';
import { accessToken } from '../domain/access-token.ts';
import type { AtomicFileWrites, FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { AuthManager, FetchFn, SecondaryTokenCommands } from './auth.ts';
import { createAuthLadder, NO_SECONDARY_TOKEN_COMMANDS } from './auth.ts';
import { resolveAuthPaths } from './auth-paths.ts';
import type { BrowserAuth } from './browser-auth.ts';
import { createBrowserAuth } from './browser-auth.ts';
import { createBunFileSystem } from './filesystem-bun.ts';
import { createNodeFileSystem } from './filesystem-node.ts';
import type { TokenCacheLock } from './token-cache-lock.ts';

// Wires the auth ladder (auth.ts) to the Playwright sign-in browser
// (browser-auth.ts).

const createAuthManagerFromApi = (
  browserAuth: BrowserAuth,
  cachePath: string,
  browserProfileDir: string,
  logger: Logger,
  fs: FileSystem & AtomicFileWrites,
  recaptureSecondaryViaBrowser: boolean = true,
  secondaryTokenCommands: SecondaryTokenCommands = NO_SECONDARY_TOKEN_COMMANDS,
  acquireBasicViaBrowser: boolean = true,
  // Elevated gets its OWN browser gate, separate from chatsvcagg / ic3. Those two
  // self-heal by redeeming the shared refresh token from INSIDE the ladder's
  // `!recaptureSecondaryViaBrowser` branch (auth.ts), so turning the shared flag on
  // would skip that headless refresh and open a browser instead — strictly worse.
  // Elevated carries no refresh token, so its only question is whether a browser is
  // permitted; an interactive session answers yes and refreshes it in ~17s of silent
  // SSO against the persistent profile. Defaults to the shared flag so every existing
  // caller keeps the behaviour it had.
  recaptureElevatedViaBrowser: boolean = recaptureSecondaryViaBrowser,
  // The token endpoint is the ONE network call the ladder makes. Injectable so
  // a test can exercise the refresh path without POSTing a bogus refresh token
  // to login.microsoftonline.com for real: five did, and passed or timed out at
  // the 5s test limit depending on Microsoft's latency (2026-08-31), which made
  // the suite's result a property of the network. Defaults to the global, so
  // every existing caller is unchanged.
  fetchFn?: FetchFn,
  // The machine-wide lock around every redemption, sign-in and sign-out.
  // Injectable so a test can stand in another process and a clock it controls.
  lock?: TokenCacheLock
): AuthManager =>
  createAuthLadder({
    browserAuth,
    cachePath,
    browserProfileDir,
    logger,
    fs,
    recaptureSecondaryViaBrowser,
    secondaryTokenCommands,
    acquireBasicViaBrowser,
    recaptureElevatedViaBrowser,
    fetchFn,
    lock,
  });

const defaultFileSystem = (): FileSystem & AtomicFileWrites => (typeof globalThis.Bun !== 'undefined' ? createBunFileSystem() : createNodeFileSystem());

// The same resolver `browser-auth.ts` and the composition root use, so `logout`
// wipes the folder `login` signs in with.
const defaultBrowserProfileDir = (): string => resolveAuthPaths(homedir(), process.env).browserProfile;

/**
 * Probe the token cache for a fresh access token. Handed to the browser
 * capture so its poll loop can short-circuit the multi-minute dance
 * when a concurrent process refreshes first (AAD rotates SPA refresh tokens,
 * so the loser of the race cannot refresh and falls into the browser leg).
 * Exported for the composition test; pure read, never writes.
 */
const createFreshCachedTokenProbe = (fs: FileSystem, cachePath: string): (() => Promise<string | null>) => {
  return async () => {
    const cached = await fs.readJson<{ access_token?: string }>(cachePath);
    if (!cached.ok || typeof cached.value.access_token !== 'string') return null;
    const validated = accessToken(cached.value.access_token);
    return validated.ok ? validated.value : null;
  };
};

// Progress lines go to stderr so "waiting on the user's sign-in" is
// distinguishable from a hang; stdout stays reserved for the JSON envelope.
const stderrProgress = (line: string): void => {
  process.stderr.write(`${line}\n`);
};

const createAuthManager = (deps: {
  cachePath: string;
  logger: Logger;
  fs?: FileSystem & AtomicFileWrites;
  browserProfileDir?: string;
  recaptureSecondaryViaBrowser?: boolean;
  secondaryTokenCommands?: SecondaryTokenCommands;
  acquireBasicViaBrowser?: boolean;
  recaptureElevatedViaBrowser?: boolean;
  // The seam a test uses to see how the browser is built; production uses Playwright.
  createBrowser?: typeof createBrowserAuth;
}): AuthManager => {
  const fs = deps.fs ?? defaultFileSystem();
  const browserProfileDir = deps.browserProfileDir ?? defaultBrowserProfileDir();
  // `profileDir` here too: the sign-in browser used its own default while
  // `logout` wiped this folder, so a custom profile was signed into in one place
  // and wiped in another.
  const browserAuth = (deps.createBrowser ?? createBrowserAuth)({
    logger: deps.logger,
    fs,
    freshCachedToken: createFreshCachedTokenProbe(fs, deps.cachePath),
    onProgress: stderrProgress,
    profileDir: browserProfileDir,
  });
  return createAuthManagerFromApi(
    browserAuth,
    deps.cachePath,
    browserProfileDir,
    deps.logger,
    fs,
    deps.recaptureSecondaryViaBrowser,
    deps.secondaryTokenCommands,
    deps.acquireBasicViaBrowser,
    deps.recaptureElevatedViaBrowser
  );
};

export { createAuthManager, createAuthManagerFromApi, createFreshCachedTokenProbe, stderrProgress };
