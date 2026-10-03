import { homedir } from 'node:os';
import type { AccessToken } from '../domain/access-token.ts';
import { accessToken } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { AtomicFileWrites, FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { AuthError, AuthManager, BrowserRungs, FetchFn, SecondaryTokenCommands, TokenCacheAccess } from './auth.ts';
import { createAuthLadder, neededByNote, NO_SECONDARY_TOKEN_COMMANDS } from './auth.ts';
import { resolveAuthPaths } from './auth-paths.ts';
import type { BrowserAuth, ElevatedFailureReason } from './browser-auth.ts';
import { createBrowserAuth } from './browser-auth.ts';
import { createBunFileSystem } from './filesystem-bun.ts';
import { createNodeFileSystem } from './filesystem-node.ts';
import type { TokenCacheLock } from './token-cache-lock.ts';

// Wires the auth ladder (auth.ts) to the Playwright sign-in browser
// (browser-auth.ts).

type BrowserRungsDeps = {
  readonly browserAuth: BrowserAuth;
  readonly logger: Logger;
  readonly secondaryTokenCommands: SecondaryTokenCommands;
  readonly recaptureElevatedViaBrowser: boolean;
  readonly recaptureSecondaryViaBrowser: boolean;
};

// The rungs of the ladder that need a browser, each saving what it captured
// through the ladder's cache access. A rung this session may not use is left out.
const createBrowserRungs = (deps: BrowserRungsDeps, cache: TokenCacheAccess): BrowserRungs => {
  const { browserAuth, logger, secondaryTokenCommands } = deps;
  const { underLock, persistElevated, persistChatsvcagg, persistIc3 } = cache;

  // Distinct error messages per failure mode (launch-hang, navigation
  // failure, silent-SSO timeout) so an LLM gets actionable remediation
  // rather than a one-size-fits-all message.
  const recoverableElevatedFailureMessage = (reason: ElevatedFailureReason): string => {
    const needed = neededByNote(secondaryTokenCommands.elevated);
    if (reason === 'launch_timeout') {
      return `elevated browser launch timed out (15s) — likely a corrupt persistent profile or filesystem lock, or endpoint-security / EDR software blocking the local DevTools (CDP) connection Playwright uses to drive the browser. Run \`ask-marcel-office logout && ask-marcel-office login\` to wipe the profile and retry; if a browser opens but no page ever loads, it is the second cause — running under Node rather than Bun avoids some EDR policies (\`npm i -g ask-marcel-office-cli\`), otherwise add a security exclusion. \`ASKMARCEL_LAUNCH_TIMEOUT_MS\` raises the budget, \`ASKMARCEL_TRACE=1\` names the failing browser.${needed}`;
    }
    if (reason === 'navigation_failed') {
      return `elevated capture failed: navigation to m365.cloud.microsoft did not complete — network issue, corp-proxy block, or tenant policy. Check connectivity and retry. If persistent, the commands that need the elevated token will be unavailable.${needed}`;
    }
    return `elevated token capture timed out — silent SSO against m365.cloud.microsoft did not yield a Bearer within 20s. The persistent browser-profile cookies are likely expired. Run \`ask-marcel-office logout && ask-marcel-office login\` — this now wipes the profile too.${needed}`;
  };

  const recaptureElevated = async (options?: { readonly awaitSignIn?: boolean }): Promise<Result<AccessToken, AuthError>> =>
    underLock('browser', async () => recaptureElevatedHoldingLock(options));
  const recaptureElevatedHoldingLock = async (options?: { readonly awaitSignIn?: boolean }): Promise<Result<AccessToken, AuthError>> => {
    try {
      const captured = await browserAuth.acquireElevatedToken(options);
      if (!captured.ok) {
        return err({ type: 'auth_failed', message: recoverableElevatedFailureMessage(captured.reason) });
      }
      const saved = await persistElevated(captured.token);
      if (!saved.ok) return saved;
      logger.info('auth.elevated.recaptured');
      return ok(captured.token);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: `elevated capture threw: ${msg}` });
    }
  };

  // Same in-flight serialization for the elevated path — concurrent callers
  // share one Playwright instance instead of racing.
  let inFlightElevatedRecapture: Promise<Result<AccessToken, AuthError>> | null = null;
  const recaptureElevatedShared = (options?: { readonly awaitSignIn?: boolean }): Promise<Result<AccessToken, AuthError>> => {
    // A caller willing to wait for a sign-in that JOINS a fail-fast capture
    // already in flight inherits the fail-fast behaviour. Deliberate: the
    // browser is already open and driving it twice is worse than one retry.
    if (inFlightElevatedRecapture !== null) {
      logger.info('auth.elevated.shared_in_flight');
      return inFlightElevatedRecapture;
    }
    const launched = recaptureElevated(options);
    inFlightElevatedRecapture = launched.finally(() => {
      inFlightElevatedRecapture = null;
    });
    return inFlightElevatedRecapture;
  };

  const recoverableChatsvcaggFailureMessage = (reason: ElevatedFailureReason): string => {
    const needed = neededByNote(secondaryTokenCommands.chatsvcagg);
    if (reason === 'launch_timeout') {
      return `chatsvcagg browser launch timed out (15s) — likely a corrupt persistent profile or filesystem lock. Run \`ask-marcel-office logout && ask-marcel-office login\` to wipe the profile and retry.${needed}`;
    }
    if (reason === 'navigation_failed') {
      return 'chatsvcagg capture failed: navigation to teams.cloud.microsoft did not complete — network issue, corp-proxy block, or tenant policy. Check connectivity and retry. If persistent, the Teams chat-content commands will be unavailable.';
    }
    return `chatsvcagg token capture timed out — silent SSO against teams.microsoft.com did not yield a Bearer within 20s. The persistent browser-profile cookies are likely expired. Run \`ask-marcel-office logout && ask-marcel-office login\` — this now wipes the profile too.${needed}`;
  };

  const recaptureChatsvcagg = async (): Promise<Result<AccessToken, AuthError>> => underLock('browser', async () => recaptureChatsvcaggHoldingLock());
  const recaptureChatsvcaggHoldingLock = async (): Promise<Result<AccessToken, AuthError>> => {
    try {
      const captured = await browserAuth.acquireChatsvcaggToken();
      if (!captured.ok) {
        return err({ type: 'auth_failed', message: recoverableChatsvcaggFailureMessage(captured.reason) });
      }
      const saved = await persistChatsvcagg(captured.token, captured.region);
      if (!saved.ok) return saved;
      logger.info('auth.chatsvcagg.recaptured', { region: captured.region });
      return ok(captured.token);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: `chatsvcagg capture threw: ${msg}` });
    }
  };

  let inFlightChatsvcaggRecapture: Promise<Result<AccessToken, AuthError>> | null = null;
  const recaptureChatsvcaggShared = (): Promise<Result<AccessToken, AuthError>> => {
    if (inFlightChatsvcaggRecapture !== null) {
      logger.info('auth.chatsvcagg.shared_in_flight');
      return inFlightChatsvcaggRecapture;
    }
    const launched = recaptureChatsvcagg();
    inFlightChatsvcaggRecapture = launched.finally(() => {
      inFlightChatsvcaggRecapture = null;
    });
    return inFlightChatsvcaggRecapture;
  };

  const recoverableIc3FailureMessage = (reason: ElevatedFailureReason): string => {
    const needed = neededByNote(secondaryTokenCommands.ic3);
    if (reason === 'launch_timeout') {
      return `ic3 browser launch timed out (15s) — likely a corrupt persistent profile or filesystem lock. Run \`ask-marcel-office logout && ask-marcel-office login\` to wipe the profile and retry.${needed}`;
    }
    if (reason === 'navigation_failed') {
      return 'ic3 capture failed: navigation to teams.cloud.microsoft did not complete — network issue, corp-proxy block, or tenant policy. Check connectivity and retry. If persistent, the chat-history command will be unavailable.';
    }
    return `ic3 token capture timed out — silent SSO against teams.microsoft.com did not yield a Bearer within 20s. The persistent browser-profile cookies are likely expired. Run \`ask-marcel-office logout && ask-marcel-office login\` — this now wipes the profile too.${needed}`;
  };

  const recaptureIc3 = async (): Promise<Result<AccessToken, AuthError>> => underLock('browser', async () => recaptureIc3HoldingLock());
  const recaptureIc3HoldingLock = async (): Promise<Result<AccessToken, AuthError>> => {
    try {
      const captured = await browserAuth.acquireIc3Token();
      if (!captured.ok) {
        return err({ type: 'auth_failed', message: recoverableIc3FailureMessage(captured.reason) });
      }
      const saved = await persistIc3(captured.token, captured.region);
      if (!saved.ok) return saved;
      logger.info('auth.ic3.recaptured', { region: captured.region });
      return ok(captured.token);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: `ic3 capture threw: ${msg}` });
    }
  };

  let inFlightIc3Recapture: Promise<Result<AccessToken, AuthError>> | null = null;
  const recaptureIc3Shared = (): Promise<Result<AccessToken, AuthError>> => {
    if (inFlightIc3Recapture !== null) {
      logger.info('auth.ic3.shared_in_flight');
      return inFlightIc3Recapture;
    }
    const launched = recaptureIc3();
    inFlightIc3Recapture = launched.finally(() => {
      inFlightIc3Recapture = null;
    });
    return inFlightIc3Recapture;
  };

  return {
    recaptureElevated: deps.recaptureElevatedViaBrowser ? recaptureElevatedShared : undefined,
    recaptureChatsvcagg: deps.recaptureSecondaryViaBrowser ? recaptureChatsvcaggShared : undefined,
    recaptureIc3: deps.recaptureSecondaryViaBrowser ? recaptureIc3Shared : undefined,
  };
};

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
  // self-heal headlessly from the shared refresh token, so a browser adds nothing
  // for them until that token is dead, and then a full sign-in is due anyway.
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
    secondaryTokenCommands,
    acquireBasicViaBrowser,
    fetchFn,
    lock,
    browserRungs: (cache) => createBrowserRungs({ browserAuth, logger, secondaryTokenCommands, recaptureElevatedViaBrowser, recaptureSecondaryViaBrowser }, cache),
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
