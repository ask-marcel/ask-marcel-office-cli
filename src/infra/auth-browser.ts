import { homedir } from 'node:os';
import type { AccessToken } from '../domain/access-token.ts';
import { accessToken } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { AtomicFileWrites, FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { AuthError, AuthLadder, BrowserRungs, ElevatedOutcome, FetchFn, SecondaryTokenCommands, TokenCacheAccess } from './auth.ts';
import { createAuthLadder, neededByNote, NO_SECONDARY_TOKEN_COMMANDS } from './auth.ts';
import { resolveAuthPaths } from './auth-paths.ts';
import type { BrowserAuth, ChatsvcaggTokenResult, ElevatedFailureReason, Ic3TokenResult } from './browser-auth.ts';
import { createBrowserAuth } from './browser-auth.ts';
import { createBunFileSystem } from './filesystem-bun.ts';
import { createNodeFileSystem } from './filesystem-node.ts';
import type { TokenCacheLock } from './token-cache-lock.ts';

// The browser half of auth: the rungs of the ladder that drive the Playwright
// sign-in browser (browser-auth.ts), and the factories that build an auth manager
// from the ladder (auth.ts) and those rungs. auth.ts never imports this file, so
// a session that may not open a browser runs on the ladder alone.

// Microsoft moved the Teams web app here; `teams.microsoft.com` now 302s to it.
// Navigating straight to the destination drops a redirect hop from every
// capture, and keeps the whole session on the host the substrate calls use
// (probed live 2026-08-31).
const TEAMS_URL = 'https://teams.cloud.microsoft/';

type BrowserRungsDeps = {
  readonly browserAuth: BrowserAuth;
  readonly logger: Logger;
  readonly secondaryTokenCommands: SecondaryTokenCommands;
  readonly acquireBasicViaBrowser: boolean;
  readonly recaptureElevatedViaBrowser: boolean;
  readonly recaptureSecondaryViaBrowser: boolean;
};

// The rungs of the ladder that need a browser, each saving what it captured
// through the ladder's cache access. A rung this session may not use is left out.
const createBrowserRungs = (deps: BrowserRungsDeps, cache: TokenCacheAccess): BrowserRungs => {
  const { browserAuth, logger, secondaryTokenCommands } = deps;
  const { underLock, persistTeams, persistElevated, persistChatsvcagg, persistIc3, redeemMissedSubstrateAtLogin } = cache;

  // Track the elevated-capture outcome from the most recent
  // browser-acquired session so the login command can surface it to the
  // user via `getLastElevatedOutcome()`. Reset to null on every fresh
  // `acquireViaBrowser` so stale outcomes don't leak across login attempts.
  let lastElevatedOutcome: ElevatedOutcome | null = null;
  let lastChatsvcaggOutcome: ElevatedOutcome | null = null;

  // The chatsvcagg / ic3 bearers a sign-in captured on the side, saved in turn;
  // the first save that fails stops the login with its error.
  const persistSubstrateCaptures = async (chatsvcagg: ChatsvcaggTokenResult, ic3: Ic3TokenResult): Promise<Result<void, AuthError>> => {
    if (chatsvcagg.ok) {
      logger.info('auth.chatsvcagg.captured_at_login', { region: chatsvcagg.region });
      lastChatsvcaggOutcome = { captured: true };
      const saved = await persistChatsvcagg(chatsvcagg.token, chatsvcagg.region);
      if (!saved.ok) return saved;
    } else {
      logger.info('auth.chatsvcagg.skipped_at_login', { reason: chatsvcagg.reason });
      lastChatsvcaggOutcome = { captured: false, reason: chatsvcagg.reason };
    }
    if (!ic3.ok) {
      logger.info('auth.ic3.skipped_at_login', { reason: ic3.reason });
      return ok(undefined);
    }
    logger.info('auth.ic3.captured_at_login', { region: ic3.region });
    return persistIc3(ic3.token, ic3.region);
  };

  type SignedIn = { readonly token: AccessToken; readonly substrateCaptured: { readonly chatsvcagg: boolean; readonly ic3: boolean } | null };

  // The browser leg, run while holding the lock: Chromium's own Singleton lock
  // files are cleared at every launch, so only the lock keeps a second browser
  // off the profile.
  const signInViaBrowser = async (force: boolean): Promise<Result<SignedIn, AuthError>> => {
    try {
      // Single-session capture: one Playwright-driven browser window does
      // every capture leg. Opening a SECOND browser at m365.cloud.microsoft
      // for the elevated step flashed a fresh sign-in prompt on federated
      // tenants because the elevated identity's silent-SSO cookies hadn't
      // settled from disk — so we reuse the same browser context: after
      // the Teams token comes through the network listener, the SAME
      // page navigates to m365.cloud.microsoft so cookies stay live in
      // memory. (An earlier auto-heal profile wipe was dropped for the
      // same reason — it wiped the freshly-authenticated Teams cookies
      // and made federated tenants strictly worse.)
      //
      // Substrate (chatsvcagg) round: same teams.microsoft.com session
      // emits the chatsvcagg-audience bearer on its initial chat-list
      // load, so the third capture leg piggy-backs on the existing
      // browser run — zero additional UI prompts.
      const { teams: result, elevated, chatsvcagg, ic3, fromCache } = await browserAuth.acquireBothTokens(TEAMS_URL, { skipCacheProbe: force });
      if (!result) return err({ type: 'auth_cancelled' });
      // The poll short-circuited because a concurrent process landed a
      // fresh token in the cache. Do NOT persist (refreshToken is null here —
      // writing would clobber the winner's rotated refresh token) and leave the
      // elevated/chatsvcagg outcomes null: no browser-tested state to report.
      if (fromCache === true) {
        logger.info('auth.ladder.rung', { rung: 'browser_cache_short_circuit' });
        return ok({ token: result.accessToken, substrateCaptured: null });
      }
      const elevatedToken: AccessToken | null = elevated.ok ? elevated.token : null;
      if (elevated.ok) {
        logger.info('auth.elevated.captured_at_login');
        lastElevatedOutcome = { captured: true };
      } else {
        logger.info('auth.elevated.skipped_at_login', { reason: elevated.reason });
        lastElevatedOutcome = { captured: false, reason: elevated.reason };
      }
      const savedTeams = await persistTeams(result.accessToken, result.refreshToken, elevatedToken);
      if (!savedTeams.ok) return savedTeams;
      const savedSubstrate = await persistSubstrateCaptures(chatsvcagg, ic3);
      if (!savedSubstrate.ok) return savedSubstrate;
      logger.info('auth.ladder.rung', { rung: 'browser' });
      return ok({ token: result.accessToken, substrateCaptured: { chatsvcagg: chatsvcagg.ok, ic3: ic3.ok } });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: msg });
    }
  };

  // A forced login redeems the substrate tokens the browser missed AFTER the
  // sign-in released the lock: the redemptions take the lock themselves.
  const acquireViaBrowser = async (force = false): Promise<Result<AccessToken, AuthError>> => {
    const signedIn = await underLock('browser', async () => signInViaBrowser(force));
    if (!signedIn.ok) return signedIn;
    const captured = signedIn.value.substrateCaptured;
    if (force && captured !== null) await redeemMissedSubstrateAtLogin(captured.chatsvcagg, captured.ic3);
    return ok(signedIn.value.token);
  };

  // Concurrent first-time auth was racing — two parallel commands would
  // both fall through to acquireViaBrowser, one would win and one would
  // return `auth_cancelled` from the lost Playwright context. Cache the
  // in-flight browser-acquire promise so concurrent callers share one
  // login attempt. Cleared on settle (success or failure) so the next call
  // re-checks the cache instead of returning a stale failure.
  let inFlightBrowserAcquire: Promise<Result<AccessToken, AuthError>> | null = null;
  const acquireViaBrowserShared = (force = false): Promise<Result<AccessToken, AuthError>> => {
    if (inFlightBrowserAcquire !== null) {
      logger.info('auth.ladder.rung', { rung: 'browser_shared_in_flight' });
      return inFlightBrowserAcquire;
    }
    const launched = acquireViaBrowser(force);
    inFlightBrowserAcquire = launched.finally(() => {
      inFlightBrowserAcquire = null;
    });
    return inFlightBrowserAcquire;
  };

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
    signIn: deps.acquireBasicViaBrowser ? acquireViaBrowserShared : undefined,
    recaptureElevated: deps.recaptureElevatedViaBrowser ? recaptureElevatedShared : undefined,
    recaptureChatsvcagg: deps.recaptureSecondaryViaBrowser ? recaptureChatsvcaggShared : undefined,
    recaptureIc3: deps.recaptureSecondaryViaBrowser ? recaptureIc3Shared : undefined,
    close: () => browserAuth.close(),
    lastElevatedOutcome: () => lastElevatedOutcome,
    lastChatsvcaggOutcome: () => lastChatsvcaggOutcome,
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
): AuthLadder =>
  createAuthLadder({
    cachePath,
    browserProfileDir,
    logger,
    fs,
    secondaryTokenCommands,
    interactive: acquireBasicViaBrowser,
    fetchFn,
    lock,
    browserRungs: (cache) =>
      createBrowserRungs({ browserAuth, logger, secondaryTokenCommands, acquireBasicViaBrowser, recaptureElevatedViaBrowser, recaptureSecondaryViaBrowser }, cache),
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

// Every browser rung is off. An unset elevated gate follows the shared one, as
// in createAuthManagerFromApi.
const opensNoBrowser = (deps: {
  readonly acquireBasicViaBrowser?: boolean;
  readonly recaptureSecondaryViaBrowser?: boolean;
  readonly recaptureElevatedViaBrowser?: boolean;
}): boolean => deps.acquireBasicViaBrowser === false && deps.recaptureSecondaryViaBrowser === false && deps.recaptureElevatedViaBrowser !== true;

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
}): AuthLadder => {
  const fs = deps.fs ?? defaultFileSystem();
  const browserProfileDir = deps.browserProfileDir ?? defaultBrowserProfileDir();
  // An agent, an MCP server or a piped run: no rung may open a browser, so none
  // is built and the ladder fails fast where a browser would have been.
  if (opensNoBrowser(deps)) {
    return createAuthLadder({ cachePath: deps.cachePath, browserProfileDir, logger: deps.logger, fs, secondaryTokenCommands: deps.secondaryTokenCommands, interactive: false });
  }
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
