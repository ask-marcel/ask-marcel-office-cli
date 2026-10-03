import type { AccessToken } from '../domain/access-token.ts';
import { accessToken, accessTokenUnsafe } from '../domain/access-token.ts';
import { decodeJwtPayload } from '../domain/jwt-utils.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { TenantId } from '../domain/tenant-id.ts';
import type { AtomicFileWrites, FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { BrowserAuth, ChatsvcaggTokenResult, ElevatedFailureReason, Ic3TokenResult } from './browser-auth.ts';
import { REQUEST_TIMEOUT_MS } from './network-error.ts';
import type { LockPurpose, TokenCacheLock } from './token-cache-lock.ts';
import { createSystemTokenCacheLock } from './token-cache-lock.ts';

type CachedToken = {
  access_token: string;
  expires_on: number;
  refresh_token: string;
  /**
   * "Elevated" Graph token captured from a Microsoft web app whose
   * first-party identity is on the ODSP `logicalPermissions` allow-list
   * (e.g., M365ChatClient, OfficeHome). Used by historical-version
   * commands to fetch streamContent the Teams web client token can't.
   * Refresh path is re-capture (no refresh_token in this flow).
   */
  elevated_access_token?: string;
  elevated_expires_on?: number;
  /**
   * chatsvcagg-audience bearer: same Teams web client identity as
   * `access_token`, but minted for `chatsvcagg.teams.microsoft.com`
   * (the Teams chat-aggregator API). Used by `list-teams-chats-with-messages`
   * and siblings — endpoints that return chat metadata WITH recent
   * message bodies inlined, which Graph's `Chat.Read*`-gated endpoints
   * can't reach with the scopes the CLI's two existing tokens carry.
   * Same refresh model as elevated: re-capture via the persistent
   * browser profile.
   */
  chatsvcagg_access_token?: string;
  chatsvcagg_expires_on?: number;
  /**
   * The Teams substrate is region-routed under
   * `teams.microsoft.com/api/csa/<region>/api/...` (post-2026-05 host
   * migration). Captured from the first `/api/csa/<region>/` URL the
   * chatsvcagg bearer rides on during login. Absent in caches written
   * before the migration; readers fall back to `DEFAULT_CHATSVCAGG_REGION`.
   * Shared by both chatsvcagg and IC3 paths (regions match per tenant).
   */
  chatsvcagg_region?: string;
  /**
   * IC3-audience bearer (Teams web client appid, aud
   * `https://ic3.teams.office.com`). Used by `list-teams-chat-history`
   * to walk the paginated chat-message substrate at
   * `teams.microsoft.com/api/chatsvc/<region>/v1/users/ME/conversations/{id}/messages`,
   * unlocking reads beyond the 200-message chatsvcagg cap. Same
   * lifecycle / refresh model as chatsvcagg: cache → silent re-capture
   * via the persistent profile.
   */
  ic3_access_token?: string;
  ic3_expires_on?: number;
  /**
   * Graph tokens issued by a PARTNER tenant's authority, keyed by that
   * tenant's GUID. A signed-in user who is a guest in another tenant cannot
   * read its SharePoint on the home token — Graph answers `401
   * invalidAudienceUri: Invalid audience Uri '00000003-0000-0ff1-ce00-000000000000'`
   * (SharePoint Online's app id), because home-tenant Graph cannot mint an
   * SPO token for a foreign tenant. Redeeming the shared (FOCI) refresh token
   * against `login.microsoftonline.com/{tenantId}` yields a token that can.
   *
   * Keyed rather than flat because one session legitimately touches several
   * partner tenants. `Partial<Record<...>>` (not `Record<...>`) so the type
   * admits the missing key; read with `Object.hasOwn`, never `in`.
   *
   * Refresh model: same shared refresh_token as basic/chatsvcagg/ic3 (verified
   * 2026-07-16 — a guest-rotated RT still refreshes the HOME tenant, so ONE
   * shared RT slot is correct and per-tenant RT chains are not needed).
   */
  guest_tokens?: Partial<Record<string, { access_token: string; expires_on: number }>>;
};
type AuthError = { type: 'auth_failed'; message: string; code?: string } | { type: 'auth_cancelled' };
/**
 * Outcome of the elevated-token capture leg of the most recent
 * browser-acquired session. Read by `login.execute` to surface a
 * `{ elevated: 'captured' | 'failed', elevatedReason?: ... }` field on
 * the login response so an LLM consumer can predict whether the
 * elevated-dependent commands (chat metadata, historical-version
 * downloads) will work without invoking them.
 */
type ElevatedOutcome = { captured: true } | { captured: false; reason: ElevatedFailureReason | 'unknown_error' };
// Decode-only preflight info for one token tier: availability, remaining runway,
// and the scopes granted to that token (decoded from its `scp` claim; empty when
// the token or the claim is absent).
type CachedTierInfo = { readonly available: boolean; readonly expiresInSeconds: number | undefined; readonly scopes: ReadonlyArray<string> };

type AuthManager = {
  getAccessToken: (options?: { force?: boolean }) => Promise<Result<AccessToken, AuthError>>;
  /**
   * Returns a Graph token issued for an app on Microsoft's ODSP
   * `logicalPermissions` allow-list. Falls through cache → re-capture
   * via headless Playwright. Used by the 3 historical-version commands.
   */
  getElevatedAccessToken: (options?: { readonly awaitSignIn?: boolean }) => Promise<Result<AccessToken, AuthError>>;
  /**
   * Returns a Graph token issued by a PARTNER tenant's authority, for a user
   * who is a guest there. Without it, every call against that tenant's
   * SharePoint dies at `401 invalidAudienceUri` — home-tenant Graph cannot mint
   * an SPO token for a foreign tenant. Cache -> headless redemption of the
   * shared refresh token; never a browser.
   */
  getGuestAccessToken: (tenantId: TenantId) => Promise<Result<AccessToken, AuthError>>;
  /**
   * Returns a chatsvcagg-audience token (same Teams web client identity
   * as `getAccessToken`, but issued for the chatsvcagg resource). Falls
   * through cache → re-capture via headless Playwright. Used by the
   * `list-teams-chats-with-messages` family of commands.
   */
  getChatsvcaggAccessToken: (options?: { readonly ignoreCache?: boolean }) => Promise<Result<AccessToken, AuthError>>;
  /**
   * Returns the regional segment used to construct chatsvcagg substrate
   * URLs (`teams.microsoft.com/api/csa/<region>/api/...`). Captured at
   * login from the first such URL the chatsvcagg bearer rides on. Falls
   * back to `DEFAULT_CHATSVCAGG_REGION` ('emea') when the cache is
   * either absent or pre-2026-05-migration. Synchronous on cache; calls
   * `getChatsvcaggAccessToken()` first if no cache exists so a region is
   * available immediately after login.
   */
  getChatsvcaggRegion: () => Promise<string>;
  /**
   * Returns an IC3-audience bearer (Teams web client identity, aud
   * `https://ic3.teams.office.com`). Falls through cache → re-capture
   * via headless Playwright. Used by `list-teams-chat-history` to walk
   * paginated chat-message history beyond the 200-message chatsvcagg cap.
   */
  getIc3AccessToken: (options?: { readonly ignoreCache?: boolean }) => Promise<Result<AccessToken, AuthError>>;
  /**
   * Redeem the shared refresh token for any COLD substrate tier, over HTTP, with
   * no browser on any path. `login` calls this so a warm-cache sign-in leaves all
   * four tiers usable; without it the two substrate tiers stay cold until some
   * Teams-chat command pays for them, and `login` reports them missing while
   * having done nothing about it.
   */
  warmSubstrateTokens?: () => Promise<void>;
  logout: () => Promise<Result<void, AuthError>>;
  /**
   * Inspect the elevated-capture outcome from the most recent
   * `acquireViaBrowser` invocation. Returns null if no browser-acquired
   * session has happened in this process (cache hit / refresh-only).
   */
  getLastElevatedOutcome: () => ElevatedOutcome | null;
  /**
   * Inspect the chatsvcagg-capture outcome from the most recent
   * `acquireViaBrowser` invocation. Same shape and lifetime semantics
   * as `getLastElevatedOutcome`.
   */
  getLastChatsvcaggOutcome: () => ElevatedOutcome | null;
  /**
   * Decode-only preflight for whether the *persisted* elevated
   * (M365ChatClient) token is present and still usable — the token the
   * historical-version download / convert commands need. Unlike
   * `getLastElevatedOutcome` (per-process, null in a fresh CLI invocation),
   * this reads the on-disk cache, so a separate `deep-scan` run can tell
   * "elevated available" from "run `login` first" without provoking a 403.
   * Optional: only the real manager implements it; a minimal fake omits it
   * and callers treat that as unavailable. Never captures or refreshes.
   */
  /**
   * The cached BASIC token, decoded by `scopes-check` and never acquired: no
   * refresh, no browser. The acquiring getter heals a dead session by opening
   * a browser and, when the persistent profile is already signed in, wiping it
   * so the grant re-fires (2026-09-02: a diagnostic did exactly that). A stale
   * token comes back as-is so its expiry can be reported; no cache at all is
   * `undefined`. Optional: a bring-your-own-token manager omits it and callers
   * fall back to `getAccessToken`, which is then the caller's own function.
   */
  getCachedBasicToken?: () => Promise<AccessToken | undefined>;
  getCachedElevatedInfo?: () => Promise<CachedTierInfo>;
  /**
   * Same decode-only preflight as `getCachedElevatedInfo`, for the chatsvcagg /
   * ic3 Teams-chat substrate tokens. `login`'s four-token status and
   * `scopes-check` read these; a minimal fake omits them and callers treat that
   * as unavailable.
   */
  getCachedChatsvcaggInfo?: () => Promise<CachedTierInfo>;
  getCachedIc3Info?: () => Promise<CachedTierInfo>;
};

const CLIENT_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const SCOPES = 'https://graph.microsoft.com/.default openid profile offline_access';
const SPA_ORIGIN = 'https://teams.microsoft.com';
// Microsoft moved the Teams web app here; `teams.microsoft.com` now 302s to it.
// Navigating straight to the destination drops a redirect hop from every
// capture, and keeps the whole session on the host the substrate calls use
// (probed live 2026-08-31).
const TEAMS_URL = 'https://teams.cloud.microsoft/';
/**
 * Fallback region when no `chatsvcagg_region` is persisted (pre-2026-05
 * caches, or a chatsvcagg capture that never saw a `/api/csa/<region>/`
 * URL). `emea` matches the only region we've empirically tested — the
 * use-case will surface a clear `HTTP 404 …` from the new substrate if
 * an AMER/APAC tenant ends up here, which is preferable to refusing to
 * issue the call at all.
 */
const DEFAULT_CHATSVCAGG_REGION = 'emea';

// Substrate resource audiences. The Teams web client (CLIENT_ID) is consented
// for all three (it mints them in-browser), so the shared refresh_token
// redeems for each by requesting `${resource}/.default` at the token endpoint.
const CHATSVCAGG_RESOURCE = 'https://chatsvcagg.teams.microsoft.com';
const IC3_RESOURCE = 'https://ic3.teams.office.com';

// Decode the scopes granted to a cached token from its `scp` claim (space-separated).
// Empty when the token is absent or carries no `scp` (decodeJwtPayload returns {} on
// any malformed input). Used by the per-tier preflight getters so scopes-check can
// list what each token can actually do.
const decodeScopes = (token: string | undefined): ReadonlyArray<string> => {
  if (!token) return [];
  const scp = decodeJwtPayload(token)['scp'];
  return typeof scp === 'string' ? scp.split(' ').filter((s) => s.length > 0) : [];
};

// Fail-fast (no browser) message for the secondary-token getters, used on the
// command path only AFTER the headless refresh (`refreshSubstrateToken`) has
// been tried and could not produce a token (no cached RT, or AAD rejected the
// redemption). A browser recapture per command is off-limits — the CLI runs one
// process per command and the recaptures open a visible window — so the remedy
// is an interactive `login`. (The elevated token has no HTTP-refresh path at
// all — different appid — so it always lands here on the command path.)
/**
 * The remedy is NOT identical across the three secondary tiers, though this
 * message used to claim it was.
 *
 * chatsvcagg / ic3 ride the shared refresh token and self-heal headlessly, so
 * they only reach this fail-fast once the RT itself is gone — and a plain
 * `login` genuinely fixes that.
 *
 * Elevated carries NO refresh token of its own: it exists only via the browser
 * dance. A plain `login` that finds a valid basic token used to return on the
 * cache rung without re-capturing it — a loop with no exit — so this message
 * once demanded `--force`. `login.execute` closed that loop: the login command
 * now inspects the cached elevated token and self-escalates to the forced
 * browser re-capture when it is missing, so a plain `login` recovers elevated
 * too. The remedy points there, and the message names `scopes-check` for
 * preflight so an unattended agent can re-auth up front rather than discover the
 * lapse mid-run.
 */
const failFastSecondaryMessage = (token: string, commands: ReadonlyArray<string>, remedy: string): string =>
  `${token} token is expired or was not captured at login. ${remedy} — the CLI does not open a browser per command for this token. Preflight token validity with \`ask-marcel-office scopes-check\` (no Graph call) before a long unattended run.${neededByNote(commands, 'it')}`;

/**
 * Command names quoted in the secondary-token error messages, per token kind.
 * The composition root derives these from the command registry
 * (`needsElevatedToken` / `needsSubstrateToken` flags) and injects them. A
 * caller that injects none gets messages that name the token only: auth knows
 * which tokens exist, not which commands need them, and the hard-coded default
 * that used to stand in had already drifted from the registry.
 */
type SecondaryTokenCommands = {
  readonly elevated: ReadonlyArray<string>;
  readonly chatsvcagg: ReadonlyArray<string>;
  readonly ic3: ReadonlyArray<string>;
};

const NO_SECONDARY_TOKEN_COMMANDS: SecondaryTokenCommands = { elevated: [], chatsvcagg: [], ic3: [] };

const commandList = (names: ReadonlyArray<string>): string => names.join(', ');

// The "(Commands that need this token: …)" tail of a token-failure message,
// empty when the caller named no commands.
const neededByNote = (names: ReadonlyArray<string>, token: string = 'this token'): string => (names.length === 0 ? '' : ` (Commands that need ${token}: ${commandList(names)}.)`);

const RECAPTURE_VIA_LOGIN = 'Run `ask-marcel-office login` to (re)capture it';
const RECAPTURE_ELEVATED_VIA_LOGIN =
  'It carries no refresh token of its own, so re-capture it with `ask-marcel-office login`: the login command self-escalates to a browser sign-in when the elevated token is missing. That sign-in needs a host with a display and may prompt, so unlike the other tiers this one cannot be refreshed non-interactively (a warm token cache does not imply warm browser-profile cookies).';

// Stable machine-readable code for the secondary-token fail-fast (elevated /
// chatsvcagg / ic3), so an agent can branch on `errorCode` instead of
// substring-matching the human message. The message names which token, which
// commands, and the tier-specific remedy.
const SECONDARY_TOKEN_UNAVAILABLE_CODE = 'secondary_token_unavailable';

// Machine-readable code + message for the BASIC-token fail-fast on the command
// path. When the cached basic token is absent/expired AND its refresh fails, the
// only remaining rung is an interactive browser sign-in — a 5-minute poll that a
// headless agent can never complete, so `get-user` (and every command) hung for
// minutes rather than erroring (reported 2026-07-19). On a non-interactive run
// (`acquireBasicViaBrowser: false`, wired from the absence of a TTY) we fail fast
// here instead; an interactive terminal keeps the auto-browser, and the explicit
// `login` command always has it.
const NOT_AUTHENTICATED_CODE = 'not_authenticated';

// A token the CLI obtained but could not save. Reported as itself, never as a
// missing sign-in: Entra single-uses the refresh token, so a redemption whose
// rotated refresh token was not saved leaves a spent one on disk, and the next
// command dead-ends in a login unless this one says why.
const CACHE_UNWRITABLE_CODE = 'token_cache_unwritable';

// Another process holds the token-cache lock (signing in, refreshing, signing
// out) past this call's wait budget. Reported as itself too: "run login" would
// start a second sign-in on top of the one in progress.
const SIGN_IN_IN_PROGRESS_CODE = 'sign_in_in_progress';

// The two failures a caller reports as they are, rather than trying the next
// rung of the ladder (a browser, or the generic "run login" message).
const mustReportAsIs = (r: Result<unknown, AuthError>): boolean =>
  !r.ok && r.error.type === 'auth_failed' && (r.error.code === CACHE_UNWRITABLE_CODE || r.error.code === SIGN_IN_IN_PROGRESS_CODE);

// How long a call waits for the lock: up to a whole sign-in where a person can
// see what is going on (a terminal, `login`), twenty seconds for an agent.
const LOCK_WAIT_INTERACTIVE_MS = 7 * 60_000;
const LOCK_WAIT_UNATTENDED_MS = 20_000;
const HOLDER_ACTIVITY: Readonly<Record<LockPurpose | 'unknown', string>> = {
  browser: 'signing in',
  refresh: 'refreshing its tokens',
  logout: 'signing out',
  unknown: 'using the token cache',
};

// Two Graph tokens of the same account: the same object id in the same tenant.
// A token without these claims never matches, so the caller replaces rather
// than merges.
const sameAccount = (cachedAccess: string | undefined, access: string): boolean => {
  if (!cachedAccess) return false;
  const before = decodeJwtPayload(cachedAccess);
  const after = decodeJwtPayload(access);
  return typeof before['oid'] === 'string' && before['oid'] === after['oid'] && typeof before['tid'] === 'string' && before['tid'] === after['tid'];
};
const NOT_AUTHENTICATED_MESSAGE =
  'Not signed in, or the cached session expired and its refresh failed. This command does not open a sign-in browser — run `ask-marcel-office login` (on a machine with a browser) first, then retry. Preflight with `ask-marcel-office scopes-check` (no Graph call).';

type FetchFn = (url: string, init?: RequestInit) => Promise<Response>;

// The rungs of the ladder that need a browser. auth-browser.ts supplies them to
// a session allowed to open one; a rung left out fails fast with the "run login"
// message instead.
type BrowserRungs = {
  readonly recaptureElevated?: (options?: { readonly awaitSignIn?: boolean }) => Promise<Result<AccessToken, AuthError>>;
  readonly recaptureChatsvcagg?: () => Promise<Result<AccessToken, AuthError>>;
  readonly recaptureIc3?: () => Promise<Result<AccessToken, AuthError>>;
};

// What a browser rung may do with the token cache: save what it captured, under
// the machine-wide lock.
type TokenCacheAccess = {
  readonly underLock: <T>(purpose: LockPurpose, task: () => Promise<Result<T, AuthError>>) => Promise<Result<T, AuthError>>;
  readonly persistElevated: (elevated: AccessToken) => Promise<Result<void, AuthError>>;
  readonly persistChatsvcagg: (chatsvcagg: AccessToken, region: string) => Promise<Result<void, AuthError>>;
  readonly persistIc3: (ic3: AccessToken, region: string) => Promise<Result<void, AuthError>>;
};

// What the ladder runs on. `fetchFn` defaults to the global and `lock` to the
// machine-wide lock beside the cache; see createAuthManagerFromApi.
type AuthLadderDeps = {
  readonly browserAuth: BrowserAuth;
  readonly cachePath: string;
  readonly browserProfileDir: string;
  readonly logger: Logger;
  readonly fs: FileSystem & AtomicFileWrites;
  readonly secondaryTokenCommands?: SecondaryTokenCommands;
  readonly acquireBasicViaBrowser: boolean;
  readonly fetchFn?: FetchFn;
  readonly lock?: TokenCacheLock;
  readonly browserRungs?: (cache: TokenCacheAccess) => BrowserRungs;
};

// The recovery ladder over the token cache: the cached token, a refresh-token
// redemption, then a browser rung or a fail-fast "run login". Built by
// auth-browser.ts, which owns the browser.
const createAuthLadder = (deps: AuthLadderDeps): AuthManager => {
  const { browserAuth, cachePath, browserProfileDir, logger, fs, acquireBasicViaBrowser } = deps;
  const secondaryTokenCommands = deps.secondaryTokenCommands ?? NO_SECONDARY_TOKEN_COMMANDS;
  const fetchFn = deps.fetchFn ?? globalThis.fetch;
  const lock = deps.lock ?? createSystemTokenCacheLock(fs, `${cachePath}.lock`);
  const lockWaitMs = acquireBasicViaBrowser ? LOCK_WAIT_INTERACTIVE_MS : LOCK_WAIT_UNATTENDED_MS;
  const underLock = async <T>(purpose: LockPurpose, task: () => Promise<Result<T, AuthError>>): Promise<Result<T, AuthError>> => {
    const locked = await lock.withLock(purpose, lockWaitMs, task);
    if (locked.ok) return locked.value;
    if (locked.error.type === 'lock_failed') {
      return err({ type: 'auth_failed', message: `the token-cache lock could not be taken (${locked.error.message})`, code: CACHE_UNWRITABLE_CODE });
    }
    const waited = Math.round(lockWaitMs / 1000);
    return err({
      type: 'auth_failed',
      message: `Another ask-marcel-office process is ${HOLDER_ACTIVITY[locked.error.purpose]}; this call waited ${waited} s for it. Retry once it finishes.`,
      code: SIGN_IN_IN_PROGRESS_CODE,
    });
  };

  const readCache = async (): Promise<CachedToken | null> => {
    const r = await fs.readJson<CachedToken>(cachePath);
    return r.ok ? r.value : null;
  };

  // The cache holds access and refresh tokens: replaced in one atomic step, so
  // a concurrent reader never sees half a file, and owner-only from the first
  // byte. A failed save is an error the caller reports (see CACHE_UNWRITABLE_CODE).
  const writeCache = async (next: CachedToken): Promise<Result<void, AuthError>> => {
    const written = await fs.writeTextAtomic(cachePath, JSON.stringify(next), 0o600);
    if (written.ok) return ok(undefined);
    const detail = 'message' in written.error ? written.error.message : written.error.type;
    return err({ type: 'auth_failed', message: `the token cache could not be saved (${detail})`, code: CACHE_UNWRITABLE_CODE });
  };

  // Entra single-uses the refresh token, so this process redeems it one call at
  // a time: a redemption queued behind another must spend the refresh token
  // that one saved, never the spent one both read before. Every task re-reads
  // the cache when its turn comes.
  const settled = (): void => undefined;
  let redemptionTail: Promise<void> = Promise.resolve();
  const oneRedemptionAtATime = <T>(task: () => Promise<T>): Promise<T> => {
    const run = redemptionTail.then(task);
    redemptionTail = run.then(settled, settled);
    return run;
  };

  // A sign-in keeps the other tokens of the SAME account (its chat, guest and
  // elevated tokens are still good) and drops everything of another account, so
  // one account's Graph token never sits next to another's chat token.
  const persistTeams = async (access: AccessToken, refresh: string | null, elevated?: AccessToken | null): Promise<Result<void, AuthError>> => {
    const existing = await readCache();
    const kept: Partial<CachedToken> = existing !== null && sameAccount(existing.access_token, access) ? existing : {};
    const claims = decodeJwtPayload(access);
    const exp = claims.exp as number | undefined;
    const cached: CachedToken = { ...kept, access_token: access, expires_on: exp ?? 0, refresh_token: refresh ?? '' };
    if (elevated) {
      const elevatedClaims = decodeJwtPayload(elevated);
      const elevatedExp = elevatedClaims.exp as number | undefined;
      cached.elevated_access_token = elevated;
      cached.elevated_expires_on = elevatedExp ?? 0;
    }
    return writeCache(cached);
  };

  const persistElevated = async (elevated: AccessToken): Promise<Result<void, AuthError>> => {
    const existing = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
    const elevatedClaims = decodeJwtPayload(elevated);
    const elevatedExp = elevatedClaims.exp as number | undefined;
    const next: CachedToken = { ...existing, elevated_access_token: elevated, elevated_expires_on: elevatedExp ?? 0 };
    return writeCache(next);
  };

  const persistChatsvcagg = async (chatsvcagg: AccessToken, region: string): Promise<Result<void, AuthError>> => {
    const existing = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
    const claims = decodeJwtPayload(chatsvcagg);
    const exp = claims.exp as number | undefined;
    const next: CachedToken = {
      ...existing,
      chatsvcagg_access_token: chatsvcagg,
      chatsvcagg_expires_on: exp ?? 0,
      chatsvcagg_region: region,
    };
    return writeCache(next);
  };

  const persistIc3 = async (ic3: AccessToken, region: string): Promise<Result<void, AuthError>> => {
    const existing = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
    const claims = decodeJwtPayload(ic3);
    const exp = claims.exp as number | undefined;
    const next: CachedToken = {
      ...existing,
      ic3_access_token: ic3,
      ic3_expires_on: exp ?? 0,
      // IC3 shares the chatsvcagg region (regions match per tenant). Persist
      // it whether or not a chatsvcagg capture also produced a region — this
      // path may run standalone (e.g. cached IC3 expired but chatsvcagg fine).
      chatsvcagg_region: region,
    };
    return writeCache(next);
  };

  /**
   * The one place this process redeems a refresh token. Three callers need the
   * identical dance against different authorities and scopes — the basic Graph
   * refresh (`/common`), the substrate tiers (`/common`, non-Graph audience),
   * and a partner tenant's guest token (`/{tenantId}`) — so the POST lives here
   * once (Rule of Three) rather than a fourth time at the next tier.
   *
   * `Origin` is NOT optional. The Teams client is registered as a Single-Page
   * Application, and Entra refuses SPA refresh-token redemption that does not
   * arrive as a cross-origin request: `AADSTS9002327`. It fails BEFORE the grant
   * is evaluated, so a missing Origin looks like "the tenant refused you" rather
   * than "the header is absent".
   *
   * The deadline is new to all three callers: `auth.ts` previously set none, so a
   * hung token endpoint hung the CLI with no upper bound (rule 29). Extracting
   * these into a single fetch made a per-caller exception arbitrary.
   */
  const redeemRefreshToken = async (
    refreshTokenValue: string,
    authority: string,
    scope: string
  ): Promise<Result<{ readonly accessToken: string; readonly expiresIn: number; readonly refreshToken: string | undefined }, AuthError>> => {
    const body = new URLSearchParams({ client_id: CLIENT_ID, grant_type: 'refresh_token', refresh_token: refreshTokenValue, scope });
    let res: Response;
    try {
      res = await fetchFn(`https://login.microsoftonline.com/${authority}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'content-type': 'application/x-www-form-urlencoded', Origin: SPA_ORIGIN },
        body,
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: msg });
    }
    if (!res.ok) return err({ type: 'auth_failed', message: `refresh failed (${res.status})` });
    const json = (await res.json()) as { access_token?: string; expires_in?: number; refresh_token?: string };
    return ok({ accessToken: json.access_token ?? '', expiresIn: json.expires_in ?? 0, refreshToken: json.refresh_token });
  };

  const refreshToken = async (cached: CachedToken): Promise<Result<AccessToken, AuthError>> => {
    const redeemed = await redeemRefreshToken(cached.refresh_token, 'common', SCOPES);
    if (!redeemed.ok) return redeemed;
    const json = { access_token: redeemed.value.accessToken, expires_in: redeemed.value.expiresIn, refresh_token: redeemed.value.refreshToken };
    const validated = accessToken(json.access_token ?? '');
    if (!validated.ok) return err({ type: 'auth_failed', message: 'invalid token from refresh' });
    // Merge into the cache as it is NOW, not the snapshot read before the
    // network call: another process may have saved a token meanwhile. Spreading
    // the existing cache keeps the elevated / chatsvcagg / ic3 / guest tokens;
    // elevated carries no refresh token of its own, so wiping it would cost a
    // forced browser re-login.
    const latest = (await readCache()) ?? cached;
    const token: CachedToken = {
      ...latest,
      access_token: validated.value,
      expires_on: Math.floor(Date.now() / 1000) + json.expires_in,
      refresh_token: json.refresh_token ?? latest.refresh_token,
    };
    const saved = await writeCache(token);
    if (!saved.ok) return saved;
    logger.info('auth.ladder.rung', { rung: 'refresh' });
    return ok(validated.value);
  };

  // Waits its turn, then looks again: a redemption that ran meanwhile may have
  // saved a fresh token, and redeeming again would spend its rotated refresh
  // token for nothing.
  const refreshBasicInTurn = (): Promise<Result<AccessToken, AuthError>> =>
    oneRedemptionAtATime(async () =>
      underLock('refresh', async () => {
        const latest = await readCache();
        const saved = accessToken(latest?.access_token ?? '');
        if (saved.ok) {
          logger.info('auth.ladder.rung', { rung: 'cache_after_wait' });
          return ok(saved.value);
        }
        if (!latest?.refresh_token) return err({ type: 'auth_failed', message: 'the cached session ended while this call waited for its turn' });
        return refreshToken(latest);
      })
    );

  // Substrate tokens (chatsvcagg / ic3) carry the SAME Teams appid as the Graph
  // token, so the shared refresh_token redeems for their audiences too — a
  // headless HTTP refresh, no browser. This lets the command path self-heal a
  // lapsed substrate token instead of dead-ending in "run login". The rotated
  // refresh_token is written back to the shared slot (AAD rotates + single-uses
  // it) so the Graph token keeps refreshing from the same RT. The region is not
  // observable here (we never hit the substrate URL) — reuse the last-known
  // region from cache; a mismatch surfaces later as a clean 404, never a wrong
  // read. Callers guard on `cached.refresh_token` being present.
  const refreshSubstrateToken = async (
    cached: CachedToken,
    resource: string,
    persist: (token: AccessToken, region: string) => Promise<Result<void, AuthError>>,
    rung: string
  ): Promise<Result<AccessToken, AuthError>> => {
    const redeemed = await redeemRefreshToken(cached.refresh_token, 'common', `${resource}/.default offline_access`);
    if (!redeemed.ok) return err({ type: 'auth_failed', message: `${rung}: ${redeemed.error.type === 'auth_failed' ? redeemed.error.message : 'cancelled'}` });
    const json = { access_token: redeemed.value.accessToken, refresh_token: redeemed.value.refreshToken };
    const raw = json.access_token ?? '';
    // Substrate tokens carry a non-Graph audience, so `accessToken()` (the Graph
    // validator) would reject them — accept any well-formed JWT AAD just minted.
    if (!raw.startsWith('eyJ')) return err({ type: 'auth_failed', message: `${rung} returned an unusable token` });
    const substrateToken = accessTokenUnsafe(raw);
    const saved = await persist(substrateToken, cached.chatsvcagg_region ?? DEFAULT_CHATSVCAGG_REGION);
    if (!saved.ok) return saved;
    if (json.refresh_token && json.refresh_token !== cached.refresh_token) {
      const latest = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
      const savedRefresh = await writeCache({ ...latest, refresh_token: json.refresh_token });
      if (!savedRefresh.ok) return savedRefresh;
    }
    logger.info('auth.ladder.rung', { rung });
    return ok(substrateToken);
  };

  type SubstrateTier = {
    readonly fresh: (cached: CachedToken | null) => string | undefined;
    readonly resource: string;
    readonly persist: (token: AccessToken, region: string) => Promise<Result<void, AuthError>>;
    readonly rung: string;
  };

  // The same turn-taking for a substrate token. A token saved while this call
  // waited is used, unless it is the very one the caller found dead.
  const refreshSubstrateInTurn = (tier: SubstrateTier, dead: string | undefined): Promise<Result<AccessToken, AuthError>> =>
    oneRedemptionAtATime(async () =>
      underLock('refresh', async () => {
        const latest = await readCache();
        const saved = tier.fresh(latest);
        if (saved !== undefined && saved.startsWith('eyJ') && saved !== dead) return ok(accessTokenUnsafe(saved));
        if (!latest?.refresh_token) return err({ type: 'auth_failed', message: `${tier.rung}: the cached session ended while this call waited for its turn` });
        return refreshSubstrateToken(latest, tier.resource, tier.persist, tier.rung);
      })
    );

  // Track the elevated-capture outcome from the most recent
  // browser-acquired session so the login command can surface it to the
  // user via `getLastElevatedOutcome()`. Reset to null on every fresh
  // `acquireViaBrowser` so stale outcomes don't leak across login attempts.
  let lastElevatedOutcome: ElevatedOutcome | null = null;
  let lastChatsvcaggOutcome: ElevatedOutcome | null = null;

  // A forced login promises to refresh all four tokens, but the browser captures the
  // chatsvcagg / ic3 substrate bearers only opportunistically (they fire from Teams
  // traffic that may not occur in the settle window — ic3 needs a chat-history load).
  // Redeem any the dance missed from the freshly-minted refresh token, headlessly —
  // the same path the on-demand getters use, so no second browser is needed.
  // One missed substrate token, redeemed in its turn and under the lock, from
  // the refresh token as it is then (`fresh` only if the cache vanished).
  const redeemMissedTier = async (resource: string, persist: SubstrateTier['persist'], rung: string, fresh: CachedToken): Promise<unknown> =>
    oneRedemptionAtATime(async () => underLock('refresh', async () => refreshSubstrateToken((await readCache()) ?? fresh, resource, persist, rung)));

  const redeemMissedSubstrateAtLogin = async (chatsvcaggCaptured: boolean, ic3Captured: boolean): Promise<void> => {
    if (chatsvcaggCaptured && ic3Captured) return;
    const fresh = await readCache();
    if (!fresh?.refresh_token) return;
    if (!chatsvcaggCaptured) await redeemMissedTier(CHATSVCAGG_RESOURCE, persistChatsvcagg, 'auth.chatsvcagg.login_rt_redeem', fresh);
    if (!ic3Captured) await redeemMissedTier(IC3_RESOURCE, persistIc3, 'auth.ic3.login_rt_redeem', fresh);
  };

  // The rungs that need a browser, when this session may open one.
  const browser = deps.browserRungs?.({ underLock, persistElevated, persistChatsvcagg, persistIc3 });

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

  const getAccessToken = async (options?: { force?: boolean }): Promise<Result<AccessToken, AuthError>> => {
    // `login --force` skips the cache + refresh rungs so a warm session still
    // re-captures every token via the browser (elevated carries no refresh_token,
    // so a cache-hit login would otherwise never refresh it).
    if (!options?.force) {
      const cached = await readCache();
      if (cached) {
        const validated = accessToken(cached.access_token);
        if (validated.ok) {
          logger.info('auth.ladder.rung', { rung: 'cache' });
          return ok(validated.value);
        }
      }
      if (cached?.refresh_token) {
        const refreshed = await refreshBasicInTurn();
        if (refreshed.ok || mustReportAsIs(refreshed)) return refreshed;
      }
    }
    // Command path: never launch an interactive browser for the basic token —
    // fail fast with a single-line "run login" error instead of the 5-minute
    // headless-hang poll. The `login` command's manager sets this true.
    if (!acquireBasicViaBrowser) return err({ type: 'auth_failed', message: NOT_AUTHENTICATED_MESSAGE, code: NOT_AUTHENTICATED_CODE });
    // Under --force, tell the browser layer to skip its concurrent-refresh probe
    // so the still-valid cached token cannot short-circuit the full re-capture.
    return acquireViaBrowserShared(options?.force ?? false);
  };

  const ELEVATED_BUFFER_SECONDS = 300;

  const freshGuestToken = (cached: CachedToken | null, tenant: TenantId): string | undefined => {
    const slot = cached?.guest_tokens;
    // `Object.hasOwn`, never `in` — the key is a tenant GUID from outside this
    // process and `in` would match inherited prototype keys.
    if (!slot || !Object.hasOwn(slot, tenant)) return undefined;
    const entry = slot[tenant];
    if (!entry?.access_token || !entry.expires_on) return undefined;
    if (Date.now() / 1000 >= entry.expires_on - ELEVATED_BUFFER_SECONDS) return undefined;
    return entry.access_token;
  };

  /**
   * A Graph token issued by a PARTNER tenant's authority, for a user who is a
   * guest there. Cache -> headless redemption of the shared refresh token; never
   * a browser (the RT already proves the identity; the partner tenant only has to
   * agree the user is a guest).
   *
   * Live-probed 2026-07-16: the guest token reads the partner tenant's whole
   * `/drives` surface (metadata 200, `/content` 302, `?format=pdf` 302), not just
   * `/shares` — which is why the drive-item family takes `--tenant-id` rather
   * than this shipping as a one-shot share-URL download command.
   */
  const noGuestCredentials = (tenant: TenantId): Result<AccessToken, AuthError> =>
    err({
      type: 'auth_failed',
      message: `no cached credentials to obtain a guest token for tenant ${tenant} — run \`ask-marcel-office login\``,
      code: SECONDARY_TOKEN_UNAVAILABLE_CODE,
    });

  const getGuestAccessToken = async (tenant: TenantId): Promise<Result<AccessToken, AuthError>> => {
    const cached = await readCache();
    const fresh = freshGuestToken(cached, tenant);
    if (fresh !== undefined) {
      logger.info('auth.guest.cache_hit', { tenant });
      return ok(accessTokenUnsafe(fresh));
    }
    if (!cached?.refresh_token) return noGuestCredentials(tenant);
    return oneRedemptionAtATime(async () => underLock('refresh', async () => redeemGuestToken(tenant)));
  };

  // Runs in its turn, so it reads the cache again: a guest token or a rotated
  // refresh token saved while it waited is the one to use.
  const redeemGuestToken = async (tenant: TenantId): Promise<Result<AccessToken, AuthError>> => {
    const current = await readCache();
    const savedMeanwhile = freshGuestToken(current, tenant);
    if (savedMeanwhile !== undefined) return ok(accessTokenUnsafe(savedMeanwhile));
    if (!current?.refresh_token) return noGuestCredentials(tenant);
    const redeemed = await redeemRefreshToken(current.refresh_token, tenant, SCOPES);
    if (!redeemed.ok) {
      // Naming the tenant matters: the caller passed a `--tenant-id` (or resolved
      // one from a sharing URL) and needs to know WHICH tenant refused, not that
      // "a refresh failed". A partner tenant refuses when the user is not a guest
      // there, or when the Teams client is not consented in it (AADSTS65001).
      const detail = redeemed.error.type === 'auth_failed' ? redeemed.error.message : 'cancelled';
      return err({
        type: 'auth_failed',
        message: `tenant ${tenant} refused a guest token (${detail}) — you may not be a guest in that tenant, or its administrator has not consented to this client`,
        code: SECONDARY_TOKEN_UNAVAILABLE_CODE,
      });
    }
    const validated = accessToken(redeemed.value.accessToken);
    if (!validated.ok) return err({ type: 'auth_failed', message: `tenant ${tenant} returned an unusable guest token` });

    // Re-read before writing: this function awaited a network call, so the cache
    // on disk may have moved under us (another tier's refresh). Merge into the
    // LATEST, never into the `cached` snapshot taken before the await — that is
    // the 2026-07-15 clobber, which silently dropped sibling tokens.
    const latest = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
    const saved = await writeCache({
      ...latest,
      guest_tokens: { ...latest.guest_tokens, [tenant]: { access_token: validated.value, expires_on: Math.floor(Date.now() / 1000) + redeemed.value.expiresIn } },
      // Entra single-uses and rotates the SPA refresh token. Dropping the rotated
      // one leaves a spent RT on disk and the next command dead-ends in a login.
      refresh_token: redeemed.value.refreshToken ?? latest.refresh_token,
    });
    if (!saved.ok) return saved;
    logger.info('auth.ladder.rung', { rung: 'guest', tenant });
    return ok(validated.value);
  };

  /**
   * Narrowing helper: returns the cached elevated token if it exists,
   * has an expiry, and is at least 5 minutes from expiring; otherwise
   * undefined. Returning the token directly (instead of a boolean)
   * lets callers skip a redundant `cached?.elevated_access_token`
   * second-check after `isElevatedFresh` returns truthy.
   */
  const freshElevatedToken = (cached: CachedToken | null): string | undefined => {
    if (!cached?.elevated_access_token || !cached.elevated_expires_on) return undefined;
    if (Date.now() / 1000 >= cached.elevated_expires_on - ELEVATED_BUFFER_SECONDS) return undefined;
    return cached.elevated_access_token;
  };

  // Decode-only preflight: does the persisted cache hold an elevated token the
  // historical-version commands could use right now? Reuses `freshElevatedToken`
  // (same 300s buffer the download path applies), so `available` never disagrees
  // with what `getElevatedAccessToken` would decide — but it never captures or
  // refreshes. `expiresInSeconds` is the raw exp − now (negative once past), so a
  // caller sees the runway even when the token is inside the buffer.
  const getCachedBasicToken = async (): Promise<AccessToken | undefined> => {
    const cached = await readCache();
    const token = cached?.access_token;
    return token === undefined || token === '' ? undefined : accessTokenUnsafe(token);
  };

  const getCachedElevatedInfo = async (): Promise<CachedTierInfo> => {
    const cached = await readCache();
    const exp = cached?.elevated_expires_on;
    const expiresInSeconds = typeof exp === 'number' ? Math.floor(exp - Date.now() / 1000) : undefined;
    return { available: freshElevatedToken(cached) !== undefined, expiresInSeconds, scopes: decodeScopes(cached?.elevated_access_token) };
  };

  const getElevatedAccessToken = async (options?: { readonly awaitSignIn?: boolean }): Promise<Result<AccessToken, AuthError>> => {
    const fresh = freshElevatedToken(await readCache());
    const validated = fresh !== undefined ? accessToken(fresh) : null;
    if (validated?.ok) {
      logger.info('auth.elevated.cache_hit');
      return ok(validated.value);
    }
    // Elevated absent, expired, or malformed; need to re-capture.
    const recapture = browser?.recaptureElevated;
    if (!recapture)
      return err({
        type: 'auth_failed',
        message: failFastSecondaryMessage('Elevated (M365)', secondaryTokenCommands.elevated, RECAPTURE_ELEVATED_VIA_LOGIN),
        code: SECONDARY_TOKEN_UNAVAILABLE_CODE,
      });
    // Normally the persistent profile cookies do the silent SSO with no UI
    // prompt. When they have lapsed the tenant serves a sign-in form instead,
    // and `awaitSignIn` decides whether we wait for it to be filled in.
    return recapture(options);
  };

  // chatsvcagg shares the elevated expiry buffer. The token itself carries no
  // refresh_token, but it shares the Teams appid with the Graph token — so the
  // shared RT redeems for it (see `refreshSubstrateToken`); browser re-capture
  // is the fallback when that RT is absent or rejected.
  const freshChatsvcaggToken = (cached: CachedToken | null): string | undefined => {
    if (!cached?.chatsvcagg_access_token || !cached.chatsvcagg_expires_on) return undefined;
    if (Date.now() / 1000 >= cached.chatsvcagg_expires_on - ELEVATED_BUFFER_SECONDS) return undefined;
    return cached.chatsvcagg_access_token;
  };

  // Decode-only preflight (mirrors getCachedElevatedInfo) so login's four-token
  // status reports the chatsvcagg substrate token without capturing or refreshing.
  const getCachedChatsvcaggInfo = async (): Promise<CachedTierInfo> => {
    const cached = await readCache();
    const exp = cached?.chatsvcagg_expires_on;
    const expiresInSeconds = typeof exp === 'number' ? Math.floor(exp - Date.now() / 1000) : undefined;
    return { available: freshChatsvcaggToken(cached) !== undefined, expiresInSeconds, scopes: decodeScopes(cached?.chatsvcagg_access_token) };
  };

  // `ignoreCache` is how a caller says the cached token is DEAD rather than
  // stale. A substrate token can be revoked server-side while still inside its
  // expiry window (a second sign-in invalidates the previous session's), and
  // nothing in the cache records that — so the freshness check happily returns
  // a token the service will 401. The substrate request path sets this after a
  // 401 to force a redemption from the shared refresh token. Observed live
  // 2026-08-31: a token minted at 07:18 was rejected at 13:47 with hours of
  // stated life left, while a freshly redeemed one worked instantly.
  const getChatsvcaggAccessToken = async (options?: { readonly ignoreCache?: boolean }): Promise<Result<AccessToken, AuthError>> => {
    // chatsvcagg tokens carry `aud=https://chatsvcagg.teams.microsoft.com`,
    // not Graph — the `accessToken()` validator's `isGraphToken` check
    // would reject every cached chatsvcagg token and force a recapture on
    // every call. `freshChatsvcaggToken` already validates expiry from the
    // JWT payload, which is the only thing we need at this boundary.
    const cached = await readCache();
    const fresh = options?.ignoreCache === true ? undefined : freshChatsvcaggToken(cached);
    if (fresh !== undefined && fresh.startsWith('eyJ')) {
      logger.info('auth.chatsvcagg.cache_hit');
      return ok(accessTokenUnsafe(fresh));
    }
    // Headless FIRST, on every manager. The shared RT mints this audience over
    // HTTP, so a browser is never the cheaper route; when this sat inside the
    // fail-fast branch below, a manager that was ALLOWED a browser skipped the
    // refresh entirely and opened a window for a token an HTTP call could mint.
    if (cached?.refresh_token) {
      const tier = { fresh: freshChatsvcaggToken, resource: CHATSVCAGG_RESOURCE, persist: persistChatsvcagg, rung: 'auth.chatsvcagg.refresh' };
      const refreshed = await refreshSubstrateInTurn(tier, options?.ignoreCache === true ? cached.chatsvcagg_access_token : undefined);
      if (refreshed.ok || mustReportAsIs(refreshed)) return refreshed;
    }
    const recapture = browser?.recaptureChatsvcagg;
    if (!recapture) {
      return err({
        type: 'auth_failed',
        message: failFastSecondaryMessage('chatsvcagg (Teams chat)', secondaryTokenCommands.chatsvcagg, RECAPTURE_VIA_LOGIN),
        code: SECONDARY_TOKEN_UNAVAILABLE_CODE,
      });
    }
    return recapture();
  };

  const getChatsvcaggRegion = async (): Promise<string> => {
    // Region MUST be paired with a live chatsvcagg bearer — the substrate
    // routes per region, and a mismatched region produces an immediate 404.
    // Trigger the token path first so a freshly-captured region lands in
    // cache before we read it (no-op when the cached token is still warm).
    await getChatsvcaggAccessToken();
    const cached = await readCache();
    return cached?.chatsvcagg_region ?? DEFAULT_CHATSVCAGG_REGION;
  };

  // IC3 shares the same expiry buffer / recovery shape as chatsvcagg: the token
  // has no refresh_token of its own but rides the shared Teams RT for a headless
  // refresh, with browser re-capture as the fallback. Region is reused from the
  // chatsvcagg slot.
  const freshIc3Token = (cached: CachedToken | null): string | undefined => {
    if (!cached?.ic3_access_token || !cached.ic3_expires_on) return undefined;
    if (Date.now() / 1000 >= cached.ic3_expires_on - ELEVATED_BUFFER_SECONDS) return undefined;
    return cached.ic3_access_token;
  };

  // Decode-only preflight (mirrors getCachedElevatedInfo) for the ic3 substrate token.
  const getCachedIc3Info = async (): Promise<CachedTierInfo> => {
    const cached = await readCache();
    const exp = cached?.ic3_expires_on;
    const expiresInSeconds = typeof exp === 'number' ? Math.floor(exp - Date.now() / 1000) : undefined;
    return { available: freshIc3Token(cached) !== undefined, expiresInSeconds, scopes: decodeScopes(cached?.ic3_access_token) };
  };

  /** `ignoreCache` as on `getChatsvcaggAccessToken`: the cached token is dead, not stale. */
  const getIc3AccessToken = async (options?: { readonly ignoreCache?: boolean }): Promise<Result<AccessToken, AuthError>> => {
    // IC3 tokens carry `aud=https://ic3.teams.office.com`, not Graph — the
    // `accessToken()` validator's `isGraphToken` check would reject every
    // cached IC3 token and force a recapture on every call. `freshIc3Token`
    // validates expiry from the JWT payload, which is the only thing we
    // need at this boundary.
    const cached = await readCache();
    const fresh = options?.ignoreCache === true ? undefined : freshIc3Token(cached);
    if (fresh !== undefined && fresh.startsWith('eyJ')) {
      logger.info('auth.ic3.cache_hit');
      return ok(accessTokenUnsafe(fresh));
    }
    // Headless first, same reasoning as chatsvcagg above.
    if (cached?.refresh_token) {
      const tier = { fresh: freshIc3Token, resource: IC3_RESOURCE, persist: persistIc3, rung: 'auth.ic3.refresh' };
      const refreshed = await refreshSubstrateInTurn(tier, options?.ignoreCache === true ? cached.ic3_access_token : undefined);
      if (refreshed.ok || mustReportAsIs(refreshed)) return refreshed;
    }
    const recapture = browser?.recaptureIc3;
    if (!recapture) {
      return err({
        type: 'auth_failed',
        message: failFastSecondaryMessage('ic3 (Teams chat history)', secondaryTokenCommands.ic3, RECAPTURE_VIA_LOGIN),
        code: SECONDARY_TOKEN_UNAVAILABLE_CODE,
      });
    }
    return recapture();
  };

  // Signing out holds the lock too, so the profile is never wiped under a
  // sign-in another process is running.
  const logout = async (): Promise<Result<void, AuthError>> => underLock('logout', async () => logoutHoldingLock());
  const logoutHoldingLock = async (): Promise<Result<void, AuthError>> => {
    try {
      await fs.deleteIfExists(cachePath);
      // Wipe the Playwright persistent browser profile too. Previously
      // `logout` only cleared the token cache, leaving stale auth cookies
      // behind — so the documented remediation
      // `ask-marcel-office logout && ask-marcel-office login` would reuse
      // the same expired cookies on the next elevated-capture attempt
      // and fail again. The profile contains only auth-flow state;
      // wiping it forces silent SSO to re-authenticate against
      // m365.cloud.microsoft on the next login. Both ops are
      // best-effort: `deleteDirIfExists` already returns ok when the
      // directory does not exist.
      await fs.deleteDirIfExists(browserProfileDir);
      await browserAuth.close();
      return ok(undefined);
    } catch (e) {
      await fs.deleteIfExists(cachePath);
      await fs.deleteDirIfExists(browserProfileDir);
      return err({ type: 'auth_failed', message: e instanceof Error ? e.message : String(e) });
    }
  };

  // `redeemMissedSubstrateAtLogin` only ever ran INSIDE the browser dance, so a
  // plain `login` on a warm cache never reached it. Same redemption, callable on
  // its own and keyed on what is actually stale.
  const warmSubstrateTokens = async (): Promise<void> => {
    const cached = await readCache();
    await redeemMissedSubstrateAtLogin(freshChatsvcaggToken(cached) !== undefined, freshIc3Token(cached) !== undefined);
  };

  const getLastElevatedOutcome = (): ElevatedOutcome | null => lastElevatedOutcome;
  const getLastChatsvcaggOutcome = (): ElevatedOutcome | null => lastChatsvcaggOutcome;

  return {
    getAccessToken,
    getElevatedAccessToken,
    getGuestAccessToken,
    getChatsvcaggAccessToken,
    getChatsvcaggRegion,
    getIc3AccessToken,
    warmSubstrateTokens,
    logout,
    getLastElevatedOutcome,
    getLastChatsvcaggOutcome,
    getCachedBasicToken,
    getCachedElevatedInfo,
    getCachedChatsvcaggInfo,
    getCachedIc3Info,
  };
};

export { createAuthLadder, neededByNote, NO_SECONDARY_TOKEN_COMMANDS };
export type { AuthError, AuthManager, BrowserRungs, CachedTierInfo, ElevatedOutcome, FetchFn, SecondaryTokenCommands, TokenCacheAccess };
