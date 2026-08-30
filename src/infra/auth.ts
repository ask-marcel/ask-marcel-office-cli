import type { AccessToken } from '../domain/access-token.ts';
import { accessToken, accessTokenUnsafe } from '../domain/access-token.ts';
import { decodeJwtPayload } from '../domain/jwt-utils.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { TenantId } from '../domain/tenant-id.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { BrowserAuth, ElevatedFailureReason } from './browser-auth.ts';
import { createBrowserAuth } from './browser-auth.ts';
import { createBunFileSystem } from './filesystem-bun.ts';
import { createNodeFileSystem } from './filesystem-node.ts';
import { REQUEST_TIMEOUT_MS } from './network-error.ts';

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
  getElevatedAccessToken: () => Promise<Result<AccessToken, AuthError>>;
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
  getChatsvcaggAccessToken: () => Promise<Result<AccessToken, AuthError>>;
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
  getIc3AccessToken: () => Promise<Result<AccessToken, AuthError>>;
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
const TEAMS_URL = 'https://teams.microsoft.com/';
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
const failFastSecondaryMessage = (token: string, commands: string, remedy: string): string =>
  `${token} token is expired or was not captured at login. ${remedy} — the CLI does not open a browser per command for this token. Preflight token validity with \`ask-marcel-office scopes-check\` (no Graph call) before a long unattended run. (Commands that need it: ${commands}.)`;

/**
 * Command names quoted in the secondary-token error messages, per token kind.
 * The composition root derives these from the command registry
 * (`needsElevatedToken` / `needsSubstrateToken` flags) and injects them; the
 * defaults below are the corrected registry sets at the time of writing, so
 * direct `createAuthManagerFromApi` callers still get accurate messages. The
 * registry sets are pinned in meta.test.ts and the wiring in
 * build-deps.test.ts, so drift surfaces there, not in a stale user message
 * (the old hardcoded elevated list omitted `get-user`).
 */
type SecondaryTokenCommands = {
  readonly elevated: ReadonlyArray<string>;
  readonly chatsvcagg: ReadonlyArray<string>;
  readonly ic3: ReadonlyArray<string>;
};

const DEFAULT_SECONDARY_TOKEN_COMMANDS: SecondaryTokenCommands = {
  elevated: ['download-drive-item-version', 'get-chat', 'get-user', 'list-chats'],
  chatsvcagg: ['find-chats-with-user', 'get-teams-chat-message', 'list-teams-chat-messages', 'list-teams-chats-with-messages'],
  ic3: ['list-teams-chat-history'],
};

const commandList = (names: ReadonlyArray<string>): string => names.join(', ');

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
const NOT_AUTHENTICATED_MESSAGE =
  'Not signed in, or the cached session expired and its refresh failed. This command does not open a sign-in browser — run `ask-marcel-office login` (on a machine with a browser) first, then retry. Preflight with `ask-marcel-office scopes-check` (no Graph call).';

const createAuthManagerFromApi = (
  browserAuth: BrowserAuth,
  cachePath: string,
  browserProfileDir: string,
  logger: Logger,
  fs: FileSystem,
  recaptureSecondaryViaBrowser: boolean = true,
  secondaryTokenCommands: SecondaryTokenCommands = DEFAULT_SECONDARY_TOKEN_COMMANDS,
  acquireBasicViaBrowser: boolean = true,
  // Elevated gets its OWN browser gate, separate from chatsvcagg / ic3. Those two
  // self-heal by redeeming the shared refresh token from INSIDE the
  // `!recaptureSecondaryViaBrowser` branch below, so turning the shared flag on
  // would skip that headless refresh and open a browser instead — strictly worse.
  // Elevated carries no refresh token, so its only question is whether a browser is
  // permitted; an interactive session answers yes and refreshes it in ~17s of silent
  // SSO against the persistent profile. Defaults to the shared flag so every existing
  // caller keeps the behaviour it had.
  recaptureElevatedViaBrowser: boolean = recaptureSecondaryViaBrowser
): AuthManager => {
  const readCache = async (): Promise<CachedToken | null> => {
    const r = await fs.readJson<CachedToken>(cachePath);
    return r.ok ? r.value : null;
  };

  const writeCache = async (next: CachedToken): Promise<void> => {
    await fs.writeText(cachePath, JSON.stringify(next));
    // The cache holds access + refresh tokens — owner-only. Best-effort:
    // a chmod failure must not fail the auth flow (the write itself succeeded).
    await fs.chmod(cachePath, 0o600);
  };

  const persistTeams = async (access: AccessToken, refresh: string | null, elevated?: AccessToken | null): Promise<void> => {
    const claims = decodeJwtPayload(access);
    const exp = claims.exp as number | undefined;
    const cached: CachedToken = { access_token: access, expires_on: exp ?? 0, refresh_token: refresh ?? '' };
    if (elevated) {
      const elevatedClaims = decodeJwtPayload(elevated);
      const elevatedExp = elevatedClaims.exp as number | undefined;
      cached.elevated_access_token = elevated;
      cached.elevated_expires_on = elevatedExp ?? 0;
    }
    await writeCache(cached);
  };

  const persistElevated = async (elevated: AccessToken): Promise<void> => {
    const existing = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
    const elevatedClaims = decodeJwtPayload(elevated);
    const elevatedExp = elevatedClaims.exp as number | undefined;
    const next: CachedToken = { ...existing, elevated_access_token: elevated, elevated_expires_on: elevatedExp ?? 0 };
    await writeCache(next);
  };

  const persistChatsvcagg = async (chatsvcagg: AccessToken, region: string): Promise<void> => {
    const existing = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
    const claims = decodeJwtPayload(chatsvcagg);
    const exp = claims.exp as number | undefined;
    const next: CachedToken = {
      ...existing,
      chatsvcagg_access_token: chatsvcagg,
      chatsvcagg_expires_on: exp ?? 0,
      chatsvcagg_region: region,
    };
    await writeCache(next);
  };

  const persistIc3 = async (ic3: AccessToken, region: string): Promise<void> => {
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
    await writeCache(next);
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
      res = await fetch(`https://login.microsoftonline.com/${authority}/oauth2/v2.0/token`, {
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
    const token: CachedToken = {
      // Spread the existing cache FIRST so a basic-token refresh preserves the
      // elevated / chatsvcagg / ic3 tokens (and chatsvcagg_region). Without this,
      // every silent refresh wiped them — and elevated, carrying no refresh token
      // of its own, could then only be recovered by a forced browser re-login.
      ...cached,
      access_token: validated.value,
      expires_on: Math.floor(Date.now() / 1000) + json.expires_in,
      refresh_token: json.refresh_token ?? cached.refresh_token,
    };
    await writeCache(token);
    logger.info('auth.ladder.rung', { rung: 'refresh' });
    return ok(validated.value);
  };

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
    persist: (token: AccessToken, region: string) => Promise<void>,
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
    await persist(substrateToken, cached.chatsvcagg_region ?? DEFAULT_CHATSVCAGG_REGION);
    if (json.refresh_token && json.refresh_token !== cached.refresh_token) {
      const latest = (await readCache()) ?? { access_token: '', expires_on: 0, refresh_token: '' };
      await writeCache({ ...latest, refresh_token: json.refresh_token });
    }
    logger.info('auth.ladder.rung', { rung });
    return ok(substrateToken);
  };

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
  const redeemMissedSubstrateAtLogin = async (chatsvcaggCaptured: boolean, ic3Captured: boolean): Promise<void> => {
    if (chatsvcaggCaptured && ic3Captured) return;
    const fresh = await readCache();
    if (!fresh?.refresh_token) return;
    if (!chatsvcaggCaptured) await refreshSubstrateToken(fresh, CHATSVCAGG_RESOURCE, persistChatsvcagg, 'auth.chatsvcagg.login_rt_redeem');
    if (!ic3Captured) await refreshSubstrateToken(fresh, IC3_RESOURCE, persistIc3, 'auth.ic3.login_rt_redeem');
  };

  const acquireViaBrowser = async (force = false): Promise<Result<AccessToken, AuthError>> => {
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
        return ok(result.accessToken);
      }
      const elevatedToken: AccessToken | null = elevated.ok ? elevated.token : null;
      if (elevated.ok) {
        logger.info('auth.elevated.captured_at_login');
        lastElevatedOutcome = { captured: true };
      } else {
        logger.info('auth.elevated.skipped_at_login', { reason: elevated.reason });
        lastElevatedOutcome = { captured: false, reason: elevated.reason };
      }
      await persistTeams(result.accessToken, result.refreshToken, elevatedToken);
      if (chatsvcagg.ok) {
        logger.info('auth.chatsvcagg.captured_at_login', { region: chatsvcagg.region });
        lastChatsvcaggOutcome = { captured: true };
        await persistChatsvcagg(chatsvcagg.token, chatsvcagg.region);
      } else {
        logger.info('auth.chatsvcagg.skipped_at_login', { reason: chatsvcagg.reason });
        lastChatsvcaggOutcome = { captured: false, reason: chatsvcagg.reason };
      }
      if (ic3.ok) {
        logger.info('auth.ic3.captured_at_login', { region: ic3.region });
        await persistIc3(ic3.token, ic3.region);
      } else {
        logger.info('auth.ic3.skipped_at_login', { reason: ic3.reason });
      }
      if (force) await redeemMissedSubstrateAtLogin(chatsvcagg.ok, ic3.ok);
      logger.info('auth.ladder.rung', { rung: 'browser' });
      return ok(result.accessToken);
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return err({ type: 'auth_failed', message: msg });
    }
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
        const refreshed = await refreshToken(cached);
        if (refreshed.ok) return refreshed;
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
  const getGuestAccessToken = async (tenant: TenantId): Promise<Result<AccessToken, AuthError>> => {
    const cached = await readCache();
    const fresh = freshGuestToken(cached, tenant);
    if (fresh !== undefined) {
      logger.info('auth.guest.cache_hit', { tenant });
      return ok(accessTokenUnsafe(fresh));
    }
    if (!cached?.refresh_token) {
      return err({
        type: 'auth_failed',
        message: `no cached credentials to obtain a guest token for tenant ${tenant} — run \`ask-marcel-office login\``,
        code: SECONDARY_TOKEN_UNAVAILABLE_CODE,
      });
    }
    const redeemed = await redeemRefreshToken(cached.refresh_token, tenant, SCOPES);
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
    await writeCache({
      ...latest,
      guest_tokens: { ...latest.guest_tokens, [tenant]: { access_token: validated.value, expires_on: Math.floor(Date.now() / 1000) + redeemed.value.expiresIn } },
      // Entra single-uses and rotates the SPA refresh token. Dropping the rotated
      // one leaves a spent RT on disk and the next command dead-ends in a login.
      refresh_token: redeemed.value.refreshToken ?? latest.refresh_token,
    });
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
  const getCachedElevatedInfo = async (): Promise<CachedTierInfo> => {
    const cached = await readCache();
    const exp = cached?.elevated_expires_on;
    const expiresInSeconds = typeof exp === 'number' ? Math.floor(exp - Date.now() / 1000) : undefined;
    return { available: freshElevatedToken(cached) !== undefined, expiresInSeconds, scopes: decodeScopes(cached?.elevated_access_token) };
  };

  // Distinct error messages per failure mode (launch-hang, navigation
  // failure, silent-SSO timeout) so an LLM gets actionable remediation
  // rather than a one-size-fits-all message.
  const recoverableElevatedFailureMessage = (reason: ElevatedFailureReason): string => {
    const elevatedCommands = commandList(secondaryTokenCommands.elevated);
    if (reason === 'launch_timeout') {
      return `elevated browser launch timed out (15s) — likely a corrupt persistent profile or filesystem lock. Run \`ask-marcel-office logout && ask-marcel-office login\` to wipe the profile and retry. (Commands that need this token: ${elevatedCommands}.)`;
    }
    if (reason === 'navigation_failed') {
      return `elevated capture failed: navigation to m365.cloud.microsoft did not complete — network issue, corp-proxy block, or tenant policy. Check connectivity and retry. If persistent, the elevated commands (${elevatedCommands}) will be unavailable.`;
    }
    return `elevated token capture timed out — silent SSO against m365.cloud.microsoft did not yield a Bearer within 20s. The persistent browser-profile cookies are likely expired. Run \`ask-marcel-office logout && ask-marcel-office login\` — this now wipes the profile too. (Commands that need this token: ${elevatedCommands}.)`;
  };

  const recaptureElevated = async (): Promise<Result<AccessToken, AuthError>> => {
    try {
      const captured = await browserAuth.acquireElevatedToken();
      if (!captured.ok) {
        return err({ type: 'auth_failed', message: recoverableElevatedFailureMessage(captured.reason) });
      }
      await persistElevated(captured.token);
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
  const recaptureElevatedShared = (): Promise<Result<AccessToken, AuthError>> => {
    if (inFlightElevatedRecapture !== null) {
      logger.info('auth.elevated.shared_in_flight');
      return inFlightElevatedRecapture;
    }
    const launched = recaptureElevated();
    inFlightElevatedRecapture = launched.finally(() => {
      inFlightElevatedRecapture = null;
    });
    return inFlightElevatedRecapture;
  };

  const getElevatedAccessToken = async (): Promise<Result<AccessToken, AuthError>> => {
    const fresh = freshElevatedToken(await readCache());
    const validated = fresh !== undefined ? accessToken(fresh) : null;
    if (validated?.ok) {
      logger.info('auth.elevated.cache_hit');
      return ok(validated.value);
    }
    // Elevated absent, expired, or malformed; need to re-capture.
    if (!recaptureElevatedViaBrowser)
      return err({
        type: 'auth_failed',
        message: failFastSecondaryMessage('Elevated (M365)', commandList(secondaryTokenCommands.elevated), RECAPTURE_ELEVATED_VIA_LOGIN),
        code: SECONDARY_TOKEN_UNAVAILABLE_CODE,
      });
    // The persistent profile cookies do the silent SSO, no UI prompt.
    return recaptureElevatedShared();
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

  const recoverableChatsvcaggFailureMessage = (reason: ElevatedFailureReason): string => {
    const chatsvcaggCommands = commandList(secondaryTokenCommands.chatsvcagg);
    if (reason === 'launch_timeout') {
      return `chatsvcagg browser launch timed out (15s) — likely a corrupt persistent profile or filesystem lock. Run \`ask-marcel-office logout && ask-marcel-office login\` to wipe the profile and retry. (Commands that need this token: ${chatsvcaggCommands}.)`;
    }
    if (reason === 'navigation_failed') {
      return 'chatsvcagg capture failed: navigation to teams.microsoft.com did not complete — network issue, corp-proxy block, or tenant policy. Check connectivity and retry. If persistent, the Teams chat-content commands will be unavailable.';
    }
    return `chatsvcagg token capture timed out — silent SSO against teams.microsoft.com did not yield a Bearer within 20s. The persistent browser-profile cookies are likely expired. Run \`ask-marcel-office logout && ask-marcel-office login\` — this now wipes the profile too. (Commands that need this token: ${chatsvcaggCommands}.)`;
  };

  const recaptureChatsvcagg = async (): Promise<Result<AccessToken, AuthError>> => {
    try {
      const captured = await browserAuth.acquireChatsvcaggToken();
      if (!captured.ok) {
        return err({ type: 'auth_failed', message: recoverableChatsvcaggFailureMessage(captured.reason) });
      }
      await persistChatsvcagg(captured.token, captured.region);
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

  const getChatsvcaggAccessToken = async (): Promise<Result<AccessToken, AuthError>> => {
    // chatsvcagg tokens carry `aud=https://chatsvcagg.teams.microsoft.com`,
    // not Graph — the `accessToken()` validator's `isGraphToken` check
    // would reject every cached chatsvcagg token and force a recapture on
    // every call. `freshChatsvcaggToken` already validates expiry from the
    // JWT payload, which is the only thing we need at this boundary.
    const cached = await readCache();
    const fresh = freshChatsvcaggToken(cached);
    if (fresh !== undefined && fresh.startsWith('eyJ')) {
      logger.info('auth.chatsvcagg.cache_hit');
      return ok(accessTokenUnsafe(fresh));
    }
    if (!recaptureSecondaryViaBrowser) {
      // Command path: self-heal via a headless refresh of the shared RT before
      // giving up. Only fail fast when there's no RT or the refresh is rejected.
      if (cached?.refresh_token) {
        const refreshed = await refreshSubstrateToken(cached, CHATSVCAGG_RESOURCE, persistChatsvcagg, 'auth.chatsvcagg.refresh');
        if (refreshed.ok) return refreshed;
      }
      return err({
        type: 'auth_failed',
        message: failFastSecondaryMessage('chatsvcagg (Teams chat)', commandList(secondaryTokenCommands.chatsvcagg), RECAPTURE_VIA_LOGIN),
        code: SECONDARY_TOKEN_UNAVAILABLE_CODE,
      });
    }
    return recaptureChatsvcaggShared();
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

  const recoverableIc3FailureMessage = (reason: ElevatedFailureReason): string => {
    const ic3Commands = commandList(secondaryTokenCommands.ic3);
    if (reason === 'launch_timeout') {
      return `ic3 browser launch timed out (15s) — likely a corrupt persistent profile or filesystem lock. Run \`ask-marcel-office logout && ask-marcel-office login\` to wipe the profile and retry. (Commands that need this token: ${ic3Commands}.)`;
    }
    if (reason === 'navigation_failed') {
      return 'ic3 capture failed: navigation to teams.microsoft.com did not complete — network issue, corp-proxy block, or tenant policy. Check connectivity and retry. If persistent, the chat-history command will be unavailable.';
    }
    return `ic3 token capture timed out — silent SSO against teams.microsoft.com did not yield a Bearer within 20s. The persistent browser-profile cookies are likely expired. Run \`ask-marcel-office logout && ask-marcel-office login\` — this now wipes the profile too. (Commands that need this token: ${ic3Commands}.)`;
  };

  const recaptureIc3 = async (): Promise<Result<AccessToken, AuthError>> => {
    try {
      const captured = await browserAuth.acquireIc3Token();
      if (!captured.ok) {
        return err({ type: 'auth_failed', message: recoverableIc3FailureMessage(captured.reason) });
      }
      await persistIc3(captured.token, captured.region);
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

  const getIc3AccessToken = async (): Promise<Result<AccessToken, AuthError>> => {
    // IC3 tokens carry `aud=https://ic3.teams.office.com`, not Graph — the
    // `accessToken()` validator's `isGraphToken` check would reject every
    // cached IC3 token and force a recapture on every call. `freshIc3Token`
    // validates expiry from the JWT payload, which is the only thing we
    // need at this boundary.
    const cached = await readCache();
    const fresh = freshIc3Token(cached);
    if (fresh !== undefined && fresh.startsWith('eyJ')) {
      logger.info('auth.ic3.cache_hit');
      return ok(accessTokenUnsafe(fresh));
    }
    if (!recaptureSecondaryViaBrowser) {
      if (cached?.refresh_token) {
        const refreshed = await refreshSubstrateToken(cached, IC3_RESOURCE, persistIc3, 'auth.ic3.refresh');
        if (refreshed.ok) return refreshed;
      }
      return err({
        type: 'auth_failed',
        message: failFastSecondaryMessage('ic3 (Teams chat history)', commandList(secondaryTokenCommands.ic3), RECAPTURE_VIA_LOGIN),
        code: SECONDARY_TOKEN_UNAVAILABLE_CODE,
      });
    }
    return recaptureIc3Shared();
  };

  const logout = async (): Promise<Result<void, AuthError>> => {
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

  const getLastElevatedOutcome = (): ElevatedOutcome | null => lastElevatedOutcome;
  const getLastChatsvcaggOutcome = (): ElevatedOutcome | null => lastChatsvcaggOutcome;

  return {
    getAccessToken,
    getElevatedAccessToken,
    getGuestAccessToken,
    getChatsvcaggAccessToken,
    getChatsvcaggRegion,
    getIc3AccessToken,
    logout,
    getLastElevatedOutcome,
    getLastChatsvcaggOutcome,
    getCachedElevatedInfo,
    getCachedChatsvcaggInfo,
    getCachedIc3Info,
  };
};

const defaultFileSystem = (): FileSystem => (typeof globalThis.Bun !== 'undefined' ? createBunFileSystem() : createNodeFileSystem());

// Matches the convention in `browser-auth.ts:defaultProfileDir`. Kept in
// sync so `logout` wipes the same directory that `acquireElevatedToken`
// reads/writes.
const defaultBrowserProfileDir = (): string => {
  const envOverride = process.env['ASKMARCEL_BROWSER_PROFILE'];
  if (envOverride) return envOverride;
  const base = process.env['USERPROFILE'] ?? process.env['HOME'] ?? '';
  return `${base}/.ask-marcel/browser-profile`;
};

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
  fs?: FileSystem;
  browserProfileDir?: string;
  recaptureSecondaryViaBrowser?: boolean;
  secondaryTokenCommands?: SecondaryTokenCommands;
  acquireBasicViaBrowser?: boolean;
  recaptureElevatedViaBrowser?: boolean;
}): AuthManager => {
  const fs = deps.fs ?? defaultFileSystem();
  const browserProfileDir = deps.browserProfileDir ?? defaultBrowserProfileDir();
  const browserAuth = createBrowserAuth({
    logger: deps.logger,
    fs,
    freshCachedToken: createFreshCachedTokenProbe(fs, deps.cachePath),
    onProgress: stderrProgress,
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
export type { AuthError, AuthManager, CachedTierInfo, ElevatedOutcome, SecondaryTokenCommands };
