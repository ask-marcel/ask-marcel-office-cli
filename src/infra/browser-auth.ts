import { join } from 'node:path';
import type { AccessToken } from '../domain/access-token.ts';
import { accessToken, accessTokenUnsafe } from '../domain/access-token.ts';
import { decodeJwtPayload, isTokenFresh } from '../domain/jwt-utils.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import { createBunFileSystem } from './filesystem-bun.ts';
import { createNodeFileSystem } from './filesystem-node.ts';
import { loadPlaywright } from './playwright-loader.ts';

type BrowserTokenResult = { accessToken: AccessToken; refreshToken: string | null };

/**
 * Discriminated outcome of an elevated-token capture attempt. Distinct
 * failure variants let the caller pick the right error message AND let
 * the auto-heal in auth.ts decide whether a retry is worthwhile (a
 * recoverable failure like `launch_timeout` or `sso_timeout` is worth
 * one retry after wiping the profile; `navigation_failed` is a network
 * issue, not worth retrying).
 *
 * was previously `AccessToken | null`, which conflated
 * "browser launch hung", "navigation broke", and "silent-SSO polling
 * timed out" into a single null and made the error message inaccurate.
 */
type ElevatedFailureReason = 'launch_timeout' | 'navigation_failed' | 'sso_timeout';
type ElevatedTokenResult = { readonly ok: true; readonly token: AccessToken } | { readonly ok: false; readonly reason: ElevatedFailureReason };

/**
 * chatsvcagg capture result. Carries the same three failure modes as
 * `ElevatedTokenResult` PLUS the parsed regional segment from the URL
 * the bearer rode on (`emea` / `amer` / `apac` / etc.) — the new
 * substrate host (`teams.microsoft.com/api/csa/<region>/api/...`) routes
 * per region, so callers must persist this alongside the token to
 * construct future URLs. See `gotcha_chatsvcagg_substrate_moved` in
 * the project memory for the 2026-05 migration that made this
 * necessary.
 */
type ChatsvcaggTokenResult = { readonly ok: true; readonly token: AccessToken; readonly region: string } | { readonly ok: false; readonly reason: ElevatedFailureReason };

/**
 * IC3 capture result. The "next-gen" Teams chat-message substrate at
 * `teams.microsoft.com/api/chatsvc/<region>/v1/users/ME/conversations/{id}/messages`
 * — the path Teams web actually uses for chat scrollback (the existing
 * chatsvcagg substrate at `/api/csa/<region>/api/v1/chats/{id}/messages`
 * is the chat-list aggregator and caps at the 200 most recent messages
 * with no working pagination cursor). The IC3 path supports `syncState`
 * + `startTime` pagination, unlocking arbitrary-depth chat history.
 *
 * Bearer audience: `https://ic3.teams.office.com`. Standard `Bearer`
 * authorization (NOT the purple-teams-documented `skype_token X`
 * header — those forks predate Teams' move to OAuth bearer on this
 * surface). Empirically discovered 2026-05-21 by widening the
 * Playwright capture listener.
 */
type Ic3TokenResult = { readonly ok: true; readonly token: AccessToken; readonly region: string } | { readonly ok: false; readonly reason: ElevatedFailureReason };

/**
 * Combined outcome of capturing all four tokens (Teams basic / M365
 * elevated / chatsvcagg substrate / IC3 substrate) inside one browser
 * session. introduced this for the basic+elevated
 * pair; chatsvcagg was added next; IC3 is the most recent leg, capturing
 * the bearer Teams web uses for unbounded chat-history reads.
 *
 * All four legs ride on the SAME `teams.microsoft.com` session, so
 * adding capture legs is essentially free — MSAL multiplexes the
 * audience-scoped bearers via silent acquisition.
 */
type BothTokensResult = {
  readonly teams: BrowserTokenResult | null;
  readonly elevated: ElevatedTokenResult;
  readonly chatsvcagg: ChatsvcaggTokenResult;
  readonly ic3: Ic3TokenResult;
  /**
   * Set when the Teams poll short-circuited because `freshCachedToken`
   * found a token written by a concurrent process. The caller
   * must NOT persist this result — the cache is already the source of
   * truth, and `refreshToken` is null here (persisting would clobber the
   * winner's rotated refresh token).
   */
  readonly fromCache?: true;
};

type BrowserAuth = {
  /**
   * Capture an "elevated" Graph access token by navigating to a
   * different Microsoft web app whose first-party app identity is on
   * Microsoft's allow-list for ODSP `logicalPermissions` (the scope our
   * Teams web client token lacks for historical-version stream
   * content). Launches a HEADED browser (`launchContext(false)`) and
   * relies, best-effort, on the persistent profile's cookies to complete
   * sign-in silently. This is NOT guaranteed: on tenants that do not
   * persist SSO cookies (or once they lapse) the navigation lands on a
   * sign-in wall and the capture times out (verified headed AND headless
   * 2026-07-19). That is why it stays headed and why the caller treats a
   * miss as non-fatal and points the user at an interactive `login`.
   *
   * Returns a discriminated union so the caller can distinguish the three
   * failure modes (browser-launch hang, navigation failure, silent-SSO
   * polling timeout). The caller decides whether the failure is fatal
   * (it isn't for most commands; only the version + chat commands need
   * this).
   *
   * Standalone fallback path: used by `getElevatedAccessToken` when the
   * cached elevated token expires and the Teams session is already in
   * the cache (no fresh sign-in needed). On `login`, use
   * `acquireBothTokens` instead so the user sees only one browser.
   *
   * `awaitSignIn` says a human asked for this and is watching, so the capture
   * gets the interactive `pollDeadlineMs` rather than the short silent cap:
   * long enough for a slow redirect chain, and long enough to fill in a
   * sign-in form if one appears. Without it the short cap applies and the
   * capture fails fast, which is what a background command needs.
   */
  acquireElevatedToken: (options?: { readonly awaitSignIn?: boolean }) => Promise<ElevatedTokenResult>;
  /**
   * Capture a chatsvcagg-audience bearer (Teams basic identity, but the
   * `chatsvcagg.teams.microsoft.com` resource instead of Graph) by
   * navigating headless to `teams.microsoft.com/v2/`. The persistent
   * profile's SSO cookies do the auth silently. Used by
   * `getChatsvcaggAccessToken()` when the cached chatsvcagg token has
   * expired and the basic Teams session is still warm.
   *
   * Same failure-mode shape as `acquireElevatedToken` — launch hang,
   * navigation failure, silent SSO timeout — so the auth-manager can
   * reuse the same auto-heal logic.
   */
  acquireChatsvcaggToken: () => Promise<ChatsvcaggTokenResult>;
  /**
   * Capture an IC3-audience bearer (aud `https://ic3.teams.office.com`,
   * same basic Teams appid). The bearer Teams web client uses to call
   * `teams.microsoft.com/api/chatsvc/<region>/v1/users/ME/conversations/{id}/messages`
   * — the chat-message substrate with proper `syncState` pagination,
   * enabling reads beyond the 200-message chatsvcagg cap. Same failure
   * modes as `acquireChatsvcaggToken`.
   */
  acquireIc3Token: () => Promise<Ic3TokenResult>;
  /**
   * capture BOTH tokens inside ONE browser session.
   * After the Teams response listener intercepts the Teams token,
   * navigate the SAME page to the elevated URL and harvest the
   * M365ChatClient bearer from outgoing request headers. Cookies are
   * live in memory, so federated SSO chains (e.g. third-party-IdP-fronted
   * tenants) work without a second visible sign-in.
   *
   * Returns `teams: null` if no Teams token came back within the full
   * `pollDeadlineMs` (5 min). Returns
   * `elevated: { ok: false, reason: ... }` if the elevated capture
   * failed inside the same session — caller decides whether to surface
   * the partial success.
   */
  acquireBothTokens: (teamsUrl: string, options?: { skipCacheProbe?: boolean }) => Promise<BothTokensResult>;
  close: () => Promise<void>;
};

type ResponseLike = {
  url(): string;
  headers(): Record<string, string>;
  text(): Promise<string>;
};

type ResponseHandler = (response: ResponseLike) => void;

type PageLike = {
  on(event: 'response', handler: ResponseHandler): void;
  on(event: 'request', handler: RequestHandler): void;
  goto(url: string, options: { waitUntil: 'domcontentloaded'; timeout: number }): Promise<unknown>;
  url(): string;
  evaluate(fn: () => void): Promise<unknown>;
  close(): Promise<void>;
};

type ContextLike = {
  newPage(): Promise<PageLike>;
  clearCookies(): Promise<void>;
  close(): Promise<void>;
};

type LaunchOptions = {
  headless: boolean;
  channel?: 'msedge' | 'chrome';
  args: string[];
};

type RequestLike = {
  url(): string;
  headers(): Record<string, string>;
};

type RequestHandler = (request: RequestLike) => void;

type BrowserAuthApi = {
  launchPersistentContext(profileDir: string, options: LaunchOptions): Promise<ContextLike>;
};

type ChromiumLike = {
  launchPersistentContext(profileDir: string, options: LaunchOptions): Promise<ContextLike>;
};

type PlaywrightLoader = () => Promise<{ readonly chromium: ChromiumLike }>;

type TraceFn = (message: string) => void;

type BrowserAuthConfig = {
  readonly logger: Logger;
  readonly fs: FileSystem;
  readonly trace?: TraceFn;
  /**
   * (login hang): probe for a fresh token that landed in the cache
   * WHILE the browser capture is polling. AAD SPA refresh tokens rotate, so
   * when two processes race, the loser's refresh fails and it falls into the
   * full multi-minute browser dance — while the winner's fresh token sits in
   * the cache the whole time. The Teams poll loop calls this every interval
   * and short-circuits (closing the browser) the moment it returns a token.
   */
  readonly freshCachedToken?: () => Promise<string | null>;
  /**
   * User-visible progress sink (stderr in production). The browser capture
   * can legitimately take minutes (federated SSO, interactive sign-in); these
   * one-liners are what separates "waiting on the user" from "hung".
   */
  readonly onProgress?: (line: string) => void;
  readonly profileDir?: string;
  readonly initialSettleMs?: number;
  readonly postReloginSettleMs?: number;
  readonly pollIntervalMs?: number;
  /**
   * How long to keep listening for the PREFERRED elevated identity after an
   * acceptable-but-weaker one has already been captured. See
   * `PREFERRED_ELEVATED_APP_ID`.
   */
  readonly elevatedPreferenceGraceMs?: number;
  readonly pollDeadlineMs?: number;
  readonly navigationTimeoutMs?: number;
  /**
   * Deadline for the SILENT elevated-token recapture flow (no user
   * interaction expected — persistent profile cookies do the SSO).
   * Defaults to 20s. The audit () flagged that reusing the
   * 5-minute interactive `pollDeadlineMs` for this silent path made
   * `list-chats` etc. hang for minutes when cookies were stale, blowing
   * the LLM tool-call window. With a tight cap, the flow either yields
   * a token quickly or fails with `auth_failed: elevated token capture
   * timed out — run `ask-marcel-office login` to refresh.`
   *
   * The cap applies only when nobody is waiting. A caller passing
   * `awaitSignIn` (today just `login`) gets `pollDeadlineMs` instead, because
   * a silent capture that merely runs SLOW would otherwise fail here and send
   * `login` into the forced dance, whose cookie wipe costs the 90-day sign-in
   * session and guarantees the next run prompts.
   */
  readonly elevatedRecaptureTimeoutMs?: number;
  /**
   * Hard deadline on `launchPersistentContext` + `newPage` for the
   * elevated capture path. Defaults to 15s. Distinct from
   * `elevatedRecaptureTimeoutMs` so the error message can name which
   * step hung — launch vs polling. Audit previously
   * unguarded, so a hung Playwright launch (corrupt persistent profile
   * with stale `Singleton*` locks, or a slow browser binary) would
   * block the whole command indefinitely.
   */
  readonly elevatedLaunchTimeoutMs?: number;
};

const TOKEN_HOSTS = ['login.microsoftonline.com', 'login.live.com', 'login.microsoft.com'];

/**
 * App identities first-party Microsoft web apps use against Graph that
 * (verified 2026-05) are enrolled in ODSP `logicalPermissions` for
 * historical-version stream content — i.e., Bearer tokens issued under
 * these appids fetch /drives/{}/items/{}/versions/{ver}/content
 * successfully where the Teams web client (5e3ce6c0-...) hits a 403.
 *
 * Listed in priority order — the first match found while waiting for
 * the elevated capture wins.
 */
const ELEVATED_APP_IDS: ReadonlyArray<string> = [
  'c0ab8ce9-e9a0-42e7-b064-33d422df41f1', // M365ChatClient
  '4765445b-32c6-49b0-83e6-1d93765276ca', // OfficeHome (Graph audience also has the right scopes)
];

/**
 * Both entries of `ELEVATED_APP_IDS` emit a Graph-audience bearer on the SAME
 * m365.cloud.microsoft page load, 0.1-0.6s apart, in a random order — two live
 * probes on 2026-08-30 came back reversed. They are not interchangeable:
 * M365ChatClient carries `Chat.ReadBasic`, `Sites.ReadWrite.All` and
 * `Group.Read.All`; OfficeHome carries none of them and adds `Contacts.Read`
 * instead. Taking whichever bearer arrived first therefore made the elevated
 * tier's scope set a coin flip, and `list-chats` / `get-chat` failed with a
 * scope 403 on roughly half of all captures.
 *
 * So the capture HOLDS a weaker match and keeps listening for this one, giving
 * up the wait after `elevatedPreferenceGraceMs`. OfficeHome remains a genuine
 * fallback: a tenant where M365ChatClient never emits still gets a token, which
 * is what it got before.
 */
const PREFERRED_ELEVATED_APP_ID = 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1';

type ElevatedPicker = {
  readonly offer: (appid: string, token: AccessToken) => void;
  /** The token to accept NOW: the preferred one, or a weaker one whose grace has run out. */
  readonly settled: () => AccessToken | null;
  /** The best token held, whether or not its grace has elapsed. For deadline expiry. */
  readonly held: () => AccessToken | null;
};

const createElevatedPicker = (graceMs: number): ElevatedPicker => {
  let preferred: AccessToken | null = null;
  let fallback: AccessToken | null = null;
  let fallbackAt = 0;
  return {
    offer: (appid, token) => {
      if (appid === PREFERRED_ELEVATED_APP_ID) {
        preferred ??= token;
        return;
      }
      if (fallback === null) {
        fallback = token;
        fallbackAt = Date.now();
      }
    },
    settled: () => preferred ?? (fallback !== null && Date.now() - fallbackAt >= graceMs ? fallback : null),
    held: () => preferred ?? fallback,
  };
};

const M365_CLOUD_URL = 'https://m365.cloud.microsoft/search';

/**
 * The Microsoft Teams web client identity. The Phase A probe at
 * `scripts/probe-elevated-scopes.ts` confirmed empirically that the same
 * appid that mints the basic Graph token ALSO mints separate-audience
 * tokens for chatsvcagg, IC3, the Skype substrate, and Outlook — the
 * Teams web app multiplexes them via MSAL's `acquireTokenSilent` with
 * different `scopes` arrays. We piggy-back on the chatsvcagg one.
 */
const BASIC_TEAMS_APP_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';

/**
 * Substrate audience the chatsvcagg capture is gated by. Adding to this
 * list would extend the capture to additional audiences from the same
 * identity (IC3 messaging, Skype, Outlook). Only chatsvcagg is needed for
 * Phase C; the others are Phase D candidates.
 */
const CHATSVCAGG_AUD = 'https://chatsvcagg.teams.microsoft.com';

/**
 * Bearer audience Teams web uses for the IC3 (next-gen messaging)
 * substrate at `teams.microsoft.com/api/chatsvc/<region>/v1/...`. The
 * appid is the same basic Teams web client identity as the chatsvcagg
 * bearer; only the audience differs. Empirically confirmed 2026-05-21
 * via Playwright bearer-trace.
 */
const IC3_AUD = 'https://ic3.teams.office.com';

/**
 * Fallback region used when no `/api/csa/<region>/` URL was observed
 * during a chatsvcagg capture (e.g. the bearer rode on a request to a
 * different host or a path that pre-dates the 2026-05 substrate move).
 * `emea` is the only region we've empirically tested against — picking
 * it as the default means a partial capture (token but no region URL)
 * still produces a usable auth state for European tenants; AMER/APAC
 * tenants on the partial path will get a clear `HTTP 404 …` error from
 * the use-case rather than a silent miscall, which is acceptable given
 * how rare the partial-capture case is.
 */
const DEFAULT_CHATSVCAGG_REGION = 'emea';

/**
 * Pulls the regional segment from URLs of the shape
 * `https://teams.microsoft.com/api/csa/<region>/api/...`. Returns null
 * if the URL doesn't match (e.g. token bearer rode on a host outside
 * the substrate, which happens during early sign-in handshakes).
 */
const parseChatsvcaggRegion = (url: string): string | null => {
  // Matches both substrate path prefixes — chatsvcagg uses `/api/csa/<region>/`
  // and IC3 uses `/api/chatsvc/<region>/`. Both ride on `teams.microsoft.com`
  // and emit the same regional segment per tenant, so a single match unlocks
  // region resolution for both.
  const match = /^https:\/\/teams\.microsoft\.com\/api\/(?:csa|chatsvc)\/([a-z0-9-]+)\//i.exec(url);
  return match ? match[1].toLowerCase() : null;
};

/**
 * Standalone re-capture target for the chatsvcagg token — the Teams web
 * client at /v2/ emits the chatsvcagg bearer on its initial chat-list
 * load, which fires automatically on page settle when the persistent
 * profile's cookies are warm. Used by `acquireChatsvcaggToken()` when
 * the cached token has expired but the basic Teams session is still
 * valid.
 */
const TEAMS_WEB_URL = 'https://teams.microsoft.com/v2/';

const PROXY_ENV_KEYS = ['HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy'];

const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

/**
 * Best-effort dual-close: close the page then the context, swallowing
 * either failure. Used at the end of the chatsvcagg paths (token capture
 * + URL tracing) which share the same launch shape — Sonar's
 * `no-identical-functions` flagged the two prior inline copies as
 * duplicates, so this is the canonical version.
 */
const closeBrowserSession = async (page: PageLike, context: ContextLike): Promise<void> => {
  try {
    await page.close();
  } catch {
    // ignore
  }
  try {
    await context.close();
  } catch {
    // ignore
  }
};

const defaultProfileDir = (): string => {
  const envOverride = process.env.ASKMARCEL_BROWSER_PROFILE;
  if (envOverride) return envOverride;
  const base = process.env.USERPROFILE ?? process.env.HOME ?? '';
  return join(base, '.ask-marcel', 'browser-profile');
};

const cleanupSingletonLocks = async (dir: string, fs: FileSystem): Promise<void> => {
  for (const name of ['SingletonLock', 'SingletonCookie', 'SingletonSocket']) {
    await fs.deleteIfExists(join(dir, name));
  }
};

const stripProxyEnv = (): void => {
  for (const k of PROXY_ENV_KEYS) delete process.env[k];
};

const createPlaywrightApi = (loader: PlaywrightLoader): BrowserAuthApi => ({
  launchPersistentContext: async (profileDir, options) => {
    stripProxyEnv();
    const { chromium } = await loader();
    return chromium.launchPersistentContext(profileDir, options);
  },
});

const createBrowserAuthFromApi = (api: BrowserAuthApi, config: BrowserAuthConfig): BrowserAuth => {
  const { logger, fs } = config;
  // Default to no-op so the CLI is silent on success. Pass an explicit
  // trace (e.g., `(m) => process.stderr.write(m)`) for verbose debugging.
  const trace: TraceFn = config.trace ?? (() => {});
  const profileDir = config.profileDir ?? defaultProfileDir();
  const initialSettleMs = config.initialSettleMs ?? 5000;
  const postReloginSettleMs = config.postReloginSettleMs ?? 3000;
  const pollIntervalMs = config.pollIntervalMs ?? 2000;
  const pollDeadlineMs = config.pollDeadlineMs ?? 5 * 60 * 1000;
  const navigationTimeoutMs = config.navigationTimeoutMs ?? 30_000;
  const elevatedRecaptureTimeoutMs = config.elevatedRecaptureTimeoutMs ?? 20_000;
  const elevatedPreferenceGraceMs = config.elevatedPreferenceGraceMs ?? 3000;
  const elevatedLaunchTimeoutMs = config.elevatedLaunchTimeoutMs ?? 15_000;
  const freshCachedToken = config.freshCachedToken ?? (async () => null);
  const onProgress = config.onProgress ?? (() => {});

  let context: ContextLike | null = null;
  let page: PageLike | null = null;

  const cleanup = async (): Promise<void> => {
    if (page) {
      try {
        await page.close();
      } catch {
        // ignore
      }
      page = null;
    }
    if (context) {
      try {
        await context.close();
      } catch {
        // ignore
      }
      context = null;
    }
  };

  const launchContext = async (headless: boolean): Promise<ContextLike> => {
    for (const channel of ['msedge', 'chrome'] as const) {
      try {
        const ctx = await api.launchPersistentContext(profileDir, {
          headless,
          channel,
          args: ['--disable-blink-features=AutomationControlled'],
        });
        logger.info('browser_launched', { channel, headless });
        trace(`[DEBUG] browser launched with channel: ${channel} (headless=${headless})\n`);
        return ctx;
      } catch {
        logger.info('browser_launch_failed', { channel });
        trace(`[DEBUG] browser launch failed for channel: ${channel}\n`);
      }
    }
    const ctx = await api.launchPersistentContext(profileDir, {
      headless,
      args: ['--disable-blink-features=AutomationControlled'],
    });
    logger.info('browser_launched', { channel: 'bundled', headless });
    trace(`[DEBUG] browser launched with channel: bundled (headless=${headless})\n`);
    return ctx;
  };

  /**
   * Capture an elevated Graph token by navigating to m365.cloud.microsoft
   * (or another candidate Microsoft web app). The persistent profile's
   * SSO cookies do the auth silently — no UI prompt — provided the user
   * has previously completed the Teams login.
   *
   * Strategy: intercept outgoing `Authorization: Bearer ...` headers,
   * decode each JWT, return the first one whose `appid` is on the
   * ODSP-elevated allow-list.
   */
  /**
   * Race a promise against a hard-deadline reject. Used to wrap
   * Playwright's `launchPersistentContext` and `newPage` which can
   * hang indefinitely on profile corruption / filesystem locks.
   * The shared sentinel error is detected by its marker message so the
   * outer handler can attribute the failure to `launch_timeout`
   * specifically (rather than the generic catch-all). Using a plain
   * Error with a literal-string message rather than a Symbol because
   * `@typescript-eslint/prefer-promise-reject-errors` (strict-lint)
   * requires rejection reasons to be Errors, and atelier rule 10
   * forbids custom Error subclasses.
   */
  const ELEVATED_LAUNCH_TIMEOUT_MESSAGE = 'ask-marcel-office:elevated_launch_timeout';
  const isLaunchTimeout = (e: unknown): boolean => e instanceof Error && e.message === ELEVATED_LAUNCH_TIMEOUT_MESSAGE;
  const withLaunchTimeout = async <T>(p: Promise<T>, ms: number): Promise<T> => {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<T>((_, reject) => {
      timer = setTimeout(() => reject(new Error(ELEVATED_LAUNCH_TIMEOUT_MESSAGE)), ms);
    });
    try {
      return await Promise.race([p, timeout]);
    } finally {
      if (timer !== undefined) clearTimeout(timer);
    }
  };

  const acquireElevatedToken = async (options?: { readonly awaitSignIn?: boolean }): Promise<ElevatedTokenResult> => {
    await cleanupSingletonLocks(profileDir, fs);

    const picker = createElevatedPicker(elevatedPreferenceGraceMs);

    // Headless flakiness with Microsoft anti-automation has been
    // observed on m365.cloud.microsoft — the SPA sometimes refuses to
    // run its OAuth dance in headless mode. Visible launch is the
    // safer default. The window opens and closes within seconds; user
    // sees a brief flash but no interaction is required.
    //
    // wrap launch + newPage in a hard timeout
    // so a hung Playwright doesn't block the whole command. Previously
    // unguarded — a corrupt profile with stale Singleton locks would
    // hang indefinitely.
    let elevatedCtx: ContextLike;
    let elevatedPage: PageLike;
    try {
      elevatedCtx = await withLaunchTimeout(launchContext(false), elevatedLaunchTimeoutMs);
    } catch (e) {
      if (isLaunchTimeout(e)) {
        trace(`[DEBUG] elevated capture: launchContext timed out after ${elevatedLaunchTimeoutMs}ms\n`);
        logger.info('elevated_token_launch_timeout', { ms: elevatedLaunchTimeoutMs });
        return { ok: false, reason: 'launch_timeout' };
      }
      throw e;
    }
    try {
      elevatedPage = await withLaunchTimeout(elevatedCtx.newPage(), elevatedLaunchTimeoutMs);
    } catch (e) {
      try {
        await elevatedCtx.close();
      } catch {
        // ignore
      }
      if (isLaunchTimeout(e)) {
        trace(`[DEBUG] elevated capture: newPage timed out after ${elevatedLaunchTimeoutMs}ms\n`);
        logger.info('elevated_token_launch_timeout', { ms: elevatedLaunchTimeoutMs, phase: 'newPage' });
        return { ok: false, reason: 'launch_timeout' };
      }
      throw e;
    }

    elevatedPage.on('request', (req) => {
      if (picker.settled()) return;
      const auth = req.headers()['authorization'];
      if (typeof auth !== 'string' || !auth.startsWith('Bearer ')) return;
      const raw = auth.slice('Bearer '.length);
      const claims = decodeJwtPayload(raw);
      const appid = typeof claims['appid'] === 'string' ? claims['appid'] : undefined;
      const aud = typeof claims['aud'] === 'string' ? claims['aud'] : undefined;
      if (!appid || !aud) return;
      if (!ELEVATED_APP_IDS.includes(appid)) return;
      if (aud !== 'https://graph.microsoft.com') return;
      const validated = accessToken(raw);
      if (!validated.ok) return;
      picker.offer(appid, validated.value);
      logger.info('elevated_token_captured', { appid, len: validated.value.length });
      trace(`[DEBUG] elevated token captured  appid=${appid}  len=${validated.value.length}\n`);
    });

    logger.info('browser_navigating', { url: M365_CLOUD_URL, purpose: 'elevated' });
    trace(`[DEBUG] elevated capture: navigating to ${M365_CLOUD_URL}\n`);
    let navigationFailed = false;
    try {
      await elevatedPage.goto(M365_CLOUD_URL, { waitUntil: 'domcontentloaded', timeout: navigationTimeoutMs });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      trace(`[DEBUG] elevated nav error: ${msg}\n`);
      navigationFailed = true;
    }

    // Settle: silent SSO capture happens within 3-8s when cookies are warm.
    // The elevated path is non-interactive — the popup is opened with the
    // persistent profile and either silent SSO completes (token captured) or
    // the cookies are stale and silent SSO can't proceed without UI prompts
    // we don't drive. Use a SHORT deadline (default 20s) and fail fast with
    // a clear discriminated failure so the caller can decide whether to
    // retry (after wiping the profile) or surface the error.
    // Two budgets, chosen by who is waiting. A background command keeps the
    // short cap and fails fast. A human who ran `login` gets the interactive
    // one, because on this tenant a SUCCESSFUL silent capture measured longer
    // than the 20s cap: it then failed, `login` escalated to the forced dance,
    // and that dance's cookie wipe destroyed the 90-day sign-in session, so the
    // next run had to prompt for real. Measured 2026-08-31, and it is the loop
    // behind "even nothing looks expired, I still get the login page".
    //
    // Deliberately NOT gated on the page sitting at an AAD sign-in URL: a
    // silent SSO redirects THROUGH those same hosts, so the URL says nothing
    // about whether a human is needed. Acting on it closed a capture that was
    // about to succeed.
    const shortDeadline = Date.now() + elevatedRecaptureTimeoutMs;
    const elevatedDeadline = options?.awaitSignIn === true ? Date.now() + pollDeadlineMs : shortDeadline;
    let announced = false;
    while (Date.now() < elevatedDeadline) {
      const settled = picker.settled();
      if (settled) {
        trace('[DEBUG] elevated capture: token found, closing\n');
        await closeBrowserSession(elevatedPage, elevatedCtx);
        return { ok: true, token: settled };
      }
      // Silence until the silent route has plainly not worked; past that the
      // window is probably showing a sign-in form and the user needs telling.
      if (options?.awaitSignIn === true && !announced && Date.now() >= shortDeadline) {
        announced = true;
        onProgress(
          `Elevated token is taking longer than usual — if a sign-in screen is showing, complete it in the browser window (waiting up to ${Math.round(pollDeadlineMs / 60000)} min)…`
        );
      }
      await sleep(pollIntervalMs);
    }

    // A weaker token that landed just before the deadline never got its grace
    // window; keep it rather than report a timeout on a capture that worked.
    const held = picker.held();
    if (held) {
      trace('[DEBUG] elevated capture: deadline reached holding a fallback, keeping it\n');
      await closeBrowserSession(elevatedPage, elevatedCtx);
      return { ok: true, token: held };
    }

    await closeBrowserSession(elevatedPage, elevatedCtx);
    if (navigationFailed) {
      trace('[DEBUG] elevated capture: deadline expired after navigation failure\n');
      logger.info('elevated_token_navigation_failed');
      return { ok: false, reason: 'navigation_failed' };
    }
    trace('[DEBUG] elevated capture: deadline expired, no elevated token captured\n');
    logger.info('elevated_token_capture_timeout');
    return { ok: false, reason: 'sso_timeout' };
  };

  /**
   * single-session capture of BOTH tokens.
   *
   * Opens ONE browser context, attaches both response (Teams token) and
   * request (elevated token) listeners up front, navigates the page
   * through the Teams sign-in, then — without closing — navigates the
   * same page to the elevated URL so the cookie chain that the user
   * just authenticated against is still live in memory. Closes the
   * browser exactly once at the end.
   *
   * Fixes the user-visible "loop of browser open/close" symptom on
   * federated tenants (ExampleCorp / a third-party IdP): the old two-session flow opened a
   * separate browser for the elevated step and silent SSO couldn't
   * re-establish the federated cookie chain quickly enough — the
   * second browser would flash through what looked like a fresh login
   * prompt. With this method, the elevated step runs on the same
   * already-authenticated page.
   */
  const acquireBothTokens = async (teamsUrl: string, options?: { skipCacheProbe?: boolean }): Promise<BothTokensResult> => {
    const elevatedUrl = M365_CLOUD_URL;
    trace('[DEBUG] acquireBothTokens: ENTER\n');
    await cleanupSingletonLocks(profileDir, fs);

    let capturedAccess: AccessToken | null = null;
    let capturedRefresh: string | null = null;
    const elevatedPicker = createElevatedPicker(elevatedPreferenceGraceMs);
    let capturedChatsvcagg: AccessToken | null = null;
    let capturedChatsvcaggRegion: string | null = null;
    let capturedIc3: AccessToken | null = null;

    let elevatedCtx: ContextLike;
    let elevatedPage: PageLike;
    try {
      elevatedCtx = await withLaunchTimeout(launchContext(false), elevatedLaunchTimeoutMs);
    } catch (e) {
      if (isLaunchTimeout(e)) {
        trace(`[DEBUG] acquireBothTokens: launch timed out after ${elevatedLaunchTimeoutMs}ms\n`);
        return {
          teams: null,
          elevated: { ok: false, reason: 'launch_timeout' },
          chatsvcagg: { ok: false, reason: 'launch_timeout' },
          ic3: { ok: false, reason: 'launch_timeout' },
        };
      }
      throw e;
    }
    try {
      elevatedPage = await withLaunchTimeout(elevatedCtx.newPage(), elevatedLaunchTimeoutMs);
    } catch (e) {
      try {
        await elevatedCtx.close();
      } catch {
        // ignore
      }
      if (isLaunchTimeout(e)) {
        trace(`[DEBUG] acquireBothTokens: newPage timed out after ${elevatedLaunchTimeoutMs}ms\n`);
        return {
          teams: null,
          elevated: { ok: false, reason: 'launch_timeout' },
          chatsvcagg: { ok: false, reason: 'launch_timeout' },
          ic3: { ok: false, reason: 'launch_timeout' },
        };
      }
      throw e;
    }
    context = elevatedCtx;
    page = elevatedPage;
    const activePage = elevatedPage;

    // Teams response listener — captures the Graph token grant response.
    const handleTeamsResponse = async (r: ResponseLike): Promise<void> => {
      if (capturedAccess) return;
      const url = r.url();
      if (!TOKEN_HOSTS.some((d) => url.includes(d))) return;
      const ct = r.headers()['content-type'] ?? '';
      if (!ct.includes('json')) return;
      try {
        const body = await r.text();
        if (!body.includes('access_token')) return;
        const data = JSON.parse(body) as { access_token?: string; refresh_token?: string };
        const raw = data.access_token ?? '';
        const validated = accessToken(raw);
        if (!validated.ok) {
          trace(`[DEBUG] acquireBothTokens: non-graph token skipped, len: ${raw.length}\n`);
          return;
        }
        capturedAccess = validated.value;
        capturedRefresh = data.refresh_token ?? null;
        logger.info('token_captured', { len: validated.value.length });
        trace(`[DEBUG] acquireBothTokens: graph token captured, len: ${validated.value.length}\n`);
      } catch {
        // ignore parse errors
      }
    };
    activePage.on('response', (r) => {
      void handleTeamsResponse(r);
    });

    // Unified request listener — captures three things from outgoing
    // Authorization bearers (multiplexed by MSAL on a single
    // teams.microsoft.com session):
    // 1. Elevated Graph bearer (ODSP-allowlisted appid + Graph aud)
    // 2. chatsvcagg-audience bearer (basic Teams appid + chatsvcagg aud)
    // 3. The substrate region from `/api/csa/<region>/` URLs the
    // chatsvcagg bearer rides on (post-2026-05 substrate move:
    // `chatsvcagg.teams.microsoft.com` was retired in favour of
    // `teams.microsoft.com/api/csa/<region>/api/...` per-region
    // routing — see `gotcha_chatsvcagg_substrate_moved` in memory).
    // Also emits a `chatsvcagg_request` trace per unique URL the bearer
    // rides on so the next substrate migration is self-diagnosable.
    const seenChatsvcaggRoutes = new Set<string>();
    activePage.on('request', (req) => {
      if (elevatedPicker.settled() && capturedChatsvcagg && capturedChatsvcaggRegion && capturedIc3) return;
      const auth = req.headers()['authorization'];
      if (typeof auth !== 'string' || !auth.startsWith('Bearer ')) return;
      const raw = auth.slice('Bearer '.length);
      const claims = decodeJwtPayload(raw);
      const appid = typeof claims['appid'] === 'string' ? claims['appid'] : undefined;
      const aud = typeof claims['aud'] === 'string' ? claims['aud'] : undefined;
      if (!appid || !aud) return;
      if (ELEVATED_APP_IDS.includes(appid) && aud === 'https://graph.microsoft.com') {
        const validated = accessToken(raw);
        if (validated.ok) {
          elevatedPicker.offer(appid, validated.value);
          logger.info('elevated_token_captured', { appid, len: validated.value.length });
          trace(`[DEBUG] acquireBothTokens: elevated token captured appid=${appid}\n`);
        }
      }
      if (appid === BASIC_TEAMS_APP_ID && aud === CHATSVCAGG_AUD) {
        // chatsvcagg tokens have aud=chatsvcagg, not Graph — can't go
        // through `accessToken()`'s Graph-audience validator. Validate
        // freshness directly and skip malformed payloads.
        if (!capturedChatsvcagg && raw.startsWith('eyJ') && isTokenFresh(raw)) {
          capturedChatsvcagg = accessTokenUnsafe(raw);
          logger.info('chatsvcagg_token_captured', { len: raw.length });
          trace(`[DEBUG] acquireBothTokens: chatsvcagg token captured len=${raw.length}\n`);
        }
        const url = req.url();
        if (!seenChatsvcaggRoutes.has(url)) {
          seenChatsvcaggRoutes.add(url);
          logger.info('chatsvcagg_request', { url });
          trace(`[DEBUG] chatsvcagg_request: ${url}\n`);
        }
        if (!capturedChatsvcaggRegion) {
          const region = parseChatsvcaggRegion(url);
          if (region !== null) {
            capturedChatsvcaggRegion = region;
            logger.info('chatsvcagg_region_captured', { region });
            trace(`[DEBUG] chatsvcagg region captured: ${region}\n`);
          }
        }
      }
      if (appid === BASIC_TEAMS_APP_ID && aud === IC3_AUD && !capturedIc3) {
        // IC3 bearer — same identity as chatsvcagg, different audience.
        // Used by Teams web for `/api/chatsvc/<region>/v1/...` calls which
        // include the paginated chat-history endpoint. Same JWT-validator
        // skip rationale as chatsvcagg (aud is not Graph).
        if (raw.startsWith('eyJ') && isTokenFresh(raw)) {
          capturedIc3 = accessTokenUnsafe(raw);
          logger.info('ic3_token_captured', { len: raw.length });
          trace(`[DEBUG] acquireBothTokens: ic3 token captured len=${raw.length}\n`);
        }
        // Region also resolvable from `/api/chatsvc/<region>/` URLs the IC3
        // bearer rides on — feeds the same `capturedChatsvcaggRegion` slot
        // (regions match per tenant; the parser handles both prefixes).
        if (!capturedChatsvcaggRegion) {
          const region = parseChatsvcaggRegion(req.url());
          if (region !== null) {
            capturedChatsvcaggRegion = region;
            logger.info('chatsvcagg_region_captured', { region });
            trace(`[DEBUG] region captured via ic3 url: ${region}\n`);
          }
        }
      }
    });

    // Navigate to teamsUrl and wait for sign-in to complete.
    trace(`[DEBUG] acquireBothTokens: navigating to ${teamsUrl}\n`);
    logger.info('browser_navigating', { url: teamsUrl });
    try {
      await activePage.goto(teamsUrl, { waitUntil: 'domcontentloaded', timeout: navigationTimeoutMs });
      logger.info('browser_navigated', { url: teamsUrl });
    } catch (navError) {
      const navMsg = navError instanceof Error ? navError.message : String(navError);
      trace(`[DEBUG] acquireBothTokens: navigation error (non-fatal): ${navMsg}\n`);
    }

    trace(`[DEBUG] acquireBothTokens: initial settle for ${initialSettleMs}ms\n`);
    await sleep(initialSettleMs);

    // Force re-login if the persistent profile dropped us into an
    // already-signed-in session — the old token grant happened before
    // our listener was attached, so we need to redo the OAuth dance.
    const currentUrl = activePage.url();
    trace(`[DEBUG] acquireBothTokens: after settle, currentUrl = ${currentUrl}\n`);
    if (!capturedAccess && !currentUrl.includes('login.microsoftonline.com') && !currentUrl.includes('login.live.com')) {
      trace('[DEBUG] acquireBothTokens: already signed in — clearing session to force fresh login\n');
      logger.info('browser_force_relogin', { url: currentUrl });
      // Said out loud because it is expensive and invisible. The wipe is needed
      // (the BASIC token is read off the grant RESPONSE, so the grant has to
      // re-fire with the listener attached), but it also takes
      // ESTSAUTHPERSISTENT, the ~90-day "Stay signed in" cookie, so the NEXT
      // capture has no session to SSO with and serves a real sign-in page. A
      // user who reached for `--force` on every prompt was re-arming the prompt
      // each time (2026-08-31). Emitted here rather than in the CLI flag
      // description so MCP and library callers hear it too, and so it can only
      // fire when a session is actually being destroyed.
      onProgress(
        'Clearing the browser session so the sign-in can be re-captured — this deletes the 90-day "Stay signed in" cookie, so you may be asked for your password. A plain `ask-marcel-office login` refreshes the elevated token without clearing it.'
      );
      await context.clearCookies();
      try {
        await activePage.evaluate(() => {
          localStorage.clear();
          sessionStorage.clear();
        });
      } catch {
        // ignore
      }
      try {
        await activePage.goto(teamsUrl, { waitUntil: 'domcontentloaded', timeout: navigationTimeoutMs });
      } catch {
        // non-fatal
      }
      await sleep(postReloginSettleMs);
    }

    // Poll until Teams token captured OR poll-deadline expires (5 min) OR a
    // concurrent process lands a fresh token in the cache.
    trace(`[DEBUG] acquireBothTokens: polling for Teams token, ${pollDeadlineMs / 1000}s deadline\n`);
    onProgress(`Browser window open — if a sign-in screen appears, complete it there (waiting up to ${Math.round(pollDeadlineMs / 60000)} min)…`);
    const teamsDeadline = Date.now() + pollDeadlineMs;
    let pollCount = 0;
    while (Date.now() < teamsDeadline && !capturedAccess) {
      // `login --force` (skipCacheProbe) wants a full re-capture, so the
      // concurrent-refresh short-circuit must NOT fire — otherwise the still-valid
      // cached basic token closes the browser before the elevated/ic3 legs run.
      const concurrent = options?.skipCacheProbe === true ? null : await freshCachedToken();
      if (concurrent !== null) {
        const validated = accessToken(concurrent);
        if (validated.ok) {
          trace('[DEBUG] acquireBothTokens: fresh cache token appeared mid-poll — short-circuiting\n');
          logger.info('browser_poll_cache_short_circuit');
          onProgress('Token was refreshed by another process — closing the browser.');
          await cleanup();
          return {
            teams: { accessToken: validated.value, refreshToken: null },
            fromCache: true,
            elevated: { ok: false, reason: 'sso_timeout' },
            chatsvcagg: { ok: false, reason: 'sso_timeout' },
            ic3: { ok: false, reason: 'sso_timeout' },
          };
        }
      }
      pollCount += 1;
      if (pollCount % 10 === 0) {
        trace(`[DEBUG] acquireBothTokens: still polling for Teams token, url=${activePage.url()}\n`);
      }
      await sleep(pollIntervalMs);
    }

    if (!capturedAccess) {
      trace('[DEBUG] acquireBothTokens: Teams token deadline expired\n');
      await cleanup();
      return { teams: null, elevated: { ok: false, reason: 'sso_timeout' }, chatsvcagg: { ok: false, reason: 'sso_timeout' }, ic3: { ok: false, reason: 'sso_timeout' } };
    }

    // Teams token captured. The chatsvcagg bearer fires from teams.microsoft.com
    // itself when the page makes its first chat-list call after sign-in — wait
    // a brief best-effort settle on the same page (uses the existing
    // postReloginSettleMs budget). If chatsvcagg hasn't fired yet, the silent
    // re-capture path picks it up on first command use. Also wait for the
    // substrate region: empirically the region URL fires a beat AFTER the
    // token mint, so a token-only exit would persist with the fallback
    // region — worse for AMER/APAC tenants.
    const chatsvcaggSettleDeadline = Date.now() + postReloginSettleMs;
    while (Date.now() < chatsvcaggSettleDeadline && !(capturedChatsvcagg && capturedChatsvcaggRegion && capturedIc3)) {
      await sleep(pollIntervalMs);
    }

    onProgress('Signed in — capturing companion tokens (Teams chat / elevated; up to ~30 s)…');
    trace(`[DEBUG] acquireBothTokens: Teams captured; navigating to elevated ${elevatedUrl}\n`);
    logger.info('browser_navigating', { url: elevatedUrl, purpose: 'elevated_same_session' });
    let navigationFailed = false;
    try {
      await activePage.goto(elevatedUrl, { waitUntil: 'domcontentloaded', timeout: navigationTimeoutMs });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      trace(`[DEBUG] acquireBothTokens: elevated navigation error: ${msg}\n`);
      navigationFailed = true;
    }

    const elevatedDeadline = Date.now() + elevatedRecaptureTimeoutMs;
    while (Date.now() < elevatedDeadline && !elevatedPicker.settled()) {
      await sleep(pollIntervalMs);
    }
    // `held`, not `settled`: a fallback captured inside the grace window at the
    // deadline is still a working token.
    const capturedElevated = elevatedPicker.held();

    await cleanup();
    onProgress('Sign-in complete — browser closed.');

    const teamsResult: BrowserTokenResult = { accessToken: capturedAccess, refreshToken: capturedRefresh };
    const elevatedFailureReason: ElevatedFailureReason = navigationFailed ? 'navigation_failed' : 'sso_timeout';
    const elevatedResult: ElevatedTokenResult = capturedElevated ? { ok: true, token: capturedElevated } : { ok: false, reason: elevatedFailureReason };
    // Region defaults to `DEFAULT_CHATSVCAGG_REGION` ('emea') when we captured a
    // bearer but never saw a substrate URL — a partial capture. Better than
    // dropping the token: European tenants (the dominant case during 2026-05
    // probing) keep working; other-region tenants get a clear HTTP 404 from
    // the use-case rather than a silent failure. The region is logged at
    // `chatsvcagg_region_captured` when it's parsed; absence in the trace is
    // the signal that the fallback was used.
    const chatsvcaggResult: ChatsvcaggTokenResult = capturedChatsvcagg
      ? { ok: true, token: capturedChatsvcagg, region: capturedChatsvcaggRegion ?? DEFAULT_CHATSVCAGG_REGION }
      : { ok: false, reason: 'sso_timeout' };
    const ic3Result: Ic3TokenResult = capturedIc3
      ? { ok: true, token: capturedIc3, region: capturedChatsvcaggRegion ?? DEFAULT_CHATSVCAGG_REGION }
      : { ok: false, reason: 'sso_timeout' };
    if (!capturedElevated) {
      trace('[DEBUG] acquireBothTokens: elevated deadline expired, Teams kept\n');
      logger.info('elevated_token_capture_timeout');
    }
    if (!capturedChatsvcagg) {
      trace('[DEBUG] acquireBothTokens: chatsvcagg settle expired, Teams kept\n');
      logger.info('chatsvcagg_token_capture_timeout');
    }
    if (!capturedIc3) {
      trace('[DEBUG] acquireBothTokens: ic3 settle expired, Teams kept\n');
      logger.info('ic3_token_capture_timeout');
    }
    return { teams: teamsResult, elevated: elevatedResult, chatsvcagg: chatsvcaggResult, ic3: ic3Result };
  };

  // Shared standalone-recapture path for chatsvcagg/IC3 — both ride on the
  // same `teams.microsoft.com/v2/` session and only differ in JWT audience.
  // `kind` parameterizes log/trace strings so a future capture-failure
  // diagnostic can tell the substrates apart without grep gymnastics.
  const acquireSubstrateToken = async (kind: 'chatsvcagg' | 'ic3', targetAud: string): Promise<ChatsvcaggTokenResult> => {
    trace(`[DEBUG] acquire${kind}Token: ENTER\n`);
    await cleanupSingletonLocks(profileDir, fs);

    let captured: AccessToken | null = null;
    let capturedRegion: string | null = null;

    let ctx: ContextLike;
    let p: PageLike;
    try {
      ctx = await withLaunchTimeout(launchContext(true), elevatedLaunchTimeoutMs);
    } catch (e) {
      if (isLaunchTimeout(e)) {
        trace(`[DEBUG] ${kind} capture: launchContext timed out after ${elevatedLaunchTimeoutMs}ms\n`);
        logger.info(`${kind}_token_launch_timeout`, { ms: elevatedLaunchTimeoutMs });
        return { ok: false, reason: 'launch_timeout' };
      }
      throw e;
    }
    try {
      p = await withLaunchTimeout(ctx.newPage(), elevatedLaunchTimeoutMs);
    } catch (e) {
      try {
        await ctx.close();
      } catch {
        // ignore
      }
      if (isLaunchTimeout(e)) {
        trace(`[DEBUG] ${kind} capture: newPage timed out after ${elevatedLaunchTimeoutMs}ms\n`);
        logger.info(`${kind}_token_launch_timeout`, { ms: elevatedLaunchTimeoutMs, phase: 'newPage' });
        return { ok: false, reason: 'launch_timeout' };
      }
      throw e;
    }

    p.on('request', (req) => {
      if (captured && capturedRegion) return;
      const auth = req.headers()['authorization'];
      if (typeof auth !== 'string' || !auth.startsWith('Bearer ')) return;
      const raw = auth.slice('Bearer '.length);
      const claims = decodeJwtPayload(raw);
      const appid = typeof claims['appid'] === 'string' ? claims['appid'] : undefined;
      const aud = typeof claims['aud'] === 'string' ? claims['aud'] : undefined;
      if (appid !== BASIC_TEAMS_APP_ID || aud !== targetAud) return;
      // Substrate-audience tokens can't use the Graph-audience validator
      // (audience is intentionally NOT graph.microsoft.com).
      if (!captured && raw.startsWith('eyJ') && isTokenFresh(raw)) {
        captured = accessTokenUnsafe(raw);
        logger.info(`${kind}_token_captured`, { len: raw.length });
        trace(`[DEBUG] ${kind} token captured len=${raw.length}\n`);
      }
      if (!capturedRegion) {
        const region = parseChatsvcaggRegion(req.url());
        if (region !== null) {
          capturedRegion = region;
          logger.info(`${kind}_region_captured`, { region });
          trace(`[DEBUG] ${kind} region captured: ${region}\n`);
        }
      }
    });

    logger.info('browser_navigating', { url: TEAMS_WEB_URL, purpose: kind });
    trace(`[DEBUG] ${kind} capture: navigating to ${TEAMS_WEB_URL}\n`);
    let navigationFailed = false;
    try {
      await p.goto(TEAMS_WEB_URL, { waitUntil: 'domcontentloaded', timeout: navigationTimeoutMs });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      trace(`[DEBUG] ${kind} nav error: ${msg}\n`);
      navigationFailed = true;
    }

    const deadline = Date.now() + elevatedRecaptureTimeoutMs;
    while (Date.now() < deadline) {
      // Region URL empirically fires a beat AFTER the token mint, so a
      // token-only early-exit would leave us with the fallback region.
      // Keep waiting until BOTH are captured, OR the deadline expires —
      // at which point we ship what we have (token + 'emea' fallback).
      if (captured && capturedRegion) {
        trace(`[DEBUG] ${kind} capture: token + region found, closing\n`);
        await closeBrowserSession(p, ctx);
        return { ok: true, token: captured, region: capturedRegion };
      }
      await sleep(pollIntervalMs);
    }
    if (captured) {
      trace(`[DEBUG] ${kind} capture: token captured but no region URL within budget; falling back to ${DEFAULT_CHATSVCAGG_REGION}\n`);
      await closeBrowserSession(p, ctx);
      return { ok: true, token: captured, region: DEFAULT_CHATSVCAGG_REGION };
    }

    await closeBrowserSession(p, ctx);
    if (navigationFailed) {
      logger.info(`${kind}_token_navigation_failed`);
      return { ok: false, reason: 'navigation_failed' };
    }
    trace(`[DEBUG] ${kind} capture: deadline expired\n`);
    logger.info(`${kind}_token_capture_timeout`);
    return { ok: false, reason: 'sso_timeout' };
  };

  const acquireChatsvcaggToken = (): Promise<ChatsvcaggTokenResult> => acquireSubstrateToken('chatsvcagg', CHATSVCAGG_AUD);
  const acquireIc3Token = (): Promise<Ic3TokenResult> => acquireSubstrateToken('ic3', IC3_AUD);

  const close = async (): Promise<void> => {
    logger.info('browser_auth_close');
    await cleanup();
  };

  return { acquireElevatedToken, acquireChatsvcaggToken, acquireIc3Token, acquireBothTokens, close };
};

const defaultFileSystem = (): FileSystem => (typeof globalThis.Bun !== 'undefined' ? createBunFileSystem() : createNodeFileSystem());

// diagnostic: when `ASKMARCEL_TRACE=1` is set in the
// environment, wire the trace function to stderr AND echo every `.info`
// logger event through it. (browser-auth only ever logs at info level —
// warn/error are pass-through.) Lets a user capture the full
// browser-auth flow (launches, navigations, captures, cleanups) for bug
// reports without editing code.
const enableTraceFromEnv = (logger: Logger): { logger: Logger; trace?: TraceFn } => {
  if (process.env['ASKMARCEL_TRACE'] !== '1') return { logger };
  const trace: TraceFn = (m) => process.stderr.write(m);
  const wrapped: Logger = {
    ...logger,
    info: (event, meta) => {
      logger.info(event, meta);
      trace(`[INFO] ${event} ${meta === undefined ? '' : JSON.stringify(meta)}\n`);
    },
  };
  return { logger: wrapped, trace };
};

const createBrowserAuth = (deps: { logger: Logger; fs?: FileSystem; freshCachedToken?: () => Promise<string | null>; onProgress?: (line: string) => void }): BrowserAuth => {
  const { logger, trace } = enableTraceFromEnv(deps.logger);
  return createBrowserAuthFromApi(createPlaywrightApi(loadPlaywright), {
    logger,
    fs: deps.fs ?? defaultFileSystem(),
    ...(trace ? { trace } : {}),
    ...(deps.freshCachedToken ? { freshCachedToken: deps.freshCachedToken } : {}),
    ...(deps.onProgress ? { onProgress: deps.onProgress } : {}),
  });
};

export { createBrowserAuth, createBrowserAuthFromApi, createPlaywrightApi };
export type {
  BothTokensResult,
  BrowserAuth,
  BrowserAuthApi,
  BrowserAuthConfig,
  BrowserTokenResult,
  ChatsvcaggTokenResult,
  Ic3TokenResult,
  ChromiumLike,
  ContextLike,
  ElevatedFailureReason,
  ElevatedTokenResult,
  PageLike,
  PlaywrightLoader,
  RequestLike,
  ResponseLike,
};
