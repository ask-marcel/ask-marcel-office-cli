import { join } from 'node:path';
import type { AccessToken } from '../domain/access-token.ts';
import { accessToken } from '../domain/access-token.ts';
import { decodeJwtPayload } from '../domain/jwt-utils.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import { createBunFileSystem } from './filesystem-bun.ts';
import { createNodeFileSystem } from './filesystem-node.ts';
import { loadPlaywright } from './playwright-loader.ts';

type BrowserTokenResult = { accessToken: AccessToken; refreshToken: string | null };

type BrowserAuth = {
  acquireToken: (scopes: string[], startUrl: string) => Promise<BrowserTokenResult | null>;
  /**
   * Capture an "elevated" Graph access token by navigating to a
   * different Microsoft web app whose first-party app identity is on
   * Microsoft's allow-list for ODSP `logicalPermissions` (the scope our
   * Teams web client token lacks for historical-version stream
   * content). Reuses the persistent profile cookies — no second
   * sign-in. Headless by default.
   *
   * Returns null on cancellation / timeout. The caller decides whether
   * the failure is fatal (it isn't for most commands; only the version
   * commands need this).
   */
  acquireElevatedToken: () => Promise<AccessToken | null>;
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
  readonly profileDir?: string;
  readonly initialSettleMs?: number;
  readonly postReloginSettleMs?: number;
  readonly pollIntervalMs?: number;
  readonly pollDeadlineMs?: number;
  readonly navigationTimeoutMs?: number;
  /**
   * Deadline for the SILENT elevated-token recapture flow (no user
   * interaction expected — persistent profile cookies do the SSO).
   * Defaults to 20s. The audit (v1.0.0 §1.1) flagged that reusing the
   * 5-minute interactive `pollDeadlineMs` for this silent path made
   * `list-chats` etc. hang for minutes when cookies were stale, blowing
   * the LLM tool-call window. With a tight cap, the flow either yields
   * a token quickly or fails with `auth_failed: elevated token capture
   * timed out — run `ask-marcel login` to refresh.`
   */
  readonly elevatedRecaptureTimeoutMs?: number;
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

const M365_CLOUD_URL = 'https://m365.cloud.microsoft/search';

const PROXY_ENV_KEYS = ['HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy'];

const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

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

  const acquireToken = async (scopes: string[], startUrl: string): Promise<BrowserTokenResult | null> => {
    await cleanupSingletonLocks(profileDir, fs);

    let capturedAccess: AccessToken | null = null;
    let capturedRefresh: string | null = null;

    context = await launchContext(false);
    page = await context.newPage();
    const activePage = page;

    const handleResponse = async (response: ResponseLike): Promise<void> => {
      if (capturedAccess) return;
      const url = response.url();
      if (!TOKEN_HOSTS.some((d) => url.includes(d))) return;
      const ct = response.headers()['content-type'] ?? '';
      if (!ct.includes('json')) return;
      try {
        const body = await response.text();
        if (!body.includes('access_token')) return;
        const data = JSON.parse(body) as { access_token?: string; refresh_token?: string };
        const raw = data.access_token ?? '';
        const validated = accessToken(raw);
        if (!validated.ok) {
          trace(`[DEBUG] non-graph token skipped, len: ${raw.length}\n`);
          return;
        }
        capturedAccess = validated.value;
        capturedRefresh = data.refresh_token ?? null;
        logger.info('token_captured', { len: validated.value.length });
        trace(`[DEBUG] graph token captured, len: ${validated.value.length}\n`);
      } catch {
        // ignore parse errors
      }
    };
    activePage.on('response', (r) => {
      void handleResponse(r);
    });

    logger.info('browser_navigating', { url: startUrl });
    trace(`[DEBUG] navigating to: ${startUrl}\n`);
    try {
      await activePage.goto(startUrl, { waitUntil: 'domcontentloaded', timeout: navigationTimeoutMs });
      logger.info('browser_navigated', { url: startUrl });
    } catch (navError) {
      const navMsg = navError instanceof Error ? navError.message : String(navError);
      trace(`[DEBUG] navigation error (non-fatal): ${navMsg}\n`);
      logger.info('browser_navigation_error', { message: navMsg });
    }

    const tryReturnCaptured = async (label: string): Promise<BrowserTokenResult | null> => {
      if (!capturedAccess) return null;
      trace(`[DEBUG] ${label}\n`);
      await cleanup();
      return { accessToken: capturedAccess, refreshToken: capturedRefresh };
    };

    await sleep(initialSettleMs);
    const settleResult = await tryReturnCaptured('token captured during initial settle');
    if (settleResult) return settleResult;

    const currentUrl = activePage.url();
    if (!currentUrl.includes('login.microsoftonline.com') && !currentUrl.includes('login.live.com')) {
      trace('[DEBUG] already signed in — clearing session to force fresh login\n');
      logger.info('browser_force_relogin', { url: currentUrl });
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
        await activePage.goto(startUrl, { waitUntil: 'domcontentloaded', timeout: navigationTimeoutMs });
      } catch {
        // non-fatal
      }
      await sleep(postReloginSettleMs);
      const reloginResult = await tryReturnCaptured('token captured after force relogin');
      if (reloginResult) return reloginResult;
    }

    trace('[DEBUG] complete sign-in in the browser window — waiting up to 5 min\n');
    const deadline = Date.now() + pollDeadlineMs;
    while (Date.now() < deadline) {
      const polledResult = await tryReturnCaptured('token found in polling loop, closing browser');
      if (polledResult) return polledResult;
      await sleep(pollIntervalMs);
    }

    trace('[DEBUG] polling loop timeout expired, no token captured\n');
    await cleanup();
    return null;
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
  const acquireElevatedToken = async (): Promise<AccessToken | null> => {
    await cleanupSingletonLocks(profileDir, fs);

    let captured: AccessToken | null = null;

    // Headless flakiness with Microsoft anti-automation has been
    // observed on m365.cloud.microsoft — the SPA sometimes refuses to
    // run its OAuth dance in headless mode. Visible launch is the
    // safer default. The window opens and closes within seconds; user
    // sees a brief flash but no interaction is required.
    const elevatedCtx = await launchContext(false);
    const elevatedPage = await elevatedCtx.newPage();

    elevatedPage.on('request', (req) => {
      if (captured) return;
      const auth = req.headers()['authorization'];
      if (typeof auth !== 'string' || !auth.startsWith('Bearer ')) return;
      const raw = auth.slice('Bearer '.length);
      const claims = decodeJwtPayload(raw);
      const appid = typeof claims['appid'] === 'string' ? (claims['appid'] as string) : undefined;
      const aud = typeof claims['aud'] === 'string' ? (claims['aud'] as string) : undefined;
      if (!appid || !aud) return;
      if (!ELEVATED_APP_IDS.includes(appid)) return;
      if (aud !== 'https://graph.microsoft.com') return;
      const validated = accessToken(raw);
      if (!validated.ok) return;
      captured = validated.value;
      logger.info('elevated_token_captured', { appid, len: validated.value.length });
      trace(`[DEBUG] elevated token captured  appid=${appid}  len=${validated.value.length}\n`);
    });

    logger.info('browser_navigating', { url: M365_CLOUD_URL, purpose: 'elevated' });
    trace(`[DEBUG] elevated capture: navigating to ${M365_CLOUD_URL}\n`);
    try {
      await elevatedPage.goto(M365_CLOUD_URL, { waitUntil: 'domcontentloaded', timeout: navigationTimeoutMs });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      trace(`[DEBUG] elevated nav error (non-fatal): ${msg}\n`);
    }

    // Settle: silent SSO capture happens within 3-8s when cookies are warm.
    // The elevated path is non-interactive — the popup is opened with the
    // persistent profile and either silent SSO completes (token captured) or
    // the cookies are stale and silent SSO can't proceed without UI prompts
    // we don't drive. Use a SHORT deadline (default 20s) and fail fast with
    // a clear error pointing at `ask-marcel login` if no token comes back —
    // audit v1.0.0 §1.1 flagged that reusing the 5-min interactive deadline
    // here hangs the chat commands for minutes and blows the LLM tool-call
    // window.
    const elevatedDeadline = Date.now() + elevatedRecaptureTimeoutMs;
    while (Date.now() < elevatedDeadline) {
      if (captured) {
        trace('[DEBUG] elevated capture: token found, closing\n');
        try {
          await elevatedPage.close();
        } catch {
          // ignore
        }
        try {
          await elevatedCtx.close();
        } catch {
          // ignore
        }
        return captured;
      }
      await sleep(pollIntervalMs);
    }

    trace('[DEBUG] elevated capture: deadline expired, no elevated token captured\n');
    logger.info('elevated_token_capture_timeout');
    try {
      await elevatedPage.close();
    } catch {
      // ignore
    }
    try {
      await elevatedCtx.close();
    } catch {
      // ignore
    }
    return null;
  };

  const close = async (): Promise<void> => {
    logger.info('browser_auth_close');
    await cleanup();
  };

  return { acquireToken, acquireElevatedToken, close };
};

const defaultFileSystem = (): FileSystem => (typeof globalThis.Bun !== 'undefined' ? createBunFileSystem() : createNodeFileSystem());

const createBrowserAuth = (deps: { logger: Logger; fs?: FileSystem }): BrowserAuth =>
  createBrowserAuthFromApi(createPlaywrightApi(loadPlaywright), { logger: deps.logger, fs: deps.fs ?? defaultFileSystem() });

export { createBrowserAuth, createBrowserAuthFromApi, createPlaywrightApi };
export type { BrowserAuth, BrowserAuthApi, BrowserAuthConfig, BrowserTokenResult, ChromiumLike, ContextLike, PageLike, PlaywrightLoader, RequestLike, ResponseLike };
