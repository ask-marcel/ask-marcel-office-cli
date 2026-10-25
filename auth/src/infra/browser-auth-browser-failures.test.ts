import { describe, expect, it } from 'bun:test';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { BrowserAuthApi, BrowserAuthConfig, ContextLike, PageLike, ResponseLike } from './browser-auth.ts';
import { createBrowserAuthFromApi } from './browser-auth.ts';

// A browser that fails in the places a real one can fail: the launch of a
// channel, the opening of a page, the wipe of the page storage, and the close.

type FakeBrowserOpts = {
  readonly failingChannels?: ReadonlyArray<'msedge' | 'chrome'>;
  readonly newPageError?: Error;
  readonly closeFails?: boolean;
  readonly evaluateFails?: boolean;
  // The 0-based navigation that times out after its answers arrive.
  readonly timedOutGoto?: number;
  readonly answersPerGoto?: ReadonlyArray<ReadonlyArray<ResponseLike>>;
  readonly urlAfterGoto?: string;
};

type FakeBrowser = { readonly api: BrowserAuthApi; readonly launchedChannels: string[]; readonly closeAttempts: string[] };

const makeFakePage = (opts: FakeBrowserOpts, closeAttempts: string[]): PageLike => {
  const answerHandlers: Array<(r: ResponseLike) => void> = [];
  let gotoCount = 0;
  let currentUrl = 'about:blank';
  return {
    on: (event, handler) => {
      if (event === 'response') answerHandlers.push(handler);
    },
    goto: async (url) => {
      const index = gotoCount;
      gotoCount += 1;
      for (const answer of opts.answersPerGoto?.[index] ?? []) for (const handler of answerHandlers) handler(answer);
      currentUrl = opts.urlAfterGoto ?? url;
      if (index === opts.timedOutGoto) throw new Error('navigation timed out');
    },
    url: () => currentUrl,
    evaluate: async () => {
      if (opts.evaluateFails === true) throw new Error('page storage is gone');
    },
    close: async () => {
      closeAttempts.push('page');
      if (opts.closeFails === true) throw new Error('page already closed');
    },
  };
};

const makeFakeBrowser = (opts: FakeBrowserOpts = {}): FakeBrowser => {
  const launchedChannels: string[] = [];
  const closeAttempts: string[] = [];
  const page = makeFakePage(opts, closeAttempts);
  const context: ContextLike = {
    newPage: async () => {
      if (opts.newPageError !== undefined) throw opts.newPageError;
      return page;
    },
    clearCookies: async () => {},
    close: async () => {
      closeAttempts.push('context');
      if (opts.closeFails === true) throw new Error('browser already closed');
    },
  };
  const api: BrowserAuthApi = {
    launchPersistentContext: async (_profileDir, options) => {
      const channel = options.channel ?? 'bundled';
      launchedChannels.push(channel);
      if (options.channel !== undefined && opts.failingChannels?.includes(options.channel) === true) throw new Error(`${channel} is not installed`);
      return context;
    },
  };
  return { api, launchedChannels, closeAttempts };
};

const fastConfig = (overrides: Partial<BrowserAuthConfig> = {}): BrowserAuthConfig => ({
  logger: createLoggerFake(),
  fs: createFileSystemFake(),
  profileDir: join(tmpdir(), 'atelier-fake-profile-browser-failures'),
  initialSettleMs: 1,
  postReloginSettleMs: 1,
  pollIntervalMs: 1,
  pollDeadlineMs: 30,
  navigationTimeoutMs: 100,
  elevatedRecaptureTimeoutMs: 5,
  elevatedPreferenceGraceMs: 1,
  elevatedLaunchTimeoutMs: 200,
  ...overrides,
});

const graphToken = (): string => unsignedJwt({ exp: Math.floor(Date.now() / 1000) + 3600, aud: 'https://graph.microsoft.com' });

const tokenAnswer = (body: () => Promise<string>): ResponseLike => ({
  url: () => 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
  headers: () => ({ 'content-type': 'application/json' }),
  text: body,
});

describe('browser auth: a machine with neither Edge nor Chrome', () => {
  it('signs in with the bundled browser after both installed channels fail to launch', async () => {
    const browser = makeFakeBrowser({ failingChannels: ['msedge', 'chrome'] });
    const logger = createLoggerFake();
    const auth = createBrowserAuthFromApi(browser.api, fastConfig({ logger }));
    const result = await auth.acquireChatsvcaggToken();
    expect(result).toEqual({ ok: false, reason: 'sso_timeout' });
    expect(browser.launchedChannels).toEqual(['msedge', 'chrome', 'bundled']);
    expect(logger.calls.filter((c) => c.event === 'browser_launched').map((c) => c.meta?.['channel'])).toEqual(['bundled']);
  });
});

describe('browser auth: a page that does not open, in a browser that does not close', () => {
  const pageCrash = new Error('page crashed on open');

  it('reports the page failure, not the close failure, for the elevated capture', async () => {
    const browser = makeFakeBrowser({ newPageError: pageCrash, closeFails: true });
    const auth = createBrowserAuthFromApi(browser.api, fastConfig());
    await expect(auth.acquireElevatedToken()).rejects.toBe(pageCrash);
    expect(browser.closeAttempts).toContain('context');
  });

  it('reports the page failure, not the close failure, for the full sign-in', async () => {
    const browser = makeFakeBrowser({ newPageError: pageCrash, closeFails: true });
    const auth = createBrowserAuthFromApi(browser.api, fastConfig());
    await expect(auth.acquireBothTokens('https://teams.microsoft.com')).rejects.toBe(pageCrash);
    expect(browser.closeAttempts).toContain('context');
  });

  it('reports the page failure, not the close failure, for the Teams chat capture', async () => {
    const browser = makeFakeBrowser({ newPageError: pageCrash, closeFails: true });
    const auth = createBrowserAuthFromApi(browser.api, fastConfig());
    await expect(auth.acquireIc3Token()).rejects.toBe(pageCrash);
    expect(browser.closeAttempts).toContain('context');
  });
});

describe('browser auth: a browser that fails to close at the end', () => {
  it('still reports the elevated capture result', async () => {
    const browser = makeFakeBrowser({ closeFails: true });
    const auth = createBrowserAuthFromApi(browser.api, fastConfig());
    expect(await auth.acquireElevatedToken()).toEqual({ ok: false, reason: 'sso_timeout' });
    expect(browser.closeAttempts).toEqual(['page', 'context']);
  });

  it('still reports the full sign-in result', async () => {
    const browser = makeFakeBrowser({ closeFails: true, urlAfterGoto: 'https://login.microsoftonline.com/common/oauth2' });
    const auth = createBrowserAuthFromApi(browser.api, fastConfig());
    const result = await auth.acquireBothTokens('https://teams.microsoft.com');
    expect(result.teams).toBeNull();
    expect(browser.closeAttempts).toEqual(['page', 'context']);
  });
});

describe('browser auth: sign-in answers and page state that go wrong', () => {
  it('ignores a token answer it cannot read and keeps the next good one', async () => {
    const token = graphToken();
    const unreadable = tokenAnswer(async () => Promise.reject(new Error('answer body lost')));
    const good = tokenAnswer(async () => JSON.stringify({ access_token: token, refresh_token: 'rt' }));
    const browser = makeFakeBrowser({ answersPerGoto: [[unreadable, good]], urlAfterGoto: 'https://login.microsoftonline.com/common/oauth2' });
    const auth = createBrowserAuthFromApi(browser.api, fastConfig());
    const result = await auth.acquireBothTokens('https://teams.microsoft.com');
    expect(String(result.teams?.accessToken)).toBe(token);
    expect(result.teams?.refreshToken).toBe('rt');
  });

  it('still signs in again when the page storage cannot be wiped and the reload times out', async () => {
    const token = graphToken();
    const good = tokenAnswer(async () => JSON.stringify({ access_token: token }));
    const browser = makeFakeBrowser({ evaluateFails: true, timedOutGoto: 1, answersPerGoto: [[], [good]] });
    const auth = createBrowserAuthFromApi(browser.api, fastConfig());
    const result = await auth.acquireBothTokens('https://teams.microsoft.com');
    expect(String(result.teams?.accessToken)).toBe(token);
    expect(result.teams?.refreshToken).toBeNull();
  });
});
