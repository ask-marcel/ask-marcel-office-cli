import { afterEach, describe, expect, it } from 'bun:test';
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import type { BrowserAuthApi, BrowserAuthConfig, ContextLike, PageLike, RequestLike, ResponseLike } from './browser-auth.ts';
import { createBrowserAuth, createBrowserAuthFromApi, createPlaywrightApi } from './browser-auth.ts';
import { createBunFileSystem } from './filesystem-bun.ts';

const makeJwt = (claims: Record<string, unknown>): string => {
  const header = btoa(JSON.stringify({ alg: 'RS256' }));
  const payload = btoa(JSON.stringify(claims));
  return `${header}.${payload}.sig`;
};

const graphTokenJwt = (): string => makeJwt({ exp: Math.floor(Date.now() / 1000) + 3600, aud: 'https://graph.microsoft.com' });

const nonGraphTokenJwt = (): string => makeJwt({ exp: Math.floor(Date.now() / 1000) + 3600, aud: 'https://api.spaces.skype.com' });

const expiredGraphTokenJwt = (): string => makeJwt({ exp: Math.floor(Date.now() / 1000) - 3600, aud: 'https://graph.microsoft.com' });

const tokenResponse = (access: string, refresh: string | null = 'rt'): ResponseLike => ({
  url: () => 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
  headers: () => ({ 'content-type': 'application/json; charset=utf-8' }),
  text: async () => JSON.stringify({ access_token: access, ...(refresh ? { refresh_token: refresh } : {}) }),
});

const customResponse = (overrides: Partial<{ url: string; contentType: string; body: string }>): ResponseLike => ({
  url: () => overrides.url ?? 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
  headers: () => ({ 'content-type': overrides.contentType ?? 'application/json' }),
  text: async () => overrides.body ?? '{}',
});

type FakePageOpts = {
  responsesPerGoto?: ReadonlyArray<ReadonlyArray<ResponseLike>>;
  requestsPerGoto?: ReadonlyArray<ReadonlyArray<RequestLike>>;
  urlsAfterGoto?: ReadonlyArray<string>;
  gotoErrors?: ReadonlyArray<unknown>;
};

type FakePageState = { closed: boolean; evaluated: boolean; gotoCount: number };

const makeFakePage = (opts: FakePageOpts): { page: PageLike; state: FakePageState } => {
  const state: FakePageState = { closed: false, evaluated: false, gotoCount: 0 };
  let currentUrl = 'about:blank';
  const responseHandlers: Array<(r: ResponseLike) => void> = [];
  const requestHandlers: Array<(r: RequestLike) => void> = [];

  const page: PageLike = {
    on: (event, handler) => {
      if (event === 'response') responseHandlers.push(handler);
      else if (event === 'request') requestHandlers.push(handler as (r: RequestLike) => void);
    },
    goto: async (url) => {
      const idx = state.gotoCount;
      state.gotoCount += 1;
      const queuedResponses = opts.responsesPerGoto?.[idx] ?? [];
      for (const r of queuedResponses) for (const h of responseHandlers) h(r);
      const queuedRequests = opts.requestsPerGoto?.[idx] ?? [];
      for (const r of queuedRequests) for (const h of requestHandlers) h(r);
      currentUrl = opts.urlsAfterGoto?.[idx] ?? url;
      const err = opts.gotoErrors?.[idx];
      if (err !== undefined) throw err;
    },
    url: () => currentUrl,
    evaluate: async (fn) => {
      state.evaluated = true;
      // Invoke the page-side function so its body counts toward coverage. In
      // a real browser it runs against `localStorage` / `sessionStorage`;
      // here those globals do not exist, so the call throws synchronously
      // and we mirror Playwright's behaviour by re-throwing.
      fn();
    },
    close: async () => {
      state.closed = true;
    },
  };

  return { page, state };
};

type ChannelOutcome = { channel: 'msedge' | 'chrome' | undefined; outcome: 'fail' | 'ok' };

type FakeApiOpts = {
  channelOutcomes?: ReadonlyArray<ChannelOutcome>;
  pageOpts?: FakePageOpts;
  contextCloseThrows?: boolean;
};

type FakeApiState = {
  contextClosed: boolean;
  cookiesCleared: boolean;
  page: FakePageState;
};

const makeFakeApi = (opts: FakeApiOpts = {}): { api: BrowserAuthApi; state: FakeApiState; getLaunchCount: () => number } => {
  const outcomes = opts.channelOutcomes ?? [{ channel: 'msedge', outcome: 'ok' }];
  let launchIdx = 0;
  const { page, state: pageState } = makeFakePage(opts.pageOpts ?? {});
  const state: FakeApiState = { contextClosed: false, cookiesCleared: false, page: pageState };

  const context: ContextLike = {
    newPage: async () => page,
    clearCookies: async () => {
      state.cookiesCleared = true;
    },
    close: async () => {
      if (opts.contextCloseThrows) throw new Error('context close failed');
      state.contextClosed = true;
    },
  };

  const api: BrowserAuthApi = {
    launchPersistentContext: async () => {
      const idx = launchIdx;
      launchIdx += 1;
      const expected = outcomes[idx];
      if (!expected) throw new Error(`fake out of channel outcomes (idx=${idx})`);
      if (expected.outcome === 'fail') throw new Error('fake launch failed');
      return context;
    },
  };

  return { api, state, getLaunchCount: () => launchIdx };
};

let profileSeq = 0;

const fastConfig = (overrides: Partial<BrowserAuthConfig> = {}): BrowserAuthConfig => {
  profileSeq += 1;
  return {
    logger: createLoggerFake(),
    fs: createFileSystemFake(),
    trace: () => {},
    profileDir: `/tmp/.atelier-fake-profile-${profileSeq}-${Date.now()}`,
    initialSettleMs: 1,
    postReloginSettleMs: 1,
    pollIntervalMs: 1,
    pollDeadlineMs: 30,
    navigationTimeoutMs: 100,
    ...overrides,
  };
};

describe('browser auth — token capture orchestration', () => {
  it('captures a Graph token fired during the initial navigation', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())]],
        urlsAfterGoto: ['https://login.microsoftonline.com/common/oauth2/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
    expect(result?.refreshToken).toBe('rt');
    expect(result?.accessToken.startsWith('eyJ')).toBe(true);
  });

  it('captures the Graph token even when the refresh token is absent', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt(), null)]],
        urlsAfterGoto: ['https://login.microsoftonline.com/common/oauth2/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
    expect(result?.refreshToken).toBeNull();
  });

  it('skips non-Graph and expired tokens before capturing the live Graph token', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(nonGraphTokenJwt()), tokenResponse(expiredGraphTokenJwt()), tokenResponse(graphTokenJwt())]],
        urlsAfterGoto: ['https://login.microsoftonline.com/common/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
    expect(result?.accessToken.startsWith('eyJ')).toBe(true);
  });

  it('forces a re-login when the page already has a Teams session, then captures the token after relogin', async () => {
    const { api, state } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[], [tokenResponse(graphTokenJwt())]],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/', 'https://login.microsoftonline.com/oauth2/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
    expect(state.cookiesCleared).toBe(true);
    expect(state.page.evaluated).toBe(true);
    expect(state.page.gotoCount).toBe(2);
  });

  it('returns null when no token arrives before the polling deadline expires', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        urlsAfterGoto: ['https://login.microsoftonline.com/oauth2/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).toBeNull();
  });

  it('falls back to the chrome channel when msedge fails to launch', async () => {
    const { api, getLaunchCount } = makeFakeApi({
      channelOutcomes: [
        { channel: 'msedge', outcome: 'fail' },
        { channel: 'chrome', outcome: 'ok' },
      ],
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())]],
        urlsAfterGoto: ['https://login.microsoftonline.com/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
    expect(getLaunchCount()).toBe(2);
  });

  it('falls back to the bundled browser when msedge and chrome both fail to launch', async () => {
    const { api, getLaunchCount } = makeFakeApi({
      channelOutcomes: [
        { channel: 'msedge', outcome: 'fail' },
        { channel: 'chrome', outcome: 'fail' },
        { channel: undefined, outcome: 'ok' },
      ],
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())]],
        urlsAfterGoto: ['https://login.microsoftonline.com/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
    expect(getLaunchCount()).toBe(3);
  });

  it('treats a non-Error navigation failure as non-fatal and returns null when no token arrives', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        gotoErrors: ['string-error'],
        urlsAfterGoto: ['https://login.microsoftonline.com/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).toBeNull();
  });

  it('treats an Error navigation failure as non-fatal and returns null when no token arrives', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        gotoErrors: [new Error('navigation timeout')],
        urlsAfterGoto: ['https://login.microsoftonline.com/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).toBeNull();
  });

  it('survives an evaluate failure during force-relogin and continues to wait for the token', async () => {
    const { page, state: pageState } = makeFakePage({
      responsesPerGoto: [[], [tokenResponse(graphTokenJwt())]],
      urlsAfterGoto: ['https://teams.microsoft.com/v2/', 'https://login.microsoftonline.com/...'],
    });
    page.evaluate = async () => {
      throw new Error('storage cleared from inside an iframe');
    };

    const apiState: FakeApiState = { contextClosed: false, cookiesCleared: false, page: pageState };
    const context: ContextLike = {
      newPage: async () => page,
      clearCookies: async () => {
        apiState.cookiesCleared = true;
      },
      close: async () => {
        apiState.contextClosed = true;
      },
    };
    const api: BrowserAuthApi = {
      launchPersistentContext: async () => context,
    };

    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
    expect(apiState.cookiesCleared).toBe(true);
  });

  it('survives a relogin-goto failure and still captures a token that arrives during the post-relogin settle', async () => {
    const { api, state } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[], [tokenResponse(graphTokenJwt())]],
        gotoErrors: [undefined, new Error('relogin nav timeout')],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/', 'https://teams.microsoft.com/v2/'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
    expect(state.page.gotoCount).toBe(2);
  });

  it('swallows a context.close failure during cleanup', async () => {
    const { api } = makeFakeApi({
      contextCloseThrows: true,
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())]],
        urlsAfterGoto: ['https://login.microsoftonline.com/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
  });

  it('swallows a page.close failure during cleanup', async () => {
    const { page, state: pageState } = makeFakePage({
      responsesPerGoto: [[tokenResponse(graphTokenJwt())]],
      urlsAfterGoto: ['https://login.microsoftonline.com/...'],
    });
    page.close = async () => {
      throw new Error('page already detached');
    };
    const apiState: FakeApiState = { contextClosed: false, cookiesCleared: false, page: pageState };
    const context: ContextLike = {
      newPage: async () => page,
      clearCookies: async () => {
        apiState.cookiesCleared = true;
      },
      close: async () => {
        apiState.contextClosed = true;
      },
    };
    const api: BrowserAuthApi = { launchPersistentContext: async () => context };
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
    expect(apiState.contextClosed).toBe(true);
  });

  it('captures the token inside the polling loop when it arrives after both settles', async () => {
    const handlers: Array<(r: ResponseLike) => void> = [];
    const currentUrl = 'https://login.microsoftonline.com/oauth2/...';
    let pageGotoCount = 0;
    const fakePage: PageLike = {
      on: (event, handler) => {
        if (event === 'response') handlers.push(handler);
      },
      goto: async () => {
        pageGotoCount += 1;
      },
      url: () => currentUrl,
      evaluate: async () => {},
      close: async () => {},
    };
    const fakeContext: ContextLike = {
      newPage: async () => fakePage,
      clearCookies: async () => {},
      close: async () => {},
    };
    const api: BrowserAuthApi = {
      launchPersistentContext: async () => fakeContext,
    };

    setTimeout(() => {
      for (const h of handlers) h(tokenResponse(graphTokenJwt()));
    }, 12);

    const browser = createBrowserAuthFromApi(api, fastConfig({ initialSettleMs: 2, postReloginSettleMs: 2, pollIntervalMs: 4, pollDeadlineMs: 200 }));
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result).not.toBeNull();
    expect(pageGotoCount).toBe(1);
  });
});

describe('browser auth — response filter', () => {
  const setup = (response: ResponseLike): ReturnType<typeof makeFakeApi> =>
    makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[response, tokenResponse(graphTokenJwt())]],
        urlsAfterGoto: ['https://login.microsoftonline.com/...'],
      },
    });

  it('ignores responses from non-Microsoft URLs', async () => {
    const { api } = setup(customResponse({ url: 'https://example.com/oauth/token' }));
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result?.accessToken.startsWith('eyJ')).toBe(true);
  });

  it('ignores responses with non-JSON content type', async () => {
    const { api } = setup(customResponse({ contentType: 'text/html' }));
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result?.accessToken.startsWith('eyJ')).toBe(true);
  });

  it('ignores responses without an access_token field', async () => {
    const { api } = setup(customResponse({ body: JSON.stringify({ id_token: 'foo' }) }));
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result?.accessToken.startsWith('eyJ')).toBe(true);
  });

  it('ignores responses with malformed JSON bodies', async () => {
    const { api } = setup(customResponse({ body: '{access_token: not-valid-json' }));
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result?.accessToken.startsWith('eyJ')).toBe(true);
  });

  it('ignores access tokens that do not start with eyJ', async () => {
    const { api } = setup(customResponse({ body: JSON.stringify({ access_token: 'opaque-token' }) }));
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result?.accessToken.startsWith('eyJ')).toBe(true);
  });

  it('ignores responses with a missing content-type header', async () => {
    const noCtType: ResponseLike = {
      url: () => 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      headers: () => ({}),
      text: async () => JSON.stringify({ access_token: graphTokenJwt() }),
    };
    const { api } = setup(noCtType);
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    expect(result?.accessToken.startsWith('eyJ')).toBe(true);
  });
});

describe('browser auth — close lifecycle', () => {
  it('logs and is a no-op when the context was never opened', async () => {
    const { api } = makeFakeApi();
    const logger = createLoggerFake();
    const browser = createBrowserAuthFromApi(api, fastConfig({ logger }));
    await browser.close();
    expect(logger.calls.some((c) => c.event === 'browser_auth_close')).toBe(true);
  });
});

describe('browser auth — profile directory & lock cleanup', () => {
  let envBackup: string | undefined;
  let tmp: string;

  afterEach(() => {
    if (envBackup === undefined) delete process.env.ASKMARCEL_BROWSER_PROFILE;
    else process.env.ASKMARCEL_BROWSER_PROFILE = envBackup;
    if (tmp && existsSync(tmp)) rmSync(tmp, { recursive: true, force: true });
  });

  it('uses ASKMARCEL_BROWSER_PROFILE as the profile directory and cleans existing singleton locks', async () => {
    tmp = mkdtempSync(join(tmpdir(), 'atelier-profile-env-'));
    envBackup = process.env.ASKMARCEL_BROWSER_PROFILE;
    process.env.ASKMARCEL_BROWSER_PROFILE = tmp;
    writeFileSync(join(tmp, 'SingletonLock'), 'lock');
    writeFileSync(join(tmp, 'SingletonCookie'), 'cookie');

    const { api } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())]],
        urlsAfterGoto: ['https://login.microsoftonline.com/...'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ profileDir: undefined, fs: createBunFileSystem() }));
    await browser.acquireToken(['scope'], 'https://teams.microsoft.com');

    expect(existsSync(join(tmp, 'SingletonLock'))).toBe(false);
    expect(existsSync(join(tmp, 'SingletonCookie'))).toBe(false);
  });

  it('falls back to HOME-derived path when ASKMARCEL_BROWSER_PROFILE is not set', async () => {
    tmp = mkdtempSync(join(tmpdir(), 'atelier-profile-home-'));
    envBackup = process.env.ASKMARCEL_BROWSER_PROFILE;
    delete process.env.ASKMARCEL_BROWSER_PROFILE;
    const homeBackup = process.env.HOME;
    const userProfileBackup = process.env.USERPROFILE;
    process.env.HOME = tmp;
    delete process.env.USERPROFILE;
    try {
      const { api } = makeFakeApi({
        pageOpts: {
          responsesPerGoto: [[tokenResponse(graphTokenJwt())]],
          urlsAfterGoto: ['https://login.microsoftonline.com/...'],
        },
      });
      const browser = createBrowserAuthFromApi(api, fastConfig({ profileDir: undefined }));
      const result = await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
      expect(result).not.toBeNull();
    } finally {
      if (homeBackup === undefined) delete process.env.HOME;
      else process.env.HOME = homeBackup;
      if (userProfileBackup !== undefined) process.env.USERPROFILE = userProfileBackup;
    }
  });
});

describe('browser auth — production wiring', () => {
  it('exposes a BrowserAuth port shape from the real createBrowserAuth factory', () => {
    const logger = createLoggerFake();
    const browser = createBrowserAuth({ logger });
    expect(typeof browser.acquireToken).toBe('function');
    expect(typeof browser.close).toBe('function');
  });

  it('strips proxy env vars before invoking the playwright loader', async () => {
    const previousHttp = process.env.HTTP_PROXY;
    const previousHttps = process.env.HTTPS_PROXY;
    const previousLowerHttp = process.env.http_proxy;
    const previousLowerHttps = process.env.https_proxy;
    process.env.HTTP_PROXY = 'http://proxy.example:8080';
    process.env.HTTPS_PROXY = 'https://proxy.example:8443';
    process.env.http_proxy = 'http://proxy.example:8080';
    process.env.https_proxy = 'https://proxy.example:8443';

    const fakeContext: ContextLike = {
      newPage: async () => ({}) as unknown as PageLike,
      clearCookies: async () => {},
      close: async () => {},
    };
    let loaderCalled = false;
    let launchedDir = '';

    const api = createPlaywrightApi(async () => {
      loaderCalled = true;
      return {
        chromium: {
          launchPersistentContext: async (dir) => {
            launchedDir = dir;
            return fakeContext;
          },
        },
      };
    });
    const probeDir = join(tmpdir(), 'atelier-fake-playwright-probe');
    const ctx = await api.launchPersistentContext(probeDir, { headless: false, args: [] });

    try {
      expect(loaderCalled).toBe(true);
      expect(launchedDir).toBe(probeDir);
      expect(ctx).toBe(fakeContext);
      expect(process.env.HTTP_PROXY).toBeUndefined();
      expect(process.env.HTTPS_PROXY).toBeUndefined();
      expect(process.env.http_proxy).toBeUndefined();
      expect(process.env.https_proxy).toBeUndefined();
    } finally {
      if (previousHttp !== undefined) process.env.HTTP_PROXY = previousHttp;
      if (previousHttps !== undefined) process.env.HTTPS_PROXY = previousHttps;
      if (previousLowerHttp !== undefined) process.env.http_proxy = previousLowerHttp;
      if (previousLowerHttps !== undefined) process.env.https_proxy = previousLowerHttps;
    }
  });

  it('logs an event when close() is called on the BrowserAuth port', async () => {
    const logger = createLoggerFake();
    const browser = createBrowserAuth({ logger });
    await browser.close();
    expect(logger.calls.some((c) => c.event === 'browser_auth_close')).toBe(true);
  });

  it('keeps stderr silent when no trace is provided in config (default trace is a no-op)', async () => {
    const original = process.stderr.write.bind(process.stderr);
    let captured = '';
    const swap = (chunk: string | Uint8Array): boolean => {
      captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
      return true;
    };
    process.stderr.write = swap;
    try {
      const { api } = makeFakeApi({
        pageOpts: {
          responsesPerGoto: [[tokenResponse(graphTokenJwt())]],
          urlsAfterGoto: ['https://login.microsoftonline.com/...'],
        },
      });
      const browser = createBrowserAuthFromApi(api, fastConfig({ trace: undefined }));
      await browser.acquireToken(['scope'], 'https://teams.microsoft.com');
    } finally {
      process.stderr.write = original;
    }
    expect(captured).not.toContain('[DEBUG]');
  });

  it('acquireElevatedToken captures the first Bearer with an ODSP-elevated appid (M365ChatClient) on Graph audience', async () => {
    const elevatedJwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://graph.microsoft.com',
      appid: 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1',
    });
    const elevatedRequest: RequestLike = {
      url: () => 'https://graph.microsoft.com/v1.0/me',
      headers: () => ({ authorization: `Bearer ${elevatedJwt}` }),
    };
    const { api } = makeFakeApi({
      pageOpts: { requestsPerGoto: [[elevatedRequest]] },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireElevatedToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.token as unknown as string).toBe(elevatedJwt);
  });

  it('acquireElevatedToken ignores Bearer tokens that do not match an elevated appid', async () => {
    const teamsJwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://graph.microsoft.com',
      appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346', // Teams web client — NOT elevated
    });
    const wrongAppRequest: RequestLike = {
      url: () => 'https://graph.microsoft.com/v1.0/me',
      headers: () => ({ authorization: `Bearer ${teamsJwt}` }),
    };
    const { api } = makeFakeApi({
      pageOpts: { requestsPerGoto: [[wrongAppRequest]] },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 100, pollIntervalMs: 20, elevatedRecaptureTimeoutMs: 100 }));
    const result = await browser.acquireElevatedToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('sso_timeout');
  });

  it('acquireElevatedToken ignores tokens whose audience is not Graph (even if appid is on the elevated list)', async () => {
    const sharepointJwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: '00000003-0000-0ff1-ce00-000000000000', // SharePoint, not Graph
      appid: 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1',
    });
    const wrongAudRequest: RequestLike = {
      url: () => 'https://examplecorp-my.sharepoint.com/_api/...',
      headers: () => ({ authorization: `Bearer ${sharepointJwt}` }),
    };
    const { api } = makeFakeApi({
      pageOpts: { requestsPerGoto: [[wrongAudRequest]] },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 100, pollIntervalMs: 20, elevatedRecaptureTimeoutMs: 100 }));
    const result = await browser.acquireElevatedToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('sso_timeout');
  });

  it('acquireElevatedToken ignores requests without an Authorization header', async () => {
    const noAuthRequest: RequestLike = {
      url: () => 'https://m365.cloud.microsoft/some/static-asset.js',
      headers: () => ({ accept: 'text/javascript' }),
    };
    const { api } = makeFakeApi({
      pageOpts: { requestsPerGoto: [[noAuthRequest]] },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 100, pollIntervalMs: 20, elevatedRecaptureTimeoutMs: 100 }));
    const result = await browser.acquireElevatedToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('sso_timeout');
  });

  it('acquireElevatedToken returns { ok: false, reason: "sso_timeout" } on poll-deadline timeout when no elevated request fires', async () => {
    const { api } = makeFakeApi({ pageOpts: {} });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 100, pollIntervalMs: 20, elevatedRecaptureTimeoutMs: 100 }));
    const result = await browser.acquireElevatedToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('sso_timeout');
  });

  it('acquireElevatedToken returns { ok: false, reason: "navigation_failed" } when goto errors AND polling yields no token (audit login-fix round-1 Wave E)', async () => {
    const { api } = makeFakeApi({
      pageOpts: { gotoErrors: [new Error('navigation timeout')] },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 100, pollIntervalMs: 20, elevatedRecaptureTimeoutMs: 100 }));
    const result = await browser.acquireElevatedToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('navigation_failed');
  });

  it('acquireElevatedToken still captures a token when goto errors but the request event fires (navigation_failed only surfaces when polling also fails)', async () => {
    const elevatedJwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://graph.microsoft.com',
      appid: 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1',
    });
    const elevatedRequest: RequestLike = {
      url: () => 'https://graph.microsoft.com/v1.0/me',
      headers: () => ({ authorization: `Bearer ${elevatedJwt}` }),
    };
    const { api } = makeFakeApi({
      pageOpts: { requestsPerGoto: [[elevatedRequest]], gotoErrors: [new Error('navigation timeout')] },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireElevatedToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.token as unknown as string).toBe(elevatedJwt);
  });

  it('acquireElevatedToken returns { ok: false, reason: "launch_timeout" } when launchPersistentContext hangs longer than elevatedLaunchTimeoutMs (audit login-fix round-1 Wave A)', async () => {
    // Build an api whose launchPersistentContext never resolves — simulate a
    // hung Playwright launch (corrupt profile / stale Singleton lock).
    const hangingApi: BrowserAuthApi = {
      launchPersistentContext: () => new Promise(() => undefined),
    };
    const browser = createBrowserAuthFromApi(hangingApi, fastConfig({ elevatedLaunchTimeoutMs: 50 }));
    const result = await browser.acquireElevatedToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('launch_timeout');
  });

  it('acquireElevatedToken returns { ok: false, reason: "launch_timeout" } when newPage hangs (Wave A — distinct hang point)', async () => {
    const hangingPageApi: BrowserAuthApi = {
      launchPersistentContext: async () => ({
        newPage: () => new Promise(() => undefined),
        clearCookies: async () => undefined,
        close: async () => undefined,
      }),
    };
    const browser = createBrowserAuthFromApi(hangingPageApi, fastConfig({ elevatedLaunchTimeoutMs: 50 }));
    const result = await browser.acquireElevatedToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('launch_timeout');
  });

  it('acquireElevatedToken propagates non-timeout errors from launchPersistentContext (e.g. Playwright not installed)', async () => {
    const erroringApi: BrowserAuthApi = {
      launchPersistentContext: async () => {
        throw new Error('playwright executable missing');
      },
    };
    const browser = createBrowserAuthFromApi(erroringApi, fastConfig({ elevatedLaunchTimeoutMs: 50 }));
    await expect(browser.acquireElevatedToken()).rejects.toThrow('playwright executable missing');
  });

  it('acquireElevatedToken propagates non-timeout errors from newPage (closes context first to avoid leak)', async () => {
    let closedCount = 0;
    const erroringPageApi: BrowserAuthApi = {
      launchPersistentContext: async () => ({
        newPage: async () => {
          throw new Error('context disposed');
        },
        clearCookies: async () => undefined,
        close: async () => {
          closedCount += 1;
        },
      }),
    };
    const browser = createBrowserAuthFromApi(erroringPageApi, fastConfig({ elevatedLaunchTimeoutMs: 50 }));
    await expect(browser.acquireElevatedToken()).rejects.toThrow('context disposed');
    expect(closedCount).toBe(1);
  });
});

describe('browser auth — single-session acquireBothTokens (login-fix round-2)', () => {
  const elevatedJwt = (): string =>
    makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://graph.microsoft.com',
      appid: 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1',
    });

  const elevatedRequest = (): RequestLike => ({
    url: () => 'https://graph.microsoft.com/v1.0/me',
    headers: () => ({ authorization: `Bearer ${elevatedJwt()}` }),
  });

  it('captures both Teams (from goto[0] response) and elevated (from goto[1] request) in one session — no second browser opens', async () => {
    const { api, getLaunchCount } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())], []],
        requestsPerGoto: [[], [elevatedRequest()]],
        urlsAfterGoto: ['https://login.microsoftonline.com/...', 'https://m365.cloud.microsoft/search'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.teams).not.toBeNull();
    expect(result.teams?.refreshToken).toBe('rt');
    expect(result.elevated.ok).toBe(true);
    if (result.elevated.ok) expect(result.elevated.token as unknown as string).toBe(elevatedJwt());
    expect(getLaunchCount()).toBe(1);
  });

  it('returns { teams: null, elevated: sso_timeout } when the Teams token never arrives within the poll deadline', async () => {
    const { api } = makeFakeApi({
      pageOpts: { urlsAfterGoto: ['https://login.microsoftonline.com/...'] },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 30, pollIntervalMs: 5 }));
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.teams).toBeNull();
    expect(result.elevated.ok).toBe(false);
    if (!result.elevated.ok) expect(result.elevated.reason).toBe('sso_timeout');
  });

  it('returns Teams ok + elevated sso_timeout when Teams captured but no elevated request fires within elevatedRecaptureTimeoutMs', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())], []],
        urlsAfterGoto: ['https://login.microsoftonline.com/...', 'https://m365.cloud.microsoft/search'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ elevatedRecaptureTimeoutMs: 30, pollIntervalMs: 5 }));
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.teams).not.toBeNull();
    expect(result.elevated.ok).toBe(false);
    if (!result.elevated.ok) expect(result.elevated.reason).toBe('sso_timeout');
  });

  it('returns Teams ok + elevated navigation_failed when the m365.cloud.microsoft goto throws and no request was captured', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())], []],
        urlsAfterGoto: ['https://login.microsoftonline.com/...', 'https://m365.cloud.microsoft/search'],
        gotoErrors: [undefined, new Error('NS_ERROR_PROXY_CONNECTION_REFUSED')],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ elevatedRecaptureTimeoutMs: 30, pollIntervalMs: 5 }));
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.teams).not.toBeNull();
    expect(result.elevated.ok).toBe(false);
    if (!result.elevated.ok) expect(result.elevated.reason).toBe('navigation_failed');
  });

  it('returns { teams: null, elevated: launch_timeout } when launchPersistentContext hangs longer than elevatedLaunchTimeoutMs', async () => {
    const hangingApi: BrowserAuthApi = {
      launchPersistentContext: () => new Promise(() => undefined),
    };
    const browser = createBrowserAuthFromApi(hangingApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.teams).toBeNull();
    expect(result.elevated.ok).toBe(false);
    if (!result.elevated.ok) expect(result.elevated.reason).toBe('launch_timeout');
  });

  it('returns { teams: null, elevated: launch_timeout } when newPage hangs longer than elevatedLaunchTimeoutMs (distinct hang point)', async () => {
    const hangingPageApi: BrowserAuthApi = {
      launchPersistentContext: async () => ({
        newPage: () => new Promise(() => undefined),
        clearCookies: async () => undefined,
        close: async () => undefined,
      }),
    };
    const browser = createBrowserAuthFromApi(hangingPageApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.teams).toBeNull();
    expect(result.elevated.ok).toBe(false);
    if (!result.elevated.ok) expect(result.elevated.reason).toBe('launch_timeout');
  });

  it('propagates non-timeout errors from launchPersistentContext (e.g. Playwright missing)', async () => {
    const erroringApi: BrowserAuthApi = {
      launchPersistentContext: async () => {
        throw new Error('playwright executable missing');
      },
    };
    const browser = createBrowserAuthFromApi(erroringApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    await expect(browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com')).rejects.toThrow('playwright executable missing');
  });

  it('propagates non-timeout errors from newPage and closes the context to avoid a leak', async () => {
    let closedCount = 0;
    const erroringPageApi: BrowserAuthApi = {
      launchPersistentContext: async () => ({
        newPage: async () => {
          throw new Error('context disposed');
        },
        clearCookies: async () => undefined,
        close: async () => {
          closedCount += 1;
        },
      }),
    };
    const browser = createBrowserAuthFromApi(erroringPageApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    await expect(browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com')).rejects.toThrow('context disposed');
    expect(closedCount).toBe(1);
  });

  it('clears cookies + storage and re-navigates when the persistent profile drops us into an already-signed-in session (forces fresh OAuth dance)', async () => {
    const { api, state } = makeFakeApi({
      pageOpts: {
        // First goto leaves the user on teams.microsoft.com (looks already-signed-in).
        // After clear+reload the Teams response comes back.
        responsesPerGoto: [[], [tokenResponse(graphTokenJwt())], []],
        requestsPerGoto: [[], [], [elevatedRequest()]],
        urlsAfterGoto: ['https://teams.microsoft.com/...', 'https://login.microsoftonline.com/...', 'https://m365.cloud.microsoft/search'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(state.cookiesCleared).toBe(true);
    expect(state.page.evaluated).toBe(true);
    expect(result.teams).not.toBeNull();
    expect(result.elevated.ok).toBe(true);
  });

  it('skips non-graph tokens in the response listener (e.g. Skype audience) and keeps polling for the Graph one', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        // First response carries a Skype-audience token (rejected), second carries the Graph one.
        responsesPerGoto: [[tokenResponse(nonGraphTokenJwt()), tokenResponse(graphTokenJwt())], []],
        requestsPerGoto: [[], [elevatedRequest()]],
        urlsAfterGoto: ['https://login.microsoftonline.com/...', 'https://m365.cloud.microsoft/search'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.teams).not.toBeNull();
    expect(result.elevated.ok).toBe(true);
  });

  it('emits a poll-progress trace line every 10 iterations while waiting for the Teams token', async () => {
    const captured: string[] = [];
    const trace = (m: string): void => {
      captured.push(m);
    };
    const { api } = makeFakeApi({
      pageOpts: { urlsAfterGoto: ['https://login.microsoftonline.com/...'] },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ trace, pollDeadlineMs: 120, pollIntervalMs: 5 }));
    await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(captured.some((m) => m.includes('still polling for Teams token'))).toBe(true);
  });

  it('ASKMARCEL_TRACE=1 wires a stderr trace that echoes .info events through it', async () => {
    const original = process.stderr.write.bind(process.stderr);
    const previousTrace = process.env['ASKMARCEL_TRACE'];
    let captured = '';
    const swap = (chunk: string | Uint8Array): boolean => {
      captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
      return true;
    };
    process.stderr.write = swap;
    process.env['ASKMARCEL_TRACE'] = '1';
    try {
      const logger = createLoggerFake();
      // close() logs 'browser_auth_close' at info — env-gated wrapper must echo it.
      const browser = createBrowserAuth({ logger });
      await browser.close();
    } finally {
      process.stderr.write = original;
      if (previousTrace === undefined) delete process.env['ASKMARCEL_TRACE'];
      else process.env['ASKMARCEL_TRACE'] = previousTrace;
    }
    expect(captured).toContain('[INFO] browser_auth_close');
  });

  it('survives a navigation error on the initial Teams goto when the response listener still fires (non-fatal)', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())], []],
        requestsPerGoto: [[], [elevatedRequest()]],
        urlsAfterGoto: ['https://login.microsoftonline.com/...', 'https://m365.cloud.microsoft/search'],
        gotoErrors: [new Error('initial navigation timeout')],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.teams).not.toBeNull();
    expect(result.elevated.ok).toBe(true);
  });
});

// chatsvcagg-tier capture. The chatsvcagg-audience bearer is minted by the
// Teams web client (same appid as the basic Teams identity, different
// audience). In login-fix round-3 the chatsvcagg leg is captured in the
// same browser session as Teams + elevated; the standalone
// `acquireChatsvcaggToken` is the re-capture fallback when only the
// chatsvcagg cache slot has expired and the persistent profile is warm.
describe('browser auth — chatsvcagg capture (Teams substrate audience)', () => {
  const chatsvcaggJwt = (): string =>
    makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://chatsvcagg.teams.microsoft.com',
      appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346',
    });

  const chatsvcaggRequest = (): RequestLike => ({
    url: () => 'https://chatsvcagg.teams.microsoft.com/api/v2/users/me/chats',
    headers: () => ({ authorization: `Bearer ${chatsvcaggJwt()}` }),
  });

  it('captures a chatsvcagg-audience Bearer in the same single-session run as Teams + elevated (no third browser)', async () => {
    const elevatedJwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://graph.microsoft.com',
      appid: 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1',
    });
    const elevatedRequest: RequestLike = {
      url: () => 'https://graph.microsoft.com/v1.0/me',
      headers: () => ({ authorization: `Bearer ${elevatedJwt}` }),
    };
    // Both the Teams token (response listener) and the chatsvcagg bearer
    // (request listener) are emitted by the SAME page run — the second
    // goto navigates to m365.cloud.microsoft for the elevated leg.
    const { api, getLaunchCount } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())], []],
        requestsPerGoto: [[chatsvcaggRequest()], [elevatedRequest]],
        urlsAfterGoto: ['https://login.microsoftonline.com/...', 'https://m365.cloud.microsoft/search'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.chatsvcagg.ok).toBe(true);
    if (result.chatsvcagg.ok) expect(result.chatsvcagg.token as unknown as string).toBe(chatsvcaggJwt());
    expect(getLaunchCount()).toBe(1);
  });

  it('captures a chatsvcagg bearer via the standalone re-capture path when no Teams response is ever seen but the request listener fires', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        requestsPerGoto: [[chatsvcaggRequest()]],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ elevatedRecaptureTimeoutMs: 30 }));
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.token as unknown as string).toBe(chatsvcaggJwt());
  });

  it('returns { ok: false, reason: "sso_timeout" } from standalone re-capture when no chatsvcagg request fires within the deadline', async () => {
    const { api } = makeFakeApi({
      pageOpts: { urlsAfterGoto: ['https://teams.microsoft.com/v2/'] },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 30, pollIntervalMs: 5, elevatedRecaptureTimeoutMs: 30 }));
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('sso_timeout');
  });

  it('returns { ok: false, reason: "navigation_failed" } from standalone re-capture when teams.microsoft.com goto throws and no request was captured', async () => {
    const { api } = makeFakeApi({
      pageOpts: {
        urlsAfterGoto: ['https://teams.microsoft.com/v2/'],
        gotoErrors: [new Error('NS_ERROR_PROXY_CONNECTION_REFUSED')],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 30, pollIntervalMs: 5, elevatedRecaptureTimeoutMs: 30 }));
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('navigation_failed');
  });

  it('returns { ok: false, reason: "launch_timeout" } from standalone re-capture when launchPersistentContext hangs', async () => {
    const hangingApi: BrowserAuthApi = {
      launchPersistentContext: () => new Promise(() => undefined),
    };
    const browser = createBrowserAuthFromApi(hangingApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('launch_timeout');
  });

  it('returns { ok: false, reason: "launch_timeout" } from standalone re-capture when newPage hangs (distinct hang point)', async () => {
    const hangingPageApi: BrowserAuthApi = {
      launchPersistentContext: async () => ({
        newPage: () => new Promise(() => undefined),
        clearCookies: async () => undefined,
        close: async () => undefined,
      }),
    };
    const browser = createBrowserAuthFromApi(hangingPageApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('launch_timeout');
  });

  it('propagates non-timeout errors from launchPersistentContext (e.g. Playwright not installed)', async () => {
    const erroringApi: BrowserAuthApi = {
      launchPersistentContext: async () => {
        throw new Error('playwright executable missing');
      },
    };
    const browser = createBrowserAuthFromApi(erroringApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    await expect(browser.acquireChatsvcaggToken()).rejects.toThrow('playwright executable missing');
  });

  it('propagates non-timeout errors from newPage and closes the context to avoid a leak', async () => {
    let closedCount = 0;
    const erroringPageApi: BrowserAuthApi = {
      launchPersistentContext: async () => ({
        newPage: async () => {
          throw new Error('context disposed');
        },
        clearCookies: async () => undefined,
        close: async () => {
          closedCount += 1;
        },
      }),
    };
    const browser = createBrowserAuthFromApi(erroringPageApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    await expect(browser.acquireChatsvcaggToken()).rejects.toThrow('context disposed');
    expect(closedCount).toBe(1);
  });

  it('ignores Bearer tokens that do not match the chatsvcagg audience (e.g. Graph-audience traffic on the same page)', async () => {
    const graphAudJwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://graph.microsoft.com',
      appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346',
    });
    const wrongAudRequest: RequestLike = {
      url: () => 'https://graph.microsoft.com/v1.0/me',
      headers: () => ({ authorization: `Bearer ${graphAudJwt}` }),
    };
    const { api } = makeFakeApi({
      pageOpts: {
        requestsPerGoto: [[wrongAudRequest]],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 30, pollIntervalMs: 5, elevatedRecaptureTimeoutMs: 30 }));
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('sso_timeout');
  });

  it('ignores chatsvcagg-audience Bearer tokens that fail JWT validation (expired)', async () => {
    const expired = makeJwt({
      exp: Math.floor(Date.now() / 1000) - 3600,
      aud: 'https://chatsvcagg.teams.microsoft.com',
      appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346',
    });
    const expiredRequest: RequestLike = {
      url: () => 'https://chatsvcagg.teams.microsoft.com/api/v2/users/me/chats',
      headers: () => ({ authorization: `Bearer ${expired}` }),
    };
    const { api } = makeFakeApi({
      pageOpts: {
        requestsPerGoto: [[expiredRequest]],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 30, pollIntervalMs: 5, elevatedRecaptureTimeoutMs: 30 }));
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('sso_timeout');
  });

  it('ignores requests without an Authorization header during chatsvcagg standalone capture', async () => {
    const noAuthRequest: RequestLike = {
      url: () => 'https://teams.microsoft.com/static/asset.js',
      headers: () => ({ accept: 'text/javascript' }),
    };
    const { api } = makeFakeApi({
      pageOpts: {
        requestsPerGoto: [[noAuthRequest]],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ pollDeadlineMs: 30, pollIntervalMs: 5, elevatedRecaptureTimeoutMs: 30 }));
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('sso_timeout');
  });

  it('captures the regional segment from acquireBothTokens too (covers the unified listener branch in the both-tokens path)', async () => {
    // Sibling of the acquireChatsvcaggToken region-capture test below, but
    // exercises the listener inside `acquireBothTokens` so the region-capture
    // branch on the all-three-tokens path is covered too. Without this, the
    // both-tokens path could regress to never capturing the region without
    // any test failing.
    const apacJwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://chatsvcagg.teams.microsoft.com',
      appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346',
    });
    const csaRequest: RequestLike = {
      url: () => 'https://teams.microsoft.com/api/csa/apac/api/v3/teams/users/me?isPrefetch=false',
      headers: () => ({ authorization: `Bearer ${apacJwt}` }),
    };
    const { api } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())], []],
        requestsPerGoto: [[csaRequest], []],
        urlsAfterGoto: ['https://login.microsoftonline.com/...', 'https://m365.cloud.microsoft/search'],
      },
    });
    // Override elevatedRecaptureTimeoutMs so the elevated deadline doesn't
    // hold the test for 20s — we don't care about the elevated leg here.
    const browser = createBrowserAuthFromApi(api, fastConfig({ elevatedRecaptureTimeoutMs: 30 }));
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.chatsvcagg.ok).toBe(true);
    if (result.chatsvcagg.ok) expect(result.chatsvcagg.region).toBe('apac');
  });

  it('acquireIc3Token (standalone) captures an IC3-audience bearer from a single navigation and returns the parsed region', async () => {
    // The standalone path is invoked when the cached IC3 token expires
    // between commands — same Playwright session shape as
    // `acquireChatsvcaggToken`, just a different aud + log prefix.
    const ic3Jwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://ic3.teams.office.com',
      appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346',
    });
    const ic3Request: RequestLike = {
      url: () => 'https://teams.microsoft.com/api/chatsvc/emea/v1/users/ME/conversations/X/messages',
      headers: () => ({ authorization: `Bearer ${ic3Jwt}` }),
    };
    const { api } = makeFakeApi({
      pageOpts: {
        requestsPerGoto: [[ic3Request]],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireIc3Token();
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.token as unknown as string).toBe(ic3Jwt);
      expect(result.region).toBe('emea');
    }
  });

  it('captures an IC3-audience Bearer + region in acquireBothTokens (the chat-history substrate path; covers the IC3 branch of the unified listener)', async () => {
    // Sibling of the chatsvcagg/region tests above, but for IC3. The IC3
    // bearer rides on `teams.microsoft.com/api/chatsvc/<region>/v1/...`
    // (different path prefix from chatsvcagg's `/api/csa/<region>/`); the
    // unified listener has a separate branch for `aud=ic3.teams.office.com`.
    const ic3Jwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://ic3.teams.office.com',
      appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346',
    });
    const ic3Request: RequestLike = {
      url: () => 'https://teams.microsoft.com/api/chatsvc/emea/v1/users/ME/conversations/19%3Aabc/messages?pageSize=200',
      headers: () => ({ authorization: `Bearer ${ic3Jwt}` }),
    };
    const { api } = makeFakeApi({
      pageOpts: {
        responsesPerGoto: [[tokenResponse(graphTokenJwt())], []],
        requestsPerGoto: [[ic3Request], []],
        urlsAfterGoto: ['https://login.microsoftonline.com/...', 'https://m365.cloud.microsoft/search'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ elevatedRecaptureTimeoutMs: 30 }));
    const result = await browser.acquireBothTokens(['scope'], 'https://teams.microsoft.com');
    expect(result.ic3.ok).toBe(true);
    if (result.ic3.ok) {
      expect(result.ic3.token as unknown as string).toBe(ic3Jwt);
      expect(result.ic3.region).toBe('emea');
    }
  });

  it('captures the regional segment from a `/api/csa/<region>/` URL the chatsvcagg bearer rides on', async () => {
    // Post-2026-05 substrate migration: chatsvcagg moved from a dedicated
    // host to `teams.microsoft.com/api/csa/<region>/api/v{N}/...`. The
    // listener parses the region out of the first such URL it sees so the
    // auth manager can persist it. Without this, every Teams substrate URL
    // would default to 'emea' and AMER/APAC tenants would 404 on every
    // chat-content command.
    const amerJwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://chatsvcagg.teams.microsoft.com',
      appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346',
    });
    const csaRequest: RequestLike = {
      url: () => 'https://teams.microsoft.com/api/csa/amer/api/v3/teams/users/me?isPrefetch=false',
      headers: () => ({ authorization: `Bearer ${amerJwt}` }),
    };
    const { api } = makeFakeApi({
      pageOpts: {
        requestsPerGoto: [[csaRequest]],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.token as unknown as string).toBe(amerJwt);
      expect(result.region).toBe('amer');
    }
  });

  it('falls back to "emea" when the chatsvcagg bearer was captured but no `/api/csa/<region>/` URL was observed within the deadline', async () => {
    // The bearer can ride on a non-substrate URL during early sign-in
    // (e.g. a token-endpoint probe). When that happens we still want to
    // ship a usable auth state — fallback default lets European tenants
    // work without a second capture pass.
    const { api } = makeFakeApi({
      pageOpts: {
        requestsPerGoto: [[chatsvcaggRequest()]],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ elevatedRecaptureTimeoutMs: 30 }));
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.region).toBe('emea');
  });

  it('traceChatsvcaggUrls records every URL a chatsvcagg-audience bearer rides on during the trace window', async () => {
    // Diagnostic helper: open a headed browser, dump every chatsvcagg URL the
    // page emits, close when the window expires. Used to discover post-2026-05
    // substrate endpoints (chat-history scrollback, search) we haven't mapped.
    const csaJwt = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 3600,
      aud: 'https://chatsvcagg.teams.microsoft.com',
      appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346',
    });
    const reqA: RequestLike = { url: () => 'https://teams.microsoft.com/api/csa/emea/api/v3/teams/users/me', headers: () => ({ authorization: `Bearer ${csaJwt}` }) };
    const reqB: RequestLike = { url: () => 'https://teams.microsoft.com/api/csa/emea/api/v1/chats/X/messages', headers: () => ({ authorization: `Bearer ${csaJwt}` }) };
    // Same URL again — should be deduped, not double-counted.
    const reqADup: RequestLike = { url: () => 'https://teams.microsoft.com/api/csa/emea/api/v3/teams/users/me', headers: () => ({ authorization: `Bearer ${csaJwt}` }) };
    const { api } = makeFakeApi({
      pageOpts: {
        requestsPerGoto: [[reqA, reqB, reqADup]],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    // 0.01 second — fastConfig durations are tiny so the test stays under 1s.
    const result = await browser.traceChatsvcaggUrls(0.01);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.urls).toHaveLength(2);
      expect(result.urls).toContain('https://teams.microsoft.com/api/csa/emea/api/v1/chats/X/messages');
      expect(result.urls).toContain('https://teams.microsoft.com/api/csa/emea/api/v3/teams/users/me');
    }
  });

  it('traceChatsvcaggUrls returns sso_timeout when the trace window expires without observing any chatsvcagg traffic', async () => {
    // Empty session: page loads, no chatsvcagg-bearer requests fire. The
    // diagnostic produced zero data, so we treat that as failure (lets the
    // CLI message the user to actually use the browser during the window).
    const { api } = makeFakeApi({ pageOpts: { urlsAfterGoto: ['https://teams.microsoft.com/v2/'] } });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.traceChatsvcaggUrls(0.01);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('sso_timeout');
  });

  it('traceChatsvcaggUrls returns navigation_failed when goto throws', async () => {
    const { api } = makeFakeApi({
      pageOpts: { urlsAfterGoto: ['https://teams.microsoft.com/v2/'], gotoErrors: [new Error('NS_ERROR_PROXY_CONNECTION_REFUSED')] },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig());
    const result = await browser.traceChatsvcaggUrls(0.01);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('navigation_failed');
  });

  it('traceChatsvcaggUrls returns launch_timeout when launchPersistentContext hangs', async () => {
    const hangingApi: BrowserAuthApi = { launchPersistentContext: () => new Promise(() => undefined) };
    const browser = createBrowserAuthFromApi(hangingApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    const result = await browser.traceChatsvcaggUrls(0.01);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('launch_timeout');
  });

  it('traceChatsvcaggUrls returns launch_timeout when newPage hangs (distinct hang point)', async () => {
    const hangingPageApi: BrowserAuthApi = {
      launchPersistentContext: async () => ({
        newPage: () => new Promise(() => undefined),
        clearCookies: async () => undefined,
        close: async () => undefined,
      }),
    };
    const browser = createBrowserAuthFromApi(hangingPageApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    const result = await browser.traceChatsvcaggUrls(0.01);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('launch_timeout');
  });

  it('traceChatsvcaggUrls propagates non-timeout errors from launchPersistentContext (e.g. Playwright not installed)', async () => {
    const erroringApi: BrowserAuthApi = {
      launchPersistentContext: async () => {
        throw new Error('playwright executable missing');
      },
    };
    const browser = createBrowserAuthFromApi(erroringApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    await expect(browser.traceChatsvcaggUrls(0.01)).rejects.toThrow('playwright executable missing');
  });

  it('traceChatsvcaggUrls propagates non-timeout errors from newPage and closes the context to avoid a leak', async () => {
    let closedCount = 0;
    const erroringPageApi: BrowserAuthApi = {
      launchPersistentContext: async () => ({
        newPage: async () => {
          throw new Error('context disposed');
        },
        clearCookies: async () => undefined,
        close: async () => {
          closedCount += 1;
        },
      }),
    };
    const browser = createBrowserAuthFromApi(erroringPageApi, fastConfig({ elevatedLaunchTimeoutMs: 30 }));
    await expect(browser.traceChatsvcaggUrls(0.01)).rejects.toThrow('context disposed');
    expect(closedCount).toBe(1);
  });

  it('keeps capturing only the FIRST chatsvcagg Bearer it sees and ignores subsequent ones (matches elevated-tier behavior)', async () => {
    const first = chatsvcaggJwt();
    const second = makeJwt({
      exp: Math.floor(Date.now() / 1000) + 7200,
      aud: 'https://chatsvcagg.teams.microsoft.com',
      appid: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346',
    });
    const firstReq: RequestLike = { url: () => 'https://chatsvcagg.teams.microsoft.com/api/v2/users/me/chats', headers: () => ({ authorization: `Bearer ${first}` }) };
    const secondReq: RequestLike = { url: () => 'https://chatsvcagg.teams.microsoft.com/api/v2/users/me/chats/x/messages', headers: () => ({ authorization: `Bearer ${second}` }) };
    const { api } = makeFakeApi({
      pageOpts: {
        requestsPerGoto: [[firstReq, secondReq]],
        urlsAfterGoto: ['https://teams.microsoft.com/v2/'],
      },
    });
    const browser = createBrowserAuthFromApi(api, fastConfig({ elevatedRecaptureTimeoutMs: 30 }));
    const result = await browser.acquireChatsvcaggToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.token as unknown as string).toBe(first);
  });
});
