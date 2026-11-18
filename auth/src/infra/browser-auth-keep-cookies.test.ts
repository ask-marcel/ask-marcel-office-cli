import { describe, expect, it } from 'bun:test';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { BrowserAuthApi, BrowserAuthConfig, ContextLike, PageLike, ResponseLike } from './browser-auth.ts';
import { createBrowserAuthFromApi } from './browser-auth.ts';

// A plain `login` captures every tier again when the session cannot renew
// itself. The persistent profile is already signed in, so the token grant must
// fire again while the listener is attached. `login --force` clears the cookies
// for that, which deletes the 90-day "Stay signed in" cookie. A capture that
// keeps the cookies clears only the page storage, so the sign-in stays silent.

type FakeBrowser = { readonly api: BrowserAuthApi; readonly wiped: string[] };

const graphToken = (): string => unsignedJwt({ exp: Math.floor(Date.now() / 1000) + 3600, aud: 'https://graph.microsoft.com' });

const tokenAnswer = (refreshToken: string): ResponseLike => ({
  url: () => 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
  headers: () => ({ 'content-type': 'application/json' }),
  text: async () => JSON.stringify({ access_token: graphToken(), refresh_token: refreshToken }),
});

// The first navigation lands in Teams, signed in, with no token grant. The
// second navigation (after the wipe) gives the grant with a refresh token.
const signedInProfile = (): FakeBrowser => {
  const wiped: string[] = [];
  const answerHandlers: Array<(r: ResponseLike) => void> = [];
  let gotoCount = 0;
  let currentUrl = 'about:blank';
  const page: PageLike = {
    on: (event, handler) => {
      if (event === 'response') answerHandlers.push(handler);
    },
    goto: async (url) => {
      gotoCount += 1;
      if (gotoCount === 2) for (const handler of answerHandlers) handler(tokenAnswer('the-new-refresh-token'));
      currentUrl = url;
    },
    url: () => currentUrl,
    evaluate: async () => {
      wiped.push('page storage');
    },
    close: async () => {},
  };
  const context: ContextLike = {
    newPage: async () => page,
    clearCookies: async () => {
      wiped.push('cookies');
    },
    close: async () => {},
  };
  return { api: { launchPersistentContext: async () => context }, wiped };
};

const fastConfig = (progress: string[]): BrowserAuthConfig => ({
  logger: createLoggerFake(),
  fs: createFileSystemFake(),
  profileDir: join(tmpdir(), 'atelier-fake-profile-keep-cookies'),
  initialSettleMs: 1,
  postReloginSettleMs: 1,
  pollIntervalMs: 1,
  pollDeadlineMs: 30,
  navigationTimeoutMs: 100,
  elevatedRecaptureTimeoutMs: 5,
  elevatedPreferenceGraceMs: 1,
  elevatedLaunchTimeoutMs: 200,
  onProgress: (line) => progress.push(line),
});

const KEEP_COOKIES_PROGRESS =
  'Clearing the stored Teams page data so the sign-in can be captured again. The browser cookies stay, so the sign-in is usually silent and asks for no password.';

describe('a full capture that keeps the browser cookies', () => {
  it('clears the page storage but not the cookies, and captures the refresh token of the new token grant', async () => {
    const browser = signedInProfile();
    const result = await createBrowserAuthFromApi(browser.api, fastConfig([])).acquireBothTokens('https://teams.cloud.microsoft/', { skipCacheProbe: true, keepCookies: true });
    expect(browser.wiped).toEqual(['page storage']);
    expect(result.teams?.refreshToken).toBe('the-new-refresh-token');
  });

  it('says that it keeps the cookies, and does not warn that the 90-day sign-in is deleted', async () => {
    const progress: string[] = [];
    await createBrowserAuthFromApi(signedInProfile().api, fastConfig(progress)).acquireBothTokens('https://teams.cloud.microsoft/', { skipCacheProbe: true, keepCookies: true });
    expect(progress.filter((line) => line.startsWith('Clearing'))).toEqual([KEEP_COOKIES_PROGRESS]);
    expect(progress.filter((line) => line.includes('Stay signed in'))).toEqual([]);
  });

  it('still clears the cookies when the caller does not ask to keep them, as login --force does', async () => {
    const browser = signedInProfile();
    await createBrowserAuthFromApi(browser.api, fastConfig([])).acquireBothTokens('https://teams.cloud.microsoft/', { skipCacheProbe: true });
    expect(browser.wiped).toEqual(['cookies', 'page storage']);
  });
});
