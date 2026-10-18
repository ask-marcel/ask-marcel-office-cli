import { describe, expect, it } from 'bun:test';
import { resolveAuthPaths } from './auth-paths.ts';

describe('resolveAuthPaths', () => {
  it('keeps the token cache and the browser profile under one home folder', () => {
    const paths = resolveAuthPaths('/home/robin', {});
    expect(paths).toEqual({
      tokenCache: '/home/robin/.ask-marcel/token-cache.json',
      browserProfile: '/home/robin/.ask-marcel/browser-profile',
    });
  });

  it('moves the browser profile, and nothing else, when ASKMARCEL_BROWSER_PROFILE names another folder', () => {
    const paths = resolveAuthPaths('/home/robin', { ASKMARCEL_BROWSER_PROFILE: '/data/profile' });
    expect(paths).toEqual({ tokenCache: '/home/robin/.ask-marcel/token-cache.json', browserProfile: '/data/profile' });
  });
});
