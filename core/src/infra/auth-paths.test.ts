import { describe, expect, it } from 'bun:test';
import { join } from 'node:path';
import { resolveAuthPaths } from './auth-paths.ts';

// The separator is the one of the system that runs the CLI: a backslash on Windows.
const TOKEN_CACHE = join('/home/robin', '.ask-marcel', 'token-cache.json');

describe('resolveAuthPaths', () => {
  it('keeps the token cache and the browser profile under one home folder', () => {
    const paths = resolveAuthPaths('/home/robin', {});
    expect(paths).toEqual({ tokenCache: TOKEN_CACHE, browserProfile: join('/home/robin', '.ask-marcel', 'browser-profile') });
  });

  it('moves the browser profile, and nothing else, when ASKMARCEL_BROWSER_PROFILE names another folder', () => {
    const paths = resolveAuthPaths('/home/robin', { ASKMARCEL_BROWSER_PROFILE: '/data/profile' });
    expect(paths).toEqual({ tokenCache: TOKEN_CACHE, browserProfile: '/data/profile' });
  });
});
