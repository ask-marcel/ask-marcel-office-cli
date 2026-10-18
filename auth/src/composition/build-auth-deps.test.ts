import { mkdirSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'bun:test';
import { createAuthManager } from '../infra/auth-browser.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import { buildAuthDeps } from './build-auth-deps.ts';

// The composition root of the auth package: the logger, the file system and
// the auth ladder with its interactive gates. It builds no graph: the auth
// package runs no Graph command.

type AuthOptions = Parameters<typeof createAuthManager>[0];

const recording = (): { readonly calls: AuthOptions[]; readonly createAuth: typeof createAuthManager } => {
  const calls: AuthOptions[] = [];
  return {
    calls,
    createAuth: (opts) => {
      calls.push(opts);
      return createAuthManager(opts);
    },
  };
};

const HOME = '/virtual/home';

describe('the composition root of the auth package', () => {
  it('exposes a login-auth factory that builds an auth manager for the interactive login flow', () => {
    const deps = buildAuthDeps({ home: HOME, logLevel: 'error', fs: createFileSystemFake() });
    expect(typeof deps.makeLoginAuth().getAccessToken).toBe('function');
  });

  it('builds the command-path auth manager with secondary browser recapture disabled so a lapsed secondary token fails fast instead of popping a per-command browser', () => {
    const { calls, createAuth } = recording();
    buildAuthDeps({ home: HOME, logLevel: 'error', fs: createFileSystemFake(), createAuth });
    expect(calls[0]?.recaptureSecondaryViaBrowser).toBe(false);
  });

  it('gates the command-path basic-token browser on the session: fails fast when not interactive, keeps the auto-browser when it is', () => {
    const { calls, createAuth } = recording();
    buildAuthDeps({ home: HOME, logLevel: 'error', fs: createFileSystemFake(), interactive: false, createAuth });
    buildAuthDeps({ home: HOME, logLevel: 'error', fs: createFileSystemFake(), interactive: true, createAuth });
    expect(calls.map((c) => c.acquireBasicViaBrowser)).toEqual([false, true]);
  });

  it('gates the command-path elevated re-capture on the session, so a TTY user refreshes transparently and an agent still fails fast', () => {
    const { calls, createAuth } = recording();
    buildAuthDeps({ home: HOME, logLevel: 'error', fs: createFileSystemFake(), interactive: false, createAuth });
    buildAuthDeps({ home: HOME, logLevel: 'error', fs: createFileSystemFake(), interactive: true, createAuth });
    expect(calls.map((c) => c.recaptureElevatedViaBrowser)).toEqual([false, true]);
    // The shared flag stays off in both cases: chatsvcagg and ic3 renew from the refresh token.
    expect(calls.map((c) => c.recaptureSecondaryViaBrowser)).toEqual([false, false]);
  });

  it('tells the command-path auth manager whether a person is at the terminal', () => {
    const { calls, createAuth } = recording();
    buildAuthDeps({ home: HOME, logLevel: 'error', fs: createFileSystemFake(), interactive: false, createAuth });
    buildAuthDeps({ home: HOME, logLevel: 'error', fs: createFileSystemFake(), interactive: true, createAuth });
    expect(calls.map((c) => c.interactive)).toEqual([false, true]);
  });

  it('hands both auth managers the token cache and the browser profile from the same home folder', () => {
    const { calls, createAuth } = recording();
    const deps = buildAuthDeps({ home: HOME, env: {}, logLevel: 'error', fs: createFileSystemFake(), createAuth });
    deps.makeLoginAuth();
    const home = ['/virtual/home/.ask-marcel/token-cache.json', '/virtual/home/.ask-marcel/browser-profile'];
    expect(calls.map((c) => [c.cachePath, c.browserProfileDir])).toEqual([home, home]);
  });

  it('names no secondary-token commands, since the auth package holds no Graph command', () => {
    const { calls, createAuth } = recording();
    buildAuthDeps({ home: HOME, logLevel: 'error', fs: createFileSystemFake(), createAuth }).makeLoginAuth();
    expect(calls.map((c) => c.secondaryTokenCommands)).toEqual([undefined, undefined]);
  });

  it('builds its parts from the defaults of this process when no config is given', () => {
    const deps = buildAuthDeps({ logLevel: 'error' });
    expect([typeof deps.logger.info, typeof deps.auth.getAccessToken, typeof deps.fs.readJson]).toEqual(['function', 'function', 'function']);
  });

  it('returns a cached token from the token cache under the home folder via the default Bun filesystem adapter', async () => {
    // A temp home per run, so parallel `bun test` processes cannot collide.
    const home = mkdtempSync(join(tmpdir(), 'atelier-build-auth-deps-'));
    const future = Math.floor(Date.now() / 1000) + 3600;
    const seededToken = unsignedJwt({ exp: future, aud: 'https://graph.microsoft.com' });
    mkdirSync(join(home, '.ask-marcel'));
    await Bun.write(join(home, '.ask-marcel', 'token-cache.json'), JSON.stringify({ access_token: seededToken, expires_on: future, refresh_token: '' }));
    try {
      const result = await buildAuthDeps({ home, env: {}, logLevel: 'error' }).auth.getAccessToken();
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.value).toBe(seededToken as never);
    } finally {
      rmSync(home, { recursive: true, force: true });
    }
  });
});
