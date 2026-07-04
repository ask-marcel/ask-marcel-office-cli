import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'bun:test';
import { createAuthManager } from '../infra/auth.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { buildDeps } from './build-deps.ts';

describe('buildDeps composition root', () => {
  it('wires logger, auth manager, graph client, and process runner when given an explicit cache path', () => {
    const fs = createFileSystemFake();
    const deps = buildDeps({ cachePath: '/virtual/cache.json', logLevel: 'error', fs });
    expect(typeof deps.logger.info).toBe('function');
    expect(typeof deps.auth.getAccessToken).toBe('function');
    expect(typeof deps.auth.logout).toBe('function');
    expect(typeof deps.graph.get).toBe('function');
    expect(typeof deps.processRunner.runInherit).toBe('function');
  });

  it('exposes a login-auth factory that builds an auth manager for the interactive login flow', () => {
    const fs = createFileSystemFake();
    const deps = buildDeps({ cachePath: '/virtual/cache.json', logLevel: 'error', fs });
    expect(typeof deps.makeLoginAuth().getAccessToken).toBe('function');
  });

  it('builds the command-path auth manager with secondary browser recapture disabled so a lapsed secondary token fails fast instead of popping a per-command browser', () => {
    const fs = createFileSystemFake();
    const calls: Array<Parameters<typeof createAuthManager>[0]> = [];
    const recordingCreateAuth: typeof createAuthManager = (opts) => {
      calls.push(opts);
      return createAuthManager(opts);
    };
    buildDeps({ cachePath: '/virtual/cache.json', logLevel: 'error', fs, createAuth: recordingCreateAuth });
    expect(calls[0]?.recaptureSecondaryViaBrowser).toBe(false); // secondary getters fail-fast, no per-command browser
  });

  it('falls back to a home-derived cache path when none is provided', () => {
    const fs = createFileSystemFake();
    const deps = buildDeps({ home: '/virtual/home', logLevel: 'error', fs });
    expect(typeof deps.auth.getAccessToken).toBe('function');
  });

  it('uses default config values when no config is provided', () => {
    const previousLevel = process.env.ASKMARCEL_LOG_LEVEL;
    process.env.ASKMARCEL_LOG_LEVEL = 'error';
    try {
      const deps = buildDeps();
      expect(typeof deps.logger.info).toBe('function');
    } finally {
      if (previousLevel === undefined) delete process.env.ASKMARCEL_LOG_LEVEL;
      else process.env.ASKMARCEL_LOG_LEVEL = previousLevel;
    }
  });

  it('returns a cached token via the default Bun filesystem adapter when run under Bun', async () => {
    // Unique temp dir per run so parallel `bun test` processes can't collide on a
    // shared path (a `Date.now()`-suffixed path repeats across processes in the
    // same millisecond, and one process's cleanup then deletes another's file).
    const dir = mkdtempSync(join(tmpdir(), 'atelier-build-deps-'));
    const tmpCache = join(dir, 'cache.json');
    const future = Math.floor(Date.now() / 1000) + 3600;
    const header = btoa(JSON.stringify({ alg: 'RS256' }));
    const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com' }));
    const seededToken = `${header}.${payload}.sig`;
    await Bun.write(tmpCache, JSON.stringify({ access_token: seededToken, expires_on: future, refresh_token: '' }));
    try {
      const deps = buildDeps({ cachePath: tmpCache, logLevel: 'error' });
      const result = await deps.auth.getAccessToken();
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.value).toBe(seededToken as never);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
