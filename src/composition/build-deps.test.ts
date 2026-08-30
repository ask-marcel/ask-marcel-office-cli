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

  it('gates the command-path basic-token browser on the session: fails fast when not interactive, keeps the auto-browser when it is', () => {
    const fs = createFileSystemFake();
    const calls: Array<Parameters<typeof createAuthManager>[0]> = [];
    const recordingCreateAuth: typeof createAuthManager = (opts) => {
      calls.push(opts);
      return createAuthManager(opts);
    };

    // Non-interactive (an agent / piped run): no auto-browser for the basic token,
    // so a cold cache fails fast instead of hanging on the interactive-login poll.
    buildDeps({ cachePath: '/virtual/cache.json', logLevel: 'error', fs, interactive: false, createAuth: recordingCreateAuth });
    expect(calls[0]?.acquireBasicViaBrowser).toBe(false);

    // Interactive (a real TTY): the first-run auto-launched sign-in browser stays.
    buildDeps({ cachePath: '/virtual/cache.json', logLevel: 'error', fs, interactive: true, createAuth: recordingCreateAuth });
    expect(calls[1]?.acquireBasicViaBrowser).toBe(true);
  });

  it('gates the command-path elevated re-capture on the session, so a TTY user refreshes transparently and an agent still fails fast', () => {
    const fs = createFileSystemFake();
    const calls: Array<Parameters<typeof createAuthManager>[0]> = [];
    const recordingCreateAuth: typeof createAuthManager = (opts) => {
      calls.push(opts);
      return createAuthManager(opts);
    };

    // Non-interactive (an agent / MCP over stdio): no browser mid-command, so a
    // lapsed elevated token still fails fast with a message naming the fix.
    buildDeps({ cachePath: '/virtual/cache.json', logLevel: 'error', fs, interactive: false, createAuth: recordingCreateAuth });
    expect(calls[0]?.recaptureElevatedViaBrowser).toBe(false);

    // Interactive (a real TTY): silent SSO against the persistent profile refreshes
    // elevated in place, so the command succeeds instead of telling the user to log in.
    buildDeps({ cachePath: '/virtual/cache.json', logLevel: 'error', fs, interactive: true, createAuth: recordingCreateAuth });
    expect(calls[1]?.recaptureElevatedViaBrowser).toBe(true);
    // The SHARED flag stays off in both cases: chatsvcagg / ic3 self-heal from the
    // refresh token inside its false branch, so enabling it would cost them that.
    expect(calls[1]?.recaptureSecondaryViaBrowser).toBe(false);
  });

  it('injects registry-derived secondary-token command lists into BOTH auth managers so the fail-fast messages track the manifest instead of a hardcoded list', () => {
    const fs = createFileSystemFake();
    const calls: Array<Parameters<typeof createAuthManager>[0]> = [];
    const recordingCreateAuth: typeof createAuthManager = (opts) => {
      calls.push(opts);
      return createAuthManager(opts);
    };
    const deps = buildDeps({ cachePath: '/virtual/cache.json', logLevel: 'error', fs, createAuth: recordingCreateAuth });
    deps.makeLoginAuth();
    expect(calls.length).toBe(2);
    for (const call of calls) {
      expect(call.secondaryTokenCommands?.elevated).toEqual(['download-drive-item-version', 'get-chat', 'get-user', 'list-chats']);
      expect(call.secondaryTokenCommands?.chatsvcagg).toEqual(['find-chats-with-user', 'get-teams-chat-message', 'list-teams-chat-messages', 'list-teams-chats-with-messages']);
      expect(call.secondaryTokenCommands?.ic3).toEqual(['list-teams-chat-history']);
    }
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
