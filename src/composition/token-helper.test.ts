import { describe, expect, it } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { err } from '../domain/result.ts';
import { tokenFingerprint } from '../domain/token-fingerprint.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createAuthLadder } from '../infra/auth.ts';
import type { TokenCacheLock } from '../infra/token-cache-lock.ts';
import type { BrowserAuth } from '../infra/browser-auth.ts';
import type { BrowserLadderOptions, TokenHelperDeps } from './token-helper.ts';
import { runTokenHelper } from './token-helper.ts';

const HOME = '/virtual/home';
const CACHE = '/virtual/home/.ask-marcel/token-cache.json';
const LOCATOR = '/virtual/home/.ask-marcel/token-helper.json';
const LOCATION = { execPath: '/usr/local/bin/node', entry: '/opt/ask-marcel/dist/token.js', version: '2.8.0' };
const ignore = (): void => undefined;
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;

const jwt = (claims: Record<string, unknown>): AccessToken => accessTokenUnsafe(`${btoa(JSON.stringify({ alg: 'RS256' }))}.${btoa(JSON.stringify(claims))}.sig`);
const graphToken = (name: string): AccessToken => jwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com', name });

const cacheHolding = (entries: Record<string, unknown>): FileSystemFake => {
  const fs = createFileSystemFake();
  fs.seed(CACHE, JSON.stringify({ access_token: '', expires_on: 0, refresh_token: 'the-refresh-token', ...entries }));
  return fs;
};

type Run = { readonly exitCode: number; readonly lines: ReadonlyArray<string>; readonly browserLadders: ReadonlyArray<string> };

// Runs the helper on a virtual home; a browser ladder, if one is asked for, is
// the plain ladder over the same cache, and the tier it was asked for is kept.
const run = async (argv: ReadonlyArray<string>, fs: FileSystemFake, extra: Partial<TokenHelperDeps> = {}): Promise<Run> => {
  const lines: string[] = [];
  const browserLadders: string[] = [];
  const browserLadder = async (options: BrowserLadderOptions): Promise<ReturnType<typeof createAuthLadder>> => {
    browserLadders.push(options.tier);
    return createAuthLadder({ cachePath: options.cachePath, browserProfileDir: options.browserProfileDir, logger: options.logger, fs: options.fs, interactive: true });
  };
  const exitCode = await runTokenHelper({ argv, location: LOCATION, home: HOME, env: {}, interactive: false, fs, print: (line) => lines.push(line), browserLadder, ...extra });
  return { exitCode, lines, browserLadders };
};

const onlyLine = (r: Run): Record<string, unknown> => {
  expect(r.lines).toHaveLength(1);
  expect(r.lines[0]?.endsWith('\n')).toBe(true);
  return JSON.parse(r.lines[0] ?? '') as Record<string, unknown>;
};

describe('the token helper entry', () => {
  it('prints the cached basic token as one JSON line and exits 0', async () => {
    const cached = graphToken('cached');
    const r = await run(['--tier', 'basic'], cacheHolding({ access_token: cached, expires_on: inAnHour() }));
    expect(r.exitCode).toBe(0);
    expect(onlyLine(r)).toMatchObject({ accessToken: cached });
  });

  it('redeems a chat token through the token endpoint and prints it with the cached region', async () => {
    const minted = jwt({ exp: inAnHour(), aud: 'https://chatsvcagg.teams.microsoft.com' });
    const fetchFn = async (): Promise<Response> => new Response(JSON.stringify({ access_token: minted, expires_in: 3600, refresh_token: 'rotated' }));
    const r = await run(['--tier', 'chatsvcagg'], cacheHolding({ chatsvcagg_region: 'amer' }), { fetchFn });
    expect(onlyLine(r)).toMatchObject({ accessToken: minted, region: 'amer' });
  });

  it('fails fast, without a browser, when no session is cached and nobody is at the terminal', async () => {
    const r = await run(['--tier', 'basic'], createFileSystemFake());
    expect(r.exitCode).toBe(1);
    expect(onlyLine(r)).toMatchObject({ errorCode: 'not_authenticated', tier: 'basic' });
    expect(r.browserLadders).toEqual([]);
  });

  it('exits 2 on arguments it refuses, still with one JSON line', async () => {
    const r = await run(['--tier', 'admin'], createFileSystemFake());
    expect(r.exitCode).toBe(2);
    expect(onlyLine(r)).toMatchObject({ errorCode: 'invalid_arguments' });
  });

  // Read and write find the helper through this file, so it is written on every
  // run, a refused one included, and only the owner can read it.
  it('records where the helper lives on every run', async () => {
    const fs = createFileSystemFake();
    await run(['--tier', 'admin'], fs);
    expect(JSON.parse(fs.snapshot(LOCATOR) ?? '')).toEqual(LOCATION);
    expect(fs.snapshotMode(LOCATOR)).toBe(0o600);
  });

  it('leaves the locator alone when it already says the same, and rewrites it after an upgrade', async () => {
    const fs = createFileSystemFake();
    const sameButSpaced = JSON.stringify(LOCATION, null, 2);
    fs.seed(LOCATOR, sameButSpaced);
    await run(['--tier', 'admin'], fs);
    expect(fs.snapshot(LOCATOR)).toBe(sameButSpaced);
    await run(['--tier', 'admin'], fs, { location: { ...LOCATION, version: '2.9.0' } });
    expect(JSON.parse(fs.snapshot(LOCATOR) ?? '')).toMatchObject({ version: '2.9.0' });
  });

  // A switch from node to bun, or a reinstall under another prefix, keeps the
  // version: the locator must follow the runtime and the entry too.
  it('rewrites the locator when the runtime or the install moved under the same version', async () => {
    const fs = createFileSystemFake();
    fs.seed(LOCATOR, JSON.stringify(LOCATION));
    const bun = { ...LOCATION, execPath: '/usr/local/bin/bun' };
    await run(['--tier', 'admin'], fs, { location: bun });
    expect(JSON.parse(fs.snapshot(LOCATOR) ?? '')).toEqual(bun);
    const moved = { ...bun, entry: '/elsewhere/dist/token.js' };
    await run(['--tier', 'admin'], fs, { location: moved });
    expect(JSON.parse(fs.snapshot(LOCATOR) ?? '')).toEqual(moved);
  });

  // The file is the user's to edit; whatever a hand or another tool left there,
  // the helper overwrites it and answers.
  it('overwrites a locator that is not a location, and still answers', async () => {
    for (const stray of ['null', '[1]', '5', 'not json']) {
      const fs = createFileSystemFake();
      fs.seed(LOCATOR, stray);
      const r = await run(['--tier', 'admin'], fs);
      expect(onlyLine(r)).toMatchObject({ errorCode: 'invalid_arguments' });
      expect(JSON.parse(fs.snapshot(LOCATOR) ?? '')).toEqual(LOCATION);
    }
  });

  it('still hands out the token when the locator cannot be written', async () => {
    const cached = graphToken('cached');
    const fs: FileSystemFake = {
      ...cacheHolding({ access_token: cached, expires_on: inAnHour() }),
      writeTextAtomic: async () => err({ type: 'io_failed', message: 'read-only home' }),
    };
    const r = await run(['--tier', 'basic'], fs);
    expect(r.exitCode).toBe(0);
    expect(onlyLine(r)).toMatchObject({ accessToken: cached });
  });

  it('in a terminal, builds the ladder that may open a browser for a basic or elevated request only', async () => {
    const cached = graphToken('cached');
    const fs = cacheHolding({ access_token: cached, expires_on: inAnHour(), elevated_access_token: cached, elevated_expires_on: inAnHour() });
    const basic = await run(['--tier', 'basic'], fs, { interactive: true });
    const elevated = await run(['--tier', 'elevated'], fs, { interactive: true });
    const chat = await run(['--tier', 'chatsvcagg'], fs, { interactive: true });
    expect([...basic.browserLadders, ...elevated.browserLadders, ...chat.browserLadders]).toEqual(['basic', 'elevated']);
    expect(onlyLine(basic)).toMatchObject({ accessToken: cached });
  });

  it('never builds the browser ladder for a replay, even in a terminal', async () => {
    const newer = graphToken('newer');
    const r = await run(['--tier', 'basic', '--reject', await tokenFingerprint(graphToken('refused'))], cacheHolding({ access_token: newer, expires_on: inAnHour() }), {
      interactive: true,
    });
    expect(r.browserLadders).toEqual([]);
    expect(onlyLine(r)).toMatchObject({ accessToken: newer });
  });

  // Another process holds the token-cache lock (a browser `login`, say). A
  // person at the terminal waits as long as a sign-in may take; an agent gives
  // up after twenty seconds. Neither case opens a browser for these requests.
  it('in a terminal, waits for a held lock as long as a sign-in may take; unattended, gives up after twenty seconds', async () => {
    const waits: number[] = [];
    const busy: TokenCacheLock = {
      withLock: async (_purpose, waitBudgetMs) => {
        waits.push(waitBudgetMs);
        return err({ type: 'lock_busy', purpose: 'browser' });
      },
    };
    const refused = graphToken('refused');
    const replay = ['--tier', 'basic', '--reject', await tokenFingerprint(refused)];
    const fs = (): FileSystemFake => cacheHolding({ access_token: refused, expires_on: inAnHour() });
    for (const argv of [['--tier', 'chatsvcagg'], replay]) {
      const attended = await run(argv, fs(), { interactive: true, lock: busy });
      expect(onlyLine(attended)).toMatchObject({ errorCode: 'sign_in_in_progress', message: expect.stringContaining('waited 420 s') });
      expect(attended.browserLadders).toEqual([]);
      const unattended = await run(argv, fs(), { interactive: false, lock: busy });
      expect(onlyLine(unattended)).toMatchObject({ errorCode: 'sign_in_in_progress', message: expect.stringContaining('waited 20 s') });
    }
    expect(waits).toEqual([420_000, 20_000, 420_000, 20_000]);
  });

  // The production browser ladder is loaded on demand; with the token cached it
  // answers from the cache and nothing launches.
  it('loads the real browser ladder only in a terminal, and answers from the cache', async () => {
    const cached = graphToken('cached-elevated');
    const fs = cacheHolding({ elevated_access_token: cached, elevated_expires_on: inAnHour() });
    const lines: string[] = [];
    const exitCode = await runTokenHelper({ argv: ['--tier', 'elevated'], location: LOCATION, home: HOME, env: {}, interactive: true, fs, print: (line) => lines.push(line) });
    expect(exitCode).toBe(0);
    expect(JSON.parse(lines[0] ?? '')).toMatchObject({ accessToken: cached });
  });

  // A Playwright stand-in that records which capture the ladder asked for and
  // finds no session, so nothing launches.
  const recordingBrowser = (): { readonly create: () => BrowserAuth; readonly asked: unknown[] } => {
    const asked: unknown[] = [];
    const noSession = { ok: false, reason: 'sso_timeout' } as const;
    const create = (): BrowserAuth => ({
      acquireElevatedToken: async (options) => {
        asked.push({ elevated: options });
        return noSession;
      },
      acquireChatsvcaggToken: async () => {
        asked.push('chatsvcagg');
        return noSession;
      },
      acquireIc3Token: async () => {
        asked.push('ic3');
        return noSession;
      },
      acquireBothTokens: async () => {
        asked.push('signIn');
        return { teams: null, elevated: noSession, chatsvcagg: noSession, ic3: noSession };
      },
      close: async () => {},
    });
    return { create, asked };
  };

  it('in a terminal with no session, a basic request opens the sign-in and an elevated one waits for the elevated sign-in', async () => {
    const browse = async (tier: string): Promise<unknown[]> => {
      const browser = recordingBrowser();
      await runTokenHelper({
        argv: ['--tier', tier],
        location: LOCATION,
        home: HOME,
        env: {},
        interactive: true,
        fs: createFileSystemFake(),
        print: ignore,
        createBrowser: browser.create,
      });
      return browser.asked;
    };
    const basic = await browse('basic');
    expect(basic).toContain('signIn');
    expect(basic).not.toContainEqual({ elevated: { awaitSignIn: true } });
    const elevated = await browse('elevated');
    expect(elevated).toContainEqual({ elevated: { awaitSignIn: true } });
    expect(elevated).not.toContain('signIn');
  });

  // The elevated sign-in runs on the real browser ladder, which opens no basic
  // sign-in. A person at the terminal still waits for another process's
  // sign-in as long as it may take, not the twenty seconds an agent gets.
  it('in a terminal, an elevated request waits for a held lock as long as a sign-in may take', async () => {
    const waits: number[] = [];
    const busy: TokenCacheLock = {
      withLock: async (_purpose, waitBudgetMs) => {
        waits.push(waitBudgetMs);
        return err({ type: 'lock_busy', purpose: 'browser' });
      },
    };
    const lines: string[] = [];
    const browser = recordingBrowser();
    const exitCode = await runTokenHelper({
      argv: ['--tier', 'elevated'],
      location: LOCATION,
      home: HOME,
      env: {},
      interactive: true,
      fs: createFileSystemFake(),
      print: (line) => lines.push(line),
      lock: busy,
      createBrowser: browser.create,
    });
    expect(exitCode).toBe(1);
    expect(JSON.parse(lines[0] ?? '')).toMatchObject({ errorCode: 'sign_in_in_progress', tier: 'elevated', message: expect.stringContaining('waited 420 s') });
    expect(waits).toEqual([420_000]);
    expect(browser.asked).toEqual([]);
  });

  // Unless told, the helper asks stdin: a terminal is a person, anything else
  // (a pipe, a spawning MCP server) is not.
  const withStdinTerminal = async <T>(isTTY: boolean, task: () => Promise<T>): Promise<T> => {
    const original = Object.getOwnPropertyDescriptor(process.stdin, 'isTTY');
    Object.defineProperty(process.stdin, 'isTTY', { value: isTTY, configurable: true });
    try {
      return await task();
    } finally {
      if (original === undefined) Reflect.deleteProperty(process.stdin, 'isTTY');
      else Object.defineProperty(process.stdin, 'isTTY', original);
    }
  };

  it('opens no browser when stdin is not a terminal and nobody said otherwise', async () => {
    const r = await withStdinTerminal(false, async () => run(['--tier', 'basic'], createFileSystemFake(), { interactive: undefined }));
    expect(r.exitCode).toBe(1);
    expect(onlyLine(r)).toMatchObject({ errorCode: 'not_authenticated', tier: 'basic' });
    expect(r.browserLadders).toEqual([]);
  });

  it('builds the browser ladder when stdin is a terminal and nobody said otherwise', async () => {
    const r = await withStdinTerminal(true, async () => run(['--tier', 'basic'], createFileSystemFake(), { interactive: undefined }));
    expect(r.browserLadders).toEqual(['basic']);
  });

  // The defaults are this process's: the real file system (on a temporary home
  // here) and stdout.
  it('writes its line to stdout and the locator to disk when given no printer and no file system', async () => {
    const home = mkdtempSync(join(tmpdir(), 'token-helper-'));
    const target = process.stdout;
    const original = target.write.bind(target);
    let captured = '';
    target.write = (chunk: string | Uint8Array): boolean => {
      captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
      return true;
    };
    try {
      await runTokenHelper({ argv: ['--tier', 'admin'], location: LOCATION, home });
    } finally {
      target.write = original;
    }
    expect(JSON.parse(captured)).toMatchObject({ errorCode: 'invalid_arguments' });
    expect(JSON.parse(readFileSync(join(home, '.ask-marcel', 'token-helper.json'), 'utf8'))).toEqual(LOCATION);
    rmSync(home, { recursive: true, force: true });
  });
});
