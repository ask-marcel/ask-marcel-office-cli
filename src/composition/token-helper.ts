import { homedir } from 'node:os';
import { dirname, join } from 'node:path';
import type { FetchFn, TokenLadder } from '../infra/auth.ts';
import { createAuthLadder } from '../infra/auth.ts';
import type { AuthPaths } from '../infra/auth-paths.ts';
import type { createBrowserAuth } from '../infra/browser-auth.ts';
import type { TokenCacheLock } from '../infra/token-cache-lock.ts';
import { resolveAuthPaths } from '../infra/auth-paths.ts';
import { createBunFileSystem } from '../infra/filesystem-bun.ts';
import { createNodeFileSystem } from '../infra/filesystem-node.ts';
import * as token from '../use-cases/commands/token.ts';
import type { TokenArgs } from '../use-cases/commands/token-args.ts';
import { parseTokenArgs } from '../use-cases/commands/token-args.ts';
import type { AtomicFileWrites, FileSystem } from '../use-cases/ports/filesystem.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { TokenIssuer } from '../use-cases/ports/token-issuer.ts';

/*
 * The `token` helper: `token --tier <tier> [--tenant <guid>] [--reject <fingerprint>]`
 * prints one JSON line on stdout, the token or why there is none, and exits 0
 * (a token), 1 (none) or 2 (arguments refused). Read and write will spawn it for
 * their bearers (docs/plans/2026-10-01-package-split.md, Token protocol).
 *
 * It starts on its own small entry (src/token.ts, dist/token.js), so it loads no
 * commander, no winston, no update notifier and no output renderer: the line is
 * written here. The ladder that may open a browser is imported only when a
 * person is at the terminal and asks for the basic or elevated tier; every other
 * request, and every replay, runs on the headless ladder and fails fast.
 */

// Where the helper lives, written next to the token cache on every run so read
// and write can spawn `execPath entry` without PATH. Holds no secret.
export type TokenHelperLocation = { readonly execPath: string; readonly entry: string; readonly version: string };

export type BrowserLadderOptions = {
  readonly cachePath: string;
  readonly browserProfileDir: string;
  readonly logger: Logger;
  readonly fs: FileSystem & AtomicFileWrites;
  readonly tier: 'basic' | 'elevated';
  // Production launches Playwright; a test records which capture was asked for.
  readonly createBrowser?: typeof createBrowserAuth;
};

// `argv` is what follows the entry (or the `token` word). The rest default to
// this process; tests pass a virtual home, file system and printer.
export type TokenHelperDeps = {
  readonly argv: ReadonlyArray<string>;
  readonly location: TokenHelperLocation;
  readonly home?: string;
  readonly env?: Readonly<Record<string, string | undefined>>;
  readonly interactive?: boolean;
  readonly fs?: FileSystem & AtomicFileWrites;
  readonly print?: (line: string) => void;
  readonly fetchFn?: FetchFn;
  readonly lock?: TokenCacheLock;
  readonly createBrowser?: typeof createBrowserAuth;
  readonly browserLadder?: (options: BrowserLadderOptions) => Promise<TokenLadder>;
};

const ISSUED = 0;
const NOT_ISSUED = 1;
const MISUSED = 2;

// The helper logs nothing: stdout carries its one line and nothing else.
const ignore = (): void => undefined;
const SILENT: Logger = { info: ignore, warn: ignore, error: ignore };

const writeStdout = (line: string): void => {
  process.stdout.write(line);
};

// The file is read as whatever it holds: `null`, a number or an array left by
// hand or by another tool is a stale locator, never a crash.
const sameLocation = (a: unknown, b: TokenHelperLocation): boolean => {
  if (typeof a !== 'object' || a === null) return false;
  const held = a as Partial<TokenHelperLocation>;
  return held.execPath === b.execPath && held.entry === b.entry && held.version === b.version;
};

// Best effort: a home the helper cannot write to costs the locator, not the token.
export const recordTokenHelperLocation = async (fs: FileSystem & AtomicFileWrites, path: string, location: TokenHelperLocation): Promise<void> => {
  const current = await fs.readJson<unknown>(path);
  if (current.ok && sameLocation(current.value, location)) return;
  await fs.writeTextAtomic(path, JSON.stringify(location), 0o600);
};

// Next to the token cache, wherever the home puts it.
export const tokenHelperLocatorPath = (paths: AuthPaths): string => join(dirname(paths.tokenCache), 'token-helper.json');

const loadBrowserLadder = async (options: BrowserLadderOptions): Promise<TokenLadder> => {
  const { createAuthManager } = await import('../infra/auth-browser.ts');
  const { cachePath, browserProfileDir, logger, fs, tier, createBrowser } = options;
  return createAuthManager({
    cachePath,
    browserProfileDir,
    logger,
    fs,
    createBrowser,
    acquireBasicViaBrowser: tier === 'basic',
    recaptureElevatedViaBrowser: tier === 'elevated',
    recaptureSecondaryViaBrowser: false,
  });
};

// A browser only for a person at the terminal, only for the two tiers that a
// browser renews, and never for a replay. The headless ladder still knows a
// person is there, so it waits for a held lock as long as a sign-in may take.
const issuerFor = async (deps: TokenHelperDeps, paths: AuthPaths, fs: FileSystem & AtomicFileWrites, args: TokenArgs): Promise<TokenIssuer> => {
  const base = { cachePath: paths.tokenCache, browserProfileDir: paths.browserProfile, logger: SILENT, fs };
  const { tier } = args.request;
  const interactive = deps.interactive ?? process.stdin.isTTY === true;
  if (interactive && args.rejected === undefined && (tier === 'basic' || tier === 'elevated'))
    return (deps.browserLadder ?? loadBrowserLadder)({ ...base, tier, createBrowser: deps.createBrowser });
  return createAuthLadder({ ...base, interactive, fetchFn: deps.fetchFn, lock: deps.lock });
};

export const runTokenHelper = async (deps: TokenHelperDeps): Promise<number> => {
  const fs = deps.fs ?? (typeof globalThis.Bun === 'undefined' ? createNodeFileSystem() : createBunFileSystem());
  const print = deps.print ?? writeStdout;
  const paths = resolveAuthPaths(deps.home ?? homedir(), deps.env ?? process.env);
  await recordTokenHelperLocation(fs, tokenHelperLocatorPath(paths), deps.location);
  const args = parseTokenArgs(deps.argv);
  if (!args.ok) {
    print(`${JSON.stringify(args.error)}\n`);
    return MISUSED;
  }
  const line = await token.execute(await issuerFor(deps, paths, fs, args.value), args.value);
  print(`${JSON.stringify(line.ok ? line.value : line.error)}\n`);
  return line.ok ? ISSUED : NOT_ISSUED;
};
