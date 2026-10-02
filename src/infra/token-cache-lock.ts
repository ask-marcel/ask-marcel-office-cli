/*
 * The token cache's lock, shared by every ask-marcel-office process on the
 * machine. Entra single-uses the refresh token, so two processes redeeming it
 * at once leave one of them with a spent token; and two browsers on one
 * persistent profile can corrupt it, because each launch first deletes
 * Chromium's own Singleton lock files. A process takes this lock to redeem, to
 * sign in, or to sign out.
 *
 * The lock is a file created only when absent, holding who took it and why.
 * A waiter takes it over when the holder is plainly gone: its process on this
 * host has exited, or it has held the lock longer than its purpose allows, or
 * the file is unreadable twice in a row (a crash between creating and writing
 * it). Release removes the file only while it is still ours.
 */
import { hostname } from 'node:os';
import { err, ok } from '../domain/result.ts';
import type { Result } from '../domain/result.ts';
import type { AtomicFileWrites, FileSystem } from '../use-cases/ports/filesystem.ts';

export type LockPurpose = 'refresh' | 'browser' | 'logout';

// How long each purpose may hold the lock before a waiter may assume its holder
// died without cleaning up: a sign-in includes up to five minutes for the
// user, a refresh two request timeouts.
const MAX_HOLD_MS: Readonly<Record<LockPurpose, number>> = { refresh: 2 * 60_000, browser: 7 * 60_000, logout: 60_000 };
const POLL_MS = 100;
const UNREADABLE_LOOKS = 2;

type LockRecord = { readonly owner: string; readonly pid: number; readonly host: string; readonly startedAt: number; readonly purpose: LockPurpose };

export type LockError = { readonly type: 'lock_busy'; readonly purpose: LockPurpose | 'unknown' } | { readonly type: 'lock_failed'; readonly message: string };

export type TokenCacheLockDeps = {
  readonly fs: FileSystem & AtomicFileWrites;
  readonly lockPath: string;
  readonly pid: number;
  readonly host: string;
  readonly now: () => number;
  readonly isProcessAlive: (pid: number) => boolean;
  readonly wait: (ms: number) => Promise<void>;
  readonly newOwnerId: () => string;
};

export type TokenCacheLock = {
  readonly withLock: <T>(purpose: LockPurpose, waitBudgetMs: number, task: () => Promise<T>) => Promise<Result<T, LockError>>;
};

const isLockRecord = (value: unknown): value is LockRecord => {
  const record = value as Partial<LockRecord> | null;
  return (
    typeof record?.owner === 'string' &&
    typeof record.pid === 'number' &&
    typeof record.host === 'string' &&
    typeof record.startedAt === 'number' &&
    typeof record.purpose === 'string' &&
    Object.hasOwn(MAX_HOLD_MS, record.purpose)
  );
};

// A process this one may not signal (EPERM) still exists; only "no such
// process" (ESRCH) means it is gone.
export const processIsAlive = (pid: number): boolean => {
  try {
    process.kill(pid, 0);
    return true;
  } catch (e) {
    return (e as NodeJS.ErrnoException).code === 'EPERM';
  }
};

export const createTokenCacheLock = (deps: TokenCacheLockDeps): TokenCacheLock => {
  const readHolder = async (): Promise<LockRecord | 'gone' | 'unreadable'> => {
    const read = await deps.fs.readJson<unknown>(deps.lockPath);
    if (!read.ok) return read.error.type === 'not_found' ? 'gone' : 'unreadable';
    return isLockRecord(read.value) ? read.value : 'unreadable';
  };

  const isStale = (held: LockRecord): boolean => (held.host === deps.host && !deps.isProcessAlive(held.pid)) || deps.now() - held.startedAt > MAX_HOLD_MS[held.purpose];

  // Removes the lock only when it still holds what was last seen there, so a
  // waiter never deletes a lock someone else took in between.
  const removeIfStill = async (owner: string | 'unreadable'): Promise<void> => {
    const current = await readHolder();
    const same = owner === 'unreadable' ? current === 'unreadable' : typeof current === 'object' && current.owner === owner;
    if (same) await deps.fs.deleteIfExists(deps.lockPath);
  };

  // One look at the lock while it is held: take it over when its holder is
  // plainly gone, otherwise report who holds it.
  const inspect = async (unreadableLooks: number): Promise<{ readonly retry: true } | { readonly retry: false; readonly purpose: LockPurpose | 'unknown' }> => {
    const held = await readHolder();
    if (held === 'gone') return { retry: true };
    if (held === 'unreadable') {
      if (unreadableLooks + 1 < UNREADABLE_LOOKS) return { retry: false, purpose: 'unknown' };
      await removeIfStill('unreadable');
      return { retry: true };
    }
    if (!isStale(held)) return { retry: false, purpose: held.purpose };
    await removeIfStill(held.owner);
    return { retry: true };
  };

  const acquire = async (purpose: LockPurpose, waitBudgetMs: number): Promise<Result<string, LockError>> => {
    const owner = deps.newOwnerId();
    const waitingSince = deps.now();
    let unreadableLooks = 0;
    for (;;) {
      const record: LockRecord = { owner, pid: deps.pid, host: deps.host, startedAt: deps.now(), purpose };
      const created = await deps.fs.createExclusive(deps.lockPath, JSON.stringify(record), 0o600);
      if (created.ok) return ok(owner);
      if (created.error.type !== 'already_exists') return err({ type: 'lock_failed', message: 'message' in created.error ? created.error.message : created.error.type });
      const look = await inspect(unreadableLooks);
      if (look.retry) continue;
      unreadableLooks = look.purpose === 'unknown' ? unreadableLooks + 1 : 0;
      if (deps.now() - waitingSince >= waitBudgetMs) return err({ type: 'lock_busy', purpose: look.purpose });
      await deps.wait(POLL_MS);
    }
  };

  const withLock = async <T>(purpose: LockPurpose, waitBudgetMs: number, task: () => Promise<T>): Promise<Result<T, LockError>> => {
    const acquired = await acquire(purpose, waitBudgetMs);
    if (!acquired.ok) return acquired;
    try {
      return ok(await task());
    } finally {
      await removeIfStill(acquired.value);
    }
  };

  return { withLock };
};

const sleep = async (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

// The lock as production uses it: this process, this host, the real clock.
export const createSystemTokenCacheLock = (fs: FileSystem & AtomicFileWrites, lockPath: string): TokenCacheLock =>
  createTokenCacheLock({ fs, lockPath, pid: process.pid, host: hostname(), now: Date.now, isProcessAlive: processIsAlive, wait: sleep, newOwnerId: () => crypto.randomUUID() });
