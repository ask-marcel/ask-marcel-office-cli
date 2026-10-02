import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '../test-helpers/filesystem-fake.ts';
import type { TokenCacheLockDeps } from './token-cache-lock.ts';
import { hostname } from 'node:os';
import { createSystemTokenCacheLock, createTokenCacheLock, processIsAlive } from './token-cache-lock.ts';

const LOCK = '/virtual/token-cache.json.lock';

// A clock that only moves when the lock waits, so a test controls time exactly
// and every wait budget runs out.
const lockDeps = (fs: FileSystemFake, overrides: Partial<TokenCacheLockDeps> = {}, startAt = 0): TokenCacheLockDeps => {
  let clock = startAt;
  return {
    fs,
    lockPath: LOCK,
    pid: 1000,
    host: 'here',
    now: () => clock,
    isProcessAlive: () => true,
    wait: async (ms) => {
      clock += ms;
    },
    newOwnerId: () => 'me',
    ...overrides,
  };
};

const holder = (fields: { pid?: number; host?: string; startedAt?: number; purpose?: string; owner?: string }): string =>
  JSON.stringify({ owner: 'other', pid: 4242, host: 'here', startedAt: 0, purpose: 'browser', ...fields });

describe('token cache lock', () => {
  it('runs the task holding the lock and removes the lock afterwards', async () => {
    const fs = createFileSystemFake();
    let heldDuringTask = false;
    const result = await createTokenCacheLock(lockDeps(fs)).withLock('refresh', 1000, async () => {
      heldDuringTask = fs.has(LOCK);
      return 'done';
    });
    expect(result).toEqual({ ok: true, value: 'done' });
    expect(heldDuringTask).toBe(true);
    expect(fs.has(LOCK)).toBe(false);
  });

  it('removes the lock when the task fails', async () => {
    const fs = createFileSystemFake();
    const failing = async (): Promise<string> => {
      throw new Error('boom');
    };
    await expect(createTokenCacheLock(lockDeps(fs)).withLock('refresh', 1000, failing)).rejects.toThrow('boom');
    expect(fs.has(LOCK)).toBe(false);
  });

  it('waits while a live process holds the lock, then takes it once released', async () => {
    const fs = createFileSystemFake();
    fs.seed(LOCK, holder({}));
    let waits = 0;
    const deps = lockDeps(fs, {
      wait: async () => {
        waits += 1;
        if (waits === 2) await fs.deleteIfExists(LOCK);
      },
    });
    const result = await createTokenCacheLock(deps).withLock('refresh', 60_000, async () => 'mine');
    expect(result).toEqual({ ok: true, value: 'mine' });
    expect(waits).toBe(2);
  });

  it('takes over a lock whose holder process on this host is dead', async () => {
    const fs = createFileSystemFake();
    fs.seed(LOCK, holder({ pid: 4242 }));
    const deps = lockDeps(fs, { isProcessAlive: (pid) => pid !== 4242 });
    expect(await createTokenCacheLock(deps).withLock('refresh', 1000, async () => 'mine')).toEqual({ ok: true, value: 'mine' });
  });

  it('takes over a lock held longer than its purpose allows, even when its pid looks alive', async () => {
    const fs = createFileSystemFake();
    fs.seed(LOCK, holder({ purpose: 'refresh', host: 'elsewhere', startedAt: 0 }));
    const deps = lockDeps(fs, {}, 3 * 60_000);
    expect(await createTokenCacheLock(deps).withLock('refresh', 1000, async () => 'mine')).toEqual({ ok: true, value: 'mine' });
  });

  it('takes over a lock file left unreadable by a crash, on the second look', async () => {
    const fs = createFileSystemFake();
    fs.seed(LOCK, '');
    let waits = 0;
    const deps = lockDeps(fs, {
      wait: async () => {
        waits += 1;
      },
    });
    expect(await createTokenCacheLock(deps).withLock('refresh', 60_000, async () => 'mine')).toEqual({ ok: true, value: 'mine' });
    expect(waits).toBe(1);
  });

  it('gives up with lock_busy after its wait budget while a live holder keeps the lock', async () => {
    const fs = createFileSystemFake();
    fs.seed(LOCK, holder({ purpose: 'browser' }));
    const result = await createTokenCacheLock(lockDeps(fs)).withLock('refresh', 1000, async () => 'mine');
    expect(result).toEqual({ ok: false, error: { type: 'lock_busy', purpose: 'browser' } });
    expect(fs.snapshot(LOCK)).toBe(holder({ purpose: 'browser' }));
  });

  it('never removes a lock another process took over while the task ran', async () => {
    const fs = createFileSystemFake();
    const result = await createTokenCacheLock(lockDeps(fs)).withLock('refresh', 1000, async () => {
      await fs.writeText(LOCK, holder({ owner: 'newcomer' }));
      return 'mine';
    });
    expect(result.ok).toBe(true);
    expect(fs.snapshot(LOCK)).toBe(holder({ owner: 'newcomer' }));
  });
});

describe('the machine-wide token cache lock', () => {
  it('waits on the real clock while this live process holds the lock, then gives up after its budget', async () => {
    const fs = createFileSystemFake();
    fs.seed(LOCK, holder({ pid: process.pid, host: hostname(), startedAt: Date.now(), purpose: 'refresh' }));
    const started = Date.now();
    const result = await createSystemTokenCacheLock(fs, LOCK).withLock('refresh', 250, async () => 'mine');
    expect(result).toEqual({ ok: false, error: { type: 'lock_busy', purpose: 'refresh' } });
    expect(Date.now() - started).toBeGreaterThanOrEqual(250);
  });
});

describe('processIsAlive', () => {
  it('sees this process as alive', () => {
    expect(processIsAlive(process.pid)).toBe(true);
  });

  it('sees a process that has exited as gone', async () => {
    const child = Bun.spawn(['true']);
    await child.exited;
    expect(processIsAlive(child.pid)).toBe(false);
  });

  it('counts a process it may not signal (pid 1) as alive', () => {
    expect(processIsAlive(1)).toBe(true);
  });
});
