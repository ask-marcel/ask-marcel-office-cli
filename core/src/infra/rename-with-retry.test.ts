import { describe, expect, it } from 'bun:test';
import { renameWithRetry } from './rename-with-retry.ts';

const refusal = (code: string): Error => Object.assign(new Error(`${code}: rename refused`), { code });

describe('renameWithRetry', () => {
  it('retries a rename refused with EPERM until it succeeds', async () => {
    let calls = 0;
    const rename = async (): Promise<void> => {
      calls += 1;
      if (calls < 3) throw refusal('EPERM');
    };
    await renameWithRetry(rename, 'cache.json.tmp', 'cache.json', [1, 1, 1]);
    expect(calls).toBe(3);
  });

  it('gives up after the last retry and reports the last refusal', async () => {
    let calls = 0;
    const rename = async (): Promise<void> => {
      calls += 1;
      throw refusal('EBUSY');
    };
    await expect(renameWithRetry(rename, 'cache.json.tmp', 'cache.json', [1, 1])).rejects.toThrow('EBUSY');
    expect(calls).toBe(3);
  });

  it('does not retry an error that waiting cannot fix', async () => {
    let calls = 0;
    const rename = async (): Promise<void> => {
      calls += 1;
      throw refusal('ENOENT');
    };
    await expect(renameWithRetry(rename, 'cache.json.tmp', 'cache.json', [1, 1])).rejects.toThrow('ENOENT');
    expect(calls).toBe(1);
  });
});
