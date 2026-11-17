import { afterEach, describe, expect, it } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { FileStat } from '../use-cases/ports/filesystem.ts';
import { createBunFileSystem } from './filesystem-bun.ts';
import { createNodeFileSystem } from './filesystem-node.ts';

/*
 * What is at a local path, from its folder entry alone: a write command checks
 * the kind and the size of a file before it reads the bytes, so a folder or a
 * file above its limit is refused without a read. Both adapters give one answer.
 */

const ADAPTERS: ReadonlyArray<readonly [string, () => FileStat]> = [
  ['the Node adapter', createNodeFileSystem],
  ['the Bun adapter', createBunFileSystem],
];

const dirs: string[] = [];
const scratch = (): string => {
  const dir = mkdtempSync(join(tmpdir(), 'file-stat-'));
  dirs.push(dir);
  return dir;
};

afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

describe.each(ADAPTERS)('%s: stat', (_name, create) => {
  it('answers a file with its size in bytes', async () => {
    const path = join(scratch(), 'report.pdf');
    writeFileSync(path, 'twelve bytes');
    expect(await create().stat(path)).toEqual({ ok: true, value: { kind: 'file', size: 12 } });
  });

  it('answers a folder as a folder', async () => {
    const path = join(scratch(), 'reports');
    mkdirSync(path);
    const result = await create().stat(path);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.kind).toBe('folder');
  });

  it('answers not_found when nothing is there, and when a file stands where a folder of the path should be', async () => {
    const dir = scratch();
    writeFileSync(join(dir, 'a-file'), 'x');
    expect(await create().stat(join(dir, 'gone.pdf'))).toEqual({ ok: false, error: { type: 'not_found' } });
    expect(await create().stat(join(dir, 'a-file', 'inside.pdf'))).toEqual({ ok: false, error: { type: 'not_found' } });
  });

  it('answers io_failed for a path the system refuses', async () => {
    const result = await create().stat(join(scratch(), 'bad\0name'));
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('io_failed');
  });
});
