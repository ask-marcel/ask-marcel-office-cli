import { afterEach, describe, expect, it } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { FolderListing } from '../use-cases/ports/filesystem.ts';
import { createBunFileSystem } from './filesystem-bun.ts';
import { createNodeFileSystem } from './filesystem-node.ts';

/*
 * The names in a folder: the run journal finds its old runs by their file
 * names. Both adapters give one answer.
 */

const ADAPTERS: ReadonlyArray<readonly [string, () => FolderListing]> = [
  ['the Node adapter', createNodeFileSystem],
  ['the Bun adapter', createBunFileSystem],
];

const dirs: string[] = [];
const scratch = (): string => {
  const dir = mkdtempSync(join(tmpdir(), 'folder-list-'));
  dirs.push(dir);
  return dir;
};

afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

describe.each(ADAPTERS)('%s: list', (_name, create) => {
  it('answers the name of each file and folder in the folder', async () => {
    const dir = scratch();
    writeFileSync(join(dir, 'a.json'), '{}');
    mkdirSync(join(dir, 'b'));
    const listed = await create().list(dir);
    expect(listed.ok ? listed.value.toSorted((x, y) => x.localeCompare(y)) : []).toEqual(['a.json', 'b']);
  });

  it('answers no name for a folder that is not there', async () => {
    expect(await create().list(join(scratch(), 'runs'))).toEqual({ ok: true, value: [] });
  });

  it('answers io_failed when a file stands where the folder should be', async () => {
    const path = join(scratch(), 'runs');
    writeFileSync(path, 'not a folder');
    const listed = await create().list(path);
    expect(listed.ok ? '' : listed.error.type).toBe('io_failed');
  });
});
