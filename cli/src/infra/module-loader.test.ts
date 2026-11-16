import { afterEach, describe, expect, it } from 'bun:test';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createModuleLoader } from './module-loader.ts';

/*
 * The real loader runs a module file in this process. The module of the test
 * writes a file next to itself when it runs, so the test sees the run and
 * shares no state with another test.
 */

const RUNS = "import { writeFileSync } from 'node:fs';\nwriteFileSync(new URL('./ran.txt', import.meta.url), 'ran');\n";

// A new `node_modules` folder under the OS temp folder: a bin is a file of an
// installed package, and the coverage table lists no file of a node_modules.
const folders: string[] = [];
const scratch = (): string => {
  const root = mkdtempSync(join(tmpdir(), 'facade-loader-'));
  folders.push(root);
  const folder = join(root, 'node_modules');
  mkdirSync(folder);
  return folder;
};

afterEach(() => {
  for (const folder of folders.splice(0)) rmSync(folder, { recursive: true, force: true });
});

describe('the loader of a module file', () => {
  it('runs the module at a file path', async () => {
    const folder = scratch();
    writeFileSync(join(folder, 'bin.mjs'), RUNS);
    const loaded = await createModuleLoader().load(join(folder, 'bin.mjs'));
    expect(loaded).toEqual({ ok: true, value: undefined });
    expect(readFileSync(join(folder, 'ran.txt'), 'utf8')).toBe('ran');
  });

  it('gives load_failed, with the cause, for a file that does not exist', async () => {
    const loaded = await createModuleLoader().load(join(scratch(), 'absent.mjs'));
    expect(loaded).toMatchObject({ ok: false, error: { type: 'load_failed', message: expect.stringContaining('absent.mjs') } });
  });
});
