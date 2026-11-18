import { afterEach, describe, expect, it } from 'bun:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { readLocalFile } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import { buildWriteFiles } from './write-files.ts';

/*
 * The write bin reads a local file that a param names (`--ids-file`) from the
 * disk of the machine, with the size limit of the command.
 */

const folders: string[] = [];

afterEach(() => {
  for (const folder of folders.splice(0)) rmSync(folder, { recursive: true, force: true });
});

const folder = (): string => {
  const made = mkdtempSync(join(tmpdir(), 'write-files-'));
  folders.push(made);
  return made;
};

describe('the local file reader of the write bin', () => {
  it('reads a file of the disk, and refuses a folder and a file above the limit', async () => {
    const at = folder();
    writeFileSync(join(at, 'ids.txt'), 'm1\nm2\n');
    const read = await readLocalFile(buildWriteFiles(), join(at, 'ids.txt'), 100);
    expect(read.ok ? new TextDecoder().decode(read.value.bytes) : read.error).toBe('m1\nm2\n');
    expect((await readLocalFile(buildWriteFiles(), at, 100)).ok).toBe(false);
    const big = await readLocalFile(buildWriteFiles(), join(at, 'ids.txt'), 3);
    expect(big.ok ? '' : big.error.code).toBe('local_file_too_large');
  });
});
