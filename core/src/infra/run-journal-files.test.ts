import { afterEach, describe, expect, it } from 'bun:test';
import { mkdirSync, mkdtempSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { err, ok } from '../domain/result.ts';
import { parseRunId } from '../domain/run-id.ts';
import type { RunId } from '../domain/run-id.ts';
import { createNodeFileSystem } from './filesystem-node.ts';
import { createRunJournal, runJournalFolder } from './run-journal-files.ts';

/*
 * The run journal on a real disk, in a temp folder: no test touches the home
 * folder of the user.
 */

const dirs: string[] = [];
const scratch = (): string => {
  const dir = mkdtempSync(join(tmpdir(), 'run-journal-'));
  dirs.push(dir);
  return dir;
};

afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

const id = (raw: string): RunId => {
  const parsed = parseRunId(raw);
  if (!parsed.ok) throw new Error(raw);
  return parsed.value;
};

const RUN = id('20261118T090507Z-0a1b2c3d');

describe('the run journal on disk', () => {
  it('keeps the journal in the runs folder of .ask-marcel under the home folder', () => {
    expect(runJournalFolder(join('home', 'robin'))).toBe(join('home', 'robin', '.ask-marcel', 'runs'));
  });

  it('saves a run as <runId>.json in a folder it makes, and loads it back', async () => {
    const folder = join(scratch(), 'runs');
    const journal = createRunJournal(folder, createNodeFileSystem());
    expect(await journal.save(RUN, { version: 1, command: 'create-mail-draft' })).toEqual(ok(undefined));
    expect(readdirSync(folder)).toEqual([`${RUN}.json`]);
    expect(await journal.load(RUN)).toEqual(ok({ version: 1, command: 'create-mail-draft' }));
  });

  it.skipIf(process.platform === 'win32')('keeps each record readable by its owner only, because it holds tenant data', async () => {
    const folder = scratch();
    await createRunJournal(folder, createNodeFileSystem()).save(RUN, { version: 1 });
    expect(statSync(join(folder, `${RUN}.json`)).mode & 0o777).toBe(0o600);
  });

  it('says not_found for a run it does not hold', async () => {
    expect(await createRunJournal(scratch(), createNodeFileSystem()).load(RUN)).toEqual(err({ type: 'not_found' }));
  });

  it('gives null for a record that is not JSON, which the undo then refuses as unreadable', async () => {
    const folder = scratch();
    writeFileSync(join(folder, `${RUN}.json`), '{ not json');
    expect(await createRunJournal(folder, createNodeFileSystem()).load(RUN)).toEqual(ok(null));
  });

  it('says io_failed when a record cannot be written', async () => {
    const blocker = join(scratch(), 'runs');
    writeFileSync(blocker, 'a file where the folder should be');
    const saved = await createRunJournal(blocker, createNodeFileSystem()).save(RUN, { version: 1 });
    expect(saved.ok ? '' : saved.error.type).toBe('io_failed');
  });

  it('says io_failed when a record cannot be read', async () => {
    const folder = scratch();
    mkdirSync(join(folder, `${RUN}.json`));
    const loaded = await createRunJournal(folder, createNodeFileSystem()).load(RUN);
    expect(loaded.ok ? '' : loaded.error.type).toBe('io_failed');
  });

  it('removes the runs that started before the cutoff, and keeps the newer runs and every other file', async () => {
    const folder = scratch();
    for (const name of ['20261019T090506Z-00000000.json', '20261019T090507Z-00000000.json', 'notes.txt', 'old.json', '20261001T000000Z-zzzzzzzz.json'])
      writeFileSync(join(folder, name), '{}');
    expect(await createRunJournal(folder, createNodeFileSystem()).removeOlderThan(new Date('2026-10-19T09:05:07.000Z'))).toEqual(ok(undefined));
    expect(readdirSync(folder).toSorted((a, b) => a.localeCompare(b))).toEqual(['20261001T000000Z-zzzzzzzz.json', '20261019T090507Z-00000000.json', 'notes.txt', 'old.json']);
  });

  it('removes nothing, with no error, when the folder is not there yet', async () => {
    expect(await createRunJournal(join(scratch(), 'runs'), createNodeFileSystem()).removeOlderThan(new Date())).toEqual(ok(undefined));
  });

  it('says io_failed when the folder cannot be listed or an old run cannot be removed', async () => {
    const blocker = join(scratch(), 'runs');
    writeFileSync(blocker, 'a file where the folder should be');
    const listed = await createRunJournal(blocker, createNodeFileSystem()).removeOlderThan(new Date());
    expect(listed.ok ? '' : listed.error.type).toBe('io_failed');
    const stuck = createRunJournal('runs', {
      ...createNodeFileSystem(),
      list: async () => ok(['20261019T090506Z-00000000.json', '20261019T090506Z-11111111.json']),
      deleteIfExists: async () => err({ type: 'io_failed', message: 'EBUSY: resource busy' }),
    });
    expect(await stuck.removeOlderThan(new Date('2026-11-18T00:00:00.000Z'))).toEqual(err({ type: 'io_failed', message: 'EBUSY: resource busy' }));
  });
});
