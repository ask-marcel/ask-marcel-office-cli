import { join } from 'node:path';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import { parseRunId, runIdTime } from '../domain/run-id.ts';
import type { AtomicFileWrites, FileSystem, FileSystemError, FolderListing } from '../use-cases/ports/filesystem.ts';
import type { JournalIoError, RunJournal } from '../use-cases/ports/run-journal.ts';

/*
 * The run journal on disk: one JSON file for each run, `<runId>.json`, in one
 * folder (`~/.ask-marcel/runs` in production). One file per run, so two runs at
 * the same time never write the same file. A file is written whole (a temp file,
 * then a rename) with mode 0600: it holds tenant data. Nothing here logs.
 */

/** The folder of the run journal under a home folder. */
export const runJournalFolder = (home: string): string => join(home, '.ask-marcel', 'runs');

/** What the journal needs of the disk. */
export type JournalDisk = Pick<FileSystem, 'readJson' | 'deleteIfExists'> & Pick<AtomicFileWrites, 'writeTextAtomic'> & FolderListing;

const RUN_FILE = /^(.+)\.json$/;

const ioFailed = (error: FileSystemError): JournalIoError => ({ type: 'io_failed', message: error.type === 'not_found' ? 'not found' : error.message });

// The run id a file name holds, when the name is that of a run file.
const runIdOf = (name: string): ReturnType<typeof parseRunId> | undefined => {
  const stem = RUN_FILE.exec(name)?.[1];
  return stem === undefined ? undefined : parseRunId(stem);
};

const removeOld = async (disk: JournalDisk, folder: string, names: ReadonlyArray<string>, cutoff: Date): Promise<Result<void, JournalIoError>> => {
  let failure: JournalIoError | undefined;
  for (const name of names) {
    const runId = runIdOf(name);
    if (runId === undefined || !runId.ok || runIdTime(runId.value) >= cutoff) continue;
    const removed = await disk.deleteIfExists(join(folder, name));
    if (!removed.ok) failure ??= ioFailed(removed.error);
  }
  return failure === undefined ? ok(undefined) : err(failure);
};

export const createRunJournal = (folder: string, disk: JournalDisk): RunJournal => ({
  save: async (runId, record) => {
    const written = await disk.writeTextAtomic(join(folder, `${runId}.json`), `${JSON.stringify(record, null, 2)}\n`, 0o600);
    return written.ok ? written : err(ioFailed(written.error));
  },
  load: async (runId) => {
    const read = await disk.readJson<unknown>(join(folder, `${runId}.json`));
    if (read.ok) return read;
    if (read.error.type === 'not_found') return err({ type: 'not_found' });
    // A file that is not JSON is a record the undo cannot read: its check refuses null.
    return read.error.type === 'parse_failed' ? ok(null) : err(ioFailed(read.error));
  },
  removeOlderThan: async (cutoff) => {
    const listed = await disk.list(folder);
    return listed.ok ? removeOld(disk, folder, listed.value, cutoff) : err(ioFailed(listed.error));
  },
});
