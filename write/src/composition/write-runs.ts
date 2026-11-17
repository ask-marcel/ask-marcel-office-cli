/*
 * The run journal of the write bin: `~/.ask-marcel/runs`, one file for each
 * run (core/src/infra/run-journal-files.ts). Each write with a plan that is not
 * a dry-run gets a run id and a record there, and removes the runs older than
 * 30 days. The CLI and the MCP server of the write bin both get it.
 */
import { homedir } from 'node:os';
import { createBunFileSystem } from '@ask-marcel/office-core/infra/filesystem-bun.ts';
import { createNodeFileSystem } from '@ask-marcel/office-core/infra/filesystem-node.ts';
import type { JournalDisk } from '@ask-marcel/office-core/infra/run-journal-files.ts';
import { createRunJournal, runJournalFolder } from '@ask-marcel/office-core/infra/run-journal-files.ts';
import type { RunRecorder } from '@ask-marcel/office-core/use-cases/commands/record-run.ts';

export type WriteRunsOptions = {
  /** The home folder; `os.homedir()` when not given. */
  readonly home?: string;
  /** The disk; this runtime's file system when not given. */
  readonly disk?: JournalDisk;
};

// 8 random hex digits: two runs in the same second get two ids.
const randomSuffix = (): string => [...crypto.getRandomValues(new Uint8Array(4))].map((byte) => byte.toString(16).padStart(2, '0')).join('');

const defaultDisk = (): JournalDisk => (typeof globalThis.Bun === 'undefined' ? createNodeFileSystem() : createBunFileSystem());

export const buildWriteRuns = (options: WriteRunsOptions = {}): RunRecorder => ({
  journal: createRunJournal(runJournalFolder(options.home ?? homedir()), options.disk ?? defaultDisk()),
  now: () => new Date(),
  randomSuffix,
});
