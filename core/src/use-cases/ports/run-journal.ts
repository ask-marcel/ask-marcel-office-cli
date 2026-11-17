import type { Result } from '../../domain/result.ts';
import type { RunId } from '../../domain/run-id.ts';

/*
 * The run journal: one record for each write run that is not a dry-run, kept
 * 30 days so that `undo <runId>` can reverse the run. The records hold tenant
 * data (ids, subjects, bodies), so no adapter writes them to a log.
 */

export type JournalIoError = { readonly type: 'io_failed'; readonly message: string };

export type JournalError = { readonly type: 'not_found' } | JournalIoError;

export type RunJournal = {
  /** Writes the record of `runId`, whole; a second save of the same id replaces the first. */
  readonly save: (runId: RunId, record: unknown) => Promise<Result<void, JournalIoError>>;
  /** The record of `runId` as JSON: the caller checks its shape, because a file on disk is outside data. */
  readonly load: (runId: RunId) => Promise<Result<unknown, JournalError>>;
  /** Removes every run that started before `cutoff`. */
  readonly removeOlderThan: (cutoff: Date) => Promise<Result<void, JournalIoError>>;
};
