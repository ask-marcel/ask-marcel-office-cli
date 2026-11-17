import type { Result } from './result.ts';
import { err, ok } from './result.ts';

/**
 * The id of one write run, for example `20261118T090507Z-0a1b2c3d`: the UTC
 * second the run started, then 8 random hex digits. The time comes first, so
 * the ids sort by time and the 30-day limit reads the age from the id alone.
 *
 * The id is a trust boundary (hard rule 12): a user types it after `undo`, and
 * the journal makes it a file name. Only the exact form passes, so no path can
 * reach the disk through it.
 */
export type RunId = string & { readonly __brand: 'RunId' };

/** Why a run cannot be read, recorded or undone. `code` is the machine-readable part. */
export type RunError = {
  readonly code: 'invalid_run_id' | 'run_not_found' | 'run_expired' | 'run_already_undone' | 'run_unreadable' | 'journal_failed' | 'dry_run_unsupported';
  readonly message: string;
};

const RUN_ID_PATTERN = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z-[0-9a-f]{8}$/;

const INVALID_RUN_ID: RunError = {
  code: 'invalid_run_id',
  message: 'A run id has the form 20261118T090507Z-0a1b2c3d: the UTC time of the run, then 8 hex digits. Copy the runId field of the write output.',
};

export const parseRunId = (raw: string): Result<RunId, RunError> => (RUN_ID_PATTERN.test(raw) ? ok(raw as RunId) : err(INVALID_RUN_ID));

/** A new run id for a run that started at `startedAt`; `suffix` is 8 random hex digits. */
export const newRunId = (startedAt: Date, suffix: string): RunId => {
  const stamp = startedAt
    .toISOString()
    .replace(/\.\d{3}Z$/, 'Z')
    .replaceAll(/[-:]/g, '');
  return `${stamp}-${suffix}` as RunId;
};

/** The second the run started, read from its id. */
export const runIdTime = (runId: RunId): Date => {
  const [, year, month, day, hour, minute, second] = RUN_ID_PATTERN.exec(runId) ?? [];
  return new Date(`${year}-${month}-${day}T${hour}:${minute}:${second}Z`);
};
