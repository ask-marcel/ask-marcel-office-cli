import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { RunError, RunId } from '../../domain/run-id.ts';
import { parseRunId, runIdTime } from '../../domain/run-id.ts';
import type { JournalError, RunJournal } from '../ports/run-journal.ts';
import { RUN_RETENTION_DAYS, retentionCutoff } from './record-run.ts';
import type { RunRecord } from './record-run.ts';

/*
 * A run record as the undo reads it back. The file on disk is outside data (a
 * person can edit it), so its shape is checked here before any undo trusts it.
 */

const idsSchema = z.record(z.string(), z.string());

const recordSchema = z.object({
  version: z.literal(1),
  runId: z.string(),
  command: z.string(),
  startedAt: z.string(),
  changes: z
    .array(
      z.object({
        action: z.string(),
        target: z.string(),
        before: idsSchema,
        after: idsSchema,
        undo: z.object({ kind: z.string(), data: z.record(z.string(), z.unknown()) }),
        undoneAt: z.string().optional(),
      })
    )
    .min(1),
});

/** A run that the undo can work on: its id and its record. */
export type LoadedRun = { readonly runId: RunId; readonly record: RunRecord };

const runError = (code: RunError['code'], message: string): Result<never, RunError> => err({ code, message });

const notLoaded = (runId: RunId, error: JournalError): Result<never, RunError> =>
  error.type === 'not_found'
    ? runError('run_not_found', `The run journal has no run ${runId}. Copy the runId field of the write output: a dry-run has no run id.`)
    : runError('journal_failed', `The run journal could not be read: ${error.message}`);

const readRecord = (runId: RunId, raw: unknown): Result<RunRecord, RunError> => {
  const parsed = recordSchema.safeParse(raw);
  if (!parsed.success || parsed.data.runId !== runId) return runError('run_unreadable', `The journal record of run ${runId} is not readable, so the run cannot be undone.`);
  return ok(parsed.data);
};

/** Reads the run `raw` names, if it can still be undone: a good id, younger than 30 days, in the journal, not fully undone. */
export const loadRun = async (journal: RunJournal, now: Date, raw: string): Promise<Result<LoadedRun, RunError>> => {
  const parsed = parseRunId(raw);
  if (!parsed.ok) return parsed;
  const runId = parsed.value;
  if (runIdTime(runId) < retentionCutoff(now))
    return runError('run_expired', `Run ${runId} is older than ${RUN_RETENTION_DAYS} days: the journal keeps a run ${RUN_RETENTION_DAYS} days only, so it cannot be undone.`);
  const loaded = await journal.load(runId);
  if (!loaded.ok) return notLoaded(runId, loaded.error);
  const record = readRecord(runId, loaded.value);
  if (!record.ok) return record;
  if (record.value.changes.every((change) => change.undoneAt !== undefined))
    return runError('run_already_undone', `Run ${runId} is already fully undone: there is nothing left to undo.`);
  return ok({ runId, record: record.value });
};
