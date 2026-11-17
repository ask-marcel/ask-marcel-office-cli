import type { GraphError } from '../../domain/graph-error.ts';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import { newRunId } from '../../domain/run-id.ts';
import type { RunJournal } from '../ports/run-journal.ts';
import type { AppliedChange, BoundPlan } from './write-plan.ts';

/*
 * The dry-run and the recorded run of a write with a plan (write-plan.ts). A
 * dry-run shows the changes and calls no write. A run makes the changes, gives
 * them a run id, saves the record that the undo reads (undo-run.ts), and
 * removes the runs older than 30 days.
 */

/** What a recorded run needs: the journal, the clock, and 8 random hex digits for each run id. */
export type RunRecorder = {
  readonly journal: RunJournal;
  readonly now: () => Date;
  readonly randomSuffix: () => string;
};

/** How long the journal keeps a run, so that `undo <runId>` can reverse it. */
export const RUN_RETENTION_DAYS = 30;

const DAY_MS = 86_400_000;

/** The time before which a run is too old to keep or to undo. */
export const retentionCutoff = (now: Date): Date => new Date(now.getTime() - RUN_RETENTION_DAYS * DAY_MS);

/** The record format that the journal keeps (version 1). */
export type RunRecord = {
  readonly version: 1;
  readonly runId: string;
  readonly command: string;
  readonly startedAt: string;
  readonly changes: ReadonlyArray<AppliedChange & { readonly undoneAt?: string }>;
};

export const noDryRun = (command: string): GraphError => ({
  type: 'validation_error',
  code: 'dry_run_unsupported',
  message: `--dry-run: ${command} has no dry-run and no undo. Only the commands whose docs show "Dry-run and undo" have them.`,
});

/** The answer of a dry-run: the changes the command would make. */
export const previewPlan = (command: string, planned: Result<BoundPlan, GraphError>): Result<unknown, GraphError> =>
  planned.ok ? ok({ dryRun: true, command, changes: planned.value.changes }) : planned;

const asObject = (value: unknown): object => (typeof value === 'object' && value !== null && !Array.isArray(value) ? value : { value });

// The changes as the output shows them: the undo data stays in the journal.
const shown = (changes: ReadonlyArray<AppliedChange>): ReadonlyArray<Omit<AppliedChange, 'undo'>> =>
  changes.map(({ action, target, before, after }) => ({ action, target, before, after }));

/** Makes the changes of the plan and records the run; the answer carries the run id and the ids before and after. */
export const recordRun = async (recorder: RunRecorder, command: string, planned: Result<BoundPlan, GraphError>): Promise<Result<unknown, GraphError>> => {
  if (!planned.ok) return planned;
  const startedAt = recorder.now();
  const applied = await planned.value.apply();
  if (!applied.ok) return applied;
  const { value, changes } = applied.value;
  if (changes.length === 0) return ok({ ...asObject(value), changes: [] });
  const runId = newRunId(startedAt, recorder.randomSuffix());
  const record: RunRecord = { version: 1, runId, command, startedAt: startedAt.toISOString(), changes };
  const saved = await recorder.journal.save(runId, record);
  await recorder.journal.removeOlderThan(retentionCutoff(startedAt));
  if (saved.ok) return ok({ ...asObject(value), runId, changes: shown(changes) });
  return ok({
    ...asObject(value),
    changes: shown(changes),
    warning: `The change is done, but the run journal could not be written, so this run has no undo: ${saved.error.message}`,
  });
};
