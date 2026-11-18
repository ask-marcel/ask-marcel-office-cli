import { z } from 'zod';
import type { GraphError } from '../../domain/graph-error.ts';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { RunError } from '../../domain/run-id.ts';
import type { WriteGraph } from '../ports/graph.ts';
import type { RunJournal } from '../ports/run-journal.ts';
import { readsOf, writeGraphOf } from './command-graph.ts';
import type { RunRecord } from './record-run.ts';
import { loadRun } from './run-record.ts';
import type { GraphReads, ItemIds } from './write-plan.ts';

/*
 * The undo of a recorded run (record-run.ts), generic over the kinds of
 * change. Each kind has a handler: `check` reads the item and says whether it
 * is as the run left it, and `apply` reverses the change. Without `apply` the
 * undo only shows what it would do. An item that changed after the run, or
 * that is gone, is passed over and reported, never reversed.
 */

/** Whether an item is as the run left it (`ready`), changed after the run, or gone. */
export type UndoCheck = 'ready' | 'changed' | 'missing';

export type UndoHandler = {
  /** Reads the item that the undo data names. No write. */
  readonly check: (reads: GraphReads, data: Readonly<Record<string, unknown>>) => Promise<Result<UndoCheck, GraphError>>;
  /**
   * Reverses the change. Called only after `check` answered `ready`. An undo
   * that gives the item new ids (a mail message moved back) answers them as
   * `{ now: { messageId: ... } }`, and the outcome shows them.
   */
  readonly apply: (graph: WriteGraph, data: Readonly<Record<string, unknown>>) => Promise<Result<unknown, GraphError>>;
};

/** The handler of each kind of undo step, by kind (`UndoStep.kind`). */
export type UndoHandlers = Readonly<Record<string, UndoHandler>>;

export type UndoDeps = {
  readonly journal: RunJournal;
  readonly graph: WriteGraph;
  readonly handlers: UndoHandlers;
  readonly now: () => Date;
};

/** What the undo did, or would do, to one change of the run. */
export type UndoStatus = 'will-undo' | 'undone' | 'already-undone' | 'changed' | 'missing' | 'failed';

export type UndoOutcome = {
  readonly action: string;
  readonly target: string;
  readonly after: ItemIds;
  readonly undo: string;
  readonly status: UndoStatus;
  readonly reason?: string;
  /** The ids of the item after the undo, when the undo gave it new ones. */
  readonly now?: ItemIds;
};

export type UndoReport = {
  readonly runId: string;
  readonly command: string;
  readonly startedAt: string;
  readonly applied: boolean;
  readonly changes: ReadonlyArray<UndoOutcome>;
  readonly warning?: string;
};

type Change = RunRecord['changes'][number];

const REASONS: Readonly<Record<'changed' | 'missing', string>> = {
  changed: 'The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.',
  missing: 'The item is no longer there, so there is nothing to undo.',
};

const outcome = (change: Change, status: UndoStatus, reason?: string): UndoOutcome => ({
  action: change.action,
  target: change.target,
  after: change.after,
  undo: change.undo.kind,
  status,
  ...(reason === undefined ? {} : { reason }),
});

const nowSchema = z.object({ now: z.record(z.string(), z.string()) });

// The new ids that the handler of an undo answers, if any.
const nowOf = (value: unknown): { readonly now?: ItemIds } => {
  const parsed = nowSchema.safeParse(value);
  return parsed.success ? { now: parsed.data.now } : {};
};

// Checks one change and, with `apply`, reverses it.
const undoChange = async (deps: UndoDeps, change: Change, apply: boolean): Promise<UndoOutcome> => {
  if (change.undoneAt !== undefined) return outcome(change, 'already-undone');
  const handler = Object.hasOwn(deps.handlers, change.undo.kind) ? deps.handlers[change.undo.kind] : undefined;
  if (handler === undefined) return outcome(change, 'failed', `This version of the CLI has no undo for "${change.undo.kind}".`);
  const checked = await handler.check(readsOf(writeGraphOf(deps.graph)), change.undo.data);
  if (!checked.ok) return outcome(change, 'failed', checked.error.message);
  if (checked.value !== 'ready') return outcome(change, checked.value, REASONS[checked.value]);
  if (!apply) return outcome(change, 'will-undo');
  // The write view: the no-send guard stands in front of each undo too.
  const reversed = await handler.apply(writeGraphOf(deps.graph), change.undo.data);
  return reversed.ok ? { ...outcome(change, 'undone'), ...nowOf(reversed.value) } : outcome(change, 'failed', reversed.error.message);
};

// The changes of a run are reversed last first; the report keeps the order of the run.
const undoChanges = async (deps: UndoDeps, changes: ReadonlyArray<Change>, apply: boolean): Promise<ReadonlyArray<UndoOutcome>> => {
  const outcomes: UndoOutcome[] = [];
  for (const change of changes.toReversed()) outcomes.unshift(await undoChange(deps, change, apply));
  return outcomes;
};

// The record with the time of the undo on each change that this undo reversed.
const markUndone = (record: RunRecord, outcomes: ReadonlyArray<UndoOutcome>, at: string): RunRecord => ({
  ...record,
  changes: record.changes.map((change, index) => (outcomes[index]?.status === 'undone' ? { ...change, undoneAt: at } : change)),
});

/**
 * Shows (`apply` false) or makes (`apply` true) the undo of the run `rawRunId`.
 * After an undo that reversed a change, the journal records it, so that change
 * is never reversed twice.
 */
export const undoRun = async (deps: UndoDeps, rawRunId: string, apply: boolean): Promise<Result<UndoReport, RunError>> => {
  const loaded = await loadRun(deps.journal, deps.now(), rawRunId);
  if (!loaded.ok) return loaded;
  const { runId, record } = loaded.value;
  const outcomes = await undoChanges(deps, record.changes, apply);
  const report: UndoReport = { runId, command: record.command, startedAt: record.startedAt, applied: apply, changes: outcomes };
  if (!outcomes.some((change) => change.status === 'undone')) return ok(report);
  const saved = await deps.journal.save(runId, markUndone(record, outcomes, deps.now().toISOString()));
  return ok(saved.ok ? report : { ...report, warning: `The undo is done, but the run journal could not record it: ${saved.error.message}` });
};
