import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { AppliedChange, GraphReads, PlannedChange, SkippedItem, WritePlan } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { isNotThere } from './mail-folder-reader.ts';

/*
 * The plan and the writes of a To Do command: one op or more (the ops file).
 * Each op says its change for the dry-run. Before its write, an op on an item
 * that exists reads it again and is passed over if it changed or is gone since
 * the plan. An op can make more than one change (a task and its steps). An op
 * that fails is reported, and the next ops go on: the changes made, a part of
 * an op too, stay in the run journal, so the undo can reverse them.
 */

/** What the write of an op made: its changes, and the error that stopped it, if any. */
export type OpOutcome = { readonly changes: ReadonlyArray<AppliedChange>; readonly error?: GraphError };

export type TodoOp = {
  readonly change: PlannedChange;
  /** Reads the item again: `true` if it is as the plan read it. Absent for an op that only creates. */
  readonly recheck?: (reads: GraphReads) => Promise<Result<boolean, GraphError>>;
  readonly write: (graph: WriteGraph) => Promise<OpOutcome>;
};

export type TodoBatch = { readonly ops: ReadonlyArray<TodoOp>; readonly skipped?: ReadonlyArray<SkippedItem> };

export const GONE = 'The item is no longer there: it was moved or deleted after the plan.';
export const CHANGED = 'The item changed after the plan, so this run passes over it. Run the command again to see it as it is now.';

type Outcome = { readonly changes: AppliedChange[]; readonly skipped: SkippedItem[]; readonly failed: SkippedItem[]; done: number; firstError?: GraphError };

const itemOf = (change: PlannedChange, reason: string): SkippedItem => ({ target: change.target, before: change.before, reason });

// Why the write passes over the op, or undefined to write it. A read error that is not a 404 stops the op.
const passOver = async (graph: WriteGraph, op: TodoOp): Promise<Result<string | undefined, GraphError>> => {
  if (op.recheck === undefined) return ok(undefined);
  const same = await op.recheck(graph);
  if (!same.ok) return isNotThere(same.error) ? ok(GONE) : same;
  return ok(same.value ? undefined : CHANGED);
};

const fail = (outcome: Outcome, op: TodoOp, error: GraphError): void => {
  outcome.firstError ??= error;
  outcome.failed.push(itemOf(op.change, error.message));
};

const applyOne = async (graph: WriteGraph, op: TodoOp, outcome: Outcome): Promise<void> => {
  const reason = await passOver(graph, op);
  if (!reason.ok) return fail(outcome, op, reason.error);
  if (reason.value !== undefined) {
    outcome.skipped.push(itemOf(op.change, reason.value));
    return;
  }
  const written = await op.write(graph);
  outcome.changes.push(...written.changes);
  if (written.error === undefined) outcome.done += 1;
  else fail(outcome, op, written.error);
};

const applyBatch = async (graph: WriteGraph, batch: TodoBatch): Promise<ReturnType<WritePlan['apply']>> => {
  const outcome: Outcome = { changes: [], skipped: [...(batch.skipped ?? [])], failed: [], done: 0 };
  for (const op of batch.ops) await applyOne(graph, op, outcome);
  // A run that changed nothing and failed is an error; with one change made, the run is recorded so its undo can run.
  if (outcome.changes.length === 0 && outcome.firstError !== undefined) return err(outcome.firstError);
  return ok({ value: { done: outcome.done, skipped: outcome.skipped, failed: outcome.failed }, changes: outcome.changes });
};

/** The plan of the ops of the batch. */
export const todoPlan = (batch: TodoBatch): WritePlan => ({
  changes: batch.ops.map((op) => op.change),
  ...(batch.skipped === undefined || batch.skipped.length === 0 ? {} : { skipped: batch.skipped }),
  apply: async (graph) => applyBatch(graph, batch),
});

/** The outcome of a write that made the change, or of one that failed before any change. */
export const outcomeOf = (written: Result<AppliedChange, GraphError>): OpOutcome => (written.ok ? { changes: [written.value] } : { changes: [], error: written.error });
