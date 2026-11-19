import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads, ItemIds, PlannedChange } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { recordOf, restoreOf, snapshotOf } from '../../domain/planner-snapshot.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import type { PlannerItem, PlannerResource } from './planner-reader.ts';
import { plannerPath, readPlannerItem } from './planner-reader.ts';

/*
 * The change of a Planner item. The plan reads the item and keeps its ETag;
 * just before the write the op reads it again and passes over an item whose
 * ETag changed. The write sends the ETag as `If-Match`, so Planner refuses a
 * change made in between (412, or 409 as a live run of 2026-11-19 got): the
 * error says so, and nothing changed. The
 * write asks for the item as Planner keeps it (`return=representation`), and
 * the run journal keeps the values that it gives for the fields of the
 * change: the undo check compares them with a later read, so a read just
 * after the write, which Planner can answer with old data, is never needed.
 */

/** The message of a change that Planner refused because the item changed after the read (412 or 409). */
export const CHANGED_SINCE_READ =
  'Planner refused the change: the item changed after this command read it (its ETag is not the same). Nothing changed. Run the command again to read the item as it is now.';

// Planner answers an old ETag with 412 (the docs) or 409 (live, 2026-11-19).
const CONFLICTS: ReadonlySet<number> = new Set([409, 412]);

/** The error of a write: a 412 or a 409 of Planner in the words of `CHANGED_SINCE_READ`, else as Graph gave it. */
export const conflictOf = (error: GraphError): GraphError =>
  error.type === 'api_error' && CONFLICTS.has(error.status) ? { ...error, code: 'planner_item_changed', message: CHANGED_SINCE_READ } : error;

// The HTML notes of a task exist on Graph beta only (decision D7): the one PATCH that goes there.
const isBeta = (body: Readonly<Record<string, unknown>>): boolean => Object.hasOwn(body, 'notes');

/** PATCH of the item with the ETag of its read; the answer is the item as Planner keeps it. */
export const patchPlanner = async (
  graph: WriteGraph,
  resource: PlannerResource,
  id: GraphId,
  body: Readonly<Record<string, unknown>>,
  etag: string
): Promise<Result<unknown, GraphError>> => {
  const written = await graph.patch(plannerPath(resource, id), body, { ifMatch: etag, representation: true, ...(isBeta(body) ? { api: 'beta' } : {}) });
  return written.ok ? written : err(conflictOf(written.error));
};

/** DELETE of the item with the ETag of its read. */
export const deletePlanner = async (graph: WriteGraph, resource: PlannerResource, id: GraphId, etag: string): Promise<Result<unknown, GraphError>> => {
  const deleted = await graph.delete(plannerPath(resource, id), { ifMatch: etag });
  return deleted.ok ? deleted : err(conflictOf(deleted.error));
};

/** The recheck of an op: whether the item still has the ETag of the plan. */
export const sameItem = async (reads: GraphReads, resource: PlannerResource, item: PlannerItem): Promise<Result<boolean, GraphError>> => {
  const now = await readPlannerItem(reads, resource, item.id);
  return now.ok ? ok(now.value.etag === item.etag) : now;
};

const isItem = (value: unknown): boolean => Object.hasOwn(recordOf(value), '@odata.etag');

/** The fields of the item after a write: the answer of Graph, or a read when the answer has no item (the undo then compares with what the read gave). */
export const fieldsAfter = async (reads: GraphReads, resource: PlannerResource, id: GraphId, answer: unknown): Promise<Readonly<Record<string, unknown>>> => {
  if (isItem(answer)) return recordOf(answer);
  const read = await readPlannerItem(reads, resource, id);
  return read.ok ? read.value.fields : {};
};

/** The ids of `before` (the keys that end in `Id`), for `after`: the other values of `before` are the values of before the change. */
const idsOf = (before: ItemIds): ItemIds => Object.fromEntries(Object.entries(before).filter(([key]) => key.endsWith('Id')));

/** What a change of an item says and does. `restore` replaces the PATCH of the undo that `restoreOf` gives, for a field that needs it. */
export type PatchSpec = {
  readonly action: string;
  readonly resource: PlannerResource;
  readonly item: PlannerItem;
  readonly body: Readonly<Record<string, unknown>>;
  readonly before: ItemIds;
  readonly restore?: Readonly<Record<string, unknown>>;
};

/** The op that changes the fields of `body` of the item, and records how to write the values of before back. */
export const patchOp = (spec: PatchSpec): TodoOp => {
  const change: PlannedChange = { action: spec.action, target: plannerPath(spec.resource, spec.item.id), before: spec.before };
  return {
    change,
    recheck: async (reads) => sameItem(reads, spec.resource, spec.item),
    write: async (graph) => {
      const written = await patchPlanner(graph, spec.resource, spec.item.id, spec.body, spec.item.etag);
      if (!written.ok) return outcomeOf(written);
      const fields = await fieldsAfter(graph, spec.resource, spec.item.id, written.value);
      const restore = { ...restoreOf(spec.item.fields, spec.body), ...spec.restore };
      const data = { resource: spec.resource, id: spec.item.id, set: snapshotOf(fields, spec.body), restore };
      return outcomeOf(ok({ ...change, after: { ...idsOf(spec.before), fields: Object.keys(spec.body).join(', ') }, undo: { kind: 'restore-planner-fields', data } }));
    },
  };
};
