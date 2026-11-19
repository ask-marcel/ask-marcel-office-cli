import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoCheck, UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { graphIdSchema } from '../../domain/graph-id.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { hintBetween } from '../../domain/planner-fields.ts';
import { snapshotOf, withNewKeys } from '../../domain/planner-snapshot.ts';
import type { PlannerItem, PlannerResource } from './planner-reader.ts';
import { countBucketTasks, idField, readBuckets, readPlannerItem } from './planner-reader.ts';
import { deletePlanner, patchPlanner } from './planner-patch.ts';
import { canonical, createTodoUndoHandler, notThere } from './todo-undo-handler.ts';

/*
 * The undo of the Planner changes. A check reads the item and compares the
 * fields and entries that the run set with the values that the write gave:
 * an item that changed after the run is `changed`, and the undo leaves it.
 * The undo writes with the ETag of its own read. The undo of a change of the
 * order of a bucket puts it back between the buckets that were next to it.
 * A created bucket is deleted only when it has no task.
 */

const resourceSchema = z.enum(['task', 'details', 'bucket', 'plan', 'planDetails']);
const valuesSchema = z.record(z.string(), z.unknown());
const restoreSchema = z.object({ resource: resourceSchema, id: graphIdSchema, set: valuesSchema, restore: valuesSchema });
const createdSchema = z.object({ resource: z.enum(['task', 'bucket']), id: graphIdSchema, set: valuesSchema });
const neighboursSchema = z.object({ after: graphIdSchema.optional(), before: graphIdSchema.optional() });

/** Whether the item has the values that the run set. */
const isAsSet = (item: PlannerItem, set: Readonly<Record<string, unknown>>): boolean => canonical(snapshotOf(item.fields, set)) === canonical(set);

const checkItem = async (reads: GraphReads, resource: PlannerResource, id: GraphId, set: Readonly<Record<string, unknown>>): Promise<Result<UndoCheck, GraphError>> => {
  const read = await readPlannerItem(reads, resource, id);
  if (!read.ok) return notThere(read.error);
  return ok(isAsSet(read.value, set) ? 'ready' : 'changed');
};

/** The order hint between the buckets `after` and `before` as they are now; an end of the plan for a bucket that is gone. */
export const hintOfPlace = async (reads: GraphReads, planId: GraphId, place: z.infer<typeof neighboursSchema>): Promise<Result<string, GraphError>> => {
  const buckets = await readBuckets(reads, planId);
  if (!buckets.ok) return buckets;
  const hintOf = (id: GraphId | undefined): string | undefined => buckets.value.find((bucket) => bucket.id === id)?.orderHint;
  return ok(hintBetween(hintOf(place.after), hintOf(place.before)));
};

// The PATCH of the undo, with what it can know only now: a new key for a checklist item that comes back, the order hint between the old neighbours.
const restoreBody = async (reads: GraphReads, now: PlannerItem, restore: Readonly<Record<string, unknown>>): Promise<Result<Record<string, unknown>, GraphError>> => {
  const body: Record<string, unknown> = { ...restore };
  const checklist = valuesSchema.safeParse(restore['checklist']);
  if (checklist.success) body['checklist'] = withNewKeys(checklist.data, valuesSchema.catch({}).parse(now.fields['checklist']), () => crypto.randomUUID());
  const place = neighboursSchema.safeParse(restore['orderHint']);
  const planId = idField(now.fields, 'planId');
  if (!place.success || planId === undefined) return ok(body);
  const hint = await hintOfPlace(reads, planId, place.data);
  return hint.ok ? ok({ ...body, orderHint: hint.value }) : hint;
};

const applyRestore = async (graph: WriteGraph, data: z.infer<typeof restoreSchema>): Promise<Result<unknown, GraphError>> => {
  const now = await readPlannerItem(graph, data.resource, data.id);
  if (!now.ok) return now;
  const body = await restoreBody(graph, now.value, data.restore);
  return body.ok ? patchPlanner(graph, data.resource, data.id, body.value, now.value.etag) : body;
};

// A created bucket is deleted only when it is as the run made it and has no task.
const checkCreated = async (reads: GraphReads, data: z.infer<typeof createdSchema>): Promise<Result<UndoCheck, GraphError>> => {
  const check = await checkItem(reads, data.resource, data.id, data.set);
  if (!check.ok || check.value !== 'ready' || data.resource === 'task') return check;
  const tasks = await countBucketTasks(reads, data.id);
  if (!tasks.ok) return notThere(tasks.error);
  return ok(tasks.value === 0 ? 'ready' : 'changed');
};

const applyCreated = async (graph: WriteGraph, data: z.infer<typeof createdSchema>): Promise<Result<unknown, GraphError>> => {
  const now = await readPlannerItem(graph, data.resource, data.id);
  return now.ok ? deletePlanner(graph, data.resource, data.id, now.value.etag) : now;
};

const bucketSchema = z.object({ planId: graphIdSchema, name: z.string().min(1), place: neighboursSchema });
const keptSchema = z.object({ bucketId: graphIdSchema, tasks: z.number().int().positive() });

const noBucketUndo = (tasks: number): Result<never, GraphError> =>
  err({
    type: 'validation_error',
    code: 'planner_bucket_no_undo',
    message: `The bucket had ${tasks} task(s) when the run deleted it. This CLI cannot make a bucket and its tasks again, so this change has no undo. Look for the tasks in Planner: they can be gone, or have no bucket.`,
  });

const applyBucket = async (graph: WriteGraph, data: z.infer<typeof bucketSchema>): Promise<Result<unknown, GraphError>> => {
  const hint = await hintOfPlace(graph, data.planId, data.place);
  if (!hint.ok) return hint;
  const created = await graph.post('/planner/buckets', { name: data.name, planId: data.planId, orderHint: hint.value });
  if (!created.ok) return created;
  const id = idField(valuesSchema.catch({}).parse(created.value), 'id');
  return id === undefined ? ok({}) : ok({ now: { bucketId: id } });
};

/** The handler of each kind of undo of the Planner changes, the task copies and the files aside. */
export const PLANNER_UNDO_HANDLERS: UndoHandlers = {
  'restore-planner-fields': createTodoUndoHandler(restoreSchema, async (reads, data) => checkItem(reads, data.resource, data.id, data.set), applyRestore),
  'delete-created-planner-item': createTodoUndoHandler(createdSchema, checkCreated, applyCreated),
  'recreate-planner-bucket': createTodoUndoHandler(
    bucketSchema,
    async (reads, data) => {
      const plan = await readPlannerItem(reads, 'plan', data.planId);
      return plan.ok ? ok('ready') : notThere(plan.error);
    },
    applyBucket
  ),
  'planner-bucket-deleted-with-tasks': createTodoUndoHandler(
    keptSchema,
    async (_reads, data) => noBucketUndo(data.tasks),
    async (_graph, data) => noBucketUndo(data.tasks)
  ),
};
