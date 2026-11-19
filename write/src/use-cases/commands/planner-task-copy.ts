import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { checklistValue } from '../../domain/planner-checklist.ts';
import type { ChecklistItem } from '../../domain/planner-checklist.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { graphIdSchema, isGraphId } from '../../domain/graph-id.ts';
import { ASSIGNMENT, byHint, hintBetween, labelsOn } from '../../domain/planner-fields.ts';
import { isWebUrl, urlOfReferenceKey } from '../../domain/planner-reference.ts';
import { checklistItemOf, recordOf, referenceOf, referenceValue } from '../../domain/planner-snapshot.ts';
import type { PlannerItem } from './planner-reader.ts';
import { readPlannerItem } from './planner-reader.ts';
import { readMe } from './planner-people.ts';
import { patchPlanner } from './planner-patch.ts';
import { unreadable } from './todo-reader.ts';
import { createTodoUndoHandler, notThere } from './todo-undo-handler.ts';

/*
 * The copy of a deleted Planner task, which the run journal keeps, and the
 * undo that makes the task again from it: its fields, the user's assignment,
 * its description, preview, checklist (in its order) and references. Planner
 * gives the task a new id. Lost: the comments, the dates of creation and of
 * completion, the order of the task in its bucket, and the HTML formatting of
 * the notes (the plain text stays). Another person is assigned again only
 * when the delete had `--notify-others true`: Planner sends that person a
 * notification.
 */

const TASK_FIELDS = ['title', 'priority', 'percentComplete', 'dueDateTime', 'startDateTime'] as const;

const taskSchema = z
  .object({
    planId: graphIdSchema,
    bucketId: graphIdSchema.optional(),
    title: z.string().min(1),
    priority: z.number().int().optional(),
    percentComplete: z.number().int().optional(),
    dueDateTime: z.string().optional(),
    startDateTime: z.string().optional(),
    appliedCategories: z.record(z.string(), z.literal(true)).optional(),
  })
  .strict();

const referenceKeySchema = z.string().refine((key) => isWebUrl(urlOfReferenceKey(key)));

const copySchema = z.object({
  task: taskSchema,
  details: z.object({
    description: z.string(),
    previewType: z.string(),
    checklist: z.array(z.object({ title: z.string().min(1), isChecked: z.boolean() })),
    references: z.record(referenceKeySchema, z.object({ alias: z.string().nullable(), type: z.string().nullable() })),
  }),
  assignedMe: z.boolean(),
  others: z.array(z.string().min(1)),
  notifyOthers: z.boolean(),
});

type Copy = z.infer<typeof copySchema>;

const entriesOf = (value: unknown): ReadonlyArray<readonly [string, unknown]> => Object.entries(recordOf(value));

const checklistOf = (details: PlannerItem): ReadonlyArray<ChecklistItem> =>
  entriesOf(details.fields['checklist'])
    .map(([, value]) => ({ item: checklistItemOf(value), hint: String(recordOf(value)['orderHint'] ?? '') }))
    .filter((entry): entry is { item: ChecklistItem; hint: string } => entry.item !== undefined)
    .toSorted((a, b) => byHint(a.hint, b.hint))
    .map((entry) => entry.item);

const referencesOf = (details: PlannerItem): Copy['details']['references'] =>
  Object.fromEntries(
    entriesOf(details.fields['references']).flatMap(([key, value]) => {
      const reference = referenceOf(value);
      return reference === undefined ? [] : [[key, reference] as const];
    })
  );

const taskFieldsOf = (task: PlannerItem, planId: GraphId): Record<string, unknown> => ({
  planId,
  ...(isGraphId(String(task.fields['bucketId'] ?? '')) ? { bucketId: task.fields['bucketId'] } : {}),
  ...Object.fromEntries(TASK_FIELDS.filter((key) => task.fields[key] !== undefined && task.fields[key] !== null).map((key) => [key, task.fields[key]])),
  appliedCategories: Object.fromEntries(labelsOn((task.fields['appliedCategories'] ?? {}) as Record<string, boolean>).map((key) => [key, true])),
});

/** The copy of the task and its details that the undo of its delete makes again. `meId` is the user; `notifyOthers` the consent of the delete. */
export const copyOf = (task: PlannerItem, details: PlannerItem, planId: GraphId, meId: string, notifyOthers: boolean): Copy => {
  const assignees = entriesOf(task.fields['assignments'])
    .filter(([, value]) => value !== null)
    .map(([id]) => id);
  return {
    task: taskFieldsOf(task, planId) as Copy['task'],
    details: {
      description: typeof details.fields['description'] === 'string' ? details.fields['description'] : '',
      previewType: typeof details.fields['previewType'] === 'string' ? details.fields['previewType'] : 'automatic',
      checklist: [...checklistOf(details)],
      references: referencesOf(details),
    },
    assignedMe: assignees.includes(meId),
    others: assignees.filter((id) => id !== meId),
    notifyOthers,
  };
};

const hasDetails = (details: Copy['details']): boolean =>
  details.description !== '' || details.previewType !== 'automatic' || details.checklist.length > 0 || Object.keys(details.references).length > 0;

const detailsBody = (details: Copy['details']): Record<string, unknown> => {
  let hint: string | undefined;
  const checklist = details.checklist.map((item) => {
    hint = hintBetween(hint, undefined);
    return [crypto.randomUUID(), { ...checklistValue(item), orderHint: hint }] as const;
  });
  return {
    description: details.description,
    previewType: details.previewType,
    checklist: Object.fromEntries(checklist),
    references: Object.fromEntries(Object.entries(details.references).map(([key, value]) => [key, referenceValue(value)])),
  };
};

// Planner makes the details of a new task a moment after the task: up to three reads.
const READS_OF_NEW_DETAILS = 3;

const readNewDetails = async (graph: WriteGraph, id: GraphId): Promise<Result<PlannerItem, GraphError>> => {
  let read = await readPlannerItem(graph, 'details', id);
  for (let count = 1; count < READS_OF_NEW_DETAILS && !read.ok; count += 1) read = await readPlannerItem(graph, 'details', id);
  return read;
};

const writeDetails = async (graph: WriteGraph, id: GraphId, details: Copy['details']): Promise<Record<string, string>> => {
  if (!hasDetails(details)) return {};
  const read = await readNewDetails(graph, id);
  const written = read.ok ? await patchPlanner(graph, 'details', id, detailsBody(details), read.value.etag) : read;
  return written.ok ? {} : { details: `not written: ${written.error.message}` };
};

const assignmentsOf = (copy: Copy, meId: string): Record<string, unknown> => ({
  ...(copy.assignedMe ? { [meId]: ASSIGNMENT } : {}),
  ...(copy.notifyOthers ? Object.fromEntries(copy.others.map((id) => [id, ASSIGNMENT])) : {}),
});

const recreate = async (graph: WriteGraph, copy: Copy): Promise<Result<unknown, GraphError>> => {
  const me = await readMe(graph);
  if (!me.ok) return me;
  const created = await graph.post('/planner/tasks', { ...copy.task, assignments: assignmentsOf(copy, me.value.id) });
  if (!created.ok) return created;
  const id = String(recordOf(created.value)['id'] ?? '');
  if (!isGraphId(id)) return err(unreadable('the Planner task it made again'));
  return ok({ now: { taskId: id, ...(await writeDetails(graph, id, copy.details)) } });
};

/** The handler of the undo of a deleted task. */
export const TASK_COPY_UNDO_HANDLERS: UndoHandlers = {
  'recreate-planner-task': createTodoUndoHandler(
    copySchema,
    async (reads, data) => {
      const plan = await readPlannerItem(reads, 'plan', data.task.planId);
      return plan.ok ? ok('ready') : notThere(plan.error);
    },
    recreate
  ),
};
