import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { graphIdSchema, isGraphId } from '../../domain/graph-id.ts';
import type { TodoTask } from './todo-reader.ts';
import { readTask, taskPath, unreadable } from './todo-reader.ts';

/*
 * The copy of a To Do task: its fields, its steps, its linked resources and its
 * file attachments with their bytes. Graph has no move of a task, so a move
 * creates the copy in the other list and then deletes the original; the undo
 * of a delete creates the task again from the copy that the run journal keeps.
 * A copy is not exact: the task, its steps and its attachments get new ids,
 * and the dates of creation and of change are those of the copy.
 */

/** The fields of a task that a copy keeps: the fields that Graph takes in a create. */
export const COPIED_FIELDS = [
  'title',
  'body',
  'importance',
  'status',
  'categories',
  'dueDateTime',
  'startDateTime',
  'reminderDateTime',
  'isReminderOn',
  'recurrence',
  'completedDateTime',
] as const;

const stepSchema = z.object({ displayName: z.string(), isChecked: z.boolean() });
const linkSchema = z.object({ webUrl: z.string(), applicationName: z.string(), displayName: z.string(), externalId: z.string() }).partial();
const fileSchema = z.object({ name: z.string(), contentType: z.string().optional(), contentBytes: z.string() });

/** The copy as the run journal keeps it. */
export const taskCopySchema = z.object({
  fields: z.record(z.string(), z.unknown()),
  steps: z.array(stepSchema),
  links: z.array(linkSchema),
  files: z.array(fileSchema),
});

export type TaskCopy = z.infer<typeof taskCopySchema>;

/** The items under a task, as a dry-run counts them. */
export type TaskChildren = { readonly steps: number; readonly links: number; readonly files: ReadonlyArray<GraphId> };

const pageOf = <T extends z.ZodType>(item: T): z.ZodObject<{ value: z.ZodArray<T>; '@odata.nextLink': z.ZodOptional<z.ZodString> }> =>
  z.object({ value: z.array(item), '@odata.nextLink': z.string().optional() });
const readStepSchema = z.object({ displayName: z.string(), isChecked: z.boolean().optional() });
const attachmentSchema = z.object({ id: z.string() });
const oneFileSchema = z.union([fileSchema, z.object({ value: fileSchema }).transform((answer) => answer.value)]);

const tooMany = (what: string): GraphError => ({
  type: 'validation_error',
  code: 'todo_copy_too_large',
  message: `The task has more ${what} than Graph gives in one page, so this CLI does not copy it: the copy would not be complete.`,
});

const readPage = async <T extends z.ZodType>(reads: GraphReads, path: string, what: string, item: T): Promise<Result<ReadonlyArray<z.infer<T>>, GraphError>> => {
  const read = await reads.get(path);
  if (!read.ok) return read;
  const page = pageOf(item).safeParse(read.value);
  if (!page.success) return err(unreadable(`the ${what} of a To Do task`));
  return page.data['@odata.nextLink'] === undefined ? ok(page.data.value) : err(tooMany(what));
};

const presentFields = (fields: Readonly<Record<string, unknown>>, keys: ReadonlyArray<string>): Record<string, unknown> =>
  Object.fromEntries(keys.flatMap((key) => (fields[key] === undefined || fields[key] === null ? [] : [[key, fields[key]]])));

const fileIds = (items: ReadonlyArray<{ readonly id: string }>): Result<ReadonlyArray<GraphId>, GraphError> => {
  const ids = items.map((item) => item.id);
  return ids.every(isGraphId) ? ok(ids) : err(unreadable('an attachment of a To Do task'));
};

/** The steps, the linked resources and the attachment ids of the task. */
export const readChildren = async (reads: GraphReads, task: TodoTask): Promise<Result<TaskChildren & { readonly copy: Omit<TaskCopy, 'files'> }, GraphError>> => {
  const path = taskPath(task.listId, task.id);
  const steps = await readPage(reads, `${path}/checklistItems`, 'steps', readStepSchema);
  if (!steps.ok) return steps;
  const links = await readPage(reads, `${path}/linkedResources`, 'linked resources', z.record(z.string(), z.unknown()));
  if (!links.ok) return links;
  const files = await readPage(reads, `${path}/attachments`, 'attachments', attachmentSchema);
  if (!files.ok) return files;
  const ids = fileIds(files.value);
  if (!ids.ok) return ids;
  const copy = {
    fields: presentFields(task.fields, COPIED_FIELDS),
    steps: steps.value.map((step) => ({ displayName: step.displayName, isChecked: step.isChecked === true })),
    links: links.value.map((link) => presentFields(link, ['webUrl', 'applicationName', 'displayName', 'externalId'])),
  };
  return ok({ steps: copy.steps.length, links: copy.links.length, files: ids.value, copy });
};

const readFile = async (reads: GraphReads, task: TodoTask, id: GraphId): Promise<Result<TaskCopy['files'][number], GraphError>> => {
  const read = await reads.get(`${taskPath(task.listId, task.id)}/attachments/${id}`);
  if (!read.ok) return read;
  const file = oneFileSchema.safeParse(read.value);
  return file.success ? ok(file.data) : err(unreadable('an attachment of a To Do task'));
};

/** The full copy of the task, the bytes of its attachments too. */
export const readCopy = async (reads: GraphReads, task: TodoTask): Promise<Result<TaskCopy, GraphError>> => {
  const children = await readChildren(reads, task);
  if (!children.ok) return children;
  const files: TaskCopy['files'] = [];
  for (const id of children.value.files) {
    const file = await readFile(reads, task, id);
    if (!file.ok) return file;
    files.push(file.value);
  }
  return ok({ ...children.value.copy, files });
};

const createdSchema = z.object({ id: graphIdSchema });

// The items of the copy under the new task, one write each. The first error stops it.
const addChildren = async (graph: WriteGraph, listId: GraphId, taskId: GraphId, copy: TaskCopy): Promise<Result<unknown, GraphError>> => {
  const path = taskPath(listId, taskId);
  for (const step of copy.steps) {
    const added = await graph.post(`${path}/checklistItems`, step);
    if (!added.ok) return added;
  }
  for (const link of copy.links) {
    const added = await graph.post(`${path}/linkedResources`, link);
    if (!added.ok) return added;
  }
  for (const file of copy.files) {
    const bytes = Uint8Array.from(Buffer.from(file.contentBytes, 'base64'));
    const added = await graph.uploadAttachment(
      { kind: 'todo-task', listId, taskId },
      { name: file.name, bytes, ...(file.contentType === undefined ? {} : { contentType: file.contentType }) }
    );
    if (!added.ok) return added;
  }
  return ok(undefined);
};

// Deletes the part copy of a copy that failed, and says the cause and what stays.
const partCopy = async (graph: WriteGraph, listId: GraphId, taskId: GraphId, cause: GraphError): Promise<GraphError> => {
  const removed = await graph.delete(taskPath(listId, taskId));
  const after = removed.ok ? 'The CLI deleted the part copy.' : `The CLI could not delete the part copy (task ${taskId} in list ${listId}): delete it in To Do.`;
  return { type: 'validation_error', code: 'todo_copy_failed', message: `The copy of the task failed: ${cause.message} ${after}` };
};

/** A new task in the list `listId` made from the copy; its id and its stamp. If a part of the copy fails, the new task is deleted again. */
export const createFromCopy = async (graph: WriteGraph, listId: GraphId, copy: TaskCopy): Promise<Result<TodoTask, GraphError>> => {
  const created = await graph.post(`/me/todo/lists/${listId}/tasks`, copy.fields);
  if (!created.ok) return created;
  const made = createdSchema.safeParse(created.value);
  if (!made.success) return err(unreadable('the task it created'));
  const id = made.data.id;
  const children = await addChildren(graph, listId, id, copy);
  if (!children.ok) return err(await partCopy(graph, listId, id, children.error));
  return readTask(graph, listId, id);
};

/**
 * Moves the task to the list `toListId`: a full copy there, then the delete of
 * the original. If the copy fails, or the delete of the original fails, the
 * copy is deleted again and the original stays: the answer is the error.
 */
export const moveTask = async (graph: WriteGraph, task: TodoTask, toListId: GraphId): Promise<Result<TodoTask, GraphError>> => {
  const copy = await readCopy(graph, task);
  if (!copy.ok) return copy;
  const created = await createFromCopy(graph, toListId, copy.value);
  if (!created.ok) return created;
  const deleted = await graph.delete(taskPath(task.listId, task.id));
  if (deleted.ok) return created;
  const removed = await graph.delete(taskPath(toListId, created.value.id));
  const after = removed.ok
    ? 'The CLI deleted the copy again, so the task stays in its list.'
    : `The CLI could not delete the copy (task ${created.value.id} in list ${toListId}): the task is now in both lists. Delete one in To Do.`;
  return err({
    type: 'validation_error',
    code: 'todo_move_failed',
    message: `The copy of the task is made, but the delete of the original failed: ${deleted.error.message} ${after}`,
  });
};
