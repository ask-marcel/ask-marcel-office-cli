import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { LocalFile } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { GraphReads, PlannedChange } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { graphIdSchema, isGraphId, parseGraphId } from '../../domain/graph-id.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import type { TodoList, TodoTask } from './todo-reader.ts';
import { taskPath, unreadable } from './todo-reader.ts';

/*
 * The ops on the linked resources and the file attachments of a task. The
 * undo of an added item deletes it. The undo of a removed attachment attaches
 * the file again from the bytes that the run journal keeps: the attachment
 * gets a new id.
 */

/** The fields of a linked resource that the command of linked resources sets. */
export type LinkFields = Readonly<Partial<Record<'webUrl' | 'applicationName' | 'displayName' | 'externalId', string>>>;

/** An attachment as the list of the attachments of a task gives it: no bytes. */
export type TodoAttachment = { readonly id: GraphId; readonly name: string; readonly size: number };

export const attachmentsPath = (listId: string, taskId: string): string => `${taskPath(listId, taskId)}/attachments`;

const linkSchema = z.object({ id: z.string() }).loose();
const attachmentSchema = z.object({ id: graphIdSchema, name: z.string(), size: z.number() });
const attachmentsSchema = z.object({ value: z.array(attachmentSchema) });
const contentSchema = z.object({ contentBytes: z.string(), contentType: z.string().optional() });
const oneContentSchema = z.union([contentSchema, z.object({ value: contentSchema }).transform((answer) => answer.value)]);

const taskIds = (list: TodoList, task: TodoTask): Readonly<Record<string, string>> => ({ listId: list.id, list: list.displayName, taskId: task.id, title: task.title });

/** Links a resource to the task. */
export const addLinkOp = (list: TodoList, task: TodoTask, link: LinkFields): TodoOp => {
  const change: PlannedChange = { action: 'add-linked-resource', target: `${taskPath(list.id, task.id)}/linkedResources`, before: taskIds(list, task) };
  return {
    change,
    write: async (graph) => {
      const added = await graph.post(change.target, link);
      if (!added.ok) return outcomeOf(added);
      const id = linkSchema.safeParse(added.value).data?.id ?? '';
      if (!isGraphId(id)) return outcomeOf(err(unreadable('the linked resource it created')));
      const ids = { listId: list.id, taskId: task.id, linkedResourceId: id };
      return outcomeOf(ok({ ...change, after: ids, undo: { kind: 'delete-created-todo-link', data: { ...ids, set: { ...link } } } }));
    },
  };
};

/** Attaches the local file to the task. */
export const addFileOp = (list: TodoList, task: TodoTask, file: LocalFile): TodoOp => {
  const change: PlannedChange = {
    action: 'add-attachment',
    target: attachmentsPath(list.id, task.id),
    before: { ...taskIds(list, task), file: file.name, size: String(file.bytes.byteLength) },
  };
  return {
    change,
    write: async (graph) => {
      const added = await graph.uploadAttachment({ kind: 'todo-task', listId: list.id, taskId: task.id }, file);
      if (!added.ok) return outcomeOf(added);
      const ids = { listId: list.id, taskId: task.id, attachmentId: added.value.id ?? '' };
      return outcomeOf(ok({ ...change, after: ids, undo: { kind: 'delete-created-todo-attachment', data: { ...ids, name: added.value.name } } }));
    },
  };
};

/** The attachments of the task, with no bytes. */
export const readAttachments = async (reads: GraphReads, task: TodoTask): Promise<Result<ReadonlyArray<TodoAttachment>, GraphError>> => {
  const read = await reads.get(attachmentsPath(task.listId, task.id));
  if (!read.ok) return read;
  const page = attachmentsSchema.safeParse(read.value);
  return page.success ? ok(page.data.value) : err(unreadable('the attachments of a To Do task'));
};

const notAttached = (): Result<never, GraphError> =>
  err({
    type: 'validation_error',
    code: 'todo_attachment_not_found',
    message: '--attachment-id: the task has no attachment with this id. get-todo-task with --expand attachments gives the ids.',
  });

/** The attachment `raw` of the task, after the check of the id. */
export const findAttachment = async (reads: GraphReads, task: TodoTask, raw: string): Promise<Result<TodoAttachment, GraphError>> => {
  const id = parseGraphId(raw, '--attachment-id');
  if (!id.ok) return id;
  const all = await readAttachments(reads, task);
  if (!all.ok) return all;
  const found = all.value.find((attachment) => attachment.id === id.value);
  return found === undefined ? notAttached() : ok(found);
};

const sameAttachment = async (reads: GraphReads, task: TodoTask, attachment: TodoAttachment): Promise<Result<boolean, GraphError>> => {
  const all = await readAttachments(reads, task);
  if (!all.ok) return all;
  const now = all.value.find((item) => item.id === attachment.id);
  return now === undefined
    ? err({ type: 'api_error', status: 404, message: 'The attachment is no longer there.' })
    : ok(now.name === attachment.name && now.size === attachment.size);
};

/** Removes the attachment. The run journal keeps its bytes, so the undo can attach the file again. */
export const removeFileOp = (list: TodoList, task: TodoTask, attachment: TodoAttachment): TodoOp => {
  const target = `${attachmentsPath(list.id, task.id)}/${attachment.id}`;
  const change: PlannedChange = {
    action: 'remove-attachment',
    target,
    before: { ...taskIds(list, task), attachmentId: attachment.id, file: attachment.name, size: String(attachment.size) },
  };
  return {
    change,
    recheck: async (reads) => sameAttachment(reads, task, attachment),
    write: async (graph) => {
      const read = await graph.get(target);
      if (!read.ok) return outcomeOf(read);
      const content = oneContentSchema.safeParse(read.value);
      if (!content.success) return outcomeOf(err(unreadable('the bytes of the attachment')));
      const deleted = await graph.delete(target);
      if (!deleted.ok) return outcomeOf(deleted);
      const file = { name: attachment.name, contentBytes: content.data.contentBytes, ...(content.data.contentType === undefined ? {} : { contentType: content.data.contentType }) };
      return outcomeOf(ok({ ...change, after: {}, undo: { kind: 'reattach-todo-file', data: { listId: list.id, taskId: task.id, file } } }));
    },
  };
};
