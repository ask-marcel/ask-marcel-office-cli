import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoCheck, UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { graphIdSchema } from '../../domain/graph-id.ts';
import { attachmentsPath, readAttachments } from './todo-item-ops.ts';
import { readListName } from './todo-list-ops.ts';
import { readTask, taskPath } from './todo-reader.ts';
import { readStep, stepsPath, stepStateSchema } from './todo-step-ops.ts';
import { canonical, checkRead, checkTask, createTodoUndoHandler, listThere, notThere } from './todo-undo-handler.ts';

/*
 * The undo of the changes of the items of a task (steps, linked resources,
 * attachments) and of a list. An item that the undo adds again gets a new id;
 * a step added again also loses the date it was checked. No command deletes a
 * list, so the undo of a new list refuses and keeps the list.
 */

const ids = { listId: graphIdSchema, taskId: graphIdSchema };
const stepSchema = z.object({ ...ids, stepId: graphIdSchema, set: stepStateSchema });
const restoreStepSchema = stepSchema.extend({ before: stepStateSchema });
const recreateStepSchema = z.object({ ...ids, step: stepStateSchema });
const linkSchema = z.object({ ...ids, linkedResourceId: graphIdSchema, set: z.record(z.string(), z.string()) });
// Graph lists an attachment with a size that is not the size of the file, so the check compares the name only: an attachment cannot change.
const fileSchema = z.object({ ...ids, attachmentId: graphIdSchema, name: z.string() });
const reattachSchema = z.object({ ...ids, file: z.object({ name: z.string(), contentType: z.string().optional(), contentBytes: z.string() }) });
const listSchema = z.object({ listId: graphIdSchema });
const listNameSchema = listSchema.extend({ before: z.string(), set: z.string() });

const linkFieldsSchema = z.record(z.string(), z.unknown());

// The step is as the run left it: the text and the check mark that the run set.
const checkStep = async (reads: GraphReads, data: z.infer<typeof stepSchema>): Promise<Result<UndoCheck, GraphError>> => {
  const task = await readTask(reads, data.listId, data.taskId);
  if (!task.ok) return notThere(task.error);
  return checkRead(await readStep(reads, task.value, data.stepId), (step) => canonical({ displayName: step.displayName, isChecked: step.isChecked }) === canonical(data.set));
};

/** The refusal of the undo of a new list. */
export const LIST_KEPT = 'No command of this CLI deletes a To Do list, so this undo keeps the list that the run created. Delete it in To Do if you do not want it.';

/** The handler of each kind of undo of an item of a task or of a list. */
export const ITEM_UNDO_HANDLERS: UndoHandlers = {
  'delete-created-todo-step': createTodoUndoHandler(
    stepSchema,
    async (reads, data) => checkStep(reads, data),
    async (graph, data) => graph.delete(`${stepsPath(data.listId, data.taskId)}/${data.stepId}`)
  ),
  'restore-todo-step': createTodoUndoHandler(
    restoreStepSchema,
    async (reads, data) => checkStep(reads, data),
    async (graph, data) => graph.patch(`${stepsPath(data.listId, data.taskId)}/${data.stepId}`, data.before)
  ),
  'recreate-todo-step': createTodoUndoHandler(
    recreateStepSchema,
    async (reads, data) => checkTask(reads, data.listId, data.taskId, () => true),
    async (graph, data) => {
      const added = await graph.post(stepsPath(data.listId, data.taskId), data.step);
      if (!added.ok) return added;
      const id = z.object({ id: z.string() }).safeParse(added.value).data?.id;
      return ok(id === undefined ? {} : { now: { listId: data.listId, taskId: data.taskId, stepId: id } });
    }
  ),
  'delete-created-todo-link': createTodoUndoHandler(
    linkSchema,
    async (reads, data) =>
      checkRead(await reads.get(`${taskPath(data.listId, data.taskId)}/linkedResources/${data.linkedResourceId}`), (link) => {
        const now = linkFieldsSchema.safeParse(link).data ?? {};
        return Object.keys(data.set).every((key) => now[key] === data.set[key]);
      }),
    async (graph, data) => graph.delete(`${taskPath(data.listId, data.taskId)}/linkedResources/${data.linkedResourceId}`)
  ),
  'delete-created-todo-attachment': createTodoUndoHandler(
    fileSchema,
    async (reads, data) => {
      const task = await readTask(reads, data.listId, data.taskId);
      if (!task.ok) return notThere(task.error);
      const all = await readAttachments(reads, task.value);
      if (!all.ok) return all;
      const now = all.value.find((item) => item.id === data.attachmentId);
      if (now === undefined) return ok('missing');
      return ok(now.name === data.name ? 'ready' : 'changed');
    },
    async (graph, data) => graph.delete(`${attachmentsPath(data.listId, data.taskId)}/${data.attachmentId}`)
  ),
  'reattach-todo-file': createTodoUndoHandler(
    reattachSchema,
    async (reads, data) => checkTask(reads, data.listId, data.taskId, () => true),
    async (graph, data) => {
      const { name, contentType, contentBytes } = data.file;
      const bytes = Uint8Array.from(Buffer.from(contentBytes, 'base64'));
      const added = await graph.uploadAttachment(
        { kind: 'todo-task', listId: data.listId, taskId: data.taskId },
        { name, bytes, ...(contentType === undefined ? {} : { contentType }) }
      );
      if (!added.ok) return added;
      return ok(added.value.id === undefined ? {} : { now: { listId: data.listId, taskId: data.taskId, attachmentId: added.value.id } });
    }
  ),
  'todo-list-created': createTodoUndoHandler(
    listSchema,
    async () => err({ type: 'validation_error', code: 'no_list_delete', message: LIST_KEPT }),
    async () => err({ type: 'validation_error', code: 'no_list_delete', message: LIST_KEPT })
  ),
  'restore-todo-list-name': createTodoUndoHandler(
    listNameSchema,
    async (reads, data) => {
      const there = await listThere(reads, data.listId);
      if (!there.ok || there.value !== 'ready') return there;
      return checkRead(await readListName(reads, data.listId), (name) => name === data.set);
    },
    async (graph, data) => graph.patch(`/me/todo/lists/${data.listId}`, { displayName: data.before })
  ),
};
