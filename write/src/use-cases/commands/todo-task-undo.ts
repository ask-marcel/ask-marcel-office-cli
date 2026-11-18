import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import { graphIdSchema } from '../../domain/graph-id.ts';
import { createFromCopy, moveTask, taskCopySchema } from './todo-copy.ts';
import { readTask, taskPath } from './todo-reader.ts';
import { canonical, checkTask, createTodoUndoHandler, listThere } from './todo-undo-handler.ts';

/*
 * The undo of the changes of a task. Each check reads the task: an undo that
 * deletes or moves a task runs only if the task has the `lastModifiedDateTime`
 * that the run left; the undo of an update compares the fields that the run
 * set, so a later change of another field or of a step does not stop it. A
 * task that the undo creates again, or moves back, gets a new id.
 */

const ids = { listId: graphIdSchema, taskId: graphIdSchema };
const stampedSchema = z.object({ ...ids, stamp: z.string() });
const fieldsSchema = z.object({ ...ids, before: z.record(z.string(), z.unknown()), set: z.record(z.string(), z.unknown()) });
const recreateSchema = z.object({ listId: graphIdSchema, copy: taskCopySchema });
const backSchema = stampedSchema.extend({ backTo: graphIdSchema });

/** The handler of each kind of undo of a task change. */
export const TASK_UNDO_HANDLERS: UndoHandlers = {
  'delete-created-todo-task': createTodoUndoHandler(
    stampedSchema,
    async (reads, data) => checkTask(reads, data.listId, data.taskId, (task) => task.stamp === data.stamp),
    async (graph, data) => graph.delete(taskPath(data.listId, data.taskId))
  ),
  'restore-todo-task-fields': createTodoUndoHandler(
    fieldsSchema,
    async (reads, data) =>
      checkTask(reads, data.listId, data.taskId, (task) => Object.keys(data.set).every((key) => canonical(task.fields[key] ?? null) === canonical(data.set[key]))),
    async (graph, data) => graph.patch(taskPath(data.listId, data.taskId), data.before)
  ),
  'recreate-todo-task': createTodoUndoHandler(
    recreateSchema,
    async (reads, data) => listThere(reads, data.listId),
    async (graph, data) => {
      const created = await createFromCopy(graph, data.listId, data.copy);
      return created.ok ? ok({ now: { listId: data.listId, taskId: created.value.id } }) : created;
    }
  ),
  'move-todo-task-back': createTodoUndoHandler(
    backSchema,
    async (reads, data) => {
      const target = await listThere(reads, data.backTo);
      if (!target.ok || target.value !== 'ready') return target;
      return checkTask(reads, data.listId, data.taskId, (task) => task.stamp === data.stamp);
    },
    async (graph, data) => {
      const task = await readTask(graph, data.listId, data.taskId);
      if (!task.ok) return task;
      const moved = await moveTask(graph, task.value, data.backTo);
      return moved.ok ? ok({ now: { listId: data.backTo, taskId: moved.value.id } }) : moved;
    }
  ),
};
