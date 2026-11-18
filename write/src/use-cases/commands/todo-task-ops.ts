import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { AppliedChange, PlannedChange } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { isGraphId } from '../../domain/graph-id.ts';
import type { TaskFields } from '../../domain/todo-fields.ts';
import type { OpOutcome, TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import { sameTask, taskBefore } from './todo-command.ts';
import type { TaskChildren } from './todo-copy.ts';
import { moveTask, readCopy } from './todo-copy.ts';
import type { TodoList, TodoTask } from './todo-reader.ts';
import { taskPath, unreadable } from './todo-reader.ts';
import { addStepWrite } from './todo-step-ops.ts';

/*
 * The ops on a task: create, update (with new steps, for the ops file),
 * move and delete. Each op says its change, writes it, and records how to
 * reverse it (the task undo handlers).
 */

const createdSchema = z.object({ id: z.string(), title: z.string(), lastModifiedDateTime: z.string() });

const taskChange = (action: string, list: TodoList, task: TodoTask, more: Readonly<Record<string, string>> = {}): PlannedChange => ({
  action,
  target: taskPath(list.id, task.id),
  before: { ...taskBefore(list, task), ...more },
});

const childCounts = (children: TaskChildren): Readonly<Record<string, string>> => ({
  steps: String(children.steps),
  linkedResources: String(children.links),
  attachments: String(children.files.length),
});

/** Creates a task in the list. */
export const createTaskOp = (list: TodoList, fields: TaskFields): TodoOp => {
  const change: PlannedChange = { action: 'create-task', target: `/me/todo/lists/${list.id}/tasks`, before: { listId: list.id, list: list.displayName } };
  return {
    change,
    write: async (graph) => {
      const created = await graph.post(change.target, fields);
      if (!created.ok) return outcomeOf(created);
      const task = createdSchema.safeParse(created.value);
      if (!task.success || !isGraphId(task.data.id)) return outcomeOf(err(unreadable('the task it created')));
      const { id, title, lastModifiedDateTime } = task.data;
      return outcomeOf(
        ok({
          ...change,
          after: { listId: list.id, taskId: id, title },
          undo: { kind: 'delete-created-todo-task', data: { listId: list.id, taskId: id, stamp: lastModifiedDateTime } },
        })
      );
    },
  };
};

// The value of each field before the run: a field that Graph did not give is empty.
const valuesOf = (from: Readonly<Record<string, unknown>>, keys: ReadonlyArray<string>): Record<string, unknown> =>
  Object.fromEntries(keys.map((key) => [key, from[key] ?? (key === 'categories' ? [] : null)]));

const patchTask = async (graph: WriteGraph, change: PlannedChange, task: TodoTask, fields: TaskFields): Promise<Result<AppliedChange, GraphError>> => {
  const keys = Object.keys(fields);
  const written = await graph.patch(change.target, fields);
  if (!written.ok) return written;
  const answer = z.record(z.string(), z.unknown()).safeParse(written.value);
  if (!answer.success) return err(unreadable('the task it changed'));
  return ok({
    ...change,
    after: { listId: task.listId, taskId: task.id, fields: keys.join(', ') },
    undo: { kind: 'restore-todo-task-fields', data: { listId: task.listId, taskId: task.id, before: valuesOf(task.fields, keys), set: valuesOf(answer.data, keys) } },
  });
};

/** Sets the fields of the task, then adds the steps `steps`. `action` names the change: `update-task` or `complete-task`. */
export const updateTaskOp = (action: string, list: TodoList, task: TodoTask, fields: TaskFields, steps: ReadonlyArray<string> = []): TodoOp => {
  const change = taskChange(action, list, task);
  return {
    change,
    recheck: async (reads) => sameTask(reads, task),
    write: async (graph) => {
      const changes: AppliedChange[] = [];
      if (Object.keys(fields).length > 0) {
        const patched = await patchTask(graph, change, task, fields);
        if (!patched.ok) return { changes, error: patched.error };
        changes.push(patched.value);
      }
      for (const title of steps) {
        const added = await addStepWrite(graph, list, task, { displayName: title, isChecked: false });
        if (!added.ok) return { changes, error: added.error };
        changes.push(added.value);
      }
      return { changes };
    },
  };
};

/** Moves the task to the list `to`: a copy there, then the delete of the original. */
export const moveTaskOp = (list: TodoList, task: TodoTask, to: TodoList, children: TaskChildren): TodoOp => {
  const change = taskChange('move-task', list, task, { ...childCounts(children), toListId: to.id, to: to.displayName });
  return {
    change,
    recheck: async (reads) => sameTask(reads, task),
    write: async (graph): Promise<OpOutcome> => {
      const moved = await moveTask(graph, task, to.id);
      if (!moved.ok) return outcomeOf(moved);
      const { id, stamp } = moved.value;
      return outcomeOf(
        ok({
          ...change,
          after: { listId: to.id, list: to.displayName, taskId: id },
          undo: { kind: 'move-todo-task-back', data: { listId: to.id, taskId: id, stamp, backTo: list.id } },
        })
      );
    },
  };
};

/** Deletes the task. The run journal keeps its full copy, so the undo can create it again. */
export const deleteTaskOp = (list: TodoList, task: TodoTask, children: TaskChildren): TodoOp => {
  const change = taskChange('delete-task', list, task, childCounts(children));
  return {
    change,
    recheck: async (reads) => sameTask(reads, task),
    write: async (graph) => {
      const copy = await readCopy(graph, task);
      if (!copy.ok) return outcomeOf(copy);
      const deleted = await graph.delete(change.target);
      if (!deleted.ok) return outcomeOf(deleted);
      return outcomeOf(ok({ ...change, after: {}, undo: { kind: 'recreate-todo-task', data: { listId: list.id, copy: copy.value } } }));
    },
  };
};
