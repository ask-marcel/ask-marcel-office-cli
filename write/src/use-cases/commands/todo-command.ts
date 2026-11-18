import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { CommandOptionMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { GraphReads, ItemIds } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { TodoList, TodoTask } from './todo-reader.ts';
import { readTask, readTaskById, resolveList } from './todo-reader.ts';

/*
 * What the To Do commands share: the options that name a list and a task, the
 * words of their output, the ids that a dry-run shows of a task, and the check
 * that a task is still as the plan read it.
 */

export const LIST_OPTION: CommandOptionMeta = {
  name: 'list',
  key: 'list',
  required: true,
  description: 'The To Do list: its name (compared with no case) or its id, as list-todo-task-lists returns them. Give the id when two lists have the same name.',
  argumentHint: { kind: 'idOrName' },
};

export const TASK_ID_OPTION: CommandOptionMeta = {
  name: 'task-id',
  key: 'taskId',
  required: true,
  description: 'The id of the task, as list-todo-tasks or get-todo-task return it.',
  argumentHint: { kind: 'idOrName' },
};

/** The words of the output that each To Do command shares. */
export const TODO_RESPONSE =
  '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids of the list and of the task before the change (and the title and the status of a task), and `after` the ids after it: a new task, step, linked resource or attachment shows its new id there. An item that changed or is gone between the plan and the write is in `skipped`; a write that Graph refused is in `failed`. With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.';

/** The ids and the state of a task that a dry-run shows. */
export const taskBefore = (list: TodoList, task: TodoTask): ItemIds => ({ listId: list.id, list: list.displayName, taskId: task.id, title: task.title, status: task.status });

/** Whether the task is as the plan read it: the same `lastModifiedDateTime`. */
export const sameTask = async (reads: GraphReads, task: TodoTask): Promise<Result<boolean, GraphError>> => {
  const now = await readTask(reads, task.listId, task.id);
  return now.ok ? ok(now.value.stamp === task.stamp) : now;
};

/** The list of `--list` and its task `--task-id`. */
export const listAndTask = async (
  reads: GraphReads,
  params: { readonly list: string; readonly taskId: string }
): Promise<Result<{ readonly list: TodoList; readonly task: TodoTask }, GraphError>> => {
  const list = await resolveList(reads, params.list, '--list');
  if (!list.ok) return list;
  const task = await readTaskById(reads, list.value, params.taskId, '--task-id');
  return task.ok ? ok({ list: list.value, task: task.value }) : task;
};
