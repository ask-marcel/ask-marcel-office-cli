import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, listAndTask, TASK_ID_OPTION, TODO_RESPONSE } from './todo-command.ts';
import { readChildren } from './todo-copy.ts';
import { resolveList } from './todo-reader.ts';
import { moveTaskOp } from './todo-task-ops.ts';

/*
 * Moves a To Do task to another list. Graph has no move of a task: the command
 * copies the task to the other list (its fields, its steps, its linked
 * resources and its attachments), and deletes the original only after the full
 * copy is made. If a part of the copy fails, it deletes the copy and the
 * original stays.
 */

const schema = z.object({ list: z.string().min(1), taskId: z.string().min(1), to: z.string().min(1) });

const sameList: GraphError = { type: 'validation_error', code: 'todo_same_list', message: '--to: the task is in this list already.' };

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const found = await listAndTask(reads, parsed.value);
  if (!found.ok) return found;
  const to = await resolveList(reads, parsed.value.to, '--to');
  if (!to.ok) return to;
  if (to.value.id === found.value.list.id) return err(sameList);
  const children = await readChildren(reads, found.value.task);
  if (!children.ok) return children;
  return ok(todoPlan({ ops: [moveTaskOp(found.value.list, found.value.task, to.value, children.value)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Move a Microsoft To Do task to another list. Graph has no move of a task: the command copies the task with its steps, linked resources and attachments to the other list, then deletes the original. The task gets a new id, which the answer gives. If the copy fails, the original stays. Over MCP it runs on run-destructive-command.',
  category: 'tasks',
  graphMethod: 'POST',
  graphPathTemplate: '/me/todo/lists/{to}/tasks, then DELETE /me/todo/lists/{list}/tasks/{task-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks',
  options: [
    LIST_OPTION,
    TASK_ID_OPTION,
    { name: 'to', key: 'to', required: true, description: 'The list to move the task to: its name (compared with no case) or its id.', argumentHint: { kind: 'idOrName' } },
  ],
  example: '{write-bin} move-todo-task --list "Tasks" --task-id "AAMkAD..." --to "Later" --dry-run',
  bodyTemplate: '{ <the fields of the task> }, then one POST for each step, linked resource and attachment',
  effect: 'delete',
  undo: 'moves the task back with a new copy (it gets a new id again), only if it did not change after the run. The dates of creation are not kept',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
