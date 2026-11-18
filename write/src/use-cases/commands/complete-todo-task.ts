import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, listAndTask, TASK_ID_OPTION, taskBefore, TODO_RESPONSE } from './todo-command.ts';
import { taskPath } from './todo-reader.ts';
import { updateTaskOp } from './todo-task-ops.ts';

/*
 * Marks a To Do task completed. A task that is completed already is passed
 * over. The undo gives the task its status of before.
 */

const schema = z.object({ list: z.string().min(1), taskId: z.string().min(1) });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const found = await listAndTask(reads, parsed.value);
  if (!found.ok) return found;
  const { list, task } = found.value;
  if (task.status === 'completed') {
    return ok(todoPlan({ ops: [], skipped: [{ target: taskPath(list.id, task.id), before: taskBefore(list, task), reason: 'The task is completed already.' }] }));
  }
  return ok(todoPlan({ ops: [updateTaskOp('complete-task', list, task, { status: 'completed' })] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Mark a Microsoft To Do task completed. A task that is completed already is passed over.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/me/todo/lists/{list}/tasks/{task-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotask-update',
  options: [LIST_OPTION, TASK_ID_OPTION],
  example: '{write-bin} complete-todo-task --list "Tasks" --task-id "AAMkAD..."',
  bodyTemplate: "{ status: 'completed' }",
  effect: 'task-write',
  undo: 'gives the task its status of before, only if it is still completed',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
