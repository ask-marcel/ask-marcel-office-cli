import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, listAndTask, TASK_ID_OPTION, TODO_RESPONSE } from './todo-command.ts';
import { readChildren } from './todo-copy.ts';
import { deleteTaskOp } from './todo-task-ops.ts';

/*
 * Deletes a To Do task. Before the delete, the run reads a full copy of the
 * task (its fields, steps, linked resources, and attachments with their bytes)
 * and the run journal keeps it, so the undo can create the task again.
 */

const schema = z.object({ list: z.string().min(1), taskId: z.string().min(1) });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const found = await listAndTask(reads, parsed.value);
  if (!found.ok) return found;
  const children = await readChildren(reads, found.value.task);
  if (!children.ok) return children;
  return ok(todoPlan({ ops: [deleteTaskOp(found.value.list, found.value.task, children.value)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Delete a Microsoft To Do task. The run journal keeps a full copy of the task (fields, steps, linked resources and attachments), so its undo can create the task again. Over MCP it runs on run-destructive-command.',
  category: 'tasks',
  graphMethod: 'DELETE',
  graphPathTemplate: '/me/todo/lists/{list}/tasks/{task-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotask-delete',
  options: [LIST_OPTION, TASK_ID_OPTION],
  example: '{write-bin} delete-todo-task --list "Tasks" --task-id "AAMkAD..." --dry-run',
  effect: 'delete',
  undo: 'creates the task again from the copy in the run journal, with its steps, linked resources and attachments. It is a new task: it gets a new id, and the dates of creation and the date a step was checked are not kept',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
