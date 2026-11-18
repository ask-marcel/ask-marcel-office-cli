import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, listAndTask, TASK_ID_OPTION, TODO_RESPONSE } from './todo-command.ts';
import { deleteStepOp, readStep } from './todo-step-ops.ts';
import { STEP_ID_OPTION } from './todo-step-options.ts';

/*
 * Deletes a step of a To Do task. The run journal keeps its text and its check
 * mark, so the undo can add the step again, with a new id.
 */

const schema = z.object({ list: z.string().min(1), taskId: z.string().min(1), stepId: z.string().min(1) });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const found = await listAndTask(reads, parsed.value);
  if (!found.ok) return found;
  const step = await readStep(reads, found.value.task, parsed.value.stepId);
  if (!step.ok) return step;
  return ok(todoPlan({ ops: [deleteStepOp(found.value.list, found.value.task, step.value)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Delete a step of a Microsoft To Do task. Over MCP it runs on run-destructive-command.',
  category: 'tasks',
  graphMethod: 'DELETE',
  graphPathTemplate: '/me/todo/lists/{list}/tasks/{task-id}/checklistItems/{step-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/checklistitem-delete',
  options: [LIST_OPTION, TASK_ID_OPTION, STEP_ID_OPTION],
  example: '{write-bin} delete-todo-task-step --list "Tasks" --task-id "AAMkAD..." --step-id "51d8..." --dry-run',
  effect: 'delete',
  undo: 'adds the step again with its text and its check mark, at the end of the steps. It gets a new id, and the date it was checked is not kept',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
