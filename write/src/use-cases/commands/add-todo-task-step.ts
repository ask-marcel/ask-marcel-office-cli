import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, listAndTask, TASK_ID_OPTION, TODO_RESPONSE } from './todo-command.ts';
import { addStepOp } from './todo-step-ops.ts';
import { CHECKED_OPTION, checkedSchema, STEP_TITLE_OPTION } from './todo-step-options.ts';

/*
 * Adds a step to a To Do task (Graph calls it a checklistItem). The undo
 * deletes the step if it did not change after the run.
 */

const schema = z.object({ list: z.string().min(1), taskId: z.string().min(1), title: z.string().trim().min(1), checked: checkedSchema.optional() });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const found = await listAndTask(reads, parsed.value);
  if (!found.ok) return found;
  const step = { displayName: parsed.value.title, isChecked: parsed.value.checked === 'true' };
  return ok(todoPlan({ ops: [addStepOp(found.value.list, found.value.task, step)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Add a step to a Microsoft To Do task (a checklistItem in Graph), checked or not.',
  category: 'tasks',
  graphMethod: 'POST',
  graphPathTemplate: '/me/todo/lists/{list}/tasks/{task-id}/checklistItems',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotask-post-checklistitems',
  options: [LIST_OPTION, TASK_ID_OPTION, { ...STEP_TITLE_OPTION, required: true }, CHECKED_OPTION],
  example: '{write-bin} add-todo-task-step --list "Tasks" --task-id "AAMkAD..." --title "Read the draft"',
  bodyTemplate: "{ displayName: '{title}', isChecked: {checked} }",
  effect: 'task-write',
  undo: 'deletes the step, only if its text and its check mark did not change after the run',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
