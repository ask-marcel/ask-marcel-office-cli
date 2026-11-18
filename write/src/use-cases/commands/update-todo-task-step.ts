import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, listAndTask, TASK_ID_OPTION, TODO_RESPONSE } from './todo-command.ts';
import type { TodoStep } from './todo-step-ops.ts';
import { readStep, stepsPath, updateStepOp } from './todo-step-ops.ts';
import { CHECKED_OPTION, checkedSchema, STEP_ID_OPTION, STEP_TITLE_OPTION } from './todo-step-options.ts';

/*
 * Changes the text or the check mark of a step of a To Do task. A step that
 * has them already is passed over. The undo gives the step its text and its
 * check mark of before.
 */

const schema = z.object({
  list: z.string().min(1),
  taskId: z.string().min(1),
  stepId: z.string().min(1),
  title: z.string().trim().min(1).optional(),
  checked: checkedSchema.optional(),
});

const noChange: GraphError = { type: 'validation_error', message: 'update-todo-task-step: give --title, --checked, or both.' };

type StepSet = { readonly displayName?: string; readonly isChecked?: boolean };

const setOf = (title: string | undefined, checked: string | undefined): StepSet => ({
  ...(title === undefined ? {} : { displayName: title }),
  ...(checked === undefined ? {} : { isChecked: checked === 'true' }),
});

// Whether the step has each value of the set already.
const hasAll = (step: TodoStep, set: StepSet): boolean => (set.displayName ?? step.displayName) === step.displayName && (set.isChecked ?? step.isChecked) === step.isChecked;

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const { title, checked } = parsed.value;
  if (title === undefined && checked === undefined) return err(noChange);
  const found = await listAndTask(reads, parsed.value);
  if (!found.ok) return found;
  const { list, task } = found.value;
  const step = await readStep(reads, task, parsed.value.stepId);
  if (!step.ok) return step;
  const set = setOf(title, checked);
  if (hasAll(step.value, set)) {
    const target = `${stepsPath(list.id, task.id)}/${step.value.id}`;
    return ok(
      todoPlan({
        ops: [],
        skipped: [{ target, before: { listId: list.id, taskId: task.id, stepId: step.value.id }, reason: 'The step has this text and this check mark already.' }],
      })
    );
  }
  return ok(todoPlan({ ops: [updateStepOp(list, task, step.value, set)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Change the text or the check mark of a step of a Microsoft To Do task. A step that has them already is passed over.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/me/todo/lists/{list}/tasks/{task-id}/checklistItems/{step-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/checklistitem-update',
  options: [LIST_OPTION, TASK_ID_OPTION, STEP_ID_OPTION, STEP_TITLE_OPTION, CHECKED_OPTION],
  example: '{write-bin} update-todo-task-step --list "Tasks" --task-id "AAMkAD..." --step-id "51d8..." --checked true',
  bodyTemplate: "{ displayName?: '{title}', isChecked?: {checked} }",
  effect: 'task-write',
  undo: 'gives the step its text and its check mark of before, only if they did not change after the run. The date the step was checked is not kept',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
