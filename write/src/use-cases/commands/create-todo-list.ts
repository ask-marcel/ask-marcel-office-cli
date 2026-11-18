import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { TODO_RESPONSE } from './todo-command.ts';
import { createListOp } from './todo-list-ops.ts';

/*
 * Creates a To Do list. No command of this CLI deletes a list: the undo of
 * this command keeps the list, and says so.
 */

const schema = z.object({ name: z.string().trim().min(1) });

const plan: PlanWrite = async (_reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  return ok(todoPlan({ ops: [createListOp(parsed.value.name)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Create a Microsoft To Do list. No command of this CLI deletes a list, so its undo keeps the list: delete it in To Do.',
  category: 'tasks',
  graphMethod: 'POST',
  graphPathTemplate: '/me/todo/lists',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todo-post-lists',
  options: [{ name: 'name', key: 'name', required: true, description: 'The name of the new list.' }],
  example: '{write-bin} create-todo-list --name "Errands"',
  bodyTemplate: "{ displayName: '{name}' }",
  effect: 'task-write',
  undo: 'does not delete the list: no command of this CLI deletes a list. The undo refuses and says so; delete the list in To Do',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
