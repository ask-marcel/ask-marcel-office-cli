import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, TODO_RESPONSE } from './todo-command.ts';
import { renameListOp } from './todo-list-ops.ts';
import { resolveList } from './todo-reader.ts';

/*
 * Gives a To Do list a new name. A list that has the name already is passed
 * over. The undo gives the list its name of before.
 */

const schema = z.object({ list: z.string().min(1), name: z.string().trim().min(1) });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const list = await resolveList(reads, parsed.value.list, '--list');
  if (!list.ok) return list;
  if (list.value.displayName === parsed.value.name) {
    const before = { listId: list.value.id, list: list.value.displayName };
    return ok(todoPlan({ ops: [], skipped: [{ target: `/me/todo/lists/${list.value.id}`, before, reason: 'The list has this name already.' }] }));
  }
  return ok(todoPlan({ ops: [renameListOp(list.value, parsed.value.name)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Give a Microsoft To Do list a new name. A list that has the name already is passed over.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/me/todo/lists/{list}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotasklist-update',
  options: [LIST_OPTION, { name: 'name', key: 'name', required: true, description: 'The new name of the list.' }],
  example: '{write-bin} update-todo-list --list "Errands" --name "Errands 2026"',
  bodyTemplate: "{ displayName: '{name}' }",
  effect: 'task-write',
  undo: 'gives the list its name of before, only if it still has the name that the run gave',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
