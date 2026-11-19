import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { createExecute } from './mail-state-command.ts';
import { removePlan as plan, removeSchema as schema } from './planner-checklist-commands.ts';
import { CHECKLIST_ITEM_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES, TASK_ID_OPTION } from './planner-texts.ts';

/*
 * Removes an item from the checklist of a task. The undo adds it again, with
 * its title and its state and a new id.
 */

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Remove an item from the checklist of a Planner task. The undo adds the item again with a new id.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/planner/tasks/{task-id}/details',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
  options: [TASK_ID_OPTION, CHECKLIST_ITEM_OPTION],
  example: '{write-bin} remove-planner-checklist-item --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --item "Book the room" --dry-run',
  bodyTemplate: '{ checklist: { <key of {item}>: null } } with If-Match: <ETag of the details>',
  effect: 'task-write',
  undo: 'adds the item again with its title, its state and a new id, only if it is still not in the checklist',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
