import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { createExecute } from './mail-state-command.ts';
import { addPlan as plan, addSchema as schema } from './planner-checklist-commands.ts';
import { CHECKLIST_OPTION_CHECKED, PLANNER_RESPONSE, PLANNER_SCOPES, TASK_ID_OPTION } from './planner-texts.ts';

/*
 * Adds an item to the checklist of a task. A task with 20 items, or with an
 * item of the same title, is refused. The undo removes the item.
 */

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Add an item to the checklist of a Planner task (at most 20 items). The undo removes the item.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/planner/tasks/{task-id}/details',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
  options: [
    TASK_ID_OPTION,
    { name: 'title', key: 'title', required: true, description: 'The title of the item, at most 100 characters. The checklist must not have an item of this title.' },
    CHECKLIST_OPTION_CHECKED,
  ],
  example: '{write-bin} add-planner-checklist-item --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --title "Book the room" --dry-run',
  bodyTemplate: "{ checklist: { <new key>: { title: '{title}', isChecked: {checked} } } } with If-Match: <ETag of the details>",
  effect: 'task-write',
  undo: 'removes the item, only if its title and its state did not change',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
