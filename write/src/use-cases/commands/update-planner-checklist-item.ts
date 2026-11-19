import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { createExecute } from './mail-state-command.ts';
import { updatePlan as plan, updateSchema as schema } from './planner-checklist-commands.ts';
import { CHECKLIST_ITEM_OPTION, CHECKLIST_OPTION_CHECKED, PLANNER_RESPONSE, PLANNER_SCOPES, TASK_ID_OPTION } from './planner-texts.ts';

/*
 * Changes the title or the state of an item of the checklist of a task. The
 * undo writes the title and the state of before back.
 */

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Change the title of an item of the checklist of a Planner task, or check or uncheck it. The undo writes the title and the state of before back.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/planner/tasks/{task-id}/details',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
  options: [
    TASK_ID_OPTION,
    CHECKLIST_ITEM_OPTION,
    { name: 'title', key: 'title', required: false, description: 'The new title of the item, at most 100 characters.' },
    CHECKLIST_OPTION_CHECKED,
  ],
  example: '{write-bin} update-planner-checklist-item --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --item "Book the room" --checked true --dry-run',
  bodyTemplate: "{ checklist: { <key of {item}>: { title: '{title}', isChecked: {checked} } } } with If-Match: <ETag of the details>",
  effect: 'task-write',
  undo: 'writes the title and the state of before back, only if the title and the state that the run set did not change',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
