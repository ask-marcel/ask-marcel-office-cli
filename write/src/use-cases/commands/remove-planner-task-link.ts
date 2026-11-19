import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { createExecute } from './mail-state-command.ts';
import { removePlan as plan, removeSchema as schema } from './planner-link-commands.ts';
import { LINK_URL_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES, TASK_ID_OPTION } from './planner-texts.ts';

/*
 * Removes a link or a file from the references of a task. The file itself
 * stays where it is. The undo adds the reference again.
 */

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Remove a link or a file from the references of a Planner task (a file stays where it is). The undo adds the reference again.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/planner/tasks/{task-id}/details',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
  options: [TASK_ID_OPTION, LINK_URL_OPTION],
  example: '{write-bin} remove-planner-task-link --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --url "https://example.com/spec" --dry-run',
  bodyTemplate: '{ references: { <{url} with %, ., :, @ and # encoded>: null } } with If-Match: <ETag of the details>',
  effect: 'task-write',
  undo: 'adds the reference again with its name and its type, only if it is still not in the references',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
