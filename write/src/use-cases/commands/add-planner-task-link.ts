import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { createExecute } from './mail-state-command.ts';
import { addPlan as plan, addSchema as schema } from './planner-link-commands.ts';
import { LINK_URL_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES, TASK_ID_OPTION } from './planner-texts.ts';

/*
 * Adds a link to a task: a reference to an http or https URL, with a name.
 * The undo removes the link.
 */

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Add a link (an http or https URL, with a name) to the references of a Planner task. The undo removes the link.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/planner/tasks/{task-id}/details',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
  options: [TASK_ID_OPTION, LINK_URL_OPTION, { name: 'alias', key: 'alias', required: false, description: 'The name of the link that Planner shows, at most 255 characters.' }],
  example: '{write-bin} add-planner-task-link --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --url "https://example.com/spec" --alias "Spec" --dry-run',
  bodyTemplate: "{ references: { <{url} with %, ., :, @ and # encoded>: { alias: '{alias}', type: 'Other' } } } with If-Match: <ETag of the details>",
  effect: 'task-write',
  undo: 'removes the link, only if its name and its type did not change',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
