import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, listAndTask, TASK_ID_OPTION, TODO_RESPONSE } from './todo-command.ts';
import { addLinkOp } from './todo-item-ops.ts';

/*
 * Links a resource to a To Do task: for example the mail or the page that the
 * task comes from. The undo deletes the linked resource.
 */

const isWebUrl = (value: string): boolean => URL.canParse(value) && ['http:', 'https:'].includes(new URL(value).protocol);

const schema = z.object({
  list: z.string().min(1),
  taskId: z.string().min(1),
  displayName: z.string().trim().min(1),
  applicationName: z.string().trim().min(1),
  webUrl: z.string().refine(isWebUrl, 'must be an http or https URL').optional(),
  externalId: z.string().min(1).optional(),
});

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const found = await listAndTask(reads, parsed.value);
  if (!found.ok) return found;
  const { displayName, applicationName, webUrl, externalId } = parsed.value;
  const link = { displayName, applicationName, ...(webUrl === undefined ? {} : { webUrl }), ...(externalId === undefined ? {} : { externalId }) };
  return ok(todoPlan({ ops: [addLinkOp(found.value.list, found.value.task, link)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Link a resource to a Microsoft To Do task, for example the mail or the web page that the task comes from. To Do shows it in the details of the task.',
  category: 'tasks',
  graphMethod: 'POST',
  graphPathTemplate: '/me/todo/lists/{list}/tasks/{task-id}/linkedResources',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotask-post-linkedresources',
  options: [
    LIST_OPTION,
    TASK_ID_OPTION,
    { name: 'display-name', key: 'displayName', required: true, description: 'The title of the linked resource, as To Do shows it.' },
    { name: 'application-name', key: 'applicationName', required: true, description: 'The name of the application that the resource comes from, for example Outlook.' },
    { name: 'web-url', key: 'webUrl', required: false, description: 'The link that opens the resource: an http or https URL.' },
    { name: 'external-id', key: 'externalId', required: false, description: 'The id of the resource in its application.' },
  ],
  example: '{write-bin} add-todo-linked-resource --list "Tasks" --task-id "AAMkAD..." --display-name "Spec" --application-name "Browser" --web-url "https://example.com/spec"',
  bodyTemplate: "{ displayName: '{display-name}', applicationName: '{application-name}', webUrl?: '{web-url}', externalId?: '{external-id}' }",
  effect: 'task-write',
  undo: 'deletes the linked resource, only if it did not change after the run',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
