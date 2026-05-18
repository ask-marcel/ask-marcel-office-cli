import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { pickODataOptions } from './odata-query.ts';

// Microsoft Graph rejects $select and $orderby on the task-list endpoint
// with the opaque `RequestBroker--ParseUri: Invalid request` (audit
// Tasks Bug 1). $filter works but enum literals trip parser quirks.
// Keep the safe passthroughs (--top, --skip, --filter, --expand) and
// drop the two that always 400.
const TODO_LISTS_ODATA_KEYS = ['top', 'skip', 'filter', 'expand'] as const;
const baseSchema = z.object({}).strict();
const { execute, schema } = buildPickODataListCommand(() => '/me/todo/lists', baseSchema, TODO_LISTS_ODATA_KEYS);

const meta: CommandMeta = {
  summary:
    "List the signed-in user's Microsoft To Do task lists (e.g. `Tasks`, `Flagged Emails`, custom lists). Note: Graph rejects `$select` and `$orderby` on this endpoint with `RequestBroker--ParseUri`, so the CLI does not expose those flags — slice / sort client-side.",
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/todo/lists',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todo-list-lists',
  options: [...pickODataOptions(TODO_LISTS_ODATA_KEYS)],
  example: 'ask-marcel list-todo-task-lists',
  responseShape: 'collection of Microsoft Graph `todoTaskList` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
