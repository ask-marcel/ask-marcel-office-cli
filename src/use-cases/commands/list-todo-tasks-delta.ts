import { z } from 'zod';
import { map } from '../../domain/result.ts';
import { buildCommand } from './build-command.ts';
import type { ReadCommand, ReadCommandMeta } from './command-types.ts';
import { withTaskLinks } from './todo-web-url.ts';

const schema = z.object({ todoTaskListId: z.string().min(1) });
const delta = buildCommand((p) => `/me/todo/lists/${p.todoTaskListId}/tasks/delta()`, schema);
const execute: ReadCommand['execute'] = async (graph, params) => map(await delta.execute(graph, params), withTaskLinks);

const meta: ReadCommandMeta = {
  summary:
    'Track incremental task changes (added / updated / completed / deleted) within a single Microsoft To Do list. The first call returns the current snapshot plus `@odata.deltaLink`; subsequent calls with that link return only what has changed since. Note: Graph rejects standard OData query parameters on this delta endpoint (the page-cap flag throws `Skip token is not provided`), so the OData passthrough is intentionally NOT exposed here. Use `next-page` with the returned `@odata.nextLink` to walk pages.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/todo/lists/{todo-task-list-id}/tasks/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotask-delta',
  options: [
    {
      name: 'todo-task-list-id',
      key: 'todoTaskListId',
      required: true,
      description: 'Microsoft To Do task list ID. Returned by `list-todo-task-lists`.',
    },
  ],
  example: "ask-marcel-office list-todo-tasks-delta --todo-task-list-id 'AAMkAD...'",
  responseShape:
    'collection of Microsoft Graph `todoTask` resources under `data.value[]`, each with `webUrl`, the link the To Do web app opens it with (`https://to-do.office.com/tasks/id/<id>/details`). Cursor tokens are hoisted to envelope level: top-level `nextLink` while paging, then top-level `deltaLink` on the final page.',
  pagination: true,
  paginationStrategy: 'deltaLink',
  effect: 'read',
  scopesRequired: ['Tasks.Read'],
};

export { execute, meta, schema };
