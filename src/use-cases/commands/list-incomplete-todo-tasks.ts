import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';

// Hardcoded `$filter=status ne 'completed'` in the path means a user-supplied
// `--filter` would cause Graph to receive two `$filter` query params and
// reject with `InvalidFilterClause`. Expose the other five OData passthrough
// flags but not `--filter`.
const noFilterShape = Object.fromEntries(Object.entries(odataQuerySchema.shape).filter(([key]) => key !== 'filter')) as Omit<typeof odataQuerySchema.shape, 'filter'>;
const noFilterOptions = odataQueryOptions.filter((o) => o.name !== 'filter');

const schema = z.object({ todoTaskListId: z.string().min(1) }).extend(noFilterShape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData(`/me/todo/lists/${parsed.data.todoTaskListId}/tasks?$filter=status ne 'completed'`, parsed.data);
  return graph.get(path);
};

const meta: CommandMeta = {
  summary:
    'List every incomplete Microsoft To Do task in a given list (status not equal to `completed`). Accepts the OData passthrough flags top/skip/select/orderby/expand. The filter passthrough is intentionally omitted — the path already pins a `$filter` for the completion-status predicate, and Graph rejects two `$filter` query params.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: "/me/todo/lists/{todo-task-list-id}/tasks?$filter=status ne 'completed'",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks',
  options: [
    {
      name: 'todo-task-list-id',
      key: 'todoTaskListId',
      required: true,
      description:
        'todoTaskList ID. Returned by `ask-marcel list-todo-task-lists`. The well-known name `tasks` (the default list) is accepted on this incomplete-tasks endpoint specifically — sibling commands like `list-todo-tasks` and `list-todo-tasks-delta` only accept resolved IDs. There is no Graph endpoint that returns incomplete tasks across every list — call this once per list.',
      aliases: [
        { name: 'task-list-id', key: 'taskListId' },
        { name: 'todo-list-id', key: 'todoListId' },
      ],
    },
    ...noFilterOptions,
  ],
  example: "ask-marcel list-incomplete-todo-tasks --todo-task-list-id 'tasks' --top 5",
  responseShape: 'collection of Microsoft Graph `todoTask` resources under `value[]` where `status != "completed"`',
  pagination: true,
};

export { execute, meta, schema };
