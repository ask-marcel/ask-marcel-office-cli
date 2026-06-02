import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';
import { rewriteTodoTitleQuirk } from './todo-parse-uri-rewrite.ts';

const schema = z.object({ todoTaskListId: z.string().min(1) }).extend(odataQuerySchema.shape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData(`/me/todo/lists/${parsed.data.todoTaskListId}/tasks`, parsed.data);
  const result = await graph.get(path);
  if (result.ok) return result;
  // Graph's RequestBroker--ParseUri title quirk on this endpoint — rewrite the
  // opaque error to an actionable hint via the shared helper.
  const rewritten = rewriteTodoTitleQuirk(result.error, parsed.data);
  return rewritten ? err(rewritten) : result;
};

const meta: CommandMeta = {
  summary:
    'List every task in a single Microsoft To Do task list, regardless of completion status. Use `list-incomplete-todo-tasks` if you only want the open ones. Known Graph quirk: certain `--select` combinations (notably any combo that includes `title`) trip `RequestBroker--ParseUri` on this endpoint; the CLI rewrites that opaque error to a hint pointing at the workaround.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/todo/lists/{todo-task-list-id}/tasks',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks',
  options: [
    {
      name: 'todo-task-list-id',
      key: 'todoTaskListId',
      required: true,
      description: 'To Do task list ID. Returned by `ask-marcel list-todo-task-lists`.',
      aliases: [
        { name: 'task-list-id', key: 'taskListId' },
        { name: 'todo-list-id', key: 'todoListId' },
      ],
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-todo-tasks --todo-task-list-id 'AAMkAGI...'",
  responseShape: 'collection of Microsoft Graph `todoTask` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
