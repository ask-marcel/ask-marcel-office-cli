import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';
import { rewriteTodoTitleQuirk } from './todo-parse-uri-rewrite.ts';

// Hardcoded `$filter=status ne 'completed'` in the path means a user-supplied
// `--filter` would cause Graph to receive two `$filter` query params and
// reject with `InvalidFilterClause`. Audit round-6 §2.7: previously --filter
// was dropped from the schema entirely so Commander rejected with the
// generic "unknown option" — the LLM didn't know why or what to use instead.
// Accept --filter at the schema layer and reject in execute with a sharp
// pointer at the sibling command that supports it.
const schema = z.object({ todoTaskListId: z.string().min(1) }).extend(odataQuerySchema.shape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  if (parsed.data.filter !== undefined) {
    return err({
      type: 'validation_error',
      message:
        "--filter is not supported on list-incomplete-todo-tasks: the path already pins `$filter=status ne 'completed'` and Graph rejects two $filter query params. To combine with your own filter, call `list-todo-tasks --filter \"status ne 'completed' and <your-predicate>\"` instead (single $filter, AND your predicate yourself).",
    });
  }
  const path = appendOData(`/me/todo/lists/${parsed.data.todoTaskListId}/tasks?$filter=status ne 'completed'`, parsed.data);
  const result = await graph.get(path);
  if (result.ok) return result;
  // Same /tasks endpoint as the all-tasks sibling, so the same RequestBroker--
  // ParseUri title quirk applies — rewrite via the shared helper.
  const rewritten = rewriteTodoTitleQuirk(result.error, parsed.data);
  return rewritten ? err(rewritten) : result;
};

const meta: CommandMeta = {
  summary:
    'List every incomplete Microsoft To Do task in a given list (status not equal to `completed`). Accepts the OData passthrough flags top/skip/select/orderby/expand. The filter passthrough is intentionally omitted — the path already pins a `$filter` for the completion-status predicate, and Graph rejects two `$filter` query params. If you supply `--filter` anyway, the CLI returns a clear pointer to `list-todo-tasks` (which lets you AND your predicate with the completion filter yourself).',
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
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-incomplete-todo-tasks --todo-task-list-id 'tasks' --top 5",
  responseShape: 'collection of Microsoft Graph `todoTask` resources under `value[]` where `status != "completed"`',
  pagination: true,
};

export { execute, meta, schema };
