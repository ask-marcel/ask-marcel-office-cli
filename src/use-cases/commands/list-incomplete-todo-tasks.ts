import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';

// Hardcoded `$filter=status ne 'completed'` in the path means a user-supplied
// `--filter` would cause Graph to receive two `$filter` query params and
// reject with `InvalidFilterClause`. Audit round-6 §2.7: previously --filter
// was dropped from the schema entirely so Commander rejected with the
// generic "unknown option" — the LLM didn't know why or what to use instead.
// Accept --filter at the schema layer and reject in execute with a sharp
// pointer at the sibling command that supports it.
const schema = z.object({ todoTaskListId: z.string().min(1) }).extend(odataQuerySchema.shape);

const PARSE_URI_NEEDLE = 'RequestBroker--ParseUri';

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
  // Audit round-8 sec 1.1: this command hits the same /tasks endpoint as
  // the all-tasks sibling, which means it triggers the same RequestBroker--
  // ParseUri quirk on certain --select / --orderby values (notably any
  // combo with `title`). Mirror the round-6 rewrite from the sibling so
  // an LLM gets the same friendly hint from either command.
  if (result.error.type === 'api_error' && result.error.message.includes(PARSE_URI_NEEDLE)) {
    if (parsed.data.select !== undefined) {
      return err({
        type: 'api_error',
        status: result.error.status,
        message: `Graph rejected --select=${parsed.data.select} on this tasks endpoint with RequestBroker--ParseUri (known quirk — some field combinations are unsupported, most reliably any combo that includes \`title\`). Drop \`title\` from --select and request it in a second call (or per-task via get-todo-task), or call this command without --select and slim the response client-side.`,
        code: 'cli_rewrite_todo_select_title',
      });
    }
    if (parsed.data.orderby !== undefined) {
      return err({
        type: 'api_error',
        status: result.error.status,
        message: `Graph rejected --orderby=${parsed.data.orderby} on this tasks endpoint with RequestBroker--ParseUri (known quirk — sorting on \`title\` is unsupported). Call this command without --orderby and sort the response client-side, or order by a numeric/date field like \`createdDateTime\` / \`importance\` instead.`,
        code: 'cli_rewrite_todo_orderby_title',
      });
    }
  }
  return result;
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
