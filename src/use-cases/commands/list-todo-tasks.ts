import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';

const schema = z.object({ todoTaskListId: z.string().min(1) }).extend(odataQuerySchema.shape);

const PARSE_URI_NEEDLE = 'RequestBroker--ParseUri';

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData(`/me/todo/lists/${parsed.data.todoTaskListId}/tasks`, parsed.data);
  const result = await graph.get(path);
  if (result.ok) return result;
  // Audit round-6 §1.5 + v1.0.0 audit Bug 3: Graph has a known quirk where
  // certain --select AND --orderby values on the tasks endpoint return the
  // opaque `RequestBroker--ParseUri: Invalid request` with no recovery
  // hint. `--select id,title` is the canonical failing combo for --select;
  // `--orderby "title asc"` trips the same parser. Both rewrite to the
  // same hint pointing at the title-quirk workaround.
  if (result.error.type === 'api_error' && result.error.message.includes(PARSE_URI_NEEDLE)) {
    if (parsed.data.select !== undefined) {
      return err({
        type: 'api_error',
        status: result.error.status,
        message: `Graph rejected --select=${parsed.data.select} on this tasks endpoint with RequestBroker--ParseUri (known quirk — some field combinations are unsupported, most reliably any combo that includes \`title\`). Drop \`title\` from --select and request it in a second call (or per-task via get-todo-task), or call this command without --select and slim the response client-side.`,
      });
    }
    if (parsed.data.orderby !== undefined) {
      return err({
        type: 'api_error',
        status: result.error.status,
        message: `Graph rejected --orderby=${parsed.data.orderby} on this tasks endpoint with RequestBroker--ParseUri (known quirk — sorting on \`title\` is unsupported). Call this command without --orderby and sort the response client-side, or order by a numeric/date field like \`createdDateTime\` / \`importance\` instead.`,
      });
    }
  }
  return result;
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
