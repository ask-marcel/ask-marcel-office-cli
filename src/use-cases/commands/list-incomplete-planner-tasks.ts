import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';

// Hardcoded `$filter=percentComplete ne 100` in the path means a user-supplied
// `--filter` would cause Graph to receive two `$filter` query params. Audit
// round-6 §2.7: accept --filter at the schema layer so we can return a
// helpful "use list-planner-tasks instead" error rather than Commander's
// generic "unknown option".
const schema = z.object({}).extend(odataQuerySchema.shape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  if (parsed.data.filter !== undefined) {
    return err({
      type: 'validation_error',
      message:
        '--filter is not supported on list-incomplete-planner-tasks: the path already pins `$filter=percentComplete ne 100` and Graph rejects two $filter query params. To combine with your own filter, call `list-planner-tasks --filter "percentComplete ne 100 and <your-predicate>"` instead (single $filter, AND your predicate yourself).',
    });
  }
  const path = appendOData('/me/planner/tasks?$filter=percentComplete ne 100', parsed.data);
  return graph.get(path);
};

const meta: CommandMeta = {
  summary:
    'List every incomplete Microsoft Planner task assigned to or owned by the signed-in user, across every plan. Accepts the OData passthrough flags top/skip/select/orderby/expand. The filter passthrough is intentionally omitted — the path already pins a `$filter` for the completion-percent predicate, and Graph rejects two `$filter` query params. If you supply `--filter` anyway, the CLI returns a clear pointer to `list-planner-tasks`.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/planner/tasks?$filter=percentComplete ne 100',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/planneruser-list-tasks',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-incomplete-planner-tasks --top 25',
  responseShape: 'collection of Microsoft Graph `plannerTask` resources under `value[]` where `percentComplete < 100`',
  pagination: true,
};

export { execute, meta, schema };
