import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';

// Hardcoded `$filter=percentComplete ne 100` in the path means a
// user-supplied `--filter` would cause Graph to receive two `$filter` query
// params. Expose the other five OData passthrough flags but not `--filter`.
const noFilterShape = Object.fromEntries(Object.entries(odataQuerySchema.shape).filter(([key]) => key !== 'filter')) as Omit<typeof odataQuerySchema.shape, 'filter'>;
const noFilterOptions = odataQueryOptions.filter((o) => o.name !== 'filter');

const schema = z.object({}).extend(noFilterShape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData('/me/planner/tasks?$filter=percentComplete ne 100', parsed.data);
  return graph.get(path);
};

const meta: CommandMeta = {
  summary:
    'List every incomplete Microsoft Planner task assigned to or owned by the signed-in user, across every plan. Accepts the OData passthrough flags top/skip/select/orderby/expand. The filter passthrough is intentionally omitted — the path already pins a `$filter` for the completion-percent predicate, and Graph rejects two `$filter` query params.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/planner/tasks?$filter=percentComplete ne 100',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/planneruser-list-tasks',
  options: [...noFilterOptions],
  example: 'ask-marcel list-incomplete-planner-tasks --top 25',
  responseShape: 'collection of Microsoft Graph `plannerTask` resources under `value[]` where `percentComplete < 100`',
  pagination: true,
};

export { execute, meta, schema };
