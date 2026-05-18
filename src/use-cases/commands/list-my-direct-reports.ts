import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';

const schema = z.object({}).strict().extend(odataQuerySchema.shape);

// Graph's directory endpoints (`/me/directReports`, `/users`, `/groups`)
// require `ConsistencyLevel: eventual` for `$orderby` / `$count` /
// advanced `$filter` clauses — without it, Graph returns
// `Request_UnsupportedQuery: ... 'ConsistencyLevel:eventual' header is
// missing`. Auto-inject the header when --orderby is supplied so the
// caller doesn't have to know about Microsoft's "advanced query" gate.
const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const headers: Record<string, string> = parsed.data.orderby !== undefined ? { ConsistencyLevel: 'eventual' } : {};
  const path = appendOData('/me/directReports', parsed.data);
  return graph.get(path, headers);
};

const meta: CommandMeta = {
  summary:
    "List the signed-in user's direct reports (employees who report to them in the directory). When `--orderby` is supplied the CLI auto-injects the `ConsistencyLevel: eventual` header Graph requires on directory endpoints — otherwise Graph rejects the sort with `Request_UnsupportedQuery`.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/directReports',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-directreports',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-my-direct-reports',
  responseShape: 'collection of Microsoft Graph `directoryObject` resources (typically `user`) under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
