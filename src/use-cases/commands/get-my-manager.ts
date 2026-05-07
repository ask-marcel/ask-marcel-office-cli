import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({}).strict();

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const result = await graph.get('/me/manager');
  if (result.ok) return result;
  if (result.error.type === 'api_error' && result.error.status === 404 && result.error.message.includes('Request_ResourceNotFound')) {
    return ok(null);
  }
  return result;
};

const meta: CommandMeta = {
  summary:
    "Return the signed-in user's manager (a single `user` resource). When no manager is set in the directory, Graph returns 404 `Request_ResourceNotFound`; this command maps that one specific 404 to `{ ok: true, data: null }` so an LLM can distinguish 'no manager' from a permission failure.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/manager',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-manager',
  options: [],
  example: 'ask-marcel get-my-manager',
  responseShape: 'single Microsoft Graph `user` resource, or `null` when no manager is set in the directory',
};

export { execute, meta, schema };
