import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, selectExpandOptions, selectExpandSchema } from './odata-query.ts';

const schema = z.object({}).extend(selectExpandSchema.shape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData('/me/manager', parsed.data);
  const result = await graph.get(path);
  if (result.ok) return result;
  if (result.error.type === 'api_error' && result.error.status === 404 && result.error.message.includes('Request_ResourceNotFound')) {
    // Audit round-8 H1: previously returned `null`, which in text mode
    // renders as the bare literal `null` — an LLM can't tell whether the
    // command meant "no manager", "permission failed", or "empty payload".
    // Wrap in a discriminated object so both text and JSON modes carry
    // enough context: `{ manager: null, note: '...' }`.
    return ok({ manager: null, note: 'signed-in user has no manager set in the directory (Graph returned 404 Request_ResourceNotFound)' });
  }
  return result;
};

const meta: CommandMeta = {
  summary:
    "Return the signed-in user's manager (a single `user` resource). When no manager is set in the directory, Graph returns 404 `Request_ResourceNotFound`; this command maps that one specific 404 to `{ ok: true, data: { manager: null, note: '...' } }` so an LLM can distinguish 'no manager' from a permission failure without parsing prose. Use `--select` to slim the response (e.g. `--select id,displayName,mail`).",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/manager',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-manager',
  options: [...selectExpandOptions],
  example: "ask-marcel get-my-manager --select 'id,displayName,mail'",
  responseShape:
    'single Microsoft Graph `user` resource on success, OR `{ manager: null, note: <string> }` when the signed-in user has no manager set. Detect the no-manager case via `data.manager === null` (also: `data.note` carries a human description).',
};

export { execute, meta, schema };
