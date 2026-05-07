import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, selectExpandOptions, selectExpandSchema } from './odata-query.ts';

const schema = z.object({ userId: z.string().min(1) }).extend(selectExpandSchema.shape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData(`/users/${parsed.data.userId}/manager`, parsed.data);
  const result = await graph.get(path);
  if (result.ok) return result;
  // Mirror `get-my-manager`: when no manager is set in the directory Graph
  // returns 404 `Request_ResourceNotFound`. Map that one specific case to
  // `{ ok: true, data: null }` so an LLM can distinguish 'no manager' from a
  // genuine 404 (e.g. unknown userId, which surfaces as
  // `Resource '<id>' does not exist`).
  if (result.error.type === 'api_error' && result.error.status === 404 && result.error.message.includes('Request_ResourceNotFound')) {
    return ok(null);
  }
  return result;
};

const meta: CommandMeta = {
  summary:
    "Return a specific user's manager (a single `user` resource). When the user has no manager set in the directory, Graph returns 404 `Request_ResourceNotFound`; this command maps that one specific 404 to `{ ok: true, data: null }` so an LLM can distinguish 'no manager' from 'unknown user'. Use `--select` to slim the response.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}/manager',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-manager',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      description:
        "Azure AD user ID or UPN — typically the user's email address. Discover via `list-relevant-people` (relevance-ranked colleagues) or `microsoft-search-query --query <name>` (federated person search across the tenant directory).",
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-user-manager --user-id 'alice@contoso.com' --select 'id,displayName,mail'",
  responseShape: 'single Microsoft Graph `user` resource, or `null` when no manager is set in the directory',
};

export { execute, meta, schema };
