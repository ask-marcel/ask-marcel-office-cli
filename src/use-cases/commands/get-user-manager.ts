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
  // Disambiguate two distinct 404 cases — both are `Request_ResourceNotFound`
  // and both say "does not exist", but they differ in WHICH resource is
  // reported as missing in the quotes:
  //   - "Resource 'manager' does not exist …"          → user exists, no manager set     → ok(null)
  //   - "Resource '<userId>' does not exist …"         → user does NOT exist              → pass through err
  // (The first audit pass v1.0.0 §1.3 caught us collapsing both to ok(null);
  // the round-2 audit then caught the over-corrected version that mapped
  // NEITHER to ok(null). The right discriminator is the literal `'manager'`
  // quoted name — it's the navigation-property name Graph reports as
  // missing when the user record exists but has no manager link.)
  if (
    result.error.type === 'api_error' &&
    result.error.status === 404 &&
    result.error.message.includes('Request_ResourceNotFound') &&
    result.error.message.includes("Resource 'manager'")
  ) {
    // Audit v1.0.0 §B7: shape matches get-my-manager so consumers can
    // detect "no manager set" with the same check on either command.
    return ok({ manager: null, note: 'target user has no manager set in the directory (Graph returned 404 Request_ResourceNotFound)' });
  }
  return result;
};

const meta: CommandMeta = {
  summary:
    "Return a specific user's manager (a single `user` resource). When the user has no manager set in the directory, Graph returns 404 `Request_ResourceNotFound`; this command maps that one specific 404 to `{ ok: true, data: { manager: null, note: '...' } }` (same shape as `get-my-manager`) so an LLM can distinguish 'no manager' from 'unknown user' with a single discriminator across both commands. Use `--select` to slim the response.",
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
  responseShape:
    'single Microsoft Graph `user` resource on success, OR `{ manager: null, note: <string> }` when the target user has no manager set. Detect the no-manager case via `data.manager === null` (same discriminator as `get-my-manager`).',
};

export { execute, meta, schema };
