import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, selectExpandOptions, selectExpandSchema } from './odata-query.ts';

const schema = z.object({ userId: z.string().min(1) }).extend(selectExpandSchema.shape);

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
// An Azure AD object id or a UPN/email is a valid `/users/{id}` path segment;
// a bare display name is not, so it routes to the People API search instead.
const isIdentifier = (value: string): boolean => GUID.test(value) || value.includes('@');

const str = (value: unknown): string | undefined => (typeof value === 'string' ? value : undefined);

type PersonHit = {
  readonly id?: unknown;
  readonly displayName?: unknown;
  readonly jobTitle?: unknown;
  readonly department?: unknown;
  readonly scoredEmailAddresses?: ReadonlyArray<{ readonly address?: unknown }>;
};

const toCandidate = (p: PersonHit): Record<string, string | undefined> => ({
  id: str(p.id),
  displayName: str(p.displayName),
  mail: Array.isArray(p.scoredEmailAddresses) ? str(p.scoredEmailAddresses[0]?.address) : undefined,
  jobTitle: str(p.jobTitle),
  department: str(p.department),
});

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { userId } = parsed.data;

  // id / UPN / email → the FULL directory profile via the elevated /users/{id} path
  // (User.Read.All carries jobTitle / department / office / phones). Fail-fasts with
  // secondary_token_unavailable when the elevated token is cold — the error says to log in.
  // encodeURIComponent is load-bearing: a guest UPN is `alice_x.com#EXT#@tenant.onmicrosoft.com`,
  // and a raw `#` is the URL fragment delimiter — fetch would drop everything from it and query
  // the wrong user. Encoding (`#`→`%23`, `@`→`%40`) preserves the whole id; GUIDs pass unchanged.
  if (isIdentifier(userId)) return graph.getElevated(appendOData(`/users/${encodeURIComponent(userId)}`, parsed.data));

  // A bare name → search the signed-in user's relevant-people graph (People API, basic
  // token, so it works even when the elevated token is cold). Returns candidates with
  // AAD ids so the caller disambiguates and re-queries by id for the full profile.
  const result = await graph.get(`/me/people?$search="${userId.replaceAll('"', '')}"`);
  if (!result.ok) return result;
  const value = (result.value as { readonly value?: ReadonlyArray<PersonHit> }).value;
  return ok({ query: userId, matches: (Array.isArray(value) ? value : []).map(toCandidate) });
};

const meta: CommandMeta = {
  summary:
    "Look up a directory user. Pass an Azure AD id, UPN, or email as --user-id and get that user's FULL profile (displayName, mail, jobTitle, department, officeLocation, phones) via GET /users/{id} on the elevated M365 token — run a forced login first if it is cold. Pass a NAME instead and it searches your relevant-people graph (GET /me/people) and returns candidate matches (id, displayName, mail, jobTitle, department) so you can pick the right person and re-query by their id. Name search covers colleagues in your people graph, not the whole tenant directory; use `microsoft-search-query` for a broader tenant-wide person search.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}  (id / UPN / email) OR /me/people?$search="{user-id}"  (a bare name)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-get',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      aliases: [{ name: 'id', key: 'id' }],
      description:
        'Azure AD user ID, UPN, or email (returns the full profile via the elevated token), OR a display name (returns relevant-people candidates on the basic token). Discover ids via `list-relevant-people` or `microsoft-search-query`.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel-office get-user --user-id 'alice@contoso.com' --select 'id,displayName,mail,jobTitle,department'",
  responseShape:
    'For an id / UPN / email: a single Microsoft Graph `user` resource (full profile; honours `--select` / `--expand`). For a name: `{ query, matches: [{ id, displayName, mail, jobTitle, department }] }` from the People API (empty `matches` when nobody in your relevant-people graph matches; re-query by a chosen `id` for the full card). The id path needs the elevated M365 token and fail-fasts with `secondary_token_unavailable` when it is cold — run `ask-marcel-office login --force`.',
  needsElevatedToken: true,
};

export { execute, meta, schema };
