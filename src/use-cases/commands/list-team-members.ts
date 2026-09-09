import { z } from 'zod';
import { buildFilterSelectListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { filterSelectOptions } from './odata-query.ts';

const baseSchema = z.object({ teamId: z.string().min(1) });
// Probed live 2026-09-09 on the basic token: `$select` and `$filter` reach the
// server; `$skip` is a 400 (`Query option 'Skip' is not allowed`) and `$top`
// mis-pages (`$top=1` answers an empty page, `$top=5` two members and a
// nextLink), so neither is advertised. Large teams page through
// `@odata.nextLink` on their own.
const { execute, schema } = buildFilterSelectListCommand((p) => `/teams/${p.teamId}/members`, baseSchema);

const meta: CommandMeta = {
  summary:
    'List the members of a Microsoft Team with their roles: `conversationMember` entries carrying `displayName`, `email`, `userId`, `tenantId` and `roles` (`owner`, `guest`, or empty for a plain member). The sibling of `list-group-members` for the team itself rather than its backing group, and the answer to "who is in this team". `--filter` narrows server-side, e.g. owners only with `--filter "(microsoft.graph.aadUserConversationMember/roles/any(r:r eq \'owner\'))"`. Only `--filter` and `--select` reach Graph here: the endpoint rejects `$skip` and mis-pages on `$top`, so a large team continues through the `next:` footer with `next-page`.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/members',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/team-list-members',
  options: [
    {
      name: 'team-id',
      key: 'teamId',
      required: true,
      description: 'Microsoft Teams team ID. Returned by `ask-marcel-office list-joined-teams`.',
    },
    ...filterSelectOptions,
  ],
  example: "ask-marcel-office list-team-members --team-id 'abc-1234-...'",
  responseShape:
    'collection of Microsoft Graph `aadUserConversationMember` resources under `value[]`: `id`, `roles[]`, `displayName`, `userId`, `email`, `tenantId`, `visibleHistoryStartDateTime`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
