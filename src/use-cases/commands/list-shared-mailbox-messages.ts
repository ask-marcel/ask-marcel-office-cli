import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ userId: z.string().min(1) });
// encodeURIComponent(userId) so a guest UPN (`…#EXT#@…`) is not cut at the raw `#` fragment delimiter.
const { execute, schema } = buildListCommand((p) => `/users/${encodeURIComponent(p.userId)}/messages`, baseSchema);

const meta: CommandMeta = {
  summary:
    "List messages from a shared or delegated mailbox the signed-in user has read access to. Same shape as `list-mail-messages` but scoped to a specific mailbox owner. Requires the delegated `Mail.Read.Shared` scope, which neither token this CLI can mint carries (verified live on two tenants, 2026-08-30), so any mailbox other than the signed-in user's own is expected to answer `ErrorAccessDenied` whatever delegation Exchange holds. Your own UPN works; a Microsoft 365 group's mailbox is the shared-mail path that does (`list-group-conversations`, `list-group-thread-posts`).",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-messages',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      description: 'Azure AD user ID or UPN of the shared mailbox or delegated user. The signed-in user must have `Mail.Read.Shared` access (granted by the mailbox owner).',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel-office list-shared-mailbox-messages --user-id 'shared-mailbox@contoso.com'",
  responseShape: 'collection of Microsoft Graph `message` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
