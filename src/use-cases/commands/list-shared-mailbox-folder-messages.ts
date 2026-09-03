import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ userId: z.string().min(1), mailFolderId: z.string().min(1) });
// encodeURIComponent(userId) so a guest UPN (`…#EXT#@…`) is not cut at the raw `#` fragment delimiter.
const { execute, schema } = buildListCommand((p) => `/users/${encodeURIComponent(p.userId)}/mailFolders/${p.mailFolderId}/messages`, baseSchema);

const meta: CommandMeta = {
  summary:
    "List messages in a single folder of a shared / delegated mailbox. Requires the delegated `Mail.Read.Shared` scope, which neither token this CLI can mint carries (verified live on two tenants, 2026-08-30), so any mailbox other than the signed-in user's own is expected to answer `ErrorAccessDenied` whatever delegation Exchange holds. Your own UPN works; a Microsoft 365 group's mailbox is the shared-mail path that does (`list-group-conversations`, `list-group-thread-posts`).",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}/mailFolders/{mail-folder-id}/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messages',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      description: 'Azure AD user ID or UPN of the shared mailbox.',
    },
    {
      name: 'mail-folder-id',
      key: 'mailFolderId',
      required: true,
      description: 'Mail folder ID or well-known name (`inbox`, `sentitems`, etc.) inside that mailbox.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel-office list-shared-mailbox-folder-messages --user-id 'shared-mailbox@contoso.com' --mail-folder-id 'inbox'",
  responseShape: 'collection of Microsoft Graph `message` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
