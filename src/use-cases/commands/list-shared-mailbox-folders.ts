import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { INCLUDE_HIDDEN_FOLDERS_OPTION } from './include-hidden-folders.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ userId: z.string().min(1), includeHiddenFolders: z.enum(['true', 'false']).optional() });
// encodeURIComponent(userId) so a guest UPN (`…#EXT#@…`) is not cut at the raw `#` fragment delimiter.
// `includeHiddenFolders` is a plain (non-OData) query param, so it is emitted here rather than by appendOData.
const { execute, schema } = buildListCommand(
  (p) => `/users/${encodeURIComponent(p.userId)}/mailFolders${p.includeHiddenFolders === 'true' ? '?includeHiddenFolders=true' : ''}`,
  baseSchema
);

const meta: CommandMeta = {
  summary:
    "List the top-level mail folders of a shared or delegated mailbox. The `/me` sibling is `list-mail-folders`. Use it to discover the folder IDs that `list-shared-mailbox-folder-messages` needs: without it only the well-known names (`inbox`, `sentitems`, `drafts`, …) are reachable, so custom folders are invisible. Requires the delegated `Mail.Read.Shared` scope, which neither token this CLI can mint carries (verified live on two tenants, 2026-08-30), so any mailbox other than the signed-in user's own is expected to answer `ErrorAccessDenied` whatever delegation Exchange holds. Your own UPN works; a Microsoft 365 group's mailbox is the shared-mail path that does (`list-group-conversations`, `list-group-thread-posts`).",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}/mailFolders',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      description: 'Azure AD user ID or UPN of the shared mailbox or delegated user. The signed-in user must have `Mail.Read.Shared` access (granted by the mailbox owner).',
    },
    INCLUDE_HIDDEN_FOLDERS_OPTION,
    ...odataQueryOptions,
  ],
  example: "ask-marcel-office list-shared-mailbox-folders --user-id 'shared-mailbox@contoso.com'",
  responseShape: 'collection of Microsoft Graph `mailFolder` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
