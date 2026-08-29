import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { INCLUDE_HIDDEN_FOLDERS_OPTION } from './include-hidden-folders.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ userId: z.string().min(1), mailFolderId: z.string().min(1), includeHiddenFolders: z.enum(['true', 'false']).optional() });
// encodeURIComponent(userId) so a guest UPN (`…#EXT#@…`) is not cut at the raw `#` fragment delimiter.
// `includeHiddenFolders` is a plain (non-OData) query param, so it is emitted here rather than by appendOData.
const { execute, schema } = buildListCommand(
  (p) => `/users/${encodeURIComponent(p.userId)}/mailFolders/${p.mailFolderId}/childFolders${p.includeHiddenFolders === 'true' ? '?includeHiddenFolders=true' : ''}`,
  baseSchema
);

const meta: CommandMeta = {
  summary:
    'List the subfolders of one mail folder in a shared or delegated mailbox. The `/me` sibling is `list-mail-child-folders`. Walk it from the folder IDs `list-shared-mailbox-folders` returns to reach nested custom folders. 403 if the signed-in user does not have shared access to that mailbox.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}/mailFolders/{mail-folder-id}/childFolders',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/mailfolder-list-childfolders',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      description: 'Azure AD user ID or UPN of the shared mailbox or delegated user. The signed-in user must have `Mail.Read.Shared` access (granted by the mailbox owner).',
    },
    {
      name: 'mail-folder-id',
      key: 'mailFolderId',
      required: true,
      description: 'Mail folder ID or well-known name (`inbox`, `sentitems`, etc.) inside that mailbox. Returned by `ask-marcel-office list-shared-mailbox-folders`.',
    },
    INCLUDE_HIDDEN_FOLDERS_OPTION,
    ...odataQueryOptions,
  ],
  example: "ask-marcel-office list-shared-mailbox-child-folders --user-id 'shared-mailbox@contoso.com' --mail-folder-id 'inbox'",
  responseShape: 'collection of Microsoft Graph `mailFolder` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
