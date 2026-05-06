import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ userId: z.string().min(1), mailFolderId: z.string().min(1) });
const { execute } = buildCommand((p) => `/users/${p.userId}/mailFolders/${p.mailFolderId}/messages`, schema);

const meta: CommandMeta = {
  summary: 'List messages in a single folder of a shared / delegated mailbox.',
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
  ],
  example: "ask-marcel list-shared-mailbox-folder-messages --user-id 'shared-mailbox@contoso.com' --mail-folder-id 'inbox'",
  responseShape: 'collection of Microsoft Graph `message` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
