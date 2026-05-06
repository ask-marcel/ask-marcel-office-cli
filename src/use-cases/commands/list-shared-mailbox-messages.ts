import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ userId: z.string().min(1) });
const { execute } = buildCommand((p) => `/users/${p.userId}/messages`, schema);

const meta: CommandMeta = {
  summary:
    'List messages from a shared or delegated mailbox the signed-in user has read access to. Same shape as `list-mail-messages` but scoped to a specific mailbox owner. 403 if the signed-in user does not have shared access to that mailbox.',
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
  ],
  example: "ask-marcel list-shared-mailbox-messages --user-id 'shared-mailbox@contoso.com'",
  responseShape: 'collection of Microsoft Graph `message` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
