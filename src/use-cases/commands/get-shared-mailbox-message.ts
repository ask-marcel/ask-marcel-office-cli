import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ userId: z.string().min(1), messageId: z.string().min(1) });
const { execute } = buildCommand((p) => `/users/${p.userId}/messages/${p.messageId}`, schema);

const meta: CommandMeta = {
  summary: 'Return a single message from a shared / delegated mailbox.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/users/{user-id}/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [
    {
      name: 'user-id',
      key: 'userId',
      required: true,
      description: 'Azure AD user ID or UPN of the shared mailbox.',
    },
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      description: 'Outlook message ID inside that mailbox.',
    },
  ],
  example: "ask-marcel get-shared-mailbox-message --user-id 'shared-mailbox@contoso.com' --message-id 'AAMkAD...'",
  responseShape: 'single Microsoft Graph `message` resource',
};

export { execute, meta, schema };
