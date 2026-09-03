import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ userId: z.string().min(1), messageId: z.string().min(1) });
// encodeURIComponent(userId) so a guest UPN (`…#EXT#@…`) is not cut at the raw `#` fragment delimiter.
const { execute, schema } = buildSelectableCommand((p) => `/users/${encodeURIComponent(p.userId)}/messages/${p.messageId}`, baseSchema);

const meta: CommandMeta = {
  summary:
    "Return a single message from a shared / delegated mailbox. Use `--select` to fetch only specific fields (e.g. `--select id,subject,from,receivedDateTime`) — sibling to `get-mail-message` for /me. Requires the delegated `Mail.Read.Shared` scope, which neither token this CLI can mint carries (verified live on two tenants, 2026-08-30), so any mailbox other than the signed-in user's own is expected to answer `ErrorAccessDenied` whatever delegation Exchange holds. Your own UPN works; a Microsoft 365 group's mailbox is the shared-mail path that does (`list-group-conversations`, `list-group-thread-posts`).",
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
    ...selectExpandOptions,
  ],
  example: "ask-marcel-office get-shared-mailbox-message --user-id 'shared-mailbox@contoso.com' --message-id 'AAMkAD...' --select 'id,subject,from'",
  responseShape: 'single Microsoft Graph `message` resource',
};

export { execute, meta, schema };
