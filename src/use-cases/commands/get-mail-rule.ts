import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ mailFolderId: z.string().min(1), messageRuleId: z.string().min(1) });
const { execute } = buildCommand((p) => `/me/mailFolders/${p.mailFolderId}/messageRules/${p.messageRuleId}`, schema);

const meta: CommandMeta = {
  summary: 'Return a single Outlook message rule by ID, including its conditions and actions. Sibling to `list-mail-rules`.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders/{mail-folder-id}/messageRules/{message-rule-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/messagerule-get',
  options: [
    {
      name: 'mail-folder-id',
      key: 'mailFolderId',
      required: true,
      description: 'Mail folder ID or well-known name (`inbox`, `sentitems`, etc.). Rules live per-folder.',
    },
    {
      name: 'message-rule-id',
      key: 'messageRuleId',
      required: true,
      description: 'Message rule ID. Returned by `list-mail-rules`.',
    },
  ],
  example: "ask-marcel get-mail-rule --mail-folder-id 'inbox' --message-rule-id 'AQAAANC...'",
  responseShape: 'single Microsoft Graph `messageRule` resource',
};

export { execute, meta, schema };
