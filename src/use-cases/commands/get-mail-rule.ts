import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ mailFolderId: z.string().min(1).default('inbox'), messageRuleId: z.string().min(1) });
const { execute } = buildCommand((p) => `/me/mailFolders/${p.mailFolderId}/messageRules/${p.messageRuleId}`, schema);

const meta: CommandMeta = {
  summary:
    'Return a single Outlook message rule by ID, including its conditions and actions. Sibling to `list-mail-rules`. `--mail-folder-id` defaults to `inbox` (the only folder where rules actually live in Graph); the flag is preserved for callers that want to pass a resolved Inbox ID explicitly.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders/{mail-folder-id}/messageRules/{message-rule-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/messagerule-get',
  options: [
    {
      name: 'mail-folder-id',
      key: 'mailFolderId',
      required: false,
      description: 'Mail folder ID or well-known name. Optional; defaults to `inbox` because that is the only folder Graph supports for message rules.',
    },
    {
      name: 'message-rule-id',
      key: 'messageRuleId',
      required: true,
      description: 'Message rule ID. Returned by `list-mail-rules`.',
      aliases: [{ name: 'rule-id', key: 'ruleId' }],
    },
  ],
  example: "ask-marcel get-mail-rule --message-rule-id 'AQAAANC...'",
  responseShape: 'single Microsoft Graph `messageRule` resource',
};

export { execute, meta, schema };
