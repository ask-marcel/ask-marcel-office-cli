import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({});
const { execute, schema } = buildSelectableCommand(() => '/me/mailboxSettings', baseSchema);

const meta: CommandMeta = {
  summary:
    "Get the signed-in user's Outlook mailbox settings (timezone, working hours, automatic replies). Use `--select` to fetch only specific fields (e.g. `--select timeZone,workingHours`).",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailboxSettings',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-get-mailboxsettings',
  options: [...selectExpandOptions],
  example: 'ask-marcel get-mailbox-settings',
  responseShape: 'single Microsoft Graph `mailboxSettings` resource',
};

export { execute, meta, schema };
