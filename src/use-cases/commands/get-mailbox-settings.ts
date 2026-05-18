import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

// Graph's `/me/mailboxSettings` silently ignores `$select` (and `$expand`):
// the full payload — including the ~3 KB autoReply HTML body — is always
// returned regardless. Don't advertise the flag.
const schema = z.object({});
const { execute } = buildCommand(() => '/me/mailboxSettings', schema);

const meta: CommandMeta = {
  summary:
    "Get the signed-in user's Outlook mailbox settings (timezone, working hours, automatic replies). Note: Graph silently ignores `$select` / `$expand` on this endpoint, so the CLI does NOT expose them — the full payload (including the auto-reply HTML body) is always returned. Slim client-side if you only need a subset.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailboxSettings',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-get-mailboxsettings',
  options: [],
  example: 'ask-marcel get-mailbox-settings',
  responseShape: 'single Microsoft Graph `mailboxSettings` resource',
};

export { execute, meta, schema };
