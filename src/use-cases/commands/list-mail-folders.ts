import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { INCLUDE_HIDDEN_FOLDERS_OPTION } from './include-hidden-folders.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ includeHiddenFolders: z.enum(['true', 'false']).optional() }).strict();
// Plain (non-OData) query param, so it cannot ride appendOData's `$`-prefixed
// builder and has to be emitted by the path itself.
// Graph's default page is ten folders: an eleventh top-level folder sat unseen
// on page two for a caller that did not follow the cursor. A hundred covers a
// normal mailbox in one call; `next:` still carries a larger one.
const { execute, schema } = buildListCommand((p) => (p.includeHiddenFolders === 'true' ? '/me/mailFolders?includeHiddenFolders=true' : '/me/mailFolders'), baseSchema, {
  defaultTop: '100',
});

const meta: CommandMeta = {
  summary:
    'List the top-level mail folders in the signed-in user’s Outlook mailbox (Inbox, Sent Items, etc.). The CLI asks for 100 folders a page (Graph’s own default is ten, which once hid an eleventh folder); a larger mailbox continues through the `next:` footer, and `--top` sets the page. Child folders come from `list-mail-child-folders`, and `list-mail-folders-delta` lists every folder at every depth in one walk.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders',
  options: [INCLUDE_HIDDEN_FOLDERS_OPTION, ...odataQueryOptions],
  example: 'ask-marcel-office list-mail-folders',
  responseShape: 'collection of Microsoft Graph `mailFolder` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Mail.Read'],
};

export { execute, meta, schema };
