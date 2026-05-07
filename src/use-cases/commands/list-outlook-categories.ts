import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/outlook/masterCategories', baseSchema);

const meta: CommandMeta = {
  summary:
    "List the signed-in user's Outlook color categories — the named tags that can be applied to mail, calendar items, and contacts. Each entry has `displayName` and a `color` from Outlook's preset palette.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/outlook/masterCategories',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/outlookuser-list-mastercategories',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-outlook-categories',
  responseShape: 'collection of Microsoft Graph `outlookCategory` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
