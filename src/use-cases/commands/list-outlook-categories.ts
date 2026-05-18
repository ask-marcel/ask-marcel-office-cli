import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

// Graph's `/me/outlook/masterCategories` silently ignores every standard
// OData passthrough (verified live — `--top 1` against a 14-category
// mailbox still returns all 14). Don't advertise flags Graph drops.
const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/outlook/masterCategories', schema);

const meta: CommandMeta = {
  summary:
    "List the signed-in user's Outlook color categories — the named tags that can be applied to mail, calendar items, and contacts. Each entry has `displayName` and a `color` from Outlook's preset palette. Note: Graph silently ignores every OData passthrough on this endpoint (`$top`, `$skip`, `$select`, `$filter`, `$orderby`, `$expand`), so the CLI does not expose any of those flags — the full collection is always returned. Slice client-side.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/outlook/masterCategories',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/outlookuser-list-mastercategories',
  options: [],
  example: 'ask-marcel list-outlook-categories',
  responseShape: 'collection of Microsoft Graph `outlookCategory` resources under `value[]`',
};

export { execute, meta, schema };
