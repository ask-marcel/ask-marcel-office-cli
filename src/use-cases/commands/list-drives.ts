import { z } from 'zod';
import { buildListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { odataQueryOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/drives', baseSchema);

const meta: ReadCommandMeta = {
  summary:
    "List all OneDrive / SharePoint drives the signed-in user has access to. On personal accounts this returns only the user's primary OneDrive (single entry in `value[]`); on tenanted accounts it includes every drive the user can reach including delegated mailboxes and shared SharePoint document libraries.",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drives',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-list',
  options: [...odataQueryOptions],
  example: '{read-bin} list-drives',
  responseShape: 'collection of Microsoft Graph `drive` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Files.Read'],
};

export { execute, meta, schema };
