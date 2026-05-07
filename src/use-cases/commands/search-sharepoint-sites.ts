import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/followedSites', baseSchema);

const meta: CommandMeta = {
  summary:
    'List the SharePoint sites the signed-in user has explicitly followed. Hits `GET /me/followedSites` (the unauthenticated `GET /sites` returns an empty collection in most tenants — for free-text discovery use `search-sharepoint-sites-by-name`).',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/me/followedSites',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-followedsites',
  options: [...odataQueryOptions],
  example: 'ask-marcel search-sharepoint-sites',
  responseShape: 'collection of Microsoft Graph `site` resources (the user’s followed sites) under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
