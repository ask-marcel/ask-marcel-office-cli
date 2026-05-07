import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ siteId: z.string().min(1), driveId: z.string().min(1) });
const { execute } = buildCommand((p) => `/sites/${p.siteId}/drives/${p.driveId}`, schema);

const meta: CommandMeta = {
  summary: 'Get the metadata of a single document library (drive) on a SharePoint site by drive ID.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/drives/{drive-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-get',
  options: [
    { name: 'site-id', key: 'siteId', required: true, description: 'SharePoint site ID. Returned by `ask-marcel search-sharepoint-sites-by-name`.' },
    { name: 'drive-id', key: 'driveId', required: true, description: 'Drive (document library) ID on the site. Returned by `ask-marcel list-sharepoint-site-drives`.' },
  ],
  example: "ask-marcel get-sharepoint-site-drive-by-id --site-id 'contoso.sharepoint.com,1234,5678' --drive-id 'b!abcd'",
  responseShape: 'single Microsoft Graph `drive` resource',
};

export { execute, meta, schema };
