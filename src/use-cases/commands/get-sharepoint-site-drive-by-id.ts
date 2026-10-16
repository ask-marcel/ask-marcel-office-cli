import { z } from 'zod';
import { buildSelectableCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { selectExpandOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({ siteId: z.string().min(1), driveId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/sites/${p.siteId}/drives/${p.driveId}`, baseSchema);

const meta: ReadCommandMeta = {
  summary: 'Get the metadata of a single document library (drive) on a SharePoint site by drive ID.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/drives/{drive-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-get',
  options: [
    { name: 'site-id', key: 'siteId', required: true, description: 'SharePoint site ID. Returned by `{read-bin} search-sharepoint-sites-by-name`.' },
    { name: 'drive-id', key: 'driveId', required: true, description: 'Drive (document library) ID on the site. Returned by `{read-bin} list-sharepoint-site-drives`.' },
    ...selectExpandOptions,
  ],
  example: "{read-bin} get-sharepoint-site-drive-by-id --site-id 'contoso.sharepoint.com,1234,5678' --drive-id 'b!abcd'",
  responseShape: 'single Microsoft Graph `drive` resource',
  effect: 'read',
  scopesRequired: ['Sites.Read.All', 'Files.Read'],
};

export { execute, meta, schema };
