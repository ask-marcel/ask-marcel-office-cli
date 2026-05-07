import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/tables`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the named tables across every worksheet in an Excel workbook.',
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/tables',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/workbook-list-tables',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID containing the workbook. Returned by `ask-marcel list-drives`.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .xlsx file. Returned by `list-folder-files` or `search-onedrive-files`.' },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-excel-tables --drive-id 'b!1234' --item-id '01XLSX'",
  responseShape: 'collection of Microsoft Graph `workbookTable` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
