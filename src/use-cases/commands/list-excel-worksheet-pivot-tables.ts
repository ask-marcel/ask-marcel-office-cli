import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), worksheetId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/worksheets/${p.worksheetId}/pivotTables`, baseSchema);

const meta: CommandMeta = {
  summary:
    'List the pivot tables on a worksheet. Each `workbookPivotTable` has `name` and a navigation to its source `workbookWorksheet`. Useful for understanding analytical structure inside a workbook.',
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/pivotTables',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/worksheet-list-pivottables',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description: 'OneDrive / SharePoint drive ID.',
    },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID of the .xlsx file.',
    },
    {
      name: 'worksheet-id',
      key: 'worksheetId',
      required: true,
      description: 'Worksheet ID or name.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-excel-worksheet-pivot-tables --drive-id 'b!1234' --item-id '01ABC' --worksheet-id 'Sheet1'",
  responseShape: 'collection of Microsoft Graph `workbookPivotTable` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
