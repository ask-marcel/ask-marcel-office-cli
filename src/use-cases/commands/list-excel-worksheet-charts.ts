import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), worksheetId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/worksheets/${p.worksheetId}/charts`, baseSchema);

const meta: CommandMeta = {
  summary:
    "List the charts on a worksheet. Each `workbookChart` has `id`, `name`, `height`, `width`, `top`, `left`. Use the chart's image endpoint (`.../charts/{id}/image()`) to render the chart as a base64 PNG.",
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/charts',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/worksheet-list-charts',
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
  example: "ask-marcel list-excel-worksheet-charts --drive-id 'b!1234' --item-id '01ABC' --worksheet-id 'Sheet1'",
  responseShape: 'collection of Microsoft Graph `workbookChart` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
