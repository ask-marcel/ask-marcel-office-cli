import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), tableId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/tables/${p.tableId}/rows`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the data rows of a named Excel table (excluding the header row).',
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}/rows',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/workbooktable-list-rows',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID containing the workbook. Returned by `ask-marcel list-drives`.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .xlsx file.' },
    { name: 'table-id', key: 'tableId', required: true, description: 'Workbook table ID or table name. Returned by `ask-marcel list-excel-tables`.' },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-excel-table-rows --drive-id 'b!1234' --item-id '01XLSX' --table-id 'Table1'",
  responseShape: 'collection of Microsoft Graph `workbookTableRow` resources (each `values` is a 2D array) under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
