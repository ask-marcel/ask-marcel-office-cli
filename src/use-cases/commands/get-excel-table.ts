import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { wrapExcelExecute } from './excel-error.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), tableId: z.string().min(1) });
const inner = buildCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/tables/${p.tableId}`, schema);
const execute = wrapExcelExecute(inner.execute);

const meta: CommandMeta = {
  summary: 'Get the metadata (style, header row, total row) of a single named Excel table.',
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/workbooktable-get',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID containing the workbook. Returned by `ask-marcel list-drives`.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .xlsx file.' },
    { name: 'table-id', key: 'tableId', required: true, description: 'Workbook table ID or table name. Returned by `ask-marcel list-excel-tables`.' },
  ],
  example: "ask-marcel get-excel-table --drive-id 'b!1234' --item-id '01XLSX' --table-id 'Table1'",
  responseShape: 'single Microsoft Graph `workbookTable` resource',
};

export { execute, meta, schema };
