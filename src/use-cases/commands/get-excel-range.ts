import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { wrapExcelExecute } from './excel-error.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), worksheetId: z.string().min(1), address: z.string().min(1) });
const inner = buildCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/worksheets/${p.worksheetId}/range(address='${p.address}')`, schema);
const execute = wrapExcelExecute(inner.execute);

const meta: CommandMeta = {
  summary: 'Get the cell values, formulas, and formats of a specific Excel range (e.g. `A1:C10`).',
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: "/drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/range(address='{address}')",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/worksheet-range',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID containing the workbook. Returned by `ask-marcel list-drives`.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .xlsx file.' },
    { name: 'worksheet-id', key: 'worksheetId', required: true, description: 'Worksheet ID or worksheet name. Returned by `ask-marcel list-excel-worksheets`.' },
    { name: 'address', key: 'address', required: true, description: 'A1-style range address, e.g. `A1:C10` or a single cell like `B7`.' },
  ],
  example: "ask-marcel get-excel-range --drive-id 'b!1234' --item-id '01XLSX' --worksheet-id 'Sheet1' --address 'A1:C10'",
  responseShape: 'single Microsoft Graph `workbookRange` resource (values, formulas, format)',
};

export { execute, meta, schema };
