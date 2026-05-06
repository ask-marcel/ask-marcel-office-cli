import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), worksheetId: z.string().min(1) });
const { execute } = buildCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/worksheets/${p.worksheetId}/usedRange()`, schema);

const meta: CommandMeta = {
  summary:
    "Return the worksheet's used range — the bounding box of every non-empty cell — as a single Excel range. Avoids fetching the entire 1M × 16K-cell sheet when only a small data island is populated.",
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}/usedRange()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/worksheet-usedrange',
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
      description: 'Worksheet ID or name. Returned by `list-excel-worksheets`.',
    },
  ],
  example: "ask-marcel get-excel-used-range --drive-id 'b!1234' --item-id '01ABC' --worksheet-id 'Sheet1'",
  responseShape: 'single Microsoft Graph `workbookRange` resource (`address`, `rowCount`, `columnCount`, `values`, `formulas`, etc.)',
};

export { execute, meta, schema };
