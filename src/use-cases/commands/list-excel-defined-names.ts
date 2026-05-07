import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { wrapExcelExecute } from './excel-error.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const inner = buildListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/names`, baseSchema);
const execute = wrapExcelExecute(inner.execute);
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    "List the workbook's defined names (named ranges, named formulas, named constants). Each `workbookNamedItem` has `name`, `value` (the formula or address), `comment`, and `scope` (workbook or worksheet). Useful for understanding workbook structure before reading ranges.",
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/names',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/workbook-list-names',
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
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-excel-defined-names --drive-id 'b!1234' --item-id '01ABC'",
  responseShape: 'collection of Microsoft Graph `workbookNamedItem` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
