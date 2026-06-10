import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { wrapExcelExecute } from './excel-error.ts';
import { pickODataOptions } from './odata-query.ts';

// Graph's `/workbook/tables/{id}/rows` honors `$top`, `$skip`, `$select`,
// `$expand` but silently ignores `$filter` and `$orderby`. Drop the
// silent no-ops.
const TABLE_ROWS_ODATA_KEYS = ['top', 'skip', 'select', 'expand'] as const;
const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), tableId: z.string().min(1) });
const inner = buildPickODataListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/tables/${p.tableId}/rows`, baseSchema, TABLE_ROWS_ODATA_KEYS);
const execute = wrapExcelExecute(inner.execute);
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    'List the data rows of a named Excel table (excluding the header row). Note: Graph silently ignores `$filter` and `$orderby` on this endpoint, so the CLI does not expose those flags — slice / sort client-side.',
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}/rows',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/workbooktable-list-rows',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID containing the workbook. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .xlsx file.' },
    { name: 'table-id', key: 'tableId', required: true, description: 'Workbook table ID or table name. Returned by `ask-marcel list-excel-tables`.' },
    ...pickODataOptions(TABLE_ROWS_ODATA_KEYS),
  ],
  example: "ask-marcel list-excel-table-rows --drive-id 'b!1234' --item-id '01XLSX' --table-id 'Table1'",
  responseShape: 'collection of Microsoft Graph `workbookTableRow` resources (each `values` is a 2D array) under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
