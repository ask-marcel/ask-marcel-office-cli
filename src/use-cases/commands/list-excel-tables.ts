import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { wrapExcelExecute } from './excel-error.ts';
import { pickODataOptions } from './odata-query.ts';

// Graph's `/workbook/tables` honors `$top`, `$skip`, `$select`, `$expand`
// but silently ignores `$filter` and `$orderby` (verified live). Drop the
// silent no-ops from the option set rather than advertising lies.
const TABLES_ODATA_KEYS = ['top', 'skip', 'select', 'expand'] as const;
const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const inner = buildPickODataListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/tables`, baseSchema, TABLES_ODATA_KEYS);
const execute = wrapExcelExecute(inner.execute);
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    'List the named tables across every worksheet in an Excel workbook. Note: Graph silently ignores `$filter` and `$orderby` on this endpoint, so the CLI does not expose those flags — slice / sort client-side.',
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/tables',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/workbook-list-tables',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID containing the workbook. Returned by `ask-marcel list-drives`.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .xlsx file. Returned by `list-folder-files` or `search-onedrive-files`.' },
    ...pickODataOptions(TABLES_ODATA_KEYS),
  ],
  example: "ask-marcel list-excel-tables --drive-id 'b!1234' --item-id '01XLSX'",
  responseShape: 'collection of Microsoft Graph `workbookTable` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
