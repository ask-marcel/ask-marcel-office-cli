import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { wrapExcelExecute } from './excel-error.ts';
import { pickODataOptions } from './odata-query.ts';

// Graph's `/workbook/worksheets` honors `$skip`, `$select`, `$expand` but
// silently ignores `$top`, `$filter`, `$orderby` (verified live —
// `--top 3` against a 14-sheet workbook still returns all 14). Drop the
// silent no-ops from the option set.
const WORKSHEETS_ODATA_KEYS = ['skip', 'select', 'expand'] as const;
const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const inner = buildPickODataListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/worksheets`, baseSchema, WORKSHEETS_ODATA_KEYS);
const execute = wrapExcelExecute(inner.execute);
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    'List the worksheets (tabs) inside an Excel workbook stored in OneDrive / SharePoint. Returns a clear "not an accessible Excel workbook" error if the item is a folder, non-.xlsx file, or sensitivity-label-blocked. Note: Graph silently ignores `$top`, `$filter`, and `$orderby` on this endpoint, so the CLI does not expose those flags — slice / sort client-side.',
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/worksheets',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/workbook-list-worksheets',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: 'Microsoft Graph drive ID containing the workbook. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.' },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .xlsx file. Returned by `list-folder-files` or `search-onedrive-files`.' },
    ...pickODataOptions(WORKSHEETS_ODATA_KEYS),
  ],
  example: "ask-marcel list-excel-worksheets --drive-id 'b!1234' --item-id '01XLSX'",
  responseShape: 'collection of Microsoft Graph `worksheet` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
