import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { wrapExcelExecute } from './excel-error.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const inner = buildListCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/workbook/comments`, baseSchema);
const execute = wrapExcelExecute(inner.execute);
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    "List the modern threaded comments anchored to cells in an Excel workbook (the New Comments feature, distinct from legacy notes). Each `workbookComment` has `content`, `contentType`, `task` state, plus replies via the comment's `replies` navigation.",
  category: 'excel',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/workbook/comments',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/workbook-list-comments',
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
  example: "ask-marcel list-excel-comments --drive-id 'b!1234' --item-id '01ABC'",
  responseShape: 'collection of Microsoft Graph `workbookComment` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
