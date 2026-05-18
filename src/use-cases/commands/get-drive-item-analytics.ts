import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });
const { execute } = buildCommand((p) => `/drives/${p.driveId}/items/${p.itemId}/analytics`, schema);

const meta: CommandMeta = {
  summary:
    'Return view / activity analytics for a OneDrive / SharePoint file — `allTime` totals (views, viewers) and `lastSevenDays` rollup. Useful for ranking files by attention or detecting stale content. **Known empty case**: returns `{ allTime: null, lastSevenDays: null }` on low-traffic items, or when the calling identity (the Teams web client basic token) lacks the analytics scope on the tenant. Do not interpret nulls as "no views" — interpret as "not available for this caller". For active files where you expect data and see nulls, escalate to a token with `Reports.Read.All`.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/analytics',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/itemanalytics-get',
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
      description: 'driveItem ID.',
    },
  ],
  example: "ask-marcel get-drive-item-analytics --drive-id 'b!1234' --item-id '01ABC'",
  responseShape: 'single Microsoft Graph `itemAnalytics` resource',
};

export { execute, meta, schema };
