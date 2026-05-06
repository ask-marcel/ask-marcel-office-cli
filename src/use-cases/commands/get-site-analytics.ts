import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ siteId: z.string().min(1) });
const { execute } = buildCommand((p) => `/sites/${p.siteId}/analytics`, schema);

const meta: CommandMeta = {
  summary:
    'Return view / activity analytics for a SharePoint site — `allTime` totals (visits, viewers) and `lastSevenDays` rollup. Site-level parallel to `get-drive-item-analytics`. Useful for ranking sites by attention or detecting stale workspaces.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/analytics',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/itemanalytics-get',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID.',
    },
  ],
  example: "ask-marcel get-site-analytics --site-id 'contoso.sharepoint.com,...'",
  responseShape: 'single Microsoft Graph `itemAnalytics` resource',
};

export { execute, meta, schema };
