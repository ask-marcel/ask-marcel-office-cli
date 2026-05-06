import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ siteId: z.string().min(1), onenotePageId: z.string().min(1) });
const { execute } = buildCommand((p) => `/sites/${p.siteId}/onenote/pages/${p.onenotePageId}/content`, schema);

const meta: CommandMeta = {
  summary: 'Return the HTML content of a single OneNote page from a SharePoint site (parallel to `get-onenote-page-content` for `/me`).',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/onenote/pages/{onenote-page-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/page-get',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID.',
    },
    {
      name: 'onenote-page-id',
      key: 'onenotePageId',
      required: true,
      description: 'OneNote page ID inside the site.',
    },
  ],
  example: "ask-marcel get-sharepoint-site-onenote-page-content --site-id 'contoso.sharepoint.com,...' --onenote-page-id 'p1'",
  responseShape: 'HTML envelope (`{ contentType: "text/html", size, text }`)',
};

export { execute, meta, schema };
