import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ siteId: z.string().min(1) });
const { execute } = buildCommand((p) => `/sites/${p.siteId}/onenote/notebooks`, schema);

const meta: CommandMeta = {
  summary: 'List OneNote notebooks attached to a SharePoint site (separate from the personal `list-onenote-notebooks` which targets `/me`).',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/onenote/notebooks',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/onenote-list-notebooks',
  options: [
    {
      name: 'site-id',
      key: 'siteId',
      required: true,
      description: 'SharePoint site ID.',
    },
  ],
  example: "ask-marcel list-sharepoint-site-onenote-notebooks --site-id 'contoso.sharepoint.com,...'",
  responseShape: 'collection of Microsoft Graph `notebook` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
