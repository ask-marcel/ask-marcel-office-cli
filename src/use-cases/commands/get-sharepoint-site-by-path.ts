import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({
  hostname: z.string().min(1),
  path: z
    .string()
    .min(1)
    .refine((v) => v.startsWith('/'), { message: 'must start with `/` (server-relative path, e.g. /sites/Marketing)' }),
});
const { execute } = buildCommand((p) => `/sites/${p.hostname}:${p.path}`, schema);

const meta: CommandMeta = {
  summary:
    'Resolve a SharePoint site by its hostname + server-relative path. Use this when you have a SharePoint URL (e.g. `https://contoso.sharepoint.com/sites/Marketing`) but no site ID.',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{hostname}:{path}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/site-get',
  options: [
    {
      name: 'hostname',
      key: 'hostname',
      required: true,
      description: 'SharePoint host, e.g. `contoso.sharepoint.com` (or `contoso-my.sharepoint.com` for personal sites). Take the host portion of your SharePoint URL.',
    },
    {
      name: 'path',
      key: 'path',
      required: true,
      description: '`--path` (server-relative) starts with `/`, e.g. `/sites/Marketing` or `/teams/Marketing/Subsite`. Take the URL pathname after the hostname.',
    },
  ],
  example: "ask-marcel-office get-sharepoint-site-by-path --hostname 'contoso.sharepoint.com' --path '/sites/Marketing'",
  responseShape: 'single Microsoft Graph `site` resource',
};

export { execute, meta, schema };
