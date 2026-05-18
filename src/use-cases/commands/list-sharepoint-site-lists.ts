import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, noSkipOptions, noSkipShape } from './odata-query.ts';

// Graph rejects $skip on `/sites/{id}/lists` with `invalidRequest: $skip is
// not supported on this API. Only URLs returned by the API can be used to
// page.` (audit v1.0.0 §2.3) — drop the flag from the advertised set;
// pagination still works through nextLink → next-page.
const schema = z.object({ siteId: z.string().min(1) }).extend(noSkipShape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData(`/sites/${parsed.data.siteId}/lists`, parsed.data);
  return graph.get(path);
};

const meta: CommandMeta = {
  summary:
    'List all SharePoint lists (custom + built-in document libraries) on a site. Note: the skip flag is intentionally omitted — Graph rejects $skip on this endpoint with invalidRequest. Paginate via the top-level `nextLink` → `next-page`. Heads-up: when `top` is small, the FIRST page may legitimately be empty (`value: []`) while still carrying a `nextLink` — Graph filters server-side after slicing. Always check `nextLink` before concluding "no lists".',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites/{site-id}/lists',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/list-list',
  options: [{ name: 'site-id', key: 'siteId', required: true, description: 'SharePoint site ID. Returned by `ask-marcel search-sharepoint-sites-by-name`.' }, ...noSkipOptions],
  example: "ask-marcel list-sharepoint-site-lists --site-id 'contoso.sharepoint.com,1234,5678'",
  responseShape: 'collection of Microsoft Graph `list` resources under `value[]`',
  pagination: true,
  paginationStrategy: 'nextLinkNoSkip',
};

export { execute, meta, schema };
