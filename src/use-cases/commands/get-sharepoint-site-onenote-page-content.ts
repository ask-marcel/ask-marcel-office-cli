import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({ siteId: z.string().min(1), onenotePageId: z.string().min(1) });

// OneNote page content is HTML, not JSON. The default `graph.get` always
// JSON-parses the response — calling it on an HTML body crashes with
// `Unexpected token '<'`. Mirror the personal sibling (`get-onenote-page-content`)
// and route through `graph.getBinary`, which respects the response's
// Content-Type and produces `{contentType: "text/html", size, text}` for
// text bodies.
const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  return graph.getBinary(`/sites/${parsed.data.siteId}/onenote/pages/${parsed.data.onenotePageId}/content`);
};

const meta: CommandMeta = {
  summary:
    'Return the HTML content of a single OneNote page from a SharePoint site (parallel to `get-onenote-page-content` for `/me`). The response is wrapped in the standard text-content envelope so the HTML survives JSON transport.',
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
  responseShape:
    '`{ contentType: "text/html", size: <chars>, text: "<html>..." }` — the rendered OneNote page body wrapped in a JSON envelope. Pair with the global `--output-path <path>` to write the raw HTML to disk.',
};

export { execute, meta, schema };
