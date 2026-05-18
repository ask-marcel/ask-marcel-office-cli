import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({ onenotePageId: z.string().min(1) });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  return graph.getBinary(`/me/onenote/pages/${parsed.data.onenotePageId}/content`);
};

const meta: CommandMeta = {
  summary:
    'Get the raw HTML body of a single OneNote page. Returned as a `text/html` payload so the HTML body is available verbatim (text mode prints the body raw; JSON mode wraps it in the standard `{contentType, size, text}` envelope). For markdown output use `get-onenote-page-as-markdown`.',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: '/me/onenote/pages/{onenote-page-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/page-get',
  options: [
    {
      name: 'onenote-page-id',
      key: 'onenotePageId',
      required: true,
      description: 'OneNote page ID. Returned by `ask-marcel list-onenote-section-pages`.',
      aliases: [{ name: 'page-id', key: 'pageId' }],
    },
  ],
  example: "ask-marcel get-onenote-page-content --onenote-page-id '1-abc...'",
  responseShape: '`{ contentType: "text/html", size: <chars>, text: "<html>..." }` — the rendered OneNote page body wrapped in a JSON envelope',
};

export { execute, meta, schema };
