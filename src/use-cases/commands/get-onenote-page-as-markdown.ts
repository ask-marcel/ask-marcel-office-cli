import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({ onenotePageId: z.string().min(1) });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  return convertToMarkdown(graph, `/me/onenote/pages/${parsed.data.onenotePageId}/content`);
};

const meta: CommandMeta = {
  summary:
    'Get the body of a single OneNote page as markdown. Graph already returns OneNote pages as HTML, so this command runs that HTML through turndown locally. Inline image references in the page survive as Graph resource URLs (they are NOT base64-embedded — that is future work). For the raw HTML use `get-onenote-page-content`.',
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
  example: "ask-marcel get-onenote-page-as-markdown --onenote-page-id '1-abc...'",
  responseShape: '`{ contentType: "text/markdown", size: <chars>, text: "..." }` — turndown-rendered markdown of the OneNote page body.',
  producesBytes: true,
};

export { execute, meta, schema };
