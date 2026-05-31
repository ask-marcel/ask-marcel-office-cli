import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { formatOnenoteMetadata, type OnenotePage } from './onenote-metadata.ts';
import { embedOnenoteResources } from './onenote-resource-embedder.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  onenotePageId: z.string().min(1),
  inlineImages: z.enum(['true', 'false']).optional(),
  includeMetadata: z.enum(['true', 'false']).optional(),
});

// Expand the parent section + notebook display names alongside the page's own
// title/timestamps — one GET, only when `--include-metadata true`.
const PAGE_METADATA_QUERY = '$select=title,createdDateTime,lastModifiedDateTime&$expand=parentSection($select=displayName),parentNotebook($select=displayName)';

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { onenotePageId } = parsed.data;
  // Default `true`: embed the page's `…/onenote/resources/{id}/$value` images as
  // data URIs so the markdown is self-contained. `--inline-images false` keeps
  // the raw Graph resource URLs.
  const embedImages = parsed.data.inlineImages !== 'false';
  const includeMetadata = parsed.data.includeMetadata === 'true';

  const md = await convertToMarkdown(graph, `/me/onenote/pages/${onenotePageId}/content`, embedImages ? { htmlTransform: (html) => embedOnenoteResources(graph, html) } : {});
  if (!md.ok) return md;
  if (!includeMetadata) return ok(md.value);

  const page = await graph.get(`/me/onenote/pages/${onenotePageId}?${PAGE_METADATA_QUERY}`);
  if (!page.ok) return page;
  const text = `${md.value.text}\n\n${formatOnenoteMetadata(page.value as OnenotePage)}`;
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(text).byteLength, text });
};

const meta: CommandMeta = {
  summary:
    'Get the body of a single OneNote page as markdown. Graph returns OneNote pages as HTML, which this command runs through turndown locally. By default the page’s inline images (its `…/onenote/resources/{id}/$value` references) are fetched and embedded as base64 `data:` URIs so the markdown is self-contained — pass `--inline-images false` to keep the raw Graph resource URLs instead. Image embedding is per-image isolated: any resource that fails to fetch, is oversize (> 2 MB), or is not an image is left as a URL rather than failing the page. Pass `--include-metadata true` to append a `## OneNote metadata` block (title, created / last-modified timestamps, parent section + notebook). For the raw HTML use `get-onenote-page-content`.',
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
    {
      name: 'inline-images',
      key: 'inlineImages',
      required: false,
      description:
        'Pass `--inline-images false` to skip fetching + embedding the page’s `onenote/resources/{id}/$value` images and keep the raw Graph resource URLs in the markdown. Default is `true` (embed as base64 `data:` URIs so the output is self-contained). Embedding is per-image isolated — a failed / oversize / non-image resource is left as a URL either way.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to append a `## OneNote metadata` block after the body: page title, created / last-modified timestamps, and the parent section + notebook display names (one extra GET, expanded). Default omits it.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel get-onenote-page-as-markdown --onenote-page-id '1-abc...'",
  responseShape:
    '`{ contentType: "text/markdown", size, text }` — turndown-rendered page body with inline images embedded as data URIs by default. With `--include-metadata true`, a `## OneNote metadata` block is appended after the body.',
  producesBytes: true,
};

export { execute, meta, schema };
