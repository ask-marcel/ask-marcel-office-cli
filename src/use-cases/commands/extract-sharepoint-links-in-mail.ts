import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { extractSharepointUrls, resolveSharepointUrls } from './sharepoint-link-extractor.ts';
import type { ResolvedLink } from './sharepoint-link-extractor.ts';
import { formatZodError } from './format-zod-error.ts';

type LinkExtractionSummary = {
  readonly messageId: string;
  readonly subject?: string;
  readonly links: ReadonlyArray<ResolvedLink>;
  readonly truncated: boolean;
  readonly skippedCount: number;
};

const schema = z.object({ messageId: z.string().min(1) });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<LinkExtractionSummary, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId } = parsed.data;

  // 1. Pull the mail subject + body.
  const message = await graph.get(`/me/messages/${messageId}?$select=subject,body`);
  if (!message.ok) return message;
  const m = message.value as { subject?: string; body?: { content?: string } };
  const body = m.body?.content ?? '';

  // 2. Find every *.sharepoint.com URL, then resolve each via /shares/{token}/driveItem.
  const { links, truncated, skippedCount } = await resolveSharepointUrls(graph, extractSharepointUrls(body));

  return ok({
    messageId,
    subject: m.subject,
    links,
    truncated,
    skippedCount,
  });
};

const meta: CommandMeta = {
  summary:
    'Find every `*.sharepoint.com` URL in the body of a single Outlook email and resolve each one to its driveItem (driveId, itemId, name, webUrl) so the agent can feed those into `download-drive-item-as-pdf` / `-as-markdown` etc. Read-only — no conversion happens here. Capped at 25 unique URLs per call to bound fan-out (returns `truncated: true` and `skippedCount` when the body has more); duplicate URLs are deduplicated. Per-link errors are captured inside each entry instead of failing the whole call.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      description: 'Outlook message ID. Returned by `list-mail-messages` or `list-mail-folder-messages`.',
    },
  ],
  example: "ask-marcel extract-sharepoint-links-in-mail --message-id 'AAMkADk0...'",
  responseShape:
    '`{ messageId, subject, links: [{ url, driveId, itemId, name, webUrl } | { url, error }], truncated, skippedCount }` — one entry per unique SharePoint URL found in the body, ordered by first occurrence.',
};

export { execute, meta, schema };
export type { LinkExtractionSummary, ResolvedLink };
