import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { extractSharepointUrls, resolveSharepointUrls } from './sharepoint-link-extractor.ts';
import type { ResolvedLink } from './sharepoint-link-extractor.ts';
import { formatZodError } from './format-zod-error.ts';

type PostLinkExtractionSummary = {
  readonly groupId: string;
  readonly threadId: string;
  readonly postId: string;
  readonly links: ReadonlyArray<ResolvedLink>;
  readonly truncated: boolean;
  readonly skippedCount: number;
};

const schema = z.object({
  groupId: z.string().min(1),
  threadId: z.string().min(1),
  postId: z.string().min(1),
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<PostLinkExtractionSummary, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { groupId, threadId, postId } = parsed.data;

  // A post has no subject of its own — the thread's `topic` is the subject and
  // lives on `list-group-threads` — so unlike the mail sibling this asks for the
  // body alone.
  const post = await graph.get(`/groups/${groupId}/threads/${threadId}/posts/${postId}?$select=body`);
  if (!post.ok) return post;
  const body = (post.value as { body?: { content?: string } }).body?.content ?? '';

  const { links, truncated, skippedCount } = await resolveSharepointUrls(graph, extractSharepointUrls(body));

  return ok({ groupId, threadId, postId, links, truncated, skippedCount });
};

const meta: CommandMeta = {
  summary:
    'Find every `*.sharepoint.com` URL in the body of one post of a unified (Microsoft 365) group thread and resolve each to its driveItem (driveId, itemId, name, webUrl), so the agent can feed those into `download-drive-item-as-pdf` / `-as-markdown`. The `extract-sharepoint-links-in-mail` sibling for a group inbox, sharing its resolver. Read-only — no conversion happens here. Capped at 25 unique URLs per call to bound fan-out (returns `truncated: true` and `skippedCount` when the body has more); duplicate URLs are deduplicated. Per-link errors are captured inside each entry instead of failing the whole call. A post whose body carries no SharePoint URL returns an empty `links` list, not an error. Access is membership-gated: a group the signed-in user does not belong to answers `ErrorAccessDenied`.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/threads/{thread-id}/posts/{post-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/post-get',
  options: [
    { name: 'group-id', key: 'groupId', required: true, description: 'Azure AD group object ID for a unified (Microsoft 365) group you belong to.' },
    { name: 'thread-id', key: 'threadId', required: true, description: 'Conversation thread ID, the `id` of a `list-group-threads` entry.' },
    { name: 'post-id', key: 'postId', required: true, description: 'Post ID inside that thread, the `id` of a `list-group-thread-posts` entry.' },
  ],
  example: "ask-marcel-office extract-sharepoint-links-in-group-post --group-id 'a1b2c3d4-...' --thread-id 'AAQkAD...' --post-id 'AAMkAD...'",
  responseShape:
    '`{ groupId, threadId, postId, links: [{ url, driveId, itemId, name, webUrl } | { url, error, location?, hint? }], truncated, skippedCount }` — one entry per unique SharePoint URL found in the body, ordered by first occurrence. A failed link names where it points in `location` (`{ kind: "onedrive", owner }` or `{ kind: "site", site }`, read from the URL), and a refused one (403) adds a `hint` on asking for access.',
  effect: 'read',
  scopesRequired: ['Group.Read.All', 'Files.Read.All'],
};

export { execute, meta, schema };
export type { PostLinkExtractionSummary };
