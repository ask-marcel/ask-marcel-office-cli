import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { formatAddress, nonEmpty, renderMessageAsMarkdown, type MailLikeResource } from './convert-mail-to-markdown.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  groupId: z.string().min(1),
  threadId: z.string().min(1),
  postId: z.string().min(1),
  inlineImages: z.enum(['true', 'false']).optional(),
  keepQuoted: z.enum(['true', 'false']).optional(),
});

// A post's `sender` is the person who wrote it while `from` is normally the
// group's own address (probed live 2026-09-03), so the author line follows
// Outlook's "X on behalf of Y" reading instead of naming the group as the
// writer. A post has no subject: the thread's `topic` is the subject.
const renderPostHeaders = (m: MailLikeResource): string => {
  const lines: string[] = [];
  const from = formatAddress(m.from?.emailAddress);
  const author = formatAddress(m.sender?.emailAddress) ?? from;
  if (author !== undefined) lines.push(from === undefined || from === author ? `**From:** ${author}` : `**From:** ${author} on behalf of ${from}`);
  if (nonEmpty(m.receivedDateTime)) lines.push(`**Date:** ${m.receivedDateTime}`);
  return lines.join('\n');
};

// No command fetches a post's attachments on their own; expanding them on the
// post is the documented route (post-list-attachments), bytes included.
const POST_ATTACHMENT_HINT = '_Use `convert-group-post-attachment-to-markdown` or `get-group-post-attachment` with the attachment id to fetch._';

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { groupId, threadId, postId } = parsed.data;
  return renderMessageAsMarkdown(graph, `/groups/${groupId}/threads/${threadId}/posts/${postId}`, {
    inlineImages: parsed.data.inlineImages === 'true',
    keepQuoted: parsed.data.keepQuoted === 'true',
    attachmentHint: POST_ATTACHMENT_HINT,
    renderHeaders: renderPostHeaders,
  });
};

const meta: CommandMeta = {
  summary:
    "Render one post of a unified (Microsoft 365) group thread as markdown, the way `convert-mail-to-markdown` renders an Outlook message: a `**From:**` line, a `**Date:**` line, then the HTML body through turndown with quoted reply chains stripped. A post arrives from the group's own address with the writer in `sender`, so the author line reads `Robin Chen <robin.chen@contoso.com> on behalf of Support <support@contoso.com>`. There is no subject line: the thread `topic` is the subject and lives on `list-group-threads`. By default no image bytes are fetched; inline `cid:` images render as `[inline image: <name>]` placeholders unless `--inline-images true`. File attachments are listed below the body by name, size and id and their bytes are never fetched here; read one with `convert-group-post-attachment-to-markdown` or fetch it with `get-group-post-attachment`. Same staged-fetch design as the mail command: one call for the post, one for the attachment list when `hasAttachments` is true or the body references a `cid:` image (Graph reports false for a post whose only attachments are inline), and with `--inline-images true` one per small inline image.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/threads/{thread-id}/posts/{post-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/post-get',
  options: [
    {
      name: 'group-id',
      key: 'groupId',
      required: true,
      description: 'Azure AD group object ID for a unified (Microsoft 365) group the signed-in user belongs to.',
    },
    {
      name: 'thread-id',
      key: 'threadId',
      required: true,
      description: 'Conversation thread ID, the `id` of a `list-group-threads` entry.',
    },
    {
      name: 'post-id',
      key: 'postId',
      required: true,
      description: 'Post ID inside that thread. Returned by `list-group-thread-posts`.',
    },
    {
      name: 'inline-images',
      key: 'inlineImages',
      required: false,
      description:
        'Pass `--inline-images true` to fetch small inline images (≤ 2 MB, `image/*` only) and embed them as base64 `data:` URIs. Default is `false`: no per-image bytes fetch, and every inline `cid:` image renders as a `[inline image: <name>]` placeholder while still appearing in the attachments list. Same rule as `convert-mail-to-markdown`.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
    {
      name: 'keep-quoted',
      key: 'keepQuoted',
      required: false,
      description:
        'Quoted reply chains and forwarded-message blocks are stripped by default and replaced with a single visible marker naming this flag; the `note` reports the share of the body text that went with them. Pass `--keep-quoted true` to preserve the full body. The markers recognised are the ones `convert-mail-to-markdown` documents.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel-office convert-group-post-to-markdown --group-id 'a1b2c3d4-...' --thread-id 'AAQkAD...' --post-id 'AQMkAD...'",
  responseShape:
    '`{ contentType: "text/markdown", size, text, note? }`, the same envelope as `convert-mail-to-markdown`: headers, the turndown-rendered body and, when present, an attachments list. The optional `note` carries the attachments-list failure hint and/or the quoted-chain notice with the share of the body text it removed.',
  producesBytes: true,
};

export { execute, meta, schema };
