import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { WriteGraph } from '../../infra/write-graph.ts';
import type { WriteCommandMeta } from './command-types.ts';
import { convertAttachmentToPdf } from './convert-mail-attachment-to-pdf.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  groupId: z.string().min(1),
  threadId: z.string().min(1),
  postId: z.string().min(1),
  attachmentId: z.string().min(1),
});

const execute = async (graph: WriteGraph, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { groupId, threadId, postId, attachmentId } = parsed.data;
  return convertAttachmentToPdf(graph, `/groups/${groupId}/threads/${threadId}/posts/${postId}/attachments/${attachmentId}`);
};

const meta: WriteCommandMeta = {
  summary:
    'Convert an attachment on one post of a unified (Microsoft 365) group thread to PDF on the fly, the `convert-mail-attachment-to-pdf` sibling for a group inbox and sharing its pipeline. fileAttachment uploads the bytes to a temp folder under the SIGNED-IN user’s /me/drive (the post is read from the group, the render happens on your own drive), runs Graph `?format=pdf`, then deletes the temp item; referenceAttachment resolves via /shares/{token}/driveItem and converts in place; plain-text and `pdf` sources short-circuit to a raw-bytes envelope (Graph’s `?format=pdf` does not accept `pdf` as an input). image attachments are rejected (Graph rejects image inputs); itemAttachment (embedded mail / event / contact) is unsupported — use `convert-group-post-attachment-to-markdown`. This is the command to reach for when a deck is posted to a group and slide layout matters to a vision-capable model.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/threads/{thread-id}/posts/{post-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
  options: [
    { name: 'group-id', key: 'groupId', required: true, description: 'Azure AD group object ID for a unified (Microsoft 365) group you belong to.' },
    { name: 'thread-id', key: 'threadId', required: true, description: 'Conversation thread ID, the `id` of a `list-group-threads` entry.' },
    { name: 'post-id', key: 'postId', required: true, description: 'Post ID inside that thread, the `id` of a `list-group-thread-posts` entry.' },
    { name: 'attachment-id', key: 'attachmentId', required: true, description: 'Attachment ID inside that post. Returned by `list-group-post-attachments`.' },
  ],
  example:
    "ask-marcel-office convert-group-post-attachment-to-pdf --group-id 'a1b2c3d4-...' --thread-id 'AAQkAD...' --post-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1' --output-path ./deck.pdf",
  responseShape:
    '`{ contentType: "application/pdf", size, base64 }` — the PDF bytes, inlined. Plain-text and pdf sources short-circuit to `{ contentType, size, base64, note }`; image attachments return api_error 415; itemAttachment returns api_error 400. Pair with the global `--output-path` to land the bytes on disk and replace `base64` with `savedTo`.',
  producesBytes: true,
  effect: 'transient-upload',
  scopesRequired: ['Group.Read.All', 'Files.ReadWrite'],
};

export { execute, meta, schema };
