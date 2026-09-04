import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { convertAttachmentToMarkdown } from './convert-mail-attachment-to-markdown.ts';
import { formatZodError } from './format-zod-error.ts';
import { keepQuotedOption, keepQuotedSchemaField } from './mail-quote-stripper.ts';
import type { ConversionHints } from './markdown-dispatch.ts';

// A group post has no PDF-rendering sibling (`?format=pdf` would mean uploading
// the bytes to the caller's own drive), so every unconvertible case routes to
// the raw bytes instead. Borrowing the mail wording here would name commands
// that cannot address a post at all.
const POST_HINTS: ConversionHints = {
  pdfNoText:
    'pdf attachment has no extractable text layer — it looks scanned / image-only (only page images, no embedded text). Fetch the bytes with `get-group-post-attachment --output-path /tmp/file.pdf`, then read the PDF with a vision-capable model, or run OCR.',
  legacyPpt:
    'ppt (legacy PowerPoint 97-2003, OLE binary) cannot be converted to markdown — there is no pure-JS parser for the format. Fetch the bytes with `get-group-post-attachment --output-path /tmp/deck.ppt` and convert them outside the CLI, or open the post in Outlook.',
  image: (ext) =>
    `${ext} attachment is an image and cannot be converted to markdown. Use \`get-group-post-attachment\` to fetch the bytes (returned base64-encoded) and feed them into a vision-capable model directly — that's the right shape for image content.`,
  generic: (ext) =>
    `${ext} attachment not supported by \`convert-group-post-attachment-to-markdown\`. Fetch the raw bytes with \`get-group-post-attachment\` (add \`--output-path\` to land them on disk) and handle the format outside the CLI.`,
};

const schema = z.object({
  groupId: z.string().min(1),
  threadId: z.string().min(1),
  postId: z.string().min(1),
  attachmentId: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
  keepQuoted: keepQuotedSchemaField,
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { groupId, threadId, postId, attachmentId } = parsed.data;
  return convertAttachmentToMarkdown(
    graph,
    `/groups/${groupId}/threads/${threadId}/posts/${postId}/attachments/${attachmentId}`,
    { includeMetadata: parsed.data.includeMetadata === 'true', keepQuoted: parsed.data.keepQuoted === 'true' },
    POST_HINTS
  );
};

const meta: CommandMeta = {
  summary:
    'Convert an attachment on one post of a unified (Microsoft 365) group thread to markdown, the `convert-mail-attachment-to-markdown` sibling for a group inbox. Polymorphic on the attachment’s `@odata.type` and sharing the mail pipeline: fileAttachment decodes the inline bytes and converts them locally (docx, xlsx, csv, odt/ods/odp, pptx as per-slide text, pdf text layer, legacy .xls/.doc, an Outlook `.msg` rendered recursively with its quoted chain stripped unless `--keep-quoted true`, plain text passed through); referenceAttachment resolves via `/shares/{token}/driveItem`; an embedded mail, event or contact is rendered locally. There is no PDF sibling here, so an image, a scanned PDF, a legacy `.ppt` and any other unsupported format return a 415 pointing at `get-group-post-attachment` for the raw bytes.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/threads/{thread-id}/posts/{post-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
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
      name: 'attachment-id',
      key: 'attachmentId',
      required: true,
      description: 'Attachment ID inside that post. Returned by `list-group-post-attachments`.',
    },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to append the Office side-channel metadata block, exactly as `convert-mail-attachment-to-markdown` documents it. No-op on other attachment types.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
    keepQuotedOption,
  ],
  example: "ask-marcel-office convert-group-post-attachment-to-markdown --group-id 'a1b2c3d4-...' --thread-id 'AAQkAD...' --post-id 'AQMkAD...' --attachment-id 'AAMkAD...'",
  responseShape:
    '`{ contentType: "text/markdown", size, text }` on success (file and reference attachments run through the conversion dispatch; an embedded item is rendered locally). Plain-text sources return the raw-bytes envelope, and a PDF source carries `pageCount`. Unsupported types return an api_error with status 415 naming `get-group-post-attachment` as the way to the bytes.',
  producesBytes: true,
};

export { execute, meta, schema };
