import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { extractAttachmentImages } from './extract-mail-attachment-images.ts';
import type { ImageExtractionHints } from './extract-mail-attachment-images.ts';
import { formatZodError } from './format-zod-error.ts';

// Both hints name a command that can address a POST. The mail wording would
// send the caller to `get-mail-attachment --message-id`, which cannot.
const POST_HINTS: ImageExtractionHints = {
  fetchHint: 'For other attachments, fetch the bytes via `get-group-post-attachment` and process locally.',
  inspectHint: 'Inspect the raw attachment with `get-group-post-attachment --select id,name,contentType`, or open the post in Outlook.',
};

const schema = z.object({
  groupId: z.string().min(1),
  threadId: z.string().min(1),
  postId: z.string().min(1),
  attachmentId: z.string().min(1),
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { groupId, threadId, postId, attachmentId } = parsed.data;
  return extractAttachmentImages(graph, `/groups/${groupId}/threads/${threadId}/posts/${postId}/attachments/${attachmentId}`, POST_HINTS);
};

const meta: CommandMeta = {
  summary:
    'Extract the embedded images from an attachment on one post of a unified (Microsoft 365) group thread, the `extract-mail-attachment-images` sibling for a group inbox and sharing its pipeline. Handles a pdf or a docx / xlsx / pptx (and their macro-enabled / template variants). OOXML reads the media parts directly (png/jpg/gif/bmp/tiff/webp/svg), including full-resolution / un-cropped originals and images on hidden slides; pdf walks every page via unpdf and re-encodes each painted image as PNG (page-oriented — not layer-hidden/unpainted/uncropped originals). fileAttachment decodes the inline bytes; referenceAttachment resolves via /shares/{token}/driveItem and fetches the content. Pair with the global output-dir flag to write every image to a folder; otherwise the bytes ride back base64-encoded. svg rides back as its XML source (which carries the diagram text labels); legacy vector (emf/wmf) and audio/video are skipped. An embedded mail (a forwarded message) is opened: its image files come back as they are and its pdf / Office files through the extractor, each path prefixed with its file name; an embedded event or contact, and unsupported formats, return a 415. This is how a diagram inside a document posted to a group survives into a searchable form.',
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
    "ask-marcel-office extract-group-post-attachment-images --group-id 'a1b2c3d4-...' --thread-id 'AAQkAD...' --post-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1' --output-dir ./post-images",
  responseShape:
    '`{ count, media: [{ path, contentType, sizeBytes, base64 }] }`. `path` is the in-package part path (e.g. `ppt/media/image3.png`). Pair with the global `--output-dir <dir>` to write each image to that folder — the response then replaces each `base64` with `savedTo` (the part path is flattened, e.g. `pdf_page2_Im0.png`). `count: 0` means the attachment embeds no extractable images (after the emf/wmf/audio/video filter).',
  producesMedia: true,
  effect: 'read',
  scopesRequired: ['Group.Read.All'],
};

export { execute, meta, schema };
