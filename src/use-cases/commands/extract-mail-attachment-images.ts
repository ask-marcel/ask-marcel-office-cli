import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { base64ToBytes, fetchRawBytes } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import type { PageRange } from '../../domain/page-range.ts';
import { extractImagesFromBytes, PAGES_OPTION, pagesField } from './image-extraction.ts';
import { buildShareToken } from './sharepoint-link-extractor.ts';

const schema = z.object({ messageId: z.string().min(1), attachmentId: z.string().min(1), pages: pagesField.optional() });

/**
 * The commands an unextractable attachment should send the caller to. Every
 * caller supplies its own: the group sibling reads a post, and a message
 * telling it to run `get-mail-attachment --message-id ...` names a command that
 * cannot address a post at all. That is the bug 2.5.0 fixed for the calendar
 * converter, and this pipeline had the same hardcoding.
 */
type ImageExtractionHints = {
  /** Where to get the bytes when this command cannot extract images. */
  readonly fetchHint: string;
  /** How to inspect an attachment whose link metadata came back incomplete. */
  readonly inspectHint: string;
};

const MAIL_HINTS: ImageExtractionHints = {
  fetchHint: 'For other attachments, fetch the bytes via `get-mail-attachment` and process locally.',
  inspectHint: 'Inspect the raw attachment with `get-mail-attachment --select id,name,contentType`, or open the message in Outlook.',
};

const fromFileAttachment = (
  attachment: { name?: string; contentBytes?: string },
  hints: ImageExtractionHints,
  pages: PageRange | undefined
): Promise<Result<unknown, GraphError>> => extractImagesFromBytes(base64ToBytes(attachment.contentBytes ?? ''), attachment.name ?? 'unnamed', hints.fetchHint, pages);

const fromReferenceAttachment = async (
  graph: GraphClient,
  attachment: { sourceUrl?: string },
  hints: ImageExtractionHints,
  pages: PageRange | undefined
): Promise<Result<unknown, GraphError>> => {
  const sourceUrl = attachment.sourceUrl;
  if (typeof sourceUrl !== 'string' || sourceUrl === '')
    return err({
      type: 'api_error',
      status: 400,
      message: `referenceAttachment missing sourceUrl — Graph returned incomplete link metadata (the linked file may have been deleted or the share revoked). ${hints.inspectHint}`,
    });
  const resolved = await graph.get(`/shares/${buildShareToken(sourceUrl)}/driveItem`);
  if (!resolved.ok) return resolved;
  const item = resolved.value as { id?: string; name?: string; parentReference?: { driveId?: string } };
  const driveId = item.parentReference?.driveId;
  const itemId = item.id;
  if (typeof driveId !== 'string' || typeof itemId !== 'string')
    return err({
      type: 'api_error',
      status: 500,
      message:
        'resolved driveItem missing id or driveId — the share link target may live in an external tenant this account cannot address through Graph. Open the attachment in Outlook / the browser instead.',
    });
  const bytes = await fetchRawBytes(graph, `/drives/${driveId}/items/${itemId}/content`);
  if (!bytes.ok) return bytes;
  return extractImagesFromBytes(bytes.value, item.name ?? '', hints.fetchHint, pages);
};

/**
 * Shared by every caller that can name an attachment by a Graph path: mail
 * here, and one post of a group thread. The path and the hints are the only
 * things that differ; `pages` narrows a PDF when the caller offers `--pages`.
 */
const extractAttachmentImages = async (graph: GraphClient, attachmentPath: string, hints: ImageExtractionHints, pages?: PageRange): Promise<Result<unknown, GraphError>> => {
  const fetched = await graph.get(attachmentPath);
  if (!fetched.ok) return fetched;
  const a = fetched.value as Record<string, unknown>;
  const odataType = a['@odata.type'];
  if (typeof odataType !== 'string') return err({ type: 'api_error', status: 400, message: 'attachment response missing @odata.type discriminator' });

  switch (odataType) {
    case '#microsoft.graph.fileAttachment':
      return fromFileAttachment(a, hints, pages);
    case '#microsoft.graph.referenceAttachment':
      return fromReferenceAttachment(graph, a, hints, pages);
    case '#microsoft.graph.itemAttachment':
      return err({ type: 'api_error', status: 415, message: 'itemAttachment (embedded mail / event / contact) has no document to extract images from.' });
    default:
      return err({ type: 'api_error', status: 400, message: `unsupported attachment type: ${odataType}` });
  }
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId, attachmentId } = parsed.data;
  return extractAttachmentImages(graph, `/me/messages/${messageId}/attachments/${attachmentId}`, MAIL_HINTS, parsed.data.pages);
};

const meta: CommandMeta = {
  summary:
    'Extract the embedded images from an Outlook mail attachment that is a pdf or a docx / xlsx / pptx (and their macro-enabled / template variants). OOXML reads the media parts directly (png/jpg/gif/bmp/tiff/webp/svg), including full-resolution / un-cropped originals and images on hidden slides; pdf walks every page via unpdf and re-encodes each painted image as PNG (page-oriented — not layer-hidden/unpainted/uncropped originals). fileAttachment decodes the inline bytes; referenceAttachment resolves via /shares/{token}/driveItem and fetches the content. Pair with the global output-dir flag to write every image to a folder; otherwise the bytes ride back base64-encoded. svg rides back as its XML source (which carries the diagram text labels); legacy vector (emf/wmf) and audio/video are skipped. itemAttachment and unsupported formats return a 415.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `list-mail-messages` or `list-mail-folder-messages`.' },
    { name: 'attachment-id', key: 'attachmentId', required: true, description: 'Attachment ID inside that message. Returned by `list-mail-attachments`.' },
    PAGES_OPTION,
  ],
  example: "ask-marcel-office extract-mail-attachment-images --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1' --output-dir ./att-images",
  responseShape:
    '`{ count, media: [{ path, contentType, sizeBytes, base64 }] }`. `path` is the in-package part path (e.g. `ppt/media/image3.png`). Pair with the global `--output-dir <dir>` to write each image to that folder — the response then replaces each `base64` with `savedTo` (the part path is flattened, e.g. `pdf_page2_Im0.png`). `count: 0` means the attachment embeds no extractable images (after the emf/wmf/audio/video filter).',
  producesMedia: true,
};

export { execute, extractAttachmentImages, meta, schema };
export type { ImageExtractionHints };
