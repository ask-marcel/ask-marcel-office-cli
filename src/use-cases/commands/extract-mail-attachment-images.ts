import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { base64ToBytes, fetchRawBytes } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { extractImagesFromBytes } from './image-extraction.ts';
import { buildShareToken } from './sharepoint-link-extractor.ts';

const schema = z.object({ messageId: z.string().min(1), attachmentId: z.string().min(1) });

const FETCH_HINT = 'For other attachments, fetch the bytes via `get-mail-attachment` and process locally.';

const fromFileAttachment = (attachment: { name?: string; contentBytes?: string }): Promise<Result<unknown, GraphError>> =>
  extractImagesFromBytes(base64ToBytes(attachment.contentBytes ?? ''), attachment.name ?? 'unnamed', FETCH_HINT);

const fromReferenceAttachment = async (graph: GraphClient, attachment: { sourceUrl?: string }): Promise<Result<unknown, GraphError>> => {
  const sourceUrl = attachment.sourceUrl;
  if (typeof sourceUrl !== 'string' || sourceUrl === '') return err({ type: 'api_error', status: 400, message: 'referenceAttachment missing sourceUrl — Graph returned incomplete link metadata (the linked file may have been deleted or the share revoked). Inspect the raw attachment with `get-mail-attachment --select id,name,contentType`, or open the message in Outlook.' });
  const resolved = await graph.get(`/shares/${buildShareToken(sourceUrl)}/driveItem`);
  if (!resolved.ok) return resolved;
  const item = resolved.value as { id?: string; name?: string; parentReference?: { driveId?: string } };
  const driveId = item.parentReference?.driveId;
  const itemId = item.id;
  if (typeof driveId !== 'string' || typeof itemId !== 'string') return err({ type: 'api_error', status: 500, message: 'resolved driveItem missing id or driveId — the share link target may live in an external tenant this account cannot address through Graph. Open the attachment in Outlook / the browser instead.' });
  const bytes = await fetchRawBytes(graph, `/drives/${driveId}/items/${itemId}/content`);
  if (!bytes.ok) return bytes;
  return extractImagesFromBytes(bytes.value, item.name ?? '', FETCH_HINT);
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId, attachmentId } = parsed.data;

  const fetched = await graph.get(`/me/messages/${messageId}/attachments/${attachmentId}`);
  if (!fetched.ok) return fetched;
  const a = fetched.value as Record<string, unknown>;
  const odataType = a['@odata.type'];
  if (typeof odataType !== 'string') return err({ type: 'api_error', status: 400, message: 'attachment response missing @odata.type discriminator' });

  switch (odataType) {
    case '#microsoft.graph.fileAttachment':
      return fromFileAttachment(a);
    case '#microsoft.graph.referenceAttachment':
      return fromReferenceAttachment(graph, a);
    case '#microsoft.graph.itemAttachment':
      return err({ type: 'api_error', status: 415, message: 'itemAttachment (embedded mail / event / contact) has no document to extract images from.' });
    default:
      return err({ type: 'api_error', status: 400, message: `unsupported attachment type: ${odataType}` });
  }
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
  ],
  example: "ask-marcel extract-mail-attachment-images --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1' --output-dir ./att-images",
  responseShape:
    '`{ count, media: [{ path, contentType, sizeBytes, base64 }] }`. `path` is the in-package part path (e.g. `ppt/media/image3.png`). Pair with the global `--output-dir <dir>` to write each image to that folder — the response then replaces each `base64` with `savedTo` (the part path is flattened, e.g. `pdf_page2_Im0.png`). `count: 0` means the attachment embeds no extractable images (after the emf/wmf/audio/video filter).',
  producesMedia: true,
};

export { execute, meta, schema };
