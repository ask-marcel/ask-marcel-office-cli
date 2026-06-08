import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { base64ToBytes } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import type { ConversionHints } from './markdown-dispatch.ts';
import { convertZipArchive } from './zip-archive-to-markdown.ts';

/**
 * Unzips a `.zip` Outlook mail attachment and converts every contained file in one
 * call — the mail-side mirror of `convert-drive-item-zip`. Without it, reading a
 * zipped vendor deck meant: `get-mail-attachment --output-path x.zip` → manual
 * `unzip` (with `-O GBK` for Chinese names) → convert each file. This collapses all
 * of that into one command: it pulls the fileAttachment bytes, unzips them (legacy
 * GBK / CP437 entry names are decoded, not mojibaked, in the shared zip reader), and
 * runs each entry through the same conversion dispatch the markdown commands use.
 */

const schema = z.object({
  messageId: z.string().min(1),
  attachmentId: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
});

// Notes for unconvertible entries. The inner files of the archive aren't directly
// addressable by a sibling command (they live inside the zip), so the guidance is
// generic — extract + read with a vision model — rather than pointing at one.
const MAIL_ZIP_HINTS: ConversionHints = {
  pdfNoText: 'pdf has no extractable text layer (scanned / image-only) — read it with a vision model',
  legacyPpt: 'ppt (legacy PowerPoint, OLE binary) has no markdown path — convert it to PDF, then read it with a vision model',
  image: (ext) => `${ext} is an image — not unpacked here; read it with a vision model`,
  generic: (ext) => `${ext} is not a convertible Office/text format (images, binaries, and nested archives are not unpacked here)`,
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId, attachmentId } = parsed.data;
  const includeMetadata = parsed.data.includeMetadata === 'true';

  const fetched = await graph.get(`/me/messages/${messageId}/attachments/${attachmentId}`);
  if (!fetched.ok) return fetched;
  const attachment = fetched.value as Record<string, unknown>;

  const odataType = attachment['@odata.type'];
  if (odataType !== '#microsoft.graph.fileAttachment') {
    return err({
      type: 'api_error',
      status: 400,
      message: `convert-mail-attachment-zip needs a fileAttachment whose bytes are a .zip (got ${typeof odataType === 'string' ? odataType : 'an attachment with no @odata.type'}). itemAttachment / referenceAttachment have no inline zip payload to unpack.`,
    });
  }
  const contentBytes = attachment['contentBytes'];
  if (typeof contentBytes !== 'string') {
    return err({ type: 'api_error', status: 400, message: 'fileAttachment has no contentBytes to unzip (pass `--select` was not used? the attachment may be empty).' });
  }
  return convertZipArchive(base64ToBytes(contentBytes), includeMetadata, MAIL_ZIP_HINTS);
};

const meta: CommandMeta = {
  summary:
    "Unzip a `.zip` Outlook mail attachment and convert every contained file in one call — the mail-side mirror of `convert-drive-item-zip`, so reading a zipped vendor deck doesn't need `get-mail-attachment` + manual `unzip` + per-file conversion. Pulls the fileAttachment bytes, unzips them (legacy GBK / CP437 entry names — Chinese vendor archives written by WinRAR / Windows Explorer — are decoded correctly, not mojibaked), and runs each file through the local pipelines: Office files (docx/xlsx/pptx/odt/ods/odp and macro-enabled / template variants) → markdown; plain-text entries decoded inline; legacy OLE .xls (sheetjs) and .doc (word-extractor, text only) extracted; an inner Outlook .msg rendered; PDFs have their text layer extracted; images, binaries, nested archives, legacy .ppt, and scanned/image-only PDFs are listed with a note (not unpacked) so one unsupported entry never fails the whole archive. Pass `--include-metadata true` to append each Office file's side-channel metadata block. Capped at 100 entries; beyond that the response is flagged `truncated`. itemAttachment / referenceAttachment are rejected (no inline zip payload).",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-get',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `list-mail-messages` or `list-mail-folder-messages`.' },
    { name: 'attachment-id', key: 'attachmentId', required: true, description: 'Attachment ID of the .zip fileAttachment. Returned by `list-mail-attachments`.' },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to append each converted Office file’s side-channel metadata block (`## DOCX metadata` / `## Workbook metadata` / `## PPTX metadata` / `## OpenDocument metadata`, etc.) after its body.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel convert-mail-attachment-zip --message-id 'AAMkAD...' --attachment-id 'AAMkAD...attach1'",
  responseShape:
    '`{ count, files: [{ path, contentType, size, text }] }` — one entry per file in the archive (sorted by path; non-mojibake names). Convertible files carry `{ contentType, size, text }` (the markdown); unsupported / failed entries carry `{ path, note }`. When the archive has more than 100 entries the response adds `truncated: true` + `totalEntries` and only the first 100 are converted. A non-fileAttachment (itemAttachment / referenceAttachment) or a non-zip payload returns an api_error.',
};

export { execute, meta, schema };
