import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { openZipEntries } from '../../infra/zip-reader.ts';
import type { ZipEntry } from '../../infra/zip-reader.ts';
import type { CommandMeta } from './command-types.ts';
import type { MarkdownEnvelope } from './docx-to-markdown.ts';
import { docxToMarkdown } from './docx-to-markdown.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { odfToMarkdown } from './odf-to-markdown.ts';
import { pdfToMarkdown } from './pdf-to-markdown.ts';
import { DOCX_FAMILY, ODF_FAMILY, PPTX_FAMILY, XLSX_FAMILY } from './office-extensions.ts';
import { pptxToMarkdown } from './pptx-to-markdown.ts';
import { decodeUtf8Text } from './text-passthrough.ts';
import { xlsxToMarkdown } from './xlsx-to-markdown.ts';

/**
 * Unzips a `.zip` from a OneDrive / SharePoint item and runs each contained
 * file through the same local conversion pipelines the `*-as-markdown` commands
 * use — so an agent reading "the project handover archive" doesn't have to shell
 * out to `unzip` and convert each file separately. Office files (docx/xlsx/pptx/
 * odt/ods/odp + variants) become markdown; plain-text entries are decoded inline;
 * PDFs have their text layer extracted; everything else (images, binaries, nested
 * archives — plus scanned/image-only PDFs with no text) is listed with a note
 * rather than failing the whole archive.
 */

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
});

// Bound the fan-out: the whole archive is buffered in memory and converted
// entry-by-entry, so a pathological archive can't run unbounded.
const MAX_ENTRIES = 100;

type FileResult = { readonly path: string; readonly contentType?: string; readonly size?: number; readonly text?: string; readonly note?: string };

const extensionOf = (name: string): string => {
  const dot = name.lastIndexOf('.');
  return dot === -1 ? '' : name.slice(dot + 1).toLowerCase();
};

// Convert one Office entry to markdown, or return undefined if the extension
// isn't an Office family (caller then tries the plain-text / skip paths).
const officeConvert = (ext: string, bytes: Uint8Array, includeMetadata: boolean): Promise<Result<MarkdownEnvelope, GraphError>> | undefined => {
  if (DOCX_FAMILY.has(ext)) return docxToMarkdown(bytes, { includeMetadata });
  if (XLSX_FAMILY.has(ext)) return xlsxToMarkdown(bytes, { includeMetadata });
  if (PPTX_FAMILY.has(ext)) return pptxToMarkdown(bytes, { includeMetadata });
  if (ODF_FAMILY.has(ext)) return odfToMarkdown(bytes, { includeMetadata });
  return undefined;
};

const convertEntry = async (entry: ZipEntry, includeMetadata: boolean): Promise<FileResult> => {
  const ext = extensionOf(entry.path);
  const office = officeConvert(ext, entry.bytes, includeMetadata);
  if (office !== undefined) {
    const r = await office;
    return r.ok
      ? { path: entry.path, contentType: r.value.contentType, size: r.value.size, text: r.value.text }
      : { path: entry.path, note: `conversion failed: ${r.error.message}` };
  }
  // A pdf entry's text layer is extracted inline; a scanned / image-only pdf (no text
  // layer) or an unparseable one is listed with a note instead of failing the archive.
  if (ext === 'pdf') {
    const pdf = await pdfToMarkdown(
      entry.bytes,
      `${entry.path}: pdf has no extractable text layer (scanned / image-only) — fetch it with \`download-drive-item-as-pdf\` and read it with a vision model`
    );
    return pdf.ok ? { path: entry.path, contentType: pdf.value.contentType, size: pdf.value.size, text: pdf.value.text } : { path: entry.path, note: pdf.error.message };
  }
  // Content-sniff: a zip entry whose bytes are valid UTF-8 is unpacked as text
  // (any text file, no extension list); binary entries are skipped with a note.
  const text = decodeUtf8Text(entry.bytes);
  if (text !== undefined) {
    return { path: entry.path, contentType: 'text/plain', size: entry.bytes.byteLength, text };
  }
  return {
    path: entry.path,
    note: `skipped — ${ext === '' ? 'no extension' : ext} is not a convertible Office/text format (images, binaries, and nested archives are not unpacked here)`,
  };
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const includeMetadata = parsed.data.includeMetadata === 'true';

  const bytes = await fetchRawBytes(graph, `/drives/${driveId}/items/${itemId}/content`);
  if (!bytes.ok) return bytes;
  const entries = await openZipEntries(bytes.value);
  if (!entries.ok) return entries;

  const capped = entries.value.slice(0, MAX_ENTRIES);
  const files = await Promise.all(capped.map((entry) => convertEntry(entry, includeMetadata)));
  if (entries.value.length > MAX_ENTRIES) {
    return ok({ count: files.length, totalEntries: entries.value.length, truncated: true, files });
  }
  return ok({ count: files.length, files });
};

const meta: CommandMeta = {
  summary:
    'Unzip a `.zip` from a OneDrive / SharePoint item and convert every contained file in one call — so "read the handover archive" doesn\'t need a separate unzip + per-file conversion. Office files (docx/xlsx/pptx/odt/ods/odp and their macro-enabled / template variants) are converted to markdown via the local pipelines; plain-text entries (txt/md/csv/json/yaml/…) are decoded inline; PDFs have their text layer extracted (text/plain); images, binaries, nested archives, and scanned/image-only PDFs (no text layer) are listed with a note (not unpacked) so one unsupported entry never fails the whole archive. Pass `--include-metadata true` to append each Office file\'s side-channel metadata block. Capped at 100 entries (the archive is buffered in memory); beyond that the response is flagged `truncated`.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description:
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .zip file. Returned by `list-folder-files` or `search-onedrive-files`.' },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to append each converted Office file’s side-channel metadata block (`## DOCX metadata` / `## Workbook metadata` / `## PPTX metadata` / `## OpenDocument metadata`, etc.) after its body.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel convert-drive-item-zip --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ count, files: [{ path, contentType, size, text }] }` — one entry per file in the archive (sorted by path). Convertible files carry `{ contentType, size, text }` (the markdown); unsupported / failed entries carry `{ path, note }` instead. When the archive has more than 100 entries the response adds `truncated: true` + `totalEntries` and only the first 100 are converted.',
};

export { execute, meta, schema };
