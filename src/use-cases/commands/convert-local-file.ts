import { basename } from 'node:path';
import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { FileSystem } from '../ports/filesystem.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { bytesToMarkdown } from './markdown-dispatch.ts';
import type { ConversionHints } from './markdown-dispatch.ts';
import { extensionOf } from './text-passthrough.ts';
import { convertZipArchive } from './zip-archive-to-markdown.ts';

/**
 * Convert a file ON DISK to markdown through the same `bytesToMarkdown`
 * dispatch every Graph-backed markdown command uses — the only command whose
 * input never touches Microsoft Graph (works offline, no login). The fetch
 * step was always the Graph-bound part; conversion is pure bytes-in/markdown-out.
 *
 * A `.zip` routes through the shared archive core (each entry converted, GBK
 * entry names decoded); everything else goes through the single-file dispatch.
 * What it can NOT do locally: `format=pdf` conversions (Graph renders those
 * server-side) and Loop/Fluid/Whiteboard (`format=html`, same reason).
 *
 * This is the one registry command executed via `executeLocal(fs, params)` —
 * the CLI wires its FileSystem in automatically (see `cli.ts`). The Graph-shaped
 * `execute` exists because the public `commands` registry type requires it; it
 * redirects library consumers to `executeLocal`.
 */

const schema = z.object({
  path: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
  inlineImages: z.enum(['true', 'false']).optional(),
  maxCells: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
});

// Local-context notes: the file is already on the caller's disk, so every
// unconvertible case points at reading it directly with a vision-capable model.
const LOCAL_HINTS: ConversionHints = {
  pdfNoText: 'pdf has no extractable text layer — it looks scanned / image-only. Read the local file directly with a vision-capable model, or run OCR.',
  legacyPpt:
    'ppt (legacy PowerPoint 97-2003, OLE binary) has no local markdown path — upload it to OneDrive and use `download-drive-item-as-pdf` (Graph renders legacy .ppt), or read the file directly with a vision-capable model.',
  image: (ext) => `${ext} is an image — read the file directly with a vision-capable model.`,
  generic: (ext) => `${ext} is not a convertible Office/text format. For formats Graph can render (rtf, …), upload the file to OneDrive and use \`download-drive-item-as-pdf\`.`,
};

const executeLocal = async (fs: FileSystem, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { path } = parsed.data;
  const includeMetadata = parsed.data.includeMetadata === 'true';
  const inlineImages = parsed.data.inlineImages === 'true';
  const maxCells = parsed.data.maxCells === undefined ? undefined : Number(parsed.data.maxCells);

  const bytes = await fs.readBytes(path);
  if (!bytes.ok) {
    if (bytes.error.type === 'not_found') return err({ type: 'api_error', status: 404, message: `local file not found: ${path}` });
    return err({ type: 'api_error', status: 500, message: `failed to read ${path}: ${bytes.error.message}` });
  }

  const name = basename(path);
  if (extensionOf(name) === 'zip') return convertZipArchive(bytes.value, includeMetadata, LOCAL_HINTS);
  return bytesToMarkdown(bytes.value, name, { includeMetadata, inlineImages, maxCells }, LOCAL_HINTS);
};

const execute = async (_graph: GraphClient, _params: Record<string, string>): Promise<Result<unknown, GraphError>> =>
  err({
    type: 'api_error',
    status: 400,
    message: 'convert-local-file reads the local filesystem, not Graph — call executeLocal(fs, params) with a FileSystem (the CLI wires this automatically).',
  });

const meta: CommandMeta = {
  summary:
    'Convert a file ON DISK to markdown — the only command that never calls Microsoft Graph (works offline, no login). Runs the same local pipelines as `download-drive-item-as-markdown`: docx (mammoth → turndown), xlsx (sheetjs tables, `--max-cells` OOM cap), pptx (per-slide text), odt/ods/odp, csv, pdf (text layer via unpdf), legacy OLE .xls / .doc, Outlook .msg (headers + body, attachments converted recursively), plain-text passthrough — and a `.zip` is unpacked with every contained file converted in one call (legacy GBK / CP437 entry names decoded, not mojibaked). What it canNOT do locally: convert TO pdf, and Loop/Fluid/Whiteboard sources — both need a Graph server round-trip (upload to OneDrive and use the drive-item siblings). Pass `--include-metadata true` for the Office side-channel metadata blocks; `--inline-images true` to embed docx images as base64 data URIs.',
  category: 'meta',
  graphMethod: 'GET',
  graphPathTemplate: '(local) reads {path} from the local filesystem; not a Graph endpoint',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
  options: [
    { name: 'path', key: 'path', required: true, description: 'Filesystem path of the file to convert (absolute, or relative to the current working directory). E.g. `./report.docx`, `/tmp/handover.zip`.' },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to append the converted Office file’s side-channel metadata block (`## DOCX metadata` / `## Workbook metadata` / `## PPTX metadata` / `## OpenDocument metadata`, etc.) after its body. Applies inside a `.zip` too.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
    {
      name: 'inline-images',
      key: 'inlineImages',
      required: false,
      description: "Pass `--inline-images true` to embed a docx's images as base64 `data:` URIs. Default `false` — each image becomes an `[image: <alt>]` placeholder. No-op on non-docx sources.",
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
    {
      name: 'max-cells',
      key: 'maxCells',
      required: false,
      description:
        'Per-sheet cell cap (positive integer; default 50 000) for xlsx/csv sources. A sheet whose used range exceeds the cap renders as a truncation hint instead of a multi-hundred-MB table. No-op on other sources.',
    },
  ],
  example: 'ask-marcel convert-local-file --path ./report.docx',
  responseShape:
    '`{ contentType: "text/markdown" | "text/plain", size, text }` for a single file; `{ count, files: [{ path, contentType, size, text } | { path, note }] }` for a `.zip` (one entry per contained file, unsupported entries noted). A missing file returns api_error 404 with the path. Pair with the global `--output-path` to land the markdown on disk.',
  producesBytes: true,
};

export { execute, executeLocal, meta, schema };
