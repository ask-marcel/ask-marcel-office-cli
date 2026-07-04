import { basename } from 'node:path';
import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { FileSystem } from '../ports/filesystem.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { extractImagesFromBytes } from './image-extraction.ts';

/**
 * Extract embedded images from a file ON DISK — the local sibling of
 * `extract-drive-item-images`, sharing its `extractImagesFromBytes` dispatch
 * (OOXML media parts / unpdf page walk). It completes two flows the drive
 * command cannot reach: a Graph-rendered PDF saved with the global
 * output-path flag (the legacy-.ppt route), and Office files unpacked from a
 * local archive. Like `convert-local-file`, it never touches Graph and is
 * executed via `executeLocal(fs, params)`.
 */

const schema = z.object({ path: z.string().min(1) });

// Local context: the file is already on the caller's disk, so the 415 tail
// points at local routes, not the Graph raw-bytes command.
const FETCH_HINT = 'The file is already on disk — read it directly with a vision-capable model, or convert its body with `convert-local-file`.';

const executeLocal = async (fs: FileSystem, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { path } = parsed.data;

  const bytes = await fs.readBytes(path);
  if (!bytes.ok) {
    if (bytes.error.type === 'not_found') return err({ type: 'api_error', status: 404, message: `local file not found: ${path}` });
    return err({ type: 'api_error', status: 500, message: `failed to read ${path}: ${bytes.error.message}` });
  }

  return extractImagesFromBytes(bytes.value, basename(path), FETCH_HINT);
};

const execute = async (_graph: GraphClient, _params: Record<string, string>): Promise<Result<unknown, GraphError>> =>
  err({
    type: 'api_error',
    status: 400,
    message: 'extract-local-file-images reads the local filesystem, not Graph — call executeLocal(fs, params) with a FileSystem (the CLI wires this automatically).',
  });

const meta: CommandMeta = {
  summary:
    'Extract the embedded images from a file ON DISK — the local sibling of `extract-drive-item-images`, and like `convert-local-file` it never calls Microsoft Graph (works offline, no login). Same per-extension dispatch: docx / xlsx / pptx (and their macro-enabled / template variants) have their OOXML media parts read directly (png/jpg/gif/bmp/tiff/webp/svg — full-resolution originals, including images on hidden slides); a pdf is walked page by page via unpdf with each painted image re-encoded as PNG. Two flows only this command completes: a Graph-rendered PDF saved locally (legacy `.ppt` → `download-drive-item-as-pdf` with the global output-path flag → this command pulls the slide images for OCR), and Office files unpacked from a local archive. Pair with the global output-dir flag to write every image to a folder; otherwise the bytes ride back base64-encoded. Any other extension returns a 415 naming the local ways out.',
  category: 'meta',
  graphMethod: 'GET',
  graphPathTemplate: '(local) reads {path} from the local filesystem; not a Graph endpoint',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
  options: [
    {
      name: 'path',
      key: 'path',
      required: true,
      description:
        'Filesystem path of the pdf / docx / xlsx / pptx file to extract from (absolute, or relative to the current working directory). E.g. `/tmp/deck.pdf`, `./report.docx`.',
    },
  ],
  example: 'ask-marcel-office extract-local-file-images --path /tmp/deck.pdf --output-dir ./deck-images',
  responseShape:
    '`{ count, media: [{ path, contentType, sizeBytes, base64 }] }`. `path` is the source part path — `ppt/media/image3.png` for OOXML, `pdf/page2/<key>.png` for PDF (every PDF image is re-encoded as PNG). Pair with the global `--output-dir <dir>` to write each image to that folder — the response then replaces each `base64` with `savedTo: <dir>/<filename>` (the part path is flattened, e.g. `pdf_page2_Im0.png`). `count: 0` with an empty `media` array means the document embeds no extractable images (after the emf/wmf/audio/video filter). A missing file returns api_error 404 with the path.',
  producesMedia: true,
};

export { execute, executeLocal, meta, schema };
