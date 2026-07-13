import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { openZipEntries } from '../../infra/zip-reader.ts';
import type { ZipEntry } from '../../infra/zip-reader.ts';
import { extractImagesFromBytes } from './image-extraction.ts';
import type { MediaEnvelope } from './media-files.ts';
import { bytesToMarkdown, NESTED_HINTS } from './markdown-dispatch.ts';

/**
 * Shared "unzip + convert every contained file" core behind
 * `convert-drive-item-zip` (a OneDrive / SharePoint .zip),
 * `convert-mail-attachment-zip` (an Outlook .zip attachment), and
 * `convert-local-file` (a .zip on disk). Each entry is run through the same
 * `bytesToMarkdown` dispatch the markdown commands use; an entry the dispatch
 * can't convert (image, binary, nested archive, scanned PDF) is LISTED with a
 * note instead of failing the whole archive. Notes use the container-neutral
 * NESTED_HINTS: entries live INSIDE the zip, so caller-specific
 * sibling-command pointers (`extract-drive-item-images`, …) cannot reach them.
 */

// Bound the fan-out: the whole archive is buffered in memory and converted
// entry-by-entry, so a pathological archive can't run unbounded.
const MAX_ENTRIES = 100;

type FileResult = {
  readonly path: string;
  readonly contentType?: string;
  readonly size?: number;
  readonly text?: string;
  readonly note?: string;
  readonly images?: MediaEnvelope['media'];
};
type ZipArchiveResult = { readonly count: number; readonly truncated?: true; readonly totalEntries?: number; readonly files: ReadonlyArray<FileResult> };

// Embedded images of one entry, best-effort: an entry type that carries none
// (or isn't a supported document) simply yields []. The fetchHint is empty —
// these images travel back inline; there is no per-entry sibling command.
const entryImages = async (entry: ZipEntry): Promise<MediaEnvelope['media']> => {
  const r = await extractImagesFromBytes(entry.bytes, entry.path, '');
  return r.ok ? r.value.media : [];
};

const convertEntry = async (entry: ZipEntry, includeMetadata: boolean, includeImages: boolean): Promise<FileResult> => {
  const r = await bytesToMarkdown(entry.bytes, entry.path, { includeMetadata }, NESTED_HINTS);
  if (!r.ok) return { path: entry.path, note: r.error.message };
  const env = r.value as { contentType: string; size: number; text: string };
  const base = { path: entry.path, contentType: env.contentType, size: env.size, text: env.text };
  if (!includeImages) return base;
  const images = await entryImages(entry);
  return images.length > 0 ? { ...base, images } : base;
};

const convertZipArchive = async (bytes: Uint8Array, includeMetadata: boolean, includeImages = false): Promise<Result<ZipArchiveResult, GraphError>> => {
  const entries = await openZipEntries(bytes);
  if (!entries.ok) return entries;
  const capped = entries.value.slice(0, MAX_ENTRIES);
  const files = await Promise.all(capped.map((entry) => convertEntry(entry, includeMetadata, includeImages)));
  if (entries.value.length > MAX_ENTRIES) {
    return ok({ count: files.length, totalEntries: entries.value.length, truncated: true, files });
  }
  return ok({ count: files.length, files });
};

export { convertZipArchive, MAX_ENTRIES };
export type { FileResult, ZipArchiveResult };
