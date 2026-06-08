import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { openZipEntries } from '../../infra/zip-reader.ts';
import type { ZipEntry } from '../../infra/zip-reader.ts';
import { bytesToMarkdown } from './markdown-dispatch.ts';
import type { ConversionHints } from './markdown-dispatch.ts';

/**
 * Shared "unzip + convert every contained file" core behind both
 * `convert-drive-item-zip` (a OneDrive / SharePoint .zip) and
 * `convert-mail-attachment-zip` (an Outlook .zip attachment). Each entry is run
 * through the same `bytesToMarkdown` dispatch the markdown commands use; an entry
 * the dispatch can't convert (image, binary, nested archive, scanned PDF) is
 * LISTED with a note instead of failing the whole archive. The caller supplies its
 * own `hints` so the notes point at the right sibling command (drive vs mail).
 */

// Bound the fan-out: the whole archive is buffered in memory and converted
// entry-by-entry, so a pathological archive can't run unbounded.
const MAX_ENTRIES = 100;

type FileResult = { readonly path: string; readonly contentType?: string; readonly size?: number; readonly text?: string; readonly note?: string };
type ZipArchiveResult = { readonly count: number; readonly truncated?: true; readonly totalEntries?: number; readonly files: ReadonlyArray<FileResult> };

const convertEntry = async (entry: ZipEntry, includeMetadata: boolean, hints: ConversionHints): Promise<FileResult> => {
  const r = await bytesToMarkdown(entry.bytes, entry.path, { includeMetadata }, hints);
  if (!r.ok) return { path: entry.path, note: r.error.message };
  const env = r.value as { contentType: string; size: number; text: string };
  return { path: entry.path, contentType: env.contentType, size: env.size, text: env.text };
};

const convertZipArchive = async (bytes: Uint8Array, includeMetadata: boolean, hints: ConversionHints): Promise<Result<ZipArchiveResult, GraphError>> => {
  const entries = await openZipEntries(bytes);
  if (!entries.ok) return entries;
  const capped = entries.value.slice(0, MAX_ENTRIES);
  const files = await Promise.all(capped.map((entry) => convertEntry(entry, includeMetadata, hints)));
  if (entries.value.length > MAX_ENTRIES) {
    return ok({ count: files.length, totalEntries: entries.value.length, truncated: true, files });
  }
  return ok({ count: files.length, files });
};

export { convertZipArchive, MAX_ENTRIES };
export type { FileResult, ZipArchiveResult };
