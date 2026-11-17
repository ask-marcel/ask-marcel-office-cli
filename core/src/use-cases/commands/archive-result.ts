import type { MediaItem } from './media-items.ts';
import { isMediaItem, isPlainRecord, withoutKey } from './media-items.ts';

/*
 * The archive result: what a zip conversion answers, one entry for each file of
 * the archive. `{ count, files: [{ path, text } | { path, note }] }`, and an
 * entry can carry `images: [{ path, base64 }]` (`--include-images true`).
 *
 * It has no single body and no top-level `media`, so the output flags did not
 * know it (QA 2026-11-16, P1-1). The functions here are pure: `output-path.ts`
 * does the writes.
 */

type ArchiveEntry = { readonly file: Record<string, unknown>; readonly path: string; readonly body: string };
type Archive = { readonly data: Record<string, unknown>; readonly entries: ReadonlyArray<ArchiveEntry> };

const toEntry = (value: unknown): ArchiveEntry | undefined => {
  if (!isPlainRecord(value)) return undefined;
  const { path, text, note } = value;
  if (typeof path !== 'string') return undefined;
  if (typeof text === 'string') return { file: value, path, body: text };
  if (typeof note === 'string') return { file: value, path, body: `Not converted: ${note}` };
  return undefined;
};

const isEntry = (entry: ArchiveEntry | undefined): entry is ArchiveEntry => entry !== undefined;

/** The archive that `data` holds, or `undefined` when `data` has a different shape. */
const readArchive = (data: unknown): Archive | undefined => {
  if (!isPlainRecord(data) || typeof data['count'] !== 'number' || !Array.isArray(data['files'])) return undefined;
  const entries = data['files'].map(toEntry);
  return entries.every(isEntry) ? { data, entries } : undefined;
};

const isArchiveResult = (data: unknown): boolean => readArchive(data) !== undefined;

const inlineImages = (entry: ArchiveEntry): ReadonlyArray<MediaItem> => {
  const { images } = entry.file;
  return Array.isArray(images) ? images.filter(isMediaItem) : [];
};

/** The images that are inline in the entries, each with the path of its entry before its own path. */
const archiveMedia = (archive: Archive): ReadonlyArray<MediaItem> =>
  archive.entries.flatMap((entry) => inlineImages(entry).map((image) => ({ ...image, path: `${entry.path}/${image.path}` })));

/** The base64 characters of all the images that are inline in the entries. */
const inlineImageLength = (archive: Archive): number => archiveMedia(archive).reduce((total, image) => total + image.base64.length, 0);

/** One markdown text for the archive: a heading and the text (or the note) of each entry. */
const archiveMarkdown = (archive: Archive): string => {
  const sections = archive.entries.map((entry) => `# File: ${entry.path}\n\n${entry.body}`);
  return `${sections.join('\n\n---\n\n')}\n`;
};

/** The result without the text of the entries, for when the text is in a file. */
const withoutTexts = (archive: Archive): Record<string, unknown> => ({ ...archive.data, files: archive.entries.map((entry) => withoutKey(entry.file, 'text')) });

/** The result with `savedTo` in the place of the `base64` of each inline image, in the order of `archiveMedia`. */
const withSavedImages = (archive: Archive, destinations: ReadonlyArray<string>): Record<string, unknown> => {
  const remaining = [...destinations];
  const saved = (image: unknown): unknown => (isMediaItem(image) ? { ...withoutKey(image, 'base64'), savedTo: remaining.shift() } : image);
  const files = archive.entries.map(({ file }) => (Array.isArray(file['images']) ? { ...file, images: file['images'].map(saved) } : file));
  return { ...archive.data, files };
};

export { archiveMarkdown, archiveMedia, inlineImageLength, isArchiveResult, readArchive, withoutTexts, withSavedImages };
export type { Archive };
