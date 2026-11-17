import { posix } from 'node:path';

/*
 * The inline image shape that the output flags know: `{ path, base64, ... }`,
 * as the image-extraction commands return it in `media` and a zip conversion
 * returns it in the `images` of an entry.
 */

type MediaItem = { readonly path: string; readonly base64: string };

const isPlainRecord = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);

const isMediaItem = (value: unknown): value is MediaItem => isPlainRecord(value) && typeof value['path'] === 'string' && typeof value['base64'] === 'string';

const withoutKey = (data: Record<string, unknown>, key: string): Record<string, unknown> => {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(data)) if (k !== key) out[k] = v;
  return out;
};

// The full media path is flattened (`pdf/page2/Im0.png` → `pdf_page2_Im0.png`),
// not reduced to its basename, because PDF page-image keys (`Im0`, …) repeat
// across pages. A name that still repeats (two `image.png` in one forwarded
// mail) is numbered, `image-2.png`, so no image overwrites another.
// The path comes from inside a document or an archive, so it is not trusted:
// `\` is flattened as `/` is, and the name stays one segment below the
// directory (`..\..\x` cannot leave it on Windows).
const withDestinations = (outputDir: string, media: ReadonlyArray<MediaItem>): ReadonlyArray<{ readonly item: MediaItem; readonly dest: string }> => {
  const used = new Set<string>();
  return media.map((item) => {
    const flat = item.path.replace(/[/\\]/g, '_');
    const dot = flat.lastIndexOf('.');
    const stem = dot > 0 ? flat.slice(0, dot) : flat;
    const ext = dot > 0 ? flat.slice(dot) : '';
    let name = flat;
    let copy = 1;
    while (used.has(name)) {
      copy += 1;
      name = `${stem}-${copy}${ext}`;
    }
    used.add(name);
    return { item, dest: posix.join(outputDir, name) };
  });
};

export { isMediaItem, isPlainRecord, withDestinations, withoutKey };
export type { MediaItem };
