import type { Result } from '../../domain/result.ts';
import { err, map, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { lineDiff, MAX_EDITS } from '../../infra/line-diff-adapter.ts';
import type { FetchOptions } from './fetch-raw-bytes.ts';
import { officeToMarkdown } from './office-to-markdown.ts';
import { OMITTED_TABLE_MARKER } from './xlsx-to-markdown.ts';

/**
 * The shared half of the two diff commands: each side is converted to markdown
 * by the same pipeline `download-drive-item-as-markdown` runs, and the answer is
 * the unified diff of the two renders, so a weekly deck or a changed document
 * costs the lines that moved instead of two full reads.
 */

type RenderOptions = FetchOptions & { readonly includeMetadata?: boolean; readonly maxCells?: number; readonly sheet?: string };

type DriveFile = { readonly name: string; readonly folder: boolean };

/** One side as rendered: the markdown, and the note its conversion attached (an empty Loop page). */
type Render = { readonly text: string; readonly note?: string };

type DiffEnvelope = {
  readonly contentType: 'text/x-diff';
  readonly size: number;
  readonly text: string;
  readonly added?: number;
  readonly removed?: number;
  readonly note?: string;
};

const readDriveFile = async (graph: GraphClient, driveId: string, itemId: string): Promise<Result<DriveFile, GraphError>> => {
  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const item = meta.value as { readonly name?: string; readonly folder?: unknown };
  return ok({ name: item.name ?? '', folder: item.folder !== undefined && item.folder !== null });
};

/** One side as markdown; a folder is refused, and a file with no text form (an image) fails in the conversion. */
const renderSide = async (graph: GraphClient, file: DriveFile, contentPath: string, opts: RenderOptions): Promise<Result<Render, GraphError>> => {
  if (file.folder) return err({ type: 'validation_error', message: `${file.name} is a folder, not a file: pick a file inside it with list-folder-files.` });
  // Every markdown answer carries `text`: converted markdown or a plain-text passthrough.
  return map(await officeToMarkdown(graph, contentPath, file.name, opts), (envelope) => {
    const { text, note } = envelope as { readonly text: string; readonly note?: unknown };
    return typeof note === 'string' ? { text, note } : { text };
  });
};

// A sheet over the cap renders as its heading plus one hint line, so two states
// of it with the same size render alike whatever their cells hold. A CSV has no
// heading: it is the table.
const cappedTables = (text: string): ReadonlyArray<string> => {
  const names: string[] = [];
  let heading = 'the table';
  for (const line of text.split('\n')) {
    if (line.startsWith('## ')) heading = `sheet ${line.slice(3)}`;
    else if (line.startsWith(OMITTED_TABLE_MARKER)) names.push(heading);
  }
  return names;
};

const notComparedNote = (tables: ReadonlyArray<string>): string =>
  `Not compared: ${tables.join(', ')} exceeded the --max-cells cap, so each side shows a one-line hint instead of the cells, and a change inside would not show. Raise --max-cells, or compare one sheet of a workbook alone with --sheet.`;

const TOO_DIFFERENT = `The two files differ in more than ${MAX_EDITS.toLocaleString('en-US')} lines, so they are different documents rather than two states of one: read each with download-drive-item-as-markdown.`;

const labelled = (label: string, side: Render): string | undefined => (side.note === undefined ? undefined : `${label}: ${side.note}`);

const joinNotes = (notes: ReadonlyArray<string | undefined>): { readonly note?: string } => {
  const present = notes.filter((n): n is string => n !== undefined);
  return present.length === 0 ? {} : { note: present.join(' ') };
};

const diffEnvelope = (fromLabel: string, from: Render, toLabel: string, to: Render, sameNote: string): DiffEnvelope => {
  const capped = [...new Set([...cappedTables(from.text), ...cappedTables(to.text)])];
  const cappedNote = capped.length === 0 ? undefined : notComparedNote(capped);
  const sideNotes = [labelled(fromLabel, from), labelled(toLabel, to)];
  const diff = lineDiff(fromLabel, from.text, toLabel, to.text);
  if (diff === undefined) return { contentType: 'text/x-diff', size: 0, text: '', ...joinNotes([TOO_DIFFERENT, cappedNote, ...sideNotes]) };
  // An empty patch means "unchanged" only when every table was compared.
  const same = diff.patch === '' && cappedNote === undefined ? sameNote : undefined;
  return {
    contentType: 'text/x-diff',
    size: new TextEncoder().encode(diff.patch).byteLength,
    text: diff.patch,
    added: diff.added,
    removed: diff.removed,
    ...joinNotes([same, cappedNote, ...sideNotes]),
  };
};

export { diffEnvelope, readDriveFile, renderSide };
export type { Render, RenderOptions };
