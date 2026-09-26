import type { Result } from '../../domain/result.ts';
import { err, map, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { lineDiff, MAX_EDITS } from '../../infra/line-diff-adapter.ts';
import type { FetchOptions } from './fetch-raw-bytes.ts';
import { officeToMarkdown } from './office-to-markdown.ts';

/**
 * The shared half of the two diff commands: each side is converted to markdown
 * by the same pipeline `download-drive-item-as-markdown` runs, and the answer is
 * the unified diff of the two renders, so a weekly deck or a changed document
 * costs the lines that moved instead of two full reads.
 */

type RenderOptions = FetchOptions & { readonly includeMetadata?: boolean; readonly maxCells?: number };

type DriveFile = { readonly name: string; readonly folder: boolean };

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
const renderSide = async (graph: GraphClient, file: DriveFile, contentPath: string, opts: RenderOptions): Promise<Result<string, GraphError>> => {
  if (file.folder) return err({ type: 'validation_error', message: `${file.name} is a folder, not a file: pick a file inside it with list-folder-files.` });
  // Every markdown answer carries `text`: converted markdown or a plain-text passthrough.
  return map(await officeToMarkdown(graph, contentPath, file.name, opts), (envelope) => (envelope as { readonly text: string }).text);
};

const TOO_DIFFERENT = `The two files differ in more than ${MAX_EDITS.toLocaleString('en-US')} lines, so they are different documents rather than two states of one: read each with download-drive-item-as-markdown.`;

const diffEnvelope = (fromLabel: string, fromText: string, toLabel: string, toText: string, sameNote: string): DiffEnvelope => {
  const diff = lineDiff(fromLabel, fromText, toLabel, toText);
  if (diff === undefined) return { contentType: 'text/x-diff', size: 0, text: '', note: TOO_DIFFERENT };
  const same = diff.patch === '' ? { note: sameNote } : {};
  return { contentType: 'text/x-diff', size: new TextEncoder().encode(diff.patch).byteLength, text: diff.patch, added: diff.added, removed: diff.removed, ...same };
};

export { diffEnvelope, readDriveFile, renderSide };
export type { DiffEnvelope, DriveFile, RenderOptions };
