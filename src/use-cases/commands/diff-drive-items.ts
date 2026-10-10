import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { ReadGraph } from '../../infra/read-graph.ts';
import type { ReadCommandMeta } from './command-types.ts';
import { diffEnvelope, readDriveFile, renderSide } from './drive-item-diff.ts';
import type { Render, RenderOptions } from './drive-item-diff.ts';
import { formatZodError } from './format-zod-error.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';
import { MAX_CELLS_OPTION, maxCellsField, SHEET_OPTION } from './xlsx-to-markdown.ts';

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  otherDriveId: z.string().min(1),
  otherItemId: z.string().min(1),
  includeMetadata: z.enum(['true', 'false']).optional(),
  maxCells: maxCellsField,
  sheet: z.string().min(1).optional(),
});

const renderFile = async (graph: ReadGraph, driveId: string, itemId: string, opts: RenderOptions): Promise<Result<{ name: string; render: Render }, GraphError>> => {
  const file = await readDriveFile(graph, driveId, itemId);
  if (!file.ok) return file;
  const render = await renderSide(graph, file.value, `/drives/${driveId}/items/${itemId}/content`, opts);
  return render.ok ? ok({ name: file.value.name, render: render.value }) : render;
};

const WORDING = { same: 'The two files render to the same markdown.', readWhole: 'read each file with download-drive-item-as-markdown' };

const execute = async (graph: ReadGraph, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId, otherDriveId, otherItemId, maxCells, sheet } = parsed.data;
  const opts = { includeMetadata: parsed.data.includeMetadata === 'true', maxCells, sheet };
  const [from, to] = await Promise.all([renderFile(graph, driveId, itemId, opts), renderFile(graph, otherDriveId, otherItemId, opts)]);
  if (!from.ok) return from;
  if (!to.ok) return to;
  return ok(diffEnvelope(`a/${from.value.name}`, from.value.render, `b/${to.value.name}`, to.value.render, WORDING));
};

const meta: ReadCommandMeta = {
  summary:
    'Compare two OneDrive or SharePoint files and answer only what differs: each is converted to markdown the way `download-drive-item-as-markdown` converts it, and the answer is a unified diff of the two renders (`--- a/<first>`, `+++ b/<second>`, hunks with three lines of context) with the count of added and removed lines. Made for a document that is saved as a new file each week (a status deck, a task list): the diff costs the lines that moved instead of two full reads. `--include-metadata true` renders comments and tracked changes too, so a comment added between the two shows up. A workbook sheet over the `--max-cells` cap is never reported as unchanged: the note names it, and `--sheet` compares one sheet alone. Past 1,000 added or removed lines (a changed line counts once each way) the answer is a note instead of a diff. For two versions of the same file, `diff-drive-item-versions`.',
  category: 'drive',
  producesBytes: true,
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content and /drives/{other-drive-id}/items/{other-item-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: `The first (older) file's drive. ${DRIVE_ID_DESCRIPTION}` },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'The first (older) file’s driveItem ID, from `list-folder-files`, `search-all-files` or `resolve-drive-share-link`.',
    },
    { name: 'other-drive-id', key: 'otherDriveId', required: true, description: 'The second (newer) file’s drive; the same as `--drive-id` when both sit in one library.' },
    { name: 'other-item-id', key: 'otherItemId', required: true, description: 'The second (newer) file’s driveItem ID.' },
    {
      name: 'include-metadata',
      key: 'includeMetadata',
      required: false,
      description:
        'Pass `--include-metadata true` to render the side-channel content of Office files (comments, tracked changes, properties, speaker notes) on both sides, so changes there appear in the diff.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
    MAX_CELLS_OPTION,
    SHEET_OPTION,
  ],
  example: "ask-marcel-office diff-drive-items --drive-id 'b!1234' --item-id '01WEEK6' --other-drive-id 'b!1234' --other-item-id '01WEEK7'",
  responseShape:
    '`{ contentType: "text/x-diff", size, text, added, removed }` — `text` is the unified diff; `note` says when the two render the same (empty `text`, zero counts), when they differ too much to diff (no counts), which sheets were over the `--max-cells` cap and not compared, and any note a conversion attached, prefixed with its side.',
  effect: 'read',
  scopesRequired: ['Files.Read'],
};

export { execute, meta, schema };
