import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { ReadGraph } from '../../infra/read-graph.ts';
import type { ReadCommandMeta } from './command-types.ts';
import { pickPreviousVersion, refuseVersionRender, resolveVersion } from './download-drive-item-version.ts';
import { diffEnvelope, readDriveFile, renderSide } from './drive-item-diff.ts';
import { formatZodError } from './format-zod-error.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';
import { rendersThroughGraph } from './office-to-markdown.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';
import { MAX_CELLS_OPTION, maxCellsField, SHEET_OPTION } from './xlsx-to-markdown.ts';

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  versionId: z.string().min(1).optional(),
  before: isoDateTimeField.optional(),
  includeMetadata: z.enum(['true', 'false']).optional(),
  maxCells: maxCellsField,
  sheet: z.string().min(1).optional(),
});

const WORDING = {
  same: 'This version and the live file render to the same markdown.',
  readWhole: 'read the version with download-drive-item-version --format markdown and the live file with download-drive-item-as-markdown',
};

const liveNote = (id: string, before: string | undefined): string =>
  before === undefined
    ? `Version ${id} is the only one saved: there is no earlier version to compare with the live file.`
    : `Nothing was saved after ${before}: version ${id} is the live file.`;

const execute = async (graph: ReadGraph, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId, before, maxCells, sheet } = parsed.data;
  const file = await readDriveFile(graph, driveId, itemId);
  if (!file.ok) return file;
  const { name } = file.value;
  if (rendersThroughGraph(name)) return refuseVersionRender(name, "Use `download-drive-item-version --format original` for an old version's raw bytes");
  const version =
    parsed.data.versionId === undefined && before === undefined
      ? await pickPreviousVersion(graph, driveId, itemId)
      : await resolveVersion(graph, driveId, itemId, parsed.data.versionId, before);
  if (!version.ok) return version;
  const { id, current } = version.value;
  // The chosen version is still the live file: nothing moved, nothing to download.
  if (current) return ok({ contentType: 'text/x-diff', size: 0, text: '', added: 0, removed: 0, versionId: id, note: liveNote(id, before) });
  const opts = { includeMetadata: parsed.data.includeMetadata === 'true', maxCells, sheet };
  const old = await renderSide(graph, file.value, `/drives/${driveId}/items/${itemId}/versions/${id}/content`, { ...opts, elevated: true });
  if (!old.ok) return old;
  const live = await renderSide(graph, file.value, `/drives/${driveId}/items/${itemId}/content`, opts);
  if (!live.ok) return live;
  const diff = diffEnvelope(`a/${name} (version ${id})`, old.value, `b/${name} (current)`, live.value, WORDING);
  return ok({ ...diff, versionId: id });
};

const meta: ReadCommandMeta = {
  summary:
    'What changed in a OneDrive or SharePoint file: the version saved before `--before`, the one named by `--version-id`, or by default the version saved just before the live one (what the last save changed), and the live file are both converted to markdown the way `download-drive-item-as-markdown` converts them, and the answer is a unified diff of the two (`--- a/<name> (version N)`, `+++ b/<name> (current)`, three lines of context) with the count of added and removed lines and the `versionId` compared. Nothing saved after the instant answers an empty diff with a note, without a download. `--include-metadata true` renders comments, tracked changes and properties on both sides, so a file whose body did not move but whose comments did still shows the change. A workbook sheet over the `--max-cells` cap is never reported as unchanged: the note names it, and `--sheet` compares one sheet alone. Loop, Fluid and Whiteboard pages are refused: Graph renders any old version of them as the current page. Needs the second token `login` captures for version history. For two different files, `diff-drive-items`.',
  category: 'drive',
  producesBytes: true,
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/versions/{version-id}/content and /drives/{drive-id}/items/{item-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitemversion-get-content',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: DRIVE_ID_DESCRIPTION },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID of the file, from `list-folder-files`, `search-all-files`, `list-changed-files` or `resolve-drive-share-link`.',
    },
    {
      name: 'version-id',
      key: 'versionId',
      required: false,
      description:
        'The older version to compare, from `list-drive-item-versions` (`4` and `4.0` both work). Not the first entry: that one is the live file. Without this or `--before`, the version saved just before the live one is compared.',
    },
    {
      name: 'before',
      key: 'before',
      required: false,
      description: `Instead of --version-id: compare the newest version saved strictly before this instant, the start of a reporting window. ${RELATIVE_DATE_DESCRIPTION}`,
    },
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
  example: "ask-marcel-office diff-drive-item-versions --drive-id 'b!1234' --item-id '01ABC' --before yesterday",
  responseShape:
    '`{ contentType: "text/x-diff", size, text, added, removed, versionId }` — `text` is the unified diff; `note` says when nothing was saved after the instant, when the two render the same, when they differ too much to diff (no counts), which sheets were over the `--max-cells` cap and not compared, and any note a conversion attached, prefixed with its side.',
  needsElevatedToken: true,
  effect: 'read',
  scopesRequired: ['Files.Read'],
};

export { execute, meta, schema };
