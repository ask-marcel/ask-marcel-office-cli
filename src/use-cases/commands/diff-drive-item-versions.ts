import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { refuseVersionRender, resolveVersion } from './download-drive-item-version.ts';
import { diffEnvelope, readDriveFile, renderSide } from './drive-item-diff.ts';
import { formatZodError } from './format-zod-error.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';
import { rendersThroughGraph } from './office-to-markdown.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';
import { MAX_CELLS_OPTION, maxCellsField } from './xlsx-to-markdown.ts';

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  versionId: z.string().min(1).optional(),
  before: isoDateTimeField.optional(),
  includeMetadata: z.enum(['true', 'false']).optional(),
  maxCells: maxCellsField,
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId, before, maxCells } = parsed.data;
  const file = await readDriveFile(graph, driveId, itemId);
  if (!file.ok) return file;
  const { name } = file.value;
  if (rendersThroughGraph(name)) return refuseVersionRender(name);
  const version = await resolveVersion(graph, driveId, itemId, parsed.data.versionId, before);
  if (!version.ok) return version;
  const { id, current } = version.value;
  // The newest version before the instant is still the live file: nothing moved.
  if (current)
    return ok({
      contentType: 'text/x-diff',
      size: 0,
      text: '',
      added: 0,
      removed: 0,
      versionId: id,
      note: `Nothing was saved after ${String(before)}: version ${id} is the live file.`,
    });
  const opts = { includeMetadata: parsed.data.includeMetadata === 'true', maxCells };
  const old = await renderSide(graph, file.value, `/drives/${driveId}/items/${itemId}/versions/${id}/content`, { ...opts, elevated: true });
  if (!old.ok) return old;
  const live = await renderSide(graph, file.value, `/drives/${driveId}/items/${itemId}/content`, opts);
  if (!live.ok) return live;
  const diff = diffEnvelope(`a/${name} (version ${id})`, old.value, `b/${name} (current)`, live.value, 'This version and the live file render to the same markdown.');
  return ok({ ...diff, versionId: id });
};

const meta: CommandMeta = {
  summary:
    'What changed in a OneDrive or SharePoint file since an instant: the version saved before `--before` (or the one named by `--version-id`) and the live file are both converted to markdown the way `download-drive-item-as-markdown` converts them, and the answer is a unified diff of the two (`--- a/<name> (version N)`, `+++ b/<name> (current)`, three lines of context) with the count of added and removed lines and the `versionId` compared. Nothing saved after the instant answers an empty diff with a note, without a download. `--include-metadata true` renders comments, tracked changes and properties on both sides, so a file whose body did not move but whose comments did still shows the change. Loop, Fluid and Whiteboard pages are refused: Graph renders any old version of them as the current page. Needs the second token `login` captures for version history. For two different files, `diff-drive-items`.',
  category: 'drive',
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
      description: 'The older version to compare, from `list-drive-item-versions` (`4` and `4.0` both work). Not the first entry: that one is the live file.',
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
  ],
  example: "ask-marcel-office diff-drive-item-versions --drive-id 'b!1234' --item-id '01ABC' --before yesterday",
  responseShape:
    '`{ contentType: "text/x-diff", size, text, added, removed, versionId }` — `text` is the unified diff; `note` instead says when nothing was saved after the instant, when the two render the same, or when they differ too much to diff (no counts).',
  needsElevatedToken: true,
};

export { execute, meta, schema };
