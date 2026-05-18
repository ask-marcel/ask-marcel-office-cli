import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ folderName: z.enum(['documents', 'photos', 'cameraroll', 'approot', 'music', 'attachments']) });
const { execute, schema } = buildSelectableCommand((p) => `/me/drive/special/${p.folderName}`, baseSchema);

const meta: CommandMeta = {
  summary:
    "Resolve a OneDrive well-known folder via `--folder-name` (one of `documents`, `photos`, `cameraroll`, `approot`, `music`, `attachments`) without having to navigate from the root. Returns the folder's driveItem (id, name, parentReference, etc.) ready to feed into `list-folder-files` or `download-onedrive-file-content`.",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/special/{folder-name}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-get-specialfolder',
  options: [
    {
      name: 'folder-name',
      key: 'folderName',
      required: true,
      description:
        'Well-known folder name. One of: `documents`, `photos`, `cameraroll`, `approot`, `music`, `attachments`. Returns the corresponding driveItem (folder) — use the returned `id` with `list-folder-files` etc.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-drive-special-folder --folder-name 'documents'",
  responseShape: 'single Microsoft Graph `driveItem` resource (folder)',
};

export { execute, meta, schema };
