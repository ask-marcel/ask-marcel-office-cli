import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/drive/sharedWithMe', schema);

const meta: CommandMeta = {
  summary:
    'List driveItems shared with the signed-in user (typically by colleagues). Each entry includes the original drive + item ID under `remoteItem` so you can chain into `get-drive-item`, `download-onedrive-file-content`, etc.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/sharedWithMe',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-sharedwithme',
  options: [],
  example: 'ask-marcel list-shared-with-me',
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]` (each with a `remoteItem` pointer)',
  pagination: true,
};

export { execute, meta, schema };
