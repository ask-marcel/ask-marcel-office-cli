import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

// Graph documents `/me/drive/sharedWithMe` as NOT supporting any OData query
// parameters (https://learn.microsoft.com/en-us/graph/api/drive-sharedwithme).
// Advertising --top / --select / --filter / etc. on this command would be a
// usability lie — Graph silently ignores them and the response is always the
// full ~500-item list. Slice client-side or use `--output-path` to land the
// raw JSON on disk for downstream processing.
const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/drive/sharedWithMe', schema);

const meta: CommandMeta = {
  summary:
    'List driveItems shared with the signed-in user (typically by colleagues). Each entry includes the original drive + item ID under `remoteItem` so you can chain into `get-drive-item`, `download-onedrive-file-content`, etc. Note: Graph does NOT honor any OData query parameters on this endpoint (top/select/filter/etc. are all silently ignored), so the CLI does not advertise them. The full collection (~500 items in a typical tenant) is always returned; slice client-side or pair with the global output-path flag to land the raw JSON on disk.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/sharedWithMe',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-sharedwithme',
  options: [],
  example: 'ask-marcel list-shared-with-me',
  responseShape: 'collection of Microsoft Graph `driveItem` resources under `value[]` (each with a `remoteItem` pointer)',
};

export { execute, meta, schema };
