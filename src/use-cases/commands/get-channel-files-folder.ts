import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ teamId: z.string().min(1), channelId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/teams/${p.teamId}/channels/${p.channelId}/filesFolder`, baseSchema);

const meta: CommandMeta = {
  summary:
    "Return the SharePoint folder that backs a Teams channel's Files tab. Returned `driveItem` includes `parentReference.driveId` and `id` so you can pivot into `list-folder-files`, `download-onedrive-file-content`, etc., and treat the channel like any other OneDrive folder. Requires that the signed-in user is a member of the channel — restricted channels return `AccessDenied`.",
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/channels/{channel-id}/filesFolder',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/channel-get-filesfolder',
  options: [
    {
      name: 'team-id',
      key: 'teamId',
      required: true,
      description: 'Microsoft Teams team ID. Returned by `list-joined-teams`.',
    },
    {
      name: 'channel-id',
      key: 'channelId',
      required: true,
      description: 'Channel ID inside the team. Returned by `list-team-channels`.',
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-channel-files-folder --team-id 'tm1' --channel-id 'ch1'",
  responseShape: 'single Microsoft Graph `driveItem` resource (folder)',
};

export { execute, meta, schema };
