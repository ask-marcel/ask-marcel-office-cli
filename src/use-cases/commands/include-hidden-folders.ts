import type { CommandOptionMeta } from './command-types.ts';

// The `includeHiddenFolders` flag shared by the four mail-folder LISTING
// commands (`list-mail-folders`, `list-mail-child-folders` and their
// shared-mailbox siblings). Kept as ONE constant so the four cannot drift into
// describing the same Graph parameter differently.
//
// Unlike every other passthrough on these commands this is a PLAIN query
// param, not an OData `$` one, so `appendOData` cannot emit it — each command
// appends it inside its own path function and the meta.test placeholder
// invariant exempts the name.
export const INCLUDE_HIDDEN_FOLDERS_OPTION: CommandOptionMeta = {
  name: 'include-hidden-folders',
  key: 'includeHiddenFolders',
  required: false,
  description:
    'Graph HIDES folders flagged as hidden from this listing by default (a mailbox typically hides system folders such as `Social Activity Notifications`, and an Outlook client can hide any folder). Pass `true` to append the plain `includeHiddenFolders=true` query parameter and get them as well; every returned folder then carries `isHidden` to tell the two apart. Omit it for the default visible-only listing.',
  argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
};
