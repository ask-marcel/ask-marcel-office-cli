# Teams commands

## get-channel-files-folder
Return the SharePoint folder that backs a Teams channel's Files tab. Returned `driveItem` includes `parentReference.driveId` and `id` so you can pivot into `list-folder-files`, `download-onedrive-file-content`, etc., and treat the channel like any other OneDrive folder. Requires that the signed-in u
Required: --team-id --channel-id
Optional: --select --expand
Example: ask-marcel get-channel-files-folder --team-id 'tm1' --channel-id 'ch1'
Graph: GET /teams/{team-id}/channels/{channel-id}/filesFolder

## get-team
Get the metadata of a single Microsoft Team (display name, settings, member-settings, owner group). Pass `--select displayName,description,visibility` to slim the response.
Required: --team-id
Optional: --select --expand
Example: ask-marcel get-team --team-id 'abc-1234-...' --select displayName,description,visibility
Graph: GET /teams/{team-id}

## get-team-channel
Get the metadata of a single channel inside a Microsoft Team. Use `--select` to slim the response (e.g. `--select id,displayName,webUrl`) — sibling to `get-team` and `get-team-primary-channel` which both expose the same flag.
Required: --team-id --channel-id
Optional: --select --expand
Example: ask-marcel get-team-channel --team-id 'abc-1234-...' --channel-id '19:def@thread.tacv2' --select 'id,displayName'
Graph: GET /teams/{team-id}/channels/{channel-id}

## get-team-primary-channel
Return the team's primary (General) channel directly without having to list-then-pick. The returned `channel` has `id`, `displayName`, `webUrl`, `email` — feed `id` into `list-team-channels` siblings or `get-channel-files-folder`.
Required: --team-id
Optional: --select --expand
Example: ask-marcel get-team-primary-channel --team-id 'tm1'
Graph: GET /teams/{team-id}/primaryChannel

## list-joined-teams
List the Microsoft Teams the signed-in user is a member of. Note: this endpoint does NOT accept the standard OData query parameters — Graph rejects `$top`/`$select`/`$filter`/etc. on `/me/joinedTeams` with `Query option 'X' is not allowed`. The CLI omits the OData passthrough on this command for tha
Example: ask-marcel list-joined-teams
Graph: GET /me/joinedTeams

## list-team-channels
List the channels (standard, private, shared) inside a single Microsoft Team. Microsoft documents this endpoint as supporting only `$filter` and `$select` — Graph returns `BadRequest` on `$top`, `$skip`, `$orderby`, `$expand`, so the CLI exposes only the two flags that actually work.
Required: --team-id
Optional: --select --filter
Example: ask-marcel list-team-channels --team-id 'abc-1234-...'
Graph: GET /teams/{team-id}/channels

## list-team-installed-apps
List the Teams apps installed in a team. The CLI hard-pins `$expand=teamsAppDefinition` so every entry includes `displayName`, `version`, and `distributionMethod` (the bare endpoint returns only opaque IDs). Useful for surfacing which integrations are wired into a given team. Graph rejects user-supp
Required: --team-id
Example: ask-marcel list-team-installed-apps --team-id 'tm1'
Graph: GET /teams/{team-id}/installedApps?$expand=teamsAppDefinition
