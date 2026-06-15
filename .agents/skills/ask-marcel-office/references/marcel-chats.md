# Chats commands

## find-chats-with-user
Find every Microsoft Teams chat that includes a member matching `--name` (substring search across display-name, email, given-name, surname, MRI, and object-id). Both sides are Unicode-folded (NFD + combining-mark strip) and lowercased before comparison, so `--name Jane` matches `Jane DOE` AND `jane.
Required: --name
Optional: --max-pages --page-size
Example: ask-marcel find-chats-with-user --name 'Jane DOE'
Graph: GET https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me/chats

## get-chat
Return metadata for a single Microsoft Teams chat (1:1, group, or meeting). The CLI ships a slim default `--select=id,topic,chatType,createdDateTime,lastUpdatedDateTime`; pass `--select id,topic,webUrl,onlineMeetingInfo` (or any other comma-separated field list) to widen. Pass `--expand members` to 
Required: --chat-id
Optional: --select --expand
Example: ask-marcel get-chat --chat-id '19:abc...@thread.v2'
Graph: GET /chats/{chat-id}

## get-teams-chat-message
Return a single Microsoft Teams chat message by its id via the chat substrate. Uses the chatsvcagg-audience bearer captured at login (same identity as the basic Teams token, different audience). **Best-effort, may break on Microsoft client updates** — the chat substrate is not in the public Microsof
Required: --chat-id --message-id
Example: ask-marcel get-teams-chat-message --chat-id '19:abc...@unq.gbl.spaces' --message-id '1700000000000'
Graph: GET https://teams.microsoft.com/api/csa/{region}/api/v1/chats/{chat-id}/messages/{message-id}

## list-chat-members
List the members of a single Microsoft Teams chat. Graph rejects `$top` / `$orderby` / `$expand` on this endpoint, so the CLI advertises only the subset Graph honours (`--skip`, `--select`, `--filter`).
Required: --chat-id
Optional: --skip --select --filter
Example: ask-marcel list-chat-members --chat-id '19:abc...@thread.v2'
Graph: GET /chats/{chat-id}/members

## list-chats
List the signed-in user's Microsoft Teams chats (1:1, group, and meeting chats). The CLI ships a slim default `--select=id,topic,chatType,createdDateTime,lastUpdatedDateTime`; pass `--select id,topic,webUrl,...` to widen. Returns chat metadata only — reading chat *messages* needs `Chat.Read*` which 
Optional: --top --skip --select --filter
Example: ask-marcel list-chats
Graph: GET /me/chats

## list-teams-chat-history
Deep read of a Microsoft Teams chat's message history via the IC3 substrate (`teams.microsoft.com/api/chatsvc/<region>/v1/...`). Unlike `list-teams-chat-messages` (which caps at the 200 most recent messages with no working pagination cursor), this command follows the server-provided `_metadata.syncS
Required: --chat-id
Optional: --sync-state --page-size --max-pages --full --max-content-chars
Example: ask-marcel list-teams-chat-history --chat-id '19:abc...@unq.gbl.spaces' --max-pages 5
Graph: GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{chat-id}/messages

## list-teams-chat-messages
List the most recent messages in a single Microsoft Teams chat via the chat substrate. Companion to `list-teams-chats-with-messages` when the inlined `lastMessage` isn't deep enough. Uses the chatsvcagg-audience bearer captured at login. **Best-effort, may break on Microsoft client updates** — the c
Required: --chat-id
Example: ask-marcel list-teams-chat-messages --chat-id '19:abc...@unq.gbl.spaces'
Graph: GET https://teams.microsoft.com/api/csa/{region}/api/v1/chats/{chat-id}/messages

## list-teams-chats-with-messages
List the signed-in user's Microsoft Teams chats with the last message body inlined per chat. Uses the chatsvcagg-audience bearer captured at login. Paginated via `continuationToken` (default page size 100; pass the response's `continuationToken` back as `--continuation-token` while `hasMoreData: tru
Optional: --page-size --continuation-token
Example: ask-marcel list-teams-chats-with-messages --page-size 100
Graph: GET https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me/chats

## resolve-teams-link
Parse a Microsoft Teams `Copy link` URL (the share link emitted by the message context menu in Teams) into its `chatId` + `messageId` components. Pure transformation — no Graph call. Pipe the result into `get-teams-chat-message` to fetch the message body, or into `list-teams-chat-history` to read th
Required: --url
Example: ask-marcel resolve-teams-link --url 'https://teams.microsoft.com/l/message/19%3A...%40unq.gbl.spaces/1700000000000?tenantId=...&groupId=...&ctx=chat'
Graph: GET {url}
