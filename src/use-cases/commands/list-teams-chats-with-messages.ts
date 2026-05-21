import { z } from 'zod';
import type { Command, CommandMeta } from './command-types.ts';

// Microsoft Teams chat substrate aggregator. Returns the signed-in user's
// chats AND inlines the LAST message body per chat in a single round-trip.
// This is the path Teams web/desktop uses to populate the chat sidebar —
// Graph's `Chat.Read*`-gated endpoints can't reach the message bodies with
// the scopes the CLI's basic Teams token carries, but the chatsvcagg-audience
// bearer captured at login CAN access them via the post-2026-05 substrate
// (`teams.microsoft.com/api/csa/<region>/api/v3/teams/users/me`).
//
// Pagination flags from the pre-2026-05 surface (`--page-size` / `--skip-token`)
// were removed: the new substrate returns ALL chats in a single ~1 MB blob,
// with no per-page cursor. See `gotcha_chatsvcagg_substrate_moved` in memory.
const QUERY_STRING = new URLSearchParams({
  isPrefetch: 'false',
  enableMembershipSummary: 'true',
  supportsAdditionalSystemGeneratedFolders: 'true',
  supportsSliceItems: 'true',
  enableEngageCommunities: 'false',
}).toString();

const schema = z.object({}).strict();

const execute: Command['execute'] = async (graph) => graph.teamsChat(`/api/v3/teams/users/me?${QUERY_STRING}`);

const meta: CommandMeta = {
  summary:
    "List the signed-in user's Microsoft Teams chats with the last message body inlined per chat. Uses the chatsvcagg-audience bearer captured at login (same identity as the basic Teams token, different audience) — this is the path Teams web/desktop uses to populate the chat sidebar. **Best-effort, may break on Microsoft client updates**: the chat substrate is not part of the public Microsoft Graph API; Microsoft can change route shapes without notice. If the response shape looks different from what you expect, run `ask-marcel logout && ask-marcel login` to refresh the captured tokens AND re-detect the substrate region. Caller Graph scopes do NOT matter here; the substrate server gates access on the appid + identity, not on Graph scopes.",
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: 'https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-list',
  options: [],
  example: 'ask-marcel list-teams-chats-with-messages',
  responseShape:
    'Substrate envelope: `{ chats: [...], teams: [...], conversationFolders, engageCommunities, metadata, privateFeeds, users }`. Each chat carries `id`, `title`, `chatType`, `threadType`, `members`, AND a `lastMessage` object with the most recent message body (`lastMessage.content`, `lastMessage.from`, `lastMessage.composeTime`, etc.) — this is the high-value field the existing `list-chats` cannot return. Returns ALL chats in one response (no server-side pagination); for deeper per-chat history call `list-teams-chat-messages`. **Schema is Microsoft-internal — fields may change without notice; treat the response as semi-structured.**',
};

export { execute, meta, schema };
