import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

// Microsoft Teams chat substrate, paginated chat-list endpoint
// (`/api/csa/<region>/api/v3/teams/users/me/chats`). Returns chats with
// `lastMessage` inlined per chat, plus a `continuationToken` for the next
// page. Probed live 2026-05-21 — supports real pagination via
// `?continuationToken=<token>&pageSize=N`. Each page returns up to ~100
// chats; iterate until `hasMoreData: false`.
//
// Discovery context: an earlier sibling endpoint at `/teams/users/me`
// (no trailing `/chats`) is the chat-list AGGREGATE — returns chats+teams
// +folders+users+metadata in one ~1.2 MB blob, capped server-side at 273
// chats with no working pagination cursor. We previously used that and
// dropped the page-size/skip-token flags as non-functional. The dedicated
// `/chats` sibling is paginated AND smaller (chats only). See
// `gotcha_chatsvcagg_substrate_moved` in memory for the substrate audit.
const schema = z.object({
  pageSize: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
  continuationToken: z.string().min(1).optional(),
});

const QUERY_BASE = 'enableMembershipSummary=true&supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true&enableEngageCommunities=false';

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const pageSize = parsed.data.pageSize ?? '100';
  const qs = new URLSearchParams({ pageSize });
  if (parsed.data.continuationToken !== undefined) qs.set('continuationToken', parsed.data.continuationToken);
  return graph.teamsChat(`/api/v3/teams/users/me/chats?${qs.toString()}&${QUERY_BASE}`);
};

const meta: CommandMeta = {
  summary:
    "List the signed-in user's Microsoft Teams chats with the last message body inlined per chat. Uses the chatsvcagg-audience bearer captured at login. Paginated via `continuationToken` (default page size 100; pass the response's `continuationToken` back as `--continuation-token` while `hasMoreData: true`). **Best-effort, may break on Microsoft client updates**: the chat substrate is not part of the public Microsoft Graph API; Microsoft can change route shapes without notice. Caller Graph scopes do NOT matter here; the substrate server gates access on the appid + identity, not on Graph scopes.",
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: 'https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me/chats',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-list',
  options: [
    {
      name: 'page-size',
      key: 'pageSize',
      required: false,
      description: 'Chats per page (positive integer; default 100, same value Teams web uses). Server may silently cap.',
    },
    {
      name: 'continuation-token',
      key: 'continuationToken',
      required: false,
      description: "Opaque pagination cursor returned in the prior response's `continuationToken` field. Omit on the first call; loop until `hasMoreData` is false.",
    },
  ],
  example: 'ask-marcel list-teams-chats-with-messages --page-size 100',
  responseShape:
    '`{ chats: [...], continuationToken?: string, hasMoreData?: boolean }`. Each chat carries `id`, `title`, `chatType`, `threadType`, `members[]` (with each member\'s `mri`, `displayName`, `email`), `createdAt`, AND `lastMessage` (the most recent message body inlined — `content`, `from`, `composeTime`, `imDisplayName`, etc.). When `hasMoreData: true`, chain a follow-up call with `--continuation-token "$(jq -r .data.continuationToken <prev>)"`. **Microsoft-internal schema — fields may change without notice; treat the response as semi-structured.**',
};

export { execute, meta, schema };
