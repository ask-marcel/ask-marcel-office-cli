import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

// Deep Teams chat history via the IC3 messaging substrate at
// `teams.microsoft.com/api/chatsvc/<region>/v1/users/ME/conversations/{id}/messages`.
// This is the substrate Teams web client uses for chat scrollback. Unlike the
// chatsvcagg `/api/csa/...` route (which caps every request at 200 messages
// and ignores all cursor parameters), the IC3 path returns a server-provided
// **syncState URL** in `_metadata.syncState` that resumes the read exactly
// where the previous page ended. Empirically verified 2026-05-21: walking
// syncState across pages returned the full 289-message history of a chat
// where chatsvcagg returned only the newest 200.
//
// Protocol:
//   First call:   `?startTime=1&pageSize=N&view=msnp24Equivalent|supportsMessageProperties`
//                 — `startTime=1` is "from the beginning of time"; server
//                 returns the most recent N messages plus a `_metadata.syncState`
//                 URL for the next (older) page.
//   Next call:    Hit the syncState URL verbatim (server encodes its own
//                 startTime + cursor token); response includes another
//                 syncState URL until the chat's earliest message is reached.
//   Stop when:    `_metadata.syncState` is absent, OR response has zero
//                 messages, OR `--max-pages` cap is hit.
const schema = z.object({
  chatId: z.string().min(1),
  syncState: z.string().url().optional(),
  pageSize: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
  maxPages: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
});

type Ic3Message = { readonly id?: string; readonly sequenceId?: number };
type Ic3MessagesResponse = {
  readonly messages?: ReadonlyArray<Ic3Message>;
  readonly _metadata?: { readonly syncState?: string };
};

// Strip the absolute-URL prefix from a syncState so the relative path can
// be passed to `graph.teamsChatIc3` (which prepends host + region).
// Throws if the URL doesn't match the expected substrate shape — that's a
// signal the server changed its response format.
const toRelativePath = (absoluteUrl: string): string => {
  const m = /^https:\/\/teams\.microsoft\.com\/api\/chatsvc\/[a-z0-9-]+(\/.+)$/i.exec(absoluteUrl);
  if (m === null) throw new Error(`unexpected syncState URL shape: ${absoluteUrl}`);
  return m[1];
};

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { chatId } = parsed.data;
  const pageSize = parsed.data.pageSize ?? '200';
  const maxPages = Number(parsed.data.maxPages ?? '20');

  const accumulated: Array<Ic3Message> = [];
  let nextRelativePath: string =
    parsed.data.syncState !== undefined
      ? toRelativePath(parsed.data.syncState)
      : `/v1/users/ME/conversations/${encodeURIComponent(chatId)}/messages?startTime=1&pageSize=${pageSize}&view=msnp24Equivalent|supportsMessageProperties`;

  let nextSyncState: string | undefined;
  let pagesFetched = 0;
  for (let page = 0; page < maxPages; page += 1) {
    const result = await graph.teamsChatIc3(nextRelativePath);
    if (!result.ok) return result;
    const body = result.value as Ic3MessagesResponse;
    const messages = body.messages ?? [];
    accumulated.push(...messages);
    pagesFetched += 1;
    const syncStateUrl = body._metadata?.syncState;
    if (messages.length === 0 || syncStateUrl === undefined) {
      // Empty page or no continuation — end of history reached.
      nextSyncState = undefined;
      break;
    }
    nextRelativePath = toRelativePath(syncStateUrl);
    nextSyncState = syncStateUrl;
  }

  return ok({
    messages: accumulated,
    hasMore: nextSyncState !== undefined,
    pagesFetched,
    nextSyncState,
  });
};

const meta: CommandMeta = {
  summary:
    "Deep read of a Microsoft Teams chat's message history via the IC3 substrate (`teams.microsoft.com/api/chatsvc/<region>/v1/...`). Unlike `list-teams-chat-messages` (which caps at the 200 most recent messages with no working pagination cursor), this command follows the server-provided `_metadata.syncState` URL backward through history, fetching up to `--page-size` * `--max-pages` messages per invocation (default 200 * 20 = 4000). Uses the IC3-audience bearer captured at login (same Teams web client identity as the basic Teams token). **Best-effort, may break on Microsoft client updates** — the IC3 substrate is not in the public Microsoft Graph API. To page beyond `--max-pages`, take the response's `nextSyncState` and pass it back as `--sync-state` on the next call.",
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: 'https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{chat-id}/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-list',
  options: [
    {
      name: 'chat-id',
      key: 'chatId',
      required: true,
      description: 'Teams chat ID — typically `19:<thread>@unq.gbl.spaces` (1:1) or `19:<thread>@thread.v2` (group). Source via `list-chats` or `list-teams-chats-with-messages`.',
    },
    {
      name: 'sync-state',
      key: 'syncState',
      required: false,
      description:
        "Opaque pagination URL returned in the prior response's `nextSyncState` field. Pass it back to continue paging backward from where the previous invocation stopped. Omit on the first call.",
    },
    {
      name: 'page-size',
      key: 'pageSize',
      required: false,
      description: 'Messages per IC3 page (positive integer; default 200). Server may silently cap.',
    },
    {
      name: 'max-pages',
      key: 'maxPages',
      required: false,
      description:
        'Safety cap on the backward walk (positive integer; default 20). Each page can return up to `--page-size` messages, so default ceiling is ~4000 messages or ~4 MB inline JSON per invocation. Raise carefully — response is buffered fully before returning.',
    },
  ],
  example: "ask-marcel list-teams-chat-history --chat-id '19:abc...@unq.gbl.spaces' --max-pages 5",
  responseShape:
    "`{ messages: [...], hasMore: boolean, pagesFetched: number, nextSyncState?: string }`. IC3 substrate message shape (camelCase, matches the chat substrate envelope): `id`, `sequenceId` (monotonic per-chat counter), `composetime`, `originalarrivaltime`, `messagetype`, `content`, `from`, `imdisplayname`, `properties.subject`, etc. **`hasMore: true`** means the safety cap was hit and there is older history beyond what was returned — chain a follow-up call with `--sync-state $(jq -r .data.nextSyncState <prev>)` to continue. **`hasMore: false`** means the chat's earliest message was reached. **Microsoft-internal schema — fields may change without notice.**",
};

export { execute, meta, schema };
