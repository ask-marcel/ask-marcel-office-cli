import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { type ChannelMessage } from './channel-message-html.ts';
import { fetchHostedImages, hostedImageSources } from './channel-message-images.ts';
import { renderThread } from './channel-message-markdown.ts';
import { markdownEnvelope, relativeGraphPath } from './channel-message-page.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { channelScopeOf, rewriteChannelScopedError } from './team-channel-errors.ts';

const schema = z.object({
  teamId: z.string().min(1),
  channelId: z.string().min(1),
  messageId: z.string().min(1),
  inlineImages: z.enum(['true', 'false']).optional(),
});

const REPLIES_PAGE = 50;
const MAX_REPLY_PAGES = 10;

type RepliesRead = { readonly replies: ReadonlyArray<ChannelMessage>; readonly notes: ReadonlyArray<string> };

// Replies come fifty a page (Graph's cap); a thread longer than ten pages is
// cut with a note rather than walked forever. A failed page leaves the post
// rendered on its own, said in the note, the way the mail renderer survives a
// failed attachments list.
const readReplies = async (graph: GraphClient, messagePath: string): Promise<RepliesRead> => {
  const replies: ChannelMessage[] = [];
  let path: string | undefined = `${messagePath}/replies?$top=${REPLIES_PAGE}`;
  let pages = 0;
  while (path !== undefined) {
    if (pages === MAX_REPLY_PAGES) return { replies, notes: [`replies truncated after ${MAX_REPLY_PAGES} pages (${replies.length} replies rendered); the thread has more`] };
    const page = await graph.get(path);
    if (!page.ok) return { replies, notes: [`replies fetch failed (${page.error.type}: ${page.error.message}); the post is rendered without its replies`] };
    pages += 1;
    const body = page.value as { readonly value?: ReadonlyArray<ChannelMessage>; readonly '@odata.nextLink'?: string };
    replies.push(...(body.value ?? []));
    path = relativeGraphPath(body['@odata.nextLink']);
  }
  return { replies, notes: [] };
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { teamId, channelId, messageId } = parsed.data;
  const messagePath = `/teams/${teamId}/channels/${channelId}/messages/${messageId}`;
  const fetched = rewriteChannelScopedError(await graph.get(messagePath), channelScopeOf(params));
  if (!fetched.ok) return fetched;
  const root = fetched.value as ChannelMessage;
  const { replies, notes } = await readReplies(graph, messagePath);
  const inline = parsed.data.inlineImages === 'true';
  const images = inline ? await fetchHostedImages(graph, hostedImageSources([{ ...root, replies }])) : new Map<string, string>();
  return ok(markdownEnvelope(renderThread(root, replies, images, inline), notes));
};

const meta: CommandMeta = {
  summary:
    'Render one post of a Microsoft Teams channel and the replies under it as a markdown thread, the way `convert-group-post-to-markdown` renders a group post: a `### date time · author` heading (`· high importance` when flagged), the subject in bold when the post has one, the body converted from Teams HTML with @mentions flattened to their text, attachment placeholders resolved to `[name](url)` links, `[meeting: …]`, `[tab: …]` or `[card: …]` summaries, then each reply as a quoted block, oldest first, with `_(edited)_`, `_message deleted_` and `_reactions: 👍 2_` markers. Replies are read fifty at a time up to ten pages. Pasted images live in Graph `hostedContents` and are shown as a placeholder unless `--inline-images true` fetches and embeds them (image/*, 2 MB or less). Reads through Graph on the basic token; an unknown message id is named in the error rather than surfacing the `403 UnknownError` Graph answers.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/channels/{channel-id}/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-get',
  options: [
    { name: 'team-id', key: 'teamId', required: true, description: 'Microsoft Teams team ID. Returned by `ask-marcel-office list-joined-teams`.' },
    { name: 'channel-id', key: 'channelId', required: true, description: 'Channel ID (`19:<thread>@thread.tacv2`). Returned by `ask-marcel-office list-team-channels`.' },
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      description: 'Root post ID, the numeric `id` of a `list-team-channel-messages` entry. A reply id renders that reply alone.',
    },
    {
      name: 'inline-images',
      key: 'inlineImages',
      required: false,
      description:
        'Pass `--inline-images true` to fetch pasted images from Graph `hostedContents` and embed them as base64 `data:` URIs (image/* only, 2 MB or less). Default is `false`: no per-image fetch, each image becomes a placeholder naming this flag.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel-office convert-team-channel-message-to-markdown --team-id 'abc-1234-...' --channel-id '19:def@thread.tacv2' --message-id '1700000000000'",
  responseShape:
    '`{ contentType: "text/markdown", size, text, note? }`, the same envelope as `convert-mail-to-markdown`: the post heading, its body, then the replies as quoted blocks. The optional `note` reports a failed or truncated replies read.',
  producesBytes: true,
};

export { execute, meta, schema };
