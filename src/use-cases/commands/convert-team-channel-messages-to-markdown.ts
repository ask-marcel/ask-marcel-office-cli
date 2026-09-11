import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { type ChannelMessage } from './channel-message-html.ts';
import { fetchHostedImages, hostedImageSources } from './channel-message-images.ts';
import { renderTranscript } from './channel-message-markdown.ts';
import { CHANNEL_MESSAGES_TOP_CAP, CHANNEL_MESSAGES_TOP_OPTION, markdownEnvelope, relativeGraphPath, withChannelMessagesTopCap } from './channel-message-page.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';
import { topOnlyShape } from './odata-query.ts';
import { channelScopeOf, rewriteChannelScopedError } from './team-channel-errors.ts';

const DEFAULT_MAX_PAGES = 10;
const MAX_PAGES_PATTERN = /^(?:[1-9]|[1-4]\d|50)$/;

const schema = z
  .object({
    teamId: z.string().min(1),
    channelId: z.string().min(1),
    since: isoDateTimeField.optional(),
    maxPages: z.string().regex(MAX_PAGES_PATTERN, 'must be a whole number from 1 to 50').optional(),
    inlineImages: z.enum(['true', 'false']).optional(),
  })
  .extend(topOnlyShape);

type Page = { readonly value?: ReadonlyArray<ChannelMessage>; readonly '@odata.nextLink'?: string };
type Walk = { readonly messages: ReadonlyArray<ChannelMessage>; readonly pages: number; readonly truncated: boolean };

const plural = (n: number, word: string): string => `${n} ${word}${n === 1 ? '' : 's'}`;

// Without a date the newest page of the channel is the transcript. With one,
// the delta route filtered on `lastModifiedDateTime` walks every root post
// touched since then (a fresh reply bumps its root, probed live 2026-09-10),
// page after page, until the deltaLink or the page cap.
const walkSince = async (graph: GraphClient, first: string, maxPages: number): Promise<Result<Walk, GraphError>> => {
  const messages: ChannelMessage[] = [];
  let path: string | undefined = first;
  let pages = 0;
  while (path !== undefined && pages < maxPages) {
    const page = await graph.get(path);
    if (!page.ok) return page;
    pages += 1;
    const body = page.value as Page;
    messages.push(...(body.value ?? []));
    path = relativeGraphPath(body['@odata.nextLink']);
  }
  return ok({ messages, pages, truncated: path !== undefined });
};

const channelNameOf = async (graph: GraphClient, channelPath: string): Promise<string | undefined> => {
  const fetched = await graph.get(`${channelPath}?$select=displayName`);
  if (!fetched.ok) return undefined;
  const name = (fetched.value as { readonly displayName?: unknown }).displayName;
  return typeof name === 'string' ? name : undefined;
};

const execute: Command['execute'] = withChannelMessagesTopCap(async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { teamId, channelId, since } = parsed.data;
  const top = parsed.data.top ?? String(CHANNEL_MESSAGES_TOP_CAP);
  const maxPages = parsed.data.maxPages === undefined ? DEFAULT_MAX_PAGES : Number(parsed.data.maxPages);
  const channelPath = `/teams/${teamId}/channels/${channelId}`;
  const channelName = await channelNameOf(graph, channelPath);
  const first =
    since === undefined
      ? `${channelPath}/messages?$top=${top}&$expand=replies`
      : `${channelPath}/messages/delta?$filter=lastModifiedDateTime%20gt%20${since}&$top=${top}&$expand=replies`;
  const walked = rewriteChannelScopedError(await walkSince(graph, first, since === undefined ? 1 : maxPages), channelScopeOf(params));
  if (!walked.ok) return walked;
  const inline = parsed.data.inlineImages === 'true';
  const images = inline ? await fetchHostedImages(graph, hostedImageSources(walked.value.messages)) : new Map<string, string>();
  const rendered = renderTranscript(walked.value.messages, { channelName, images, inline });
  const notes = [
    since === undefined
      ? `${plural(rendered.posts, 'post')} on the newest page`
      : `${plural(rendered.posts, 'post')} touched since ${since}, read in ${plural(walked.value.pages, 'page')}`,
  ];
  if (rendered.systemEventsOmitted > 0) notes.push(`${plural(rendered.systemEventsOmitted, 'system event')} omitted`);
  if (since !== undefined && walked.value.truncated) notes.push(`truncated at --max-pages ${maxPages}, older posts remain: raise --max-pages or narrow --since`);
  return ok(markdownEnvelope(rendered.text, notes));
});

const meta: CommandMeta = {
  summary:
    'Render a Microsoft Teams channel as one markdown transcript, oldest post first, each post with its replies quoted beneath it: the catch-up read for "what was said in this channel". Without `--since` it renders the newest page (`--top`, up to 50 root posts). With `--since` (an ISO instant or a relative date such as `7d`) it walks the delta route filtered on `lastModifiedDateTime`, so every root post touched since then comes back with its replies, up to `--max-pages` pages of `--top` posts (defaults 10 and 50, so 500 posts). Membership and channel events are omitted and counted in the `note`; a truncated walk is said there too. Bodies render as in `convert-team-channel-message-to-markdown` (@mentions, attachment links, card summaries, edit, delete and reaction markers); pasted images are placeholders unless `--inline-images true` embeds them. Reads through Graph on the basic token, no chat-substrate warm-up; an unknown channel id is named in the error.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/teams/{team-id}/channels/{channel-id}/messages/delta?$filter=lastModifiedDateTime gt {since}&$expand=replies',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-delta',
  options: [
    { name: 'team-id', key: 'teamId', required: true, description: 'Microsoft Teams team ID. Returned by `ask-marcel-office list-joined-teams`.' },
    {
      name: 'channel-id',
      key: 'channelId',
      required: true,
      description: 'Channel ID (`19:<thread>@thread.tacv2`). Returned by `ask-marcel-office list-team-channels`, or `get-team-primary-channel` for General.',
    },
    {
      name: 'since',
      key: 'since',
      required: false,
      description: `Lower bound on \`lastModifiedDateTime\`: every root post created, edited or replied to after this instant is included. Omit it to render only the newest page. ${RELATIVE_DATE_DESCRIPTION}`,
    },
    CHANNEL_MESSAGES_TOP_OPTION,
    {
      name: 'max-pages',
      key: 'maxPages',
      required: false,
      description: 'With `--since`: how many pages of `--top` posts to walk before stopping, 1 to 50 (default 10). A cut walk is reported in the `note`.',
    },
    {
      name: 'inline-images',
      key: 'inlineImages',
      required: false,
      description:
        'Pass `--inline-images true` to fetch pasted images from Graph `hostedContents` and embed them as base64 `data:` URIs (image/* only, 2 MB or less). Default is `false`: each image becomes a placeholder naming this flag.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel-office convert-team-channel-messages-to-markdown --team-id 'abc-1234-...' --channel-id '19:def@thread.tacv2' --since '7d'",
  responseShape:
    '`{ contentType: "text/markdown", size, text, note }`: a `# Transcript of <channel>` heading, then every post as `### date time · author` with its body and quoted replies. The `note` counts the posts and pages read, the system events omitted, and says when the walk was cut at `--max-pages`.',
  producesBytes: true,
};

export { execute, meta, schema };
