import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { detectSiblingResolver } from './link-shape.ts';

// Deep-link parser for Microsoft Teams message links of the shape:
//   https://teams.microsoft.com/l/message/<url-encoded-chat-id>/<message-id>?<query>
// These are the links emitted when a user clicks "Copy link" on a message
// in Teams. The two path segments after `/l/message/` are the chat-id (URL-
// encoded — typically `19%3A...%40unq.gbl.spaces` for 1:1 chats) and the
// message-id (typically the millisecond-epoch composeTime — `1700000000000`).
// Optional query params Teams adds: `groupId`, `tenantId`, `ctx`, `parentMessageId`.
//
// Pure transformation — no HTTP. Pair with `get-teams-chat-message` to
// fetch the message body once the link is resolved.
const schema = z.object({
  url: z.url(),
});

const PREFIX = 'https://teams.microsoft.com/l/message/';

type Resolved = {
  readonly chatId: string;
  readonly messageId: string;
  readonly tenantId?: string;
  readonly groupId?: string;
  readonly parentMessageId?: string;
  readonly context?: string;
};

const parse = (raw: string): Resolved | null => {
  if (!raw.startsWith(PREFIX)) return null;
  const remainder = raw.slice(PREFIX.length);
  // Split on '?' to separate the path from the query string.
  const [pathPart, queryPart] = remainder.includes('?') ? [remainder.slice(0, remainder.indexOf('?')), remainder.slice(remainder.indexOf('?') + 1)] : [remainder, ''];
  const segments = pathPart.split('/');
  if (segments.length < 2) return null;
  const chatIdRaw = segments[0];
  const messageIdRaw = segments[1];
  if (chatIdRaw === undefined || messageIdRaw === undefined || chatIdRaw === '' || messageIdRaw === '') return null;
  const query = new URLSearchParams(queryPart);
  const optional: { tenantId?: string; groupId?: string; parentMessageId?: string; context?: string } = {};
  const tenantId = query.get('tenantId');
  if (tenantId !== null) optional.tenantId = tenantId;
  const groupId = query.get('groupId');
  if (groupId !== null) optional.groupId = groupId;
  const parentMessageId = query.get('parentMessageId');
  if (parentMessageId !== null) optional.parentMessageId = parentMessageId;
  const context = query.get('ctx');
  if (context !== null) optional.context = context;
  return {
    chatId: decodeURIComponent(chatIdRaw),
    messageId: decodeURIComponent(messageIdRaw),
    ...optional,
  };
};

const execute: Command['execute'] = async (_graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  // v1.4.0 re-audit Nit 1 (drive-share + sharepoint + outlook gaps): a
  // OneDrive / SharePoint share URL or an Outlook web URL wrongly passed
  // to resolve-teams-link used to fall through to the generic "not a
  // Teams message link" rejection. Detect them early and emit a
  // cross-pointer so the LLM lands on the right sibling resolver.
  const sibling = detectSiblingResolver(parsed.data.url);
  if (sibling === 'drive-share') {
    return err({
      type: 'validation_error',
      message: '--url looks like a OneDrive / SharePoint sharing URL, not a Teams message link.',
      code: 'cli_reject_drive_share_link_on_teams_resolver',
    });
  }
  if (sibling === 'mail') {
    return err({
      type: 'validation_error',
      message: '--url looks like an Outlook mail message link, not a Teams message link.',
      code: 'cli_reject_mail_link_on_teams_resolver',
    });
  }
  if (sibling === 'calendar') {
    return err({
      type: 'validation_error',
      message: '--url looks like an Outlook calendar item link, not a Teams message link.',
      code: 'cli_reject_calendar_link_on_teams_resolver',
    });
  }
  const resolved = parse(parsed.data.url);
  if (resolved === null) {
    return err({
      type: 'validation_error',
      message: `--url: not a Teams message link. Expected shape: ${PREFIX}<chat-id>/<message-id>[?tenantId=...&groupId=...]`,
    });
  }
  return ok(resolved);
};

const meta: CommandMeta = {
  summary:
    'Parse a Microsoft Teams `Copy link` URL (the share link emitted by the message context menu in Teams) into its `chatId` + `messageId` components. Pure transformation — no Graph call. Pipe the result into `get-teams-chat-message` to fetch the message body, or into `list-teams-chat-history` to read the chat that contains it.',
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: '{url}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-get',
  options: [
    {
      name: 'url',
      key: 'url',
      required: true,
      description:
        "Teams message link, copied from Teams web/desktop via the message's `Copy link` action. Expected shape: `https://teams.microsoft.com/l/message/<url-encoded-chat-id>/<message-id>?tenantId=...&groupId=...&ctx=...`",
    },
  ],
  example: "ask-marcel resolve-teams-link --url 'https://teams.microsoft.com/l/message/19%3A...%40unq.gbl.spaces/1700000000000?tenantId=...&groupId=...&ctx=chat'",
  responseShape:
    '`{ chatId: string, messageId: string, tenantId?: string, groupId?: string, parentMessageId?: string, context?: string }`. `chatId` and `messageId` are URL-decoded and ready to pass to other commands. Optional fields are included only when the source URL carried them.',
};

export { execute, meta, schema };
