import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';

/**
 * The ids a channel-scoped command was called with, so a rewritten error can
 * name them. `messageId` is set only by the message routes.
 */
type ChannelScope = { readonly channelId: string; readonly messageId?: string };

const idOf = (params: Record<string, string>, key: string): string => params[key] ?? '<unknown>';

/** Reads the channel (and message, when present) ids out of raw CLI params. */
const channelScopeOf = (params: Record<string, string>): ChannelScope => ({ channelId: idOf(params, 'channelId'), messageId: params['messageId'] });

const LIST_CHANNELS_HINT = 'Verify it exists in this team via `ask-marcel-office list-team-channels --team-id <team-id>`.';

const channelNotFound = (status: number, channelId: string, cause = ''): GraphError => ({
  type: 'api_error',
  status,
  message: `NotFound: Microsoft Teams channel not found (channel-id: "${channelId}"). ${cause}${LIST_CHANNELS_HINT}`,
  code: 'cli_rewrite_channel_not_found',
});

const messageNotFound = (status: number, scope: ChannelScope): GraphError => ({
  type: 'api_error',
  status,
  message:
    `NotFound: Microsoft Teams channel message not found (channel-id: "${scope.channelId}", message-id: "${scope.messageId}"). ` +
    `Graph answers \`${status} Forbidden: UnknownError\` for a message id it cannot resolve in this channel; this is not a missing scope. ` +
    'Source the id via `ask-marcel-office list-team-channel-messages --team-id <team-id> --channel-id <channel-id>`.',
  code: 'cli_rewrite_channel_message_not_found',
});

const badId = (status: number, scope: ChannelScope): GraphError => {
  const ids = scope.messageId === undefined ? `channel-id: "${scope.channelId}"` : `channel-id: "${scope.channelId}", message-id: "${scope.messageId}"`;
  return {
    type: 'api_error',
    status,
    message:
      `BadRequest: Microsoft Graph could not parse an id on this channel route (${ids}). ` +
      `Graph answers \`${status} UnknownError\` for a malformed channel or message id: a channel id looks like \`19:<thread>@thread.tacv2\` (from \`ask-marcel-office list-team-channels\`) ` +
      'and a message id is the numeric id from `ask-marcel-office list-team-channel-messages`.',
    code: 'cli_rewrite_channel_bad_id',
  };
};

/**
 * Graph answers the channel routes with errors that name nothing. Probed live
 * 2026-09-09 on the basic token: a channel id it cannot resolve is a bare
 * `1: NotFound` on `/members` (the `1:` is the thread-id segment, echoed
 * without context), a Skype backend `GetThreadRequest` failure on `/tabs`, and
 * a `410 Gone: UnknownError` on `/messages`; a malformed channel or message
 * id is a `400 BadRequest: UnknownError`; and an unknown MESSAGE id is a
 * `403 Forbidden: UnknownError`, which an agent would read as a scope failure.
 * Each is rewritten to name the id at fault and where to source it. Every
 * other outcome, success included, passes through untouched.
 */
const rewriteChannelScopedError = <T>(result: Result<T, GraphError>, scope: ChannelScope): Result<T, GraphError> => {
  if (result.ok || result.error.type !== 'api_error') return result;
  const { status, message } = result.error;
  if (/^1:\s*NotFound/i.test(message) || message.includes('GetThreadRequest')) return err(channelNotFound(status, scope.channelId));
  if (!message.includes('UnknownError')) return result;
  if (status === 400) return err(badId(status, scope));
  if (status === 403 && scope.messageId !== undefined) return err(messageNotFound(status, scope));
  return err(channelNotFound(status, scope.channelId, `Graph answered \`${status} UnknownError\`, its reply for a channel id it cannot resolve. `));
};

export { channelScopeOf, rewriteChannelScopedError };
export type { ChannelScope };
