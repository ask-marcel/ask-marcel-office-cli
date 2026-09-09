import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { channelScopeOf, rewriteChannelScopedError } from './team-channel-errors.ts';

const api = (status: number, message: string): ReturnType<typeof err<GraphError>> => err({ type: 'api_error', status, message });
const CHANNEL = '19:bogus@thread.tacv2';

describe('naming the channel or message behind an opaque Graph error', () => {
  it('leaves a successful result alone', () => {
    const result = ok({ value: [] });
    expect(rewriteChannelScopedError(result, { channelId: CHANNEL })).toBe(result);
  });

  it('leaves a non-API error alone', () => {
    const result = err<GraphError>({ type: 'validation_error', message: 'bad flag' });
    expect(rewriteChannelScopedError(result, { channelId: CHANNEL })).toBe(result);
  });

  it('rewrites the bare `1: NotFound` that Graph answers for an unknown channel id', () => {
    const rewritten = rewriteChannelScopedError(api(404, '1: NotFound'), { channelId: CHANNEL });
    expect(rewritten.ok).toBe(false);
    if (rewritten.ok || rewritten.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(rewritten.error.status).toBe(404);
    expect(rewritten.error.code).toBe('cli_rewrite_channel_not_found');
    expect(rewritten.error.message).toBe(
      `NotFound: Microsoft Teams channel not found (channel-id: "${CHANNEL}"). Verify it exists in this team via \`ask-marcel-office list-team-channels --team-id <team-id>\`.`
    );
  });

  it('does not rewrite a `1: NotFound` that is not anchored at the start of the message', () => {
    const result = api(404, '11: NotFound');
    expect(rewriteChannelScopedError(result, { channelId: CHANNEL })).toBe(result);
  });

  it('rewrites the Skype backend `GetThreadRequest` failure the tabs route answers for an unknown channel', () => {
    const rewritten = rewriteChannelScopedError(api(404, 'ItemNotFound: Failed to execute Skype backend request GetThreadRequest.'), { channelId: CHANNEL });
    if (rewritten.ok || rewritten.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(rewritten.error.code).toBe('cli_rewrite_channel_not_found');
    expect(rewritten.error.message).toContain(`channel-id: "${CHANNEL}"`);
    expect(rewritten.error.message).toContain('list-team-channels');
  });

  it('names the channel behind the `410 Gone: UnknownError` the messages route answers for an unknown channel', () => {
    const rewritten = rewriteChannelScopedError(api(410, 'Gone: UnknownError'), { channelId: CHANNEL });
    if (rewritten.ok || rewritten.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(rewritten.error.status).toBe(410);
    expect(rewritten.error.code).toBe('cli_rewrite_channel_not_found');
    expect(rewritten.error.message).toContain(`channel-id: "${CHANNEL}"`);
    expect(rewritten.error.message).toContain('410');
  });

  it('names the message behind the `403 Forbidden: UnknownError` an unknown message id answers, so it does not read as a scope failure', () => {
    const rewritten = rewriteChannelScopedError(api(403, 'Forbidden: UnknownError'), { channelId: CHANNEL, messageId: '1000000000000' });
    if (rewritten.ok || rewritten.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(rewritten.error.status).toBe(403);
    expect(rewritten.error.code).toBe('cli_rewrite_channel_message_not_found');
    expect(rewritten.error.message).toContain('message-id: "1000000000000"');
    expect(rewritten.error.message).toContain('not a missing scope');
    expect(rewritten.error.message).toContain('list-team-channel-messages');
  });

  it('calls a `400 BadRequest: UnknownError` a malformed id, naming both ids when a message id is in play', () => {
    const rewritten = rewriteChannelScopedError(api(400, 'BadRequest: UnknownError'), { channelId: 'not-a-channel', messageId: 'abc' });
    if (rewritten.ok || rewritten.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(rewritten.error.status).toBe(400);
    expect(rewritten.error.code).toBe('cli_rewrite_channel_bad_id');
    expect(rewritten.error.message).toContain('channel-id: "not-a-channel"');
    expect(rewritten.error.message).toContain('message-id: "abc"');
  });

  it('a `403 UnknownError` with no message id in play still points at the channel, never at scopes', () => {
    const rewritten = rewriteChannelScopedError(api(403, 'Forbidden: UnknownError'), { channelId: CHANNEL });
    if (rewritten.ok || rewritten.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(rewritten.error.code).toBe('cli_rewrite_channel_not_found');
    expect(rewritten.error.message).toContain(`channel-id: "${CHANNEL}"`);
  });

  it('passes every other API error through untouched', () => {
    const result = api(429, 'TooManyRequests: slow down');
    expect(rewriteChannelScopedError(result, { channelId: CHANNEL, messageId: 'm1' })).toBe(result);
  });

  it('leaves a non-API error alone even when its text looks like the opaque channel answer', () => {
    const result = err<GraphError>({ type: 'validation_error', message: '1: NotFound' });
    expect(rewriteChannelScopedError(result, { channelId: CHANNEL })).toBe(result);
  });

  it('accepts the `1:NotFound` spelling with no space, the thread-id prefix is all that anchors it', () => {
    const rewritten = rewriteChannelScopedError(api(404, '1:NotFound'), { channelId: CHANNEL });
    if (rewritten.ok || rewritten.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(rewritten.error.code).toBe('cli_rewrite_channel_not_found');
  });

  it('a `410 UnknownError` names the channel even when a message id is in play: 410 answers for the channel, 403 for the message', () => {
    const rewritten = rewriteChannelScopedError(api(410, 'Gone: UnknownError'), { channelId: CHANNEL, messageId: '1000000000000' });
    if (rewritten.ok || rewritten.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(rewritten.error.code).toBe('cli_rewrite_channel_not_found');
    expect(rewritten.error.message).not.toContain('message-id');
  });

  it('the malformed-id message teaches the channel id shape and where a message id comes from', () => {
    const rewritten = rewriteChannelScopedError(api(400, 'BadRequest: UnknownError'), { channelId: 'x' });
    if (rewritten.ok || rewritten.error.type !== 'api_error') throw new Error('expected an api_error');
    expect(rewritten.error.message).toContain('channel-id: "x"');
    expect(rewritten.error.message).not.toContain('message-id');
    expect(rewritten.error.message).toContain('19:<thread>@thread.tacv2');
    expect(rewritten.error.message).toContain('numeric id from `ask-marcel-office list-team-channel-messages`');
  });

  it('reads the channel and message ids out of raw params, and marks a missing channel id as unknown', () => {
    expect(channelScopeOf({ channelId: CHANNEL, messageId: 'm1' })).toEqual({ channelId: CHANNEL, messageId: 'm1' });
    expect(channelScopeOf({ teamId: 'tm1' })).toEqual({ channelId: '<unknown>', messageId: undefined });
  });
});
