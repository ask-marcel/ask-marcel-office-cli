import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './create-reply-draft.ts';

describe('create-reply-draft', () => {
  it('creates a reply-all draft threaded on the message, then patches the reply body in above the quoted history', async () => {
    const posts: Array<{ path: string; body: unknown }> = [];
    const patches: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      post: async (path, body) => {
        posts.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true, subject: 'RE: TEMPO PATH // Transition' });
      },
      patch: async (path, body) => {
        patches.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true });
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'Confirmed for Concur.' });

    expect(result.ok).toBe(true);
    expect(posts).toEqual([{ path: '/me/messages/msg-1/createReplyAll', body: {} }]);
    expect(patches).toEqual([{ path: '/me/messages/draft-9', body: { body: { contentType: 'Text', content: 'Confirmed for Concur.' } } }]);
  });

  it('overrides the inherited RE: subject only when asked', async () => {
    let patchedBody: Record<string, unknown> = {};
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      patch: async (_path, body) => {
        patchedBody = body as Record<string, unknown>;
        return ok({ id: 'draft-9' });
      },
    });

    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'x', subject: 'RE: TEMPO PATH - Concur confirmed' });
    expect(patchedBody.subject).toBe('RE: TEMPO PATH - Concur confirmed');

    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'x' });
    expect(patchedBody.subject).toBeUndefined();
  });

  it('patches an HTML body with the HTML content type', async () => {
    let patchedBody: Record<string, unknown> = {};
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      patch: async (_path, body) => {
        patchedBody = body as Record<string, unknown>;
        return ok({ id: 'draft-9' });
      },
    });

    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: '<p>Confirmed.</p>', bodyContentType: 'HTML' });

    expect(patchedBody.body).toEqual({ contentType: 'HTML', content: '<p>Confirmed.</p>' });
  });

  it('refuses a createReplyAll response that is not an unsent draft - and never patches it', async () => {
    let patchCalls = 0;
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'sent-1', isDraft: false }),
      patch: async () => {
        patchCalls += 1;
        return ok({});
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'x' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('api_error');
    expect(patchCalls).toBe(0);
  });

  it('passes a createReplyAll failure through untouched, without patching', async () => {
    let patchCalls = 0;
    const graph = fakeGraphClient({
      post: async () => err({ type: 'api_error', status: 502, message: 'InvalidReplyAll' }),
      patch: async () => {
        patchCalls += 1;
        return ok({});
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-gone', bodyContent: 'x' });

    expect(result).toEqual(err({ type: 'api_error', status: 502, message: 'InvalidReplyAll' }));
    expect(patchCalls).toBe(0);
  });

  it('returns a validation_error when reply-to-message-id is missing', async () => {
    const result = await execute(fakeGraphClient(), { bodyContent: 'x' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});
