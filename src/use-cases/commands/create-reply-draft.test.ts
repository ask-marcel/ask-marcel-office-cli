import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './create-reply-draft.ts';

describe('create-reply-draft', () => {
  it('puts the reply text in the createReplyAll POST so the quoted thread survives, and patches nothing more', async () => {
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
    // The reply text travels via `comment` (Graph places it above the quote); a
    // body PATCH would replace the draft body and drop the quoted thread.
    expect(posts).toEqual([{ path: '/me/messages/msg-1/createReplyAll', body: { comment: 'Confirmed for Concur.' } }]);
    expect(patches).toEqual([]);
  });

  it('overrides the RE: subject via a body-free PATCH, and skips the PATCH when no override is given', async () => {
    const patches: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      patch: async (path, body) => {
        patches.push({ path, body });
        return ok({ id: 'draft-9' });
      },
    });

    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'x', subject: 'RE: TEMPO PATH - Concur confirmed' });
    // Body-free: the PATCH carries ONLY the subject. A `body` key would clobber the
    // reply text + quoted thread, so `toEqual` pins its absence.
    expect(patches).toEqual([{ path: '/me/messages/draft-9', body: { subject: 'RE: TEMPO PATH - Concur confirmed' } }]);

    patches.length = 0;
    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'x' });
    expect(patches).toEqual([]);
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

  it('refuses any createReplyAll response that is not a well-formed unsent draft - and never patches it', async () => {
    // One malformed shape per conjunct of the isUnsentDraft guard: non-object,
    // null, missing id, non-string id, and isDraft-absent must each be rejected.
    const malformed: unknown[] = ['not-an-object', null, { isDraft: true }, { id: 42, isDraft: true }, { id: 'draft-9' }];
    for (const value of malformed) {
      let patchCalls = 0;
      const graph = fakeGraphClient({
        post: async () => ok(value),
        patch: async () => {
          patchCalls += 1;
          return ok({});
        },
      });

      const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'x' });

      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('api_error');
      if (!result.ok && result.error.type === 'api_error') expect(result.error.code).toBe('not_an_unsent_draft');
      expect(patchCalls).toBe(0);
    }
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

  it('replies to the sender alone when reply-all is turned off, and to everyone when it is on or left unsaid', async () => {
    const posts: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      post: async (path, body) => {
        posts.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true });
      },
    });

    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'Confirmed.', replyAll: 'false' });
    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'Confirmed.', replyAll: 'true' });
    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'Confirmed.' });

    // Only the action differs; the comment payload is identical on both paths,
    // and anything other than an explicit `false` still replies to everyone.
    expect(posts).toEqual([
      { path: '/me/messages/msg-1/createReply', body: { comment: 'Confirmed.' } },
      { path: '/me/messages/msg-1/createReplyAll', body: { comment: 'Confirmed.' } },
      { path: '/me/messages/msg-1/createReplyAll', body: { comment: 'Confirmed.' } },
    ]);
  });

  it('refuses a reply-all value that is neither true nor false, before any draft is created', async () => {
    let postCalls = 0;
    const graph = fakeGraphClient({
      post: async () => {
        postCalls += 1;
        return ok({ id: 'draft-9', isDraft: true });
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'x', replyAll: 'maybe' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(postCalls).toBe(0);
  });
});
