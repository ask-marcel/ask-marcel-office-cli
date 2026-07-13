import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './create-forward-draft.ts';

describe('create-forward-draft', () => {
  it('puts the comment and recipients in the createForward POST so the quoted original survives, and patches nothing more', async () => {
    const posts: Array<{ path: string; body: unknown }> = [];
    const patches: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      post: async (path, body) => {
        posts.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true, subject: 'FW: TEMPO PATH' });
      },
      patch: async (path, body) => {
        patches.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true });
      },
    });

    const result = await execute(graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com, carol@example.com', bodyContent: 'Bob owns this, forwarding.' });

    expect(result.ok).toBe(true);
    // The comment travels via `comment` (Graph places it above the quote); a body
    // PATCH would replace the draft body and drop the forwarded original.
    expect(posts).toEqual([
      {
        path: '/me/messages/msg-1/createForward',
        body: { comment: 'Bob owns this, forwarding.', toRecipients: [{ emailAddress: { address: 'bob@example.com' } }, { emailAddress: { address: 'carol@example.com' } }] },
      },
    ]);
    expect(patches).toEqual([]);
  });

  it('sets cc recipients and a subject override via a body-free PATCH, and skips the PATCH entirely when neither is given', async () => {
    const patches: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      patch: async (path, body) => {
        patches.push({ path, body });
        return ok({ id: 'draft-9' });
      },
    });

    await execute(graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', ccRecipients: 'carol@example.com', bodyContent: 'x', subject: 'FW: reassigned' });
    // Body-free: the PATCH carries ONLY cc/subject. A `body` key would clobber the
    // createForward comment + quoted original, so `toEqual` pins its absence.
    expect(patches).toEqual([{ path: '/me/messages/draft-9', body: { ccRecipients: [{ emailAddress: { address: 'carol@example.com' } }], subject: 'FW: reassigned' } }]);

    patches.length = 0;
    await execute(graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', bodyContent: 'x' });
    expect(patches).toEqual([]);
  });

  it('refuses any createForward response that is not a well-formed unsent draft - and never patches it', async () => {
    // One malformed shape per conjunct of the isUnsentDraft guard, plus a sent
    // message: non-object, null, missing id, non-string id, isDraft-absent, and
    // isDraft:false must each be rejected before any PATCH.
    const malformed: unknown[] = ['not-an-object', null, { isDraft: true }, { id: 42, isDraft: true }, { id: 'draft-9' }, { id: 'sent-1', isDraft: false }];
    for (const value of malformed) {
      let patchCalls = 0;
      const graph = fakeGraphClient({
        post: async () => ok(value),
        patch: async () => {
          patchCalls += 1;
          return ok({});
        },
      });

      const result = await execute(graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', bodyContent: 'x' });

      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('api_error');
      if (!result.ok && result.error.type === 'api_error') expect(result.error.code).toBe('not_an_unsent_draft');
      expect(patchCalls).toBe(0);
    }
  });

  it('passes a createForward failure through untouched, without patching', async () => {
    let patchCalls = 0;
    const graph = fakeGraphClient({
      post: async () => err({ type: 'api_error', status: 502, message: 'InvalidForward' }),
      patch: async () => {
        patchCalls += 1;
        return ok({});
      },
    });

    const result = await execute(graph, { forwardMessageId: 'msg-gone', toRecipients: 'bob@example.com', bodyContent: 'x' });

    expect(result).toEqual(err({ type: 'api_error', status: 502, message: 'InvalidForward' }));
    expect(patchCalls).toBe(0);
  });

  it('passes a cc/subject PATCH failure through untouched (no swallow)', async () => {
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      patch: async () => err({ type: 'api_error', status: 400, message: 'ErrorInvalidRecipients' }),
    });

    // A cc override triggers the body-free PATCH; its failure must propagate.
    const result = await execute(graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', ccRecipients: 'bad-addr', bodyContent: 'x' });

    expect(result).toEqual(err({ type: 'api_error', status: 400, message: 'ErrorInvalidRecipients' }));
  });

  it('returns a validation_error when forward-message-id is missing', async () => {
    const result = await execute(fakeGraphClient(), { toRecipients: 'bob@example.com', bodyContent: 'x' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('returns a validation_error when to-recipients is missing (a forward with no recipient is not actionable)', async () => {
    const result = await execute(fakeGraphClient(), { forwardMessageId: 'msg-1', bodyContent: 'x' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('returns a validation_error when body-content is missing', async () => {
    const result = await execute(fakeGraphClient(), { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});
