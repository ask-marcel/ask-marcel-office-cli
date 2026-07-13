import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './create-forward-draft.ts';

describe('create-forward-draft', () => {
  it('creates a forward draft, then patches the comment, Text body, and parsed recipients into it', async () => {
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
    expect(posts).toEqual([{ path: '/me/messages/msg-1/createForward', body: {} }]);
    expect(patches).toEqual([
      {
        path: '/me/messages/draft-9',
        body: {
          body: { contentType: 'Text', content: 'Bob owns this, forwarding.' },
          toRecipients: [{ emailAddress: { address: 'bob@example.com' } }, { emailAddress: { address: 'carol@example.com' } }],
        },
      },
    ]);
  });

  it('adds cc recipients and a subject override to the patch only when given', async () => {
    let patchedBody: Record<string, unknown> = {};
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      patch: async (_path, body) => {
        patchedBody = body as Record<string, unknown>;
        return ok({ id: 'draft-9' });
      },
    });

    await execute(graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', ccRecipients: 'carol@example.com', bodyContent: 'x', subject: 'FW: reassigned' });
    expect(patchedBody.ccRecipients).toEqual([{ emailAddress: { address: 'carol@example.com' } }]);
    expect(patchedBody.subject).toBe('FW: reassigned');

    await execute(graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', bodyContent: 'x' });
    // Key-absence, not merely `=== undefined`: an `if (true)` regression would
    // set the key to undefined, which still passes `toBeUndefined()`.
    expect('ccRecipients' in patchedBody).toBe(false);
    expect('subject' in patchedBody).toBe(false);
  });

  it('patches an HTML comment body with the HTML content type', async () => {
    let patchedBody: Record<string, unknown> = {};
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      patch: async (_path, body) => {
        patchedBody = body as Record<string, unknown>;
        return ok({ id: 'draft-9' });
      },
    });

    await execute(graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', bodyContent: '<p>Forwarding.</p>', bodyContentType: 'HTML' });
    expect(patchedBody.body).toEqual({ contentType: 'HTML', content: '<p>Forwarding.</p>' });
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

  it('passes a PATCH failure through untouched (no swallow)', async () => {
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      patch: async () => err({ type: 'api_error', status: 400, message: 'ErrorInvalidRecipients' }),
    });

    const result = await execute(graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', bodyContent: 'x' });

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

  it('returns a validation_error for an invalid body-content-type', async () => {
    const result = await execute(fakeGraphClient(), { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', bodyContent: 'x', bodyContentType: 'Markdown' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});
