import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './create-reply-draft.ts';

// The shape Graph returns from createReplyAll, in the parts the splice must
// respect: head styles that render the quoted tail, the empty div Graph parks
// an empty comment in, and the quoted thread itself.
const HEAD = '<html><head><style>p {margin-top:0;margin-bottom:0}</style></head>';
const BODY_OPEN = '<body dir="ltr">';
const EMPTY_COMMENT_DIV = '<div class="elementToProof"></div>';
const QUOTE_TAIL = '<div id="appendonsend"></div><div id="divRplyFwdMsg"><b>From:</b> Robin Chen<br><b>Sent:</b> Monday, May 5, 2026</div><div>the original</div></body></html>';
const DRAFT_BODY = `${HEAD}${BODY_OPEN}${EMPTY_COMMENT_DIV}${QUOTE_TAIL}`;

describe('create-reply-draft', () => {
  it('puts the reply text in the createReplyAll POST so the quoted thread survives, and patches nothing more', async () => {
    const posts: Array<{ path: string; body: unknown }> = [];
    const patches: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      post: async (path, body) => {
        posts.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true, subject: 'RE: Contoso migration' });
      },
      patch: async (path, body) => {
        patches.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true });
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'Confirmed for Contoso.' });

    expect(result.ok).toBe(true);
    // The reply text travels via `comment` (Graph places it above the quote); a
    // body PATCH would replace the draft body and drop the quoted thread.
    expect(posts).toEqual([{ path: '/me/messages/msg-1/createReplyAll', body: { comment: 'Confirmed for Contoso.' } }]);
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

    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'x', subject: 'RE: Contoso migration - scope confirmed' });
    // Body-free: the PATCH carries ONLY the subject. A `body` key would clobber the
    // reply text + quoted thread, so `toEqual` pins its absence.
    expect(patches).toEqual([{ path: '/me/messages/draft-9', body: { subject: 'RE: Contoso migration - scope confirmed' } }]);

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

  it('splices an HTML reply above the quoted history in one patch, keeping the head styles and the whole quoted thread', async () => {
    const posts: Array<{ path: string; body: unknown }> = [];
    const patches: Array<{ path: string; body: unknown }> = [];
    const gets: string[] = [];
    const graph = fakeGraphClient({
      post: async (path, body) => {
        posts.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true, body: { contentType: 'html', content: DRAFT_BODY } });
      },
      get: async (path) => {
        gets.push(path);
        return ok({});
      },
      patch: async (path, body) => {
        patches.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true });
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: '<p><b>bold</b> &amp; entity</p>', bodyContentType: 'HTML' });

    expect(result.ok).toBe(true);
    // The create carries an EMPTY comment: Graph mints the scaffolding, we splice
    // the reply in. Sending it as `comment` would HTML-escape the markup.
    expect(posts).toEqual([{ path: '/me/messages/msg-1/createReplyAll', body: { comment: '' } }]);
    // The body rode back on the create response, so there is no second read.
    expect(gets).toEqual([]);
    // Exactly one PATCH, and the quote is INSIDE the body it carries.
    expect(patches).toEqual([
      { path: '/me/messages/draft-9', body: { body: { contentType: 'HTML', content: `${HEAD}${BODY_OPEN}${EMPTY_COMMENT_DIV}<p><b>bold</b> &amp; entity</p>${QUOTE_TAIL}` } } },
    ]);
  });

  it('reads the draft body back when the create response omits it, asking only for the body', async () => {
    const gets: string[] = [];
    const patches: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      get: async (path) => {
        gets.push(path);
        return ok({ body: { contentType: 'html', content: DRAFT_BODY } });
      },
      patch: async (path, body) => {
        patches.push({ path, body });
        return ok({});
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: '<p>hi</p>', bodyContentType: 'HTML' });

    expect(result.ok).toBe(true);
    expect(gets).toEqual(['/me/messages/draft-9?$select=body']);
    expect(patches).toEqual([{ path: '/me/messages/draft-9', body: { body: { contentType: 'HTML', content: `${HEAD}${BODY_OPEN}${EMPTY_COMMENT_DIV}<p>hi</p>${QUOTE_TAIL}` } } }]);
  });

  it('passes a failed draft-body read through untouched, never patching a body it could not see', async () => {
    let patchCalls = 0;
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      get: async () => err({ type: 'api_error', status: 404, message: 'ItemNotFound' }),
      patch: async () => {
        patchCalls += 1;
        return ok({});
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: '<p>hi</p>', bodyContentType: 'HTML' });

    expect(result).toEqual(err({ type: 'api_error', status: 404, message: 'ItemNotFound' }));
    expect(patchCalls).toBe(0);
  });

  it('names the draft it created when the body read back is unusable, so the caller can still find what it left behind', async () => {
    let patchCalls = 0;
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true }),
      // A 200 carrying a shape the body splice cannot work with.
      get: async () => ok({ body: { contentType: 'html' } }),
      patch: async () => {
        patchCalls += 1;
        return ok({});
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: '<p>hi</p>', bodyContentType: 'HTML' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('draft-9');
    expect(patchCalls).toBe(0);
  });

  it('refuses to splice HTML into a draft Graph minted as plain text, naming the draft it has already created', async () => {
    let patchCalls = 0;
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true, body: { contentType: 'text', content: 'plain reply\n\nFrom: Robin Chen\nSent: Monday' } }),
      patch: async () => {
        patchCalls += 1;
        return ok({});
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: '<p>hi</p>', bodyContentType: 'HTML' });

    expect(result.ok).toBe(false);
    // The draft EXISTS by the time this fails, so the message must name it or the
    // caller is left with an orphan they cannot find.
    if (!result.ok) expect(result.error.message).toContain('draft-9');
    expect(patchCalls).toBe(0);
  });

  it('places the reply just inside the body tag when Graph mints a draft that quotes nothing, keeping the head', async () => {
    const patches: Array<{ path: string; body: unknown }> = [];
    const noQuote = `${HEAD}${BODY_OPEN}<div>nothing quoted</div></body></html>`;
    const graph = fakeGraphClient({
      post: async () => ok({ id: 'draft-9', isDraft: true, body: { contentType: 'html', content: noQuote } }),
      patch: async (path, body) => {
        patches.push({ path, body });
        return ok({});
      },
    });

    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: '<p>hi</p>', bodyContentType: 'HTML' });

    // Nothing is ever dropped for want of a boundary: the reply goes to the top
    // of the body and the rest of the document survives.
    expect(patches).toEqual([
      { path: '/me/messages/draft-9', body: { body: { contentType: 'HTML', content: `${HEAD}${BODY_OPEN}<p>hi</p><div>nothing quoted</div></body></html>` } } },
    ]);
  });

  it('refuses a reply that itself carries a quote boundary marker, before any draft is created', async () => {
    const posts: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      post: async (path, body) => {
        posts.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true, body: { contentType: 'html', content: DRAFT_BODY } });
      },
    });

    const result = await execute(graph, { replyToMessageId: 'msg-1', bodyContent: '<div class="gmail_quote">pasted history</div>', bodyContentType: 'HTML' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    // Rejected before the create, so there is no orphan draft to clean up: a
    // marker inside the reply would make the next --comment edit cut there.
    expect(posts).toEqual([]);
  });

  it('sends an explicit Text reply exactly as it sends one with no content type at all', async () => {
    const calls: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      post: async (path, body) => {
        calls.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true, body: { contentType: 'html', content: DRAFT_BODY } });
      },
      patch: async (path, body) => {
        calls.push({ path, body });
        return ok({});
      },
    });

    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'plain reply', bodyContentType: 'Text' });
    const explicitText = [...calls];
    calls.length = 0;
    await execute(graph, { replyToMessageId: 'msg-1', bodyContent: 'plain reply' });

    expect(explicitText).toEqual(calls);
    // And that shared path is still today's: the reply rides in the create, and
    // no PATCH touches the body.
    expect(calls).toEqual([{ path: '/me/messages/msg-1/createReplyAll', body: { comment: 'plain reply' } }]);
  });
});
