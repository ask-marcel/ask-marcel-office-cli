import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { execute } from './create-forward-draft.ts';

// The shape Graph returns from createForward, in the parts the splice must
// respect: head styles that render the quoted original, the empty div Graph
// parks an empty comment in, and the forwarded message itself.
const HEAD = '<html><head><style>p {margin-top:0;margin-bottom:0}</style></head>';
const BODY_OPEN = '<body dir="ltr">';
const EMPTY_COMMENT_DIV = '<div class="elementToProof"></div>';
const QUOTE_TAIL = '<div id="appendonsend"></div><div id="divRplyFwdMsg"><b>From:</b> Robin Chen<br><b>Sent:</b> Monday, May 5, 2026</div><div>the original</div></body></html>';
const DRAFT_BODY = `${HEAD}${BODY_OPEN}${EMPTY_COMMENT_DIV}${QUOTE_TAIL}`;

type RecordedCall = { path: string; body: unknown };
type Recording = { graph: GraphClient; posts: RecordedCall[]; patches: RecordedCall[]; gets: string[] };

describe('create-forward-draft', () => {
  it('puts the comment and recipients in the createForward POST so the quoted original survives, and patches nothing more', async () => {
    const posts: Array<{ path: string; body: unknown }> = [];
    const patches: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      post: async (path, body) => {
        posts.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true, subject: 'FW: Contoso migration' });
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

  // Records every call, so the HTML tests below can assert the NEGATIVE (no
  // second read, no PATCH, no draft created at all) as readily as the positive.
  const recordingGraph = (created: unknown, read?: Awaited<ReturnType<GraphClient['get']>>): Recording => {
    const posts: RecordedCall[] = [];
    const patches: RecordedCall[] = [];
    const gets: string[] = [];
    const graph = fakeGraphClient({
      post: async (path, body) => {
        posts.push({ path, body });
        return ok(created);
      },
      get: async (path) => {
        gets.push(path);
        return read ?? ok({});
      },
      patch: async (path, body) => {
        patches.push({ path, body });
        return ok({ id: 'draft-9', isDraft: true });
      },
    });
    return { graph, posts, patches, gets };
  };

  const htmlDraft = { id: 'draft-9', isDraft: true, body: { contentType: 'html', content: DRAFT_BODY } };
  const forwardAsHtml = { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', bodyContent: '<p>hi</p>', bodyContentType: 'HTML' };

  it('splices an HTML comment above the forwarded original in one patch, keeping the head styles and the whole quoted original', async () => {
    const { graph, posts, patches, gets } = recordingGraph(htmlDraft);

    const result = await execute(graph, { ...forwardAsHtml, bodyContent: '<p><b>Bob</b> owns this now.</p>' });

    expect(result.ok).toBe(true);
    // The recipients still ride in the create; only the comment is emptied, so the
    // markup can be spliced in below rather than HTML-escaped by Graph.
    expect(posts).toEqual([{ path: '/me/messages/msg-1/createForward', body: { comment: '', toRecipients: [{ emailAddress: { address: 'bob@example.com' } }] } }]);
    expect(gets).toEqual([]);
    expect(patches).toEqual([
      { path: '/me/messages/draft-9', body: { body: { contentType: 'HTML', content: `${HEAD}${BODY_OPEN}<p><b>Bob</b> owns this now.</p>${EMPTY_COMMENT_DIV}${QUOTE_TAIL}` } } },
    ]);
  });

  it('carries the body, the cc list, and the subject override in a single patch when forwarding as HTML', async () => {
    const { graph, patches } = recordingGraph(htmlDraft);

    await execute(graph, { ...forwardAsHtml, ccRecipients: 'carol@example.com', subject: 'FW: Q3 planning' });

    expect(patches).toEqual([
      {
        path: '/me/messages/draft-9',
        body: {
          body: { contentType: 'HTML', content: `${HEAD}${BODY_OPEN}<p>hi</p>${EMPTY_COMMENT_DIV}${QUOTE_TAIL}` },
          ccRecipients: [{ emailAddress: { address: 'carol@example.com' } }],
          subject: 'FW: Q3 planning',
        },
      },
    ]);
  });

  it('still sends a body-free patch for cc on the text path, so the comment and quoted original Graph wrote are never rewritten', async () => {
    const { graph, patches } = recordingGraph(htmlDraft);

    await execute(graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', ccRecipients: 'carol@example.com', bodyContent: 'over to you' });

    // `toEqual` pins the ABSENCE of a body key: today's behaviour, unchanged.
    expect(patches).toEqual([{ path: '/me/messages/draft-9', body: { ccRecipients: [{ emailAddress: { address: 'carol@example.com' } }] } }]);
  });

  it('reads the draft body back when the create response omits it, asking only for the body', async () => {
    const { graph, patches, gets } = recordingGraph({ id: 'draft-9', isDraft: true }, ok({ body: { contentType: 'html', content: DRAFT_BODY } }));

    const result = await execute(graph, forwardAsHtml);

    expect(result.ok).toBe(true);
    expect(gets).toEqual(['/me/messages/draft-9?$select=body']);
    expect(patches).toEqual([{ path: '/me/messages/draft-9', body: { body: { contentType: 'HTML', content: `${HEAD}${BODY_OPEN}<p>hi</p>${EMPTY_COMMENT_DIV}${QUOTE_TAIL}` } } }]);
  });

  it('passes a failed draft-body read through untouched, never patching a body it could not see', async () => {
    const { graph, patches } = recordingGraph({ id: 'draft-9', isDraft: true }, err({ type: 'api_error', status: 404, message: 'ItemNotFound' }));

    const result = await execute(graph, forwardAsHtml);

    expect(result).toEqual(err({ type: 'api_error', status: 404, message: 'ItemNotFound' }));
    expect(patches).toEqual([]);
  });

  it('names the draft it created when the body read back is unusable, so the caller can still find what it left behind', async () => {
    const { graph, patches } = recordingGraph({ id: 'draft-9', isDraft: true }, ok({ body: { contentType: 'html' } }));

    const result = await execute(graph, forwardAsHtml);

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('draft-9');
    expect(patches).toEqual([]);
  });

  it('refuses to splice HTML into a draft Graph minted as plain text, naming the draft it has already created', async () => {
    const { graph, patches } = recordingGraph({ id: 'draft-9', isDraft: true, body: { contentType: 'text', content: 'plain\n\nFrom: Robin Chen\nSent: Monday' } });

    const result = await execute(graph, forwardAsHtml);

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('draft-9');
    expect(patches).toEqual([]);
  });

  it('places the comment just inside the body tag when Graph mints a forward that quotes nothing, keeping the head', async () => {
    const noQuote = `${HEAD}${BODY_OPEN}<div>nothing quoted</div></body></html>`;
    const { graph, patches } = recordingGraph({ id: 'draft-9', isDraft: true, body: { contentType: 'html', content: noQuote } });

    await execute(graph, forwardAsHtml);

    // Nothing is dropped for want of a boundary: the comment goes to the top of
    // the body and the rest of the document survives.
    expect(patches).toEqual([
      { path: '/me/messages/draft-9', body: { body: { contentType: 'HTML', content: `${HEAD}${BODY_OPEN}<p>hi</p><div>nothing quoted</div></body></html>` } } },
    ]);
  });

  it('refuses a comment that itself carries a quote boundary marker, before any draft is created', async () => {
    const { graph, posts } = recordingGraph(htmlDraft);

    const result = await execute(graph, { ...forwardAsHtml, bodyContent: '<div class="gmail_quote">pasted history</div>' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    // Rejected before the create, so there is no orphan draft to clean up.
    expect(posts).toEqual([]);
  });

  it('sends an explicit Text comment exactly as it sends one with no content type at all', async () => {
    const explicit = recordingGraph(htmlDraft);
    const implicit = recordingGraph(htmlDraft);

    await execute(explicit.graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', bodyContent: 'plain', bodyContentType: 'Text' });
    await execute(implicit.graph, { forwardMessageId: 'msg-1', toRecipients: 'bob@example.com', bodyContent: 'plain' });

    expect(explicit.posts).toEqual(implicit.posts);
    expect(explicit.patches).toEqual(implicit.patches);
    // And that shared path is still today's: the comment rides in the create, and
    // no PATCH touches the body.
    expect(implicit.posts).toEqual([{ path: '/me/messages/msg-1/createForward', body: { comment: 'plain', toRecipients: [{ emailAddress: { address: 'bob@example.com' } }] } }]);
    expect(implicit.patches).toEqual([]);
  });
});
