import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './update-mail-draft.ts';

// A threaded draft as Graph stores it: the author's reply, then the quote.
const BODY_OPEN = '<body dir="ltr">';
const QUOTE_TAIL = '<div id="appendonsend"></div><div id="divRplyFwdMsg"><b>From:</b> Robin Chen<br><b>Sent:</b> Monday</div><div>the original</div></body></html>';
const htmlDraft = (reply: string): string => `<html><head><style>p {margin:0}</style></head>${BODY_OPEN}${reply}${QUOTE_TAIL}`;
const TEXT_DRAFT = 'the first attempt\n\n_______________________________\nFrom: Robin Chen\nSent: Monday\n\nthe original';

type Recording = { graph: GraphClient; patches: Array<{ path: string; body: unknown }>; gets: string[] };

// Records the read and the write so each test can pin the GET it makes and
// assert the NEGATIVE (nothing was patched) when the command refuses.
const recordingGraph = (read: Awaited<ReturnType<GraphClient['get']>>): Recording => {
  const patches: Array<{ path: string; body: unknown }> = [];
  const gets: string[] = [];
  const graph = fakeGraphClient({
    get: async (path) => {
      gets.push(path);
      return read;
    },
    patch: async (path, body) => {
      patches.push({ path, body });
      return ok({ id: 'AAMk1' });
    },
  });
  return { graph, patches, gets };
};

describe('update-mail-draft', () => {
  it('PATCHes /me/messages/{id} with only the subject when that is the sole field changed', async () => {
    let capturedPath = '';
    let capturedBody: Record<string, unknown> = {};
    const graph = fakeGraphClient({
      patch: async (path, body) => {
        capturedPath = path;
        capturedBody = body as Record<string, unknown>;
        return ok({ id: 'msg-1' });
      },
    });
    const result = await execute(graph, { messageId: 'AAMk1', subject: 'Revised' });
    expect(result).toEqual(ok({ id: 'msg-1' }));
    expect(capturedPath).toBe('/me/messages/AAMk1');
    expect(capturedBody).toEqual({ subject: 'Revised' });
  });

  it('sends the HTML body, recipients, cc, bcc, and importance when every updatable field is supplied', async () => {
    let capturedBody: Record<string, unknown> = {};
    const graph = fakeGraphClient({
      // A whole-body replace reads the draft first, to refuse when it would
      // drop a quote. This one is quote-free, so the write goes through.
      get: async () => ok({ isDraft: true, body: { contentType: 'html', content: '<html><body>a plain draft</body></html>' } }),
      patch: async (_path, body) => {
        capturedBody = body as Record<string, unknown>;
        return ok({});
      },
    });
    const result = await execute(graph, {
      messageId: 'AAMk2',
      bodyContent: '<p>hi</p>',
      bodyContentType: 'HTML',
      toRecipients: 'alice@example.com',
      ccRecipients: 'carol@example.com',
      bccRecipients: 'dave@example.com',
      importance: 'Low',
    });
    expect(result.ok).toBe(true);
    expect(capturedBody.body).toEqual({ contentType: 'HTML', content: '<p>hi</p>' });
    expect(capturedBody.toRecipients).toEqual([{ emailAddress: { address: 'alice@example.com' } }]);
    expect(capturedBody.ccRecipients).toEqual([{ emailAddress: { address: 'carol@example.com' } }]);
    expect(capturedBody.bccRecipients).toEqual([{ emailAddress: { address: 'dave@example.com' } }]);
    expect(capturedBody.importance).toBe('Low');
  });

  it('returns a validation_error when the required message id is missing', async () => {
    const result = await execute(fakeGraphClient(), { subject: 'x' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('returns a validation_error when no updatable field is provided (id alone is not enough)', async () => {
    const result = await execute(fakeGraphClient(), { messageId: 'AAMk3' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });

  it('clears a recipient list when handed an empty string, which is the only way to drop the recipients a reply inherited', async () => {
    const patches: Array<{ path: string; body: unknown }> = [];
    const graph = fakeGraphClient({
      patch: async (path, body) => {
        patches.push({ path, body });
        return ok({});
      },
    });

    await execute(graph, { messageId: 'AAMk1', ccRecipients: '' });
    await execute(graph, { messageId: 'AAMk1', toRecipients: '' });
    await execute(graph, { messageId: 'AAMk1', bccRecipients: '' });

    // An empty list is Graph's clear payload. Each empty string also satisfies
    // the at-least-one guard ALONE: clearing a list is a change like any other.
    expect(patches).toEqual([
      { path: '/me/messages/AAMk1', body: { ccRecipients: [] } },
      { path: '/me/messages/AAMk1', body: { toRecipients: [] } },
      { path: '/me/messages/AAMk1', body: { bccRecipients: [] } },
    ]);
  });

  it('tells a caller who passed nothing to change how to clear a list, since an empty string is no longer nothing', async () => {
    const result = await execute(fakeGraphClient(), { messageId: 'AAMk3' });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.message).toBe(
        'At least one field must be provided to update (--subject, --body-content, --comment, --to-recipients, --cc-recipients, --bcc-recipients, or --importance). Pass an empty string to a recipient flag to clear that list.'
      );
    }
  });

  it('rewrites the reply text above the quote and leaves the quoted thread untouched, escaping what the author typed', async () => {
    const { graph, patches, gets } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: htmlDraft('<div>the first attempt</div>') } }));

    const result = await execute(graph, { messageId: 'AAMk1', comment: 'Tom & <b>Jerry</b>\nsecond line' });

    expect(result.ok).toBe(true);
    expect(gets).toEqual(['/me/messages/AAMk1?$select=body,isDraft']);
    // The old reply is replaced, not appended to, and the author's angle brackets
    // show as characters rather than taking effect.
    expect(patches).toEqual([
      {
        path: '/me/messages/AAMk1',
        body: { body: { contentType: 'html', content: htmlDraft('Tom &amp; &lt;b&gt;Jerry&lt;/b&gt;<br>second line') } },
      },
    ]);
  });

  it('keeps the author’s markup as markup when they say it is HTML', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: htmlDraft('<div>old</div>') } }));

    await execute(graph, { messageId: 'AAMk1', comment: '<p><b>bold</b></p>', bodyContentType: 'HTML' });

    expect(patches).toEqual([{ path: '/me/messages/AAMk1', body: { body: { contentType: 'html', content: htmlDraft('<p><b>bold</b></p>') } } }]);
  });

  it('passes the draft’s own content type back verbatim rather than forcing a spelling of its own', async () => {
    // Graph answers `html` lowercase; echoing a hardcoded `HTML` would be a
    // gratuitous change to a field the caller never asked to touch.
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: htmlDraft('<div>old</div>') } }));

    await execute(graph, { messageId: 'AAMk1', comment: 'revised' });

    expect(patches[0]?.body).toEqual({ body: { contentType: 'html', content: htmlDraft('revised') } });
  });

  it('revises a plain-text draft above its quote, separating the reply from it by a blank line', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'text', content: TEXT_DRAFT } }));

    await execute(graph, { messageId: 'AAMk1', comment: 'the revised reply' });

    expect(patches).toEqual([
      {
        path: '/me/messages/AAMk1',
        body: { body: { contentType: 'text', content: 'the revised reply\n\n_______________________________\nFrom: Robin Chen\nSent: Monday\n\nthe original' } },
      },
    ]);
  });

  it('refuses HTML on a plain-text draft, where markup would show as literal characters', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'text', content: TEXT_DRAFT } }));

    const result = await execute(graph, { messageId: 'AAMk1', comment: '<p>hi</p>', bodyContentType: 'HTML' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(patches).toEqual([]);
  });

  it('refuses to revise a draft that has no quoted history, pointing at the flag that replaces a whole body', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: '<html><body>a plain draft</body></html>' } }));

    const result = await execute(graph, { messageId: 'AAMk1', comment: 'revised' });

    expect(result.ok).toBe(false);
    // Revising a body with no quote to preserve would mean overwriting all of it,
    // which is --body-content's job and must be asked for explicitly.
    if (!result.ok) {
      expect(result.error.message).toBe(
        'Draft AAMk1 has no quoted reply history to preserve, so there is nothing for --comment to sit above. Use --body-content to replace the whole body instead.'
      );
    }
    expect(patches).toEqual([]);
  });

  it('refuses to revise a plain-text draft that has no quoted history, exactly as it refuses an HTML one', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'text', content: 'a plain draft, nothing quoted' } }));

    const result = await execute(graph, { messageId: 'AAMk1', comment: 'revised' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('no quoted reply history');
    expect(patches).toEqual([]);
  });

  it('escapes marker-shaped text in a plain-text comment rather than refusing it, since escaping already makes it inert', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: htmlDraft('<div>old</div>') } }));

    // The boundary-marker refusal is an HTML-mode concern: in Text mode the same
    // characters are escaped into `&lt;div ...`, which no boundary scan matches.
    const result = await execute(graph, { messageId: 'AAMk1', comment: '<div class="gmail_quote">quoting a colleague</div>' });

    expect(result.ok).toBe(true);
    expect(patches).toEqual([
      { path: '/me/messages/AAMk1', body: { body: { contentType: 'html', content: htmlDraft('&lt;div class=&quot;gmail_quote&quot;&gt;quoting a colleague&lt;/div&gt;') } } },
    ]);
  });

  it('refuses to revise a message that is not a draft, before writing to it', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: false, body: { contentType: 'html', content: htmlDraft('<div>sent</div>') } }));

    const result = await execute(graph, { messageId: 'AAMk1', comment: 'revised' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(patches).toEqual([]);
  });

  // 2026-07-23 bug report: `--body-content` means "the text above the quote" on
  // create-reply-draft and "replace the whole body, quote included" here. A
  // caller who learned the flag on the create command and reused it here wiped
  // the quoted history, with nothing warning at call time.
  it('refuses to replace the whole body of a threaded draft, pointing at the flag that keeps the quote', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: htmlDraft('<div>the reply</div>') } }));

    const result = await execute(graph, { messageId: 'AAMk1', bodyContent: '<p>whole new body</p>', bodyContentType: 'HTML' });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.message).toBe(
        'Draft AAMk1 carries quoted reply history, and --body-content replaces the entire body, quote included. Revise only your own text with --comment, which keeps the quote byte-identical, or pass --replace-quoted-history true to drop the quote deliberately.'
      );
    }
    expect(patches).toEqual([]);
  });

  it('refuses to replace the whole body of a threaded plain-text draft, exactly as it refuses an HTML one', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'text', content: TEXT_DRAFT } }));

    const result = await execute(graph, { messageId: 'AAMk1', bodyContent: 'whole new body' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('carries quoted reply history');
    expect(patches).toEqual([]);
  });

  it('replaces the whole body of a threaded draft when the caller asks for it explicitly, without reading the draft first', async () => {
    const { graph, patches, gets } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: htmlDraft('<div>the reply</div>') } }));

    const result = await execute(graph, { messageId: 'AAMk1', bodyContent: '<p>whole new body</p>', bodyContentType: 'HTML', replaceQuotedHistory: 'true' });

    expect(result.ok).toBe(true);
    expect(patches).toEqual([{ path: '/me/messages/AAMk1', body: { body: { contentType: 'HTML', content: '<p>whole new body</p>' } } }]);
    // The escape is the caller stating intent, so the guard's read is skipped.
    expect(gets).toEqual([]);
  });

  it('replaces the whole body of a quote-free draft with no escape flag, since there is no history to lose', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: '<html><body>a plain draft</body></html>' } }));

    const result = await execute(graph, { messageId: 'AAMk1', bodyContent: '<p>whole new body</p>', bodyContentType: 'HTML' });

    expect(result.ok).toBe(true);
    expect(patches).toEqual([{ path: '/me/messages/AAMk1', body: { body: { contentType: 'HTML', content: '<p>whole new body</p>' } } }]);
  });

  it('refuses to replace the body of a message that is not a draft, before writing to it', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: false, body: { contentType: 'html', content: '<html><body>sent</body></html>' } }));

    const result = await execute(graph, { messageId: 'AAMk1', bodyContent: '<p>whole new body</p>' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.message).toContain('is not a draft');
    expect(patches).toEqual([]);
  });

  it('refuses a comment and a body-content together, before reading anything', async () => {
    const { graph, gets } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: htmlDraft('<div>old</div>') } }));

    const result = await execute(graph, { messageId: 'AAMk1', comment: 'revised', bodyContent: '<p>whole new body</p>' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    // The two mean opposite things, so the CLI must not guess which was meant.
    expect(gets).toEqual([]);
  });

  it('refuses a comment that itself carries a quote boundary marker, before reading anything', async () => {
    const { graph, gets } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: htmlDraft('<div>old</div>') } }));

    const result = await execute(graph, { messageId: 'AAMk1', comment: '<div class="gmail_quote">pasted</div>', bodyContentType: 'HTML' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
    expect(gets).toEqual([]);
  });

  it('carries a revised reply and a subject change in a single patch', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html', content: htmlDraft('<div>old</div>') } }));

    await execute(graph, { messageId: 'AAMk1', comment: 'revised', subject: 'RE: Q3 planning' });

    expect(patches).toEqual([{ path: '/me/messages/AAMk1', body: { subject: 'RE: Q3 planning', body: { contentType: 'html', content: htmlDraft('revised') } } }]);
  });

  it('passes a failed draft read through untouched, never patching a body it could not see', async () => {
    const { graph, patches } = recordingGraph(err({ type: 'api_error', status: 404, message: 'ItemNotFound' }));

    const result = await execute(graph, { messageId: 'AAMk1', comment: 'revised' });

    expect(result).toEqual(err({ type: 'api_error', status: 404, message: 'ItemNotFound' }));
    expect(patches).toEqual([]);
  });

  it('refuses a draft whose read-back shape it cannot revise, rather than patching blind', async () => {
    const { graph, patches } = recordingGraph(ok({ isDraft: true, body: { contentType: 'html' } }));

    const result = await execute(graph, { messageId: 'AAMk1', comment: 'revised' });

    expect(result.ok).toBe(false);
    expect(patches).toEqual([]);
  });
});
