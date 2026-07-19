import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { slimDraftResult } from './draft-response.ts';

const fullDraft = {
  id: 'draft-9',
  subject: 'RE: Contoso migration',
  toRecipients: [{ emailAddress: { address: 'robin.chen@contoso.com' } }],
  ccRecipients: [],
  bccRecipients: [],
  importance: 'normal',
  bodyPreview: 'Confirmed for Contoso.',
  isDraft: true,
  webLink: 'https://outlook.office365.com/mail/draft-9',
  conversationId: 'conv-1',
  body: { contentType: 'html', content: '<html>…180 KB of quoted history…</html>' },
  '@odata.etag': 'W/"abc"',
  changeKey: 'abc',
  parentFolderId: 'folder-1',
};

describe('answering a draft write without reading the caller’s own text back to them', () => {
  it('confirms the write with the fields that identify the draft, and drops the body it just wrote', () => {
    const result = slimDraftResult(ok(fullDraft));

    expect(result).toEqual(
      ok({
        id: 'draft-9',
        subject: 'RE: Contoso migration',
        toRecipients: [{ emailAddress: { address: 'robin.chen@contoso.com' } }],
        ccRecipients: [],
        bccRecipients: [],
        importance: 'normal',
        bodyPreview: 'Confirmed for Contoso.',
        isDraft: true,
        webLink: 'https://outlook.office365.com/mail/draft-9',
        conversationId: 'conv-1',
      })
    );
  });

  it('keeps bodyPreview so the caller can still see WHICH draft answered, without the quoted history behind it', () => {
    const result = slimDraftResult(ok(fullDraft));

    if (result.ok) {
      const value = result.value as Record<string, unknown>;
      expect(value.bodyPreview).toBe('Confirmed for Contoso.');
      expect(Object.hasOwn(value, 'body')).toBe(false);
    }
  });

  it('passes a 204 acknowledgement straight through, since it is not a message to project', () => {
    // Graph answers PATCH with 204 No Content and the client surfaces `{ ok: true }`.
    // Projecting that against the message field list would empty it to `{}`.
    expect(slimDraftResult(ok({ ok: true }))).toEqual(ok({ ok: true }));
  });

  it('leaves a failure untouched, so an error envelope is never mistaken for a draft', () => {
    const failure = err({ type: 'api_error' as const, status: 404, message: 'ErrorItemNotFound' });

    expect(slimDraftResult(failure)).toEqual(failure);
  });

  it('reports only the fields Graph actually sent, rather than padding the answer with empty ones', () => {
    const result = slimDraftResult(ok({ id: 'draft-9', body: { content: 'x' } }));

    expect(result).toEqual(ok({ id: 'draft-9' }));
    // The KEY SET, not just the value: `toEqual` treats `{id, subject: undefined}`
    // as equal to `{id}`, so it cannot see an answer padded with empty fields.
    if (result.ok) expect(Object.keys(result.value as object)).toEqual(['id']);
  });

  it('passes a response with no id through untouched rather than emptying it', () => {
    // Anything without an id is not a message; blanking it would destroy
    // whatever Graph was actually saying.
    expect(slimDraftResult(ok({ unexpected: 'shape' }))).toEqual(ok({ unexpected: 'shape' }));
    expect(slimDraftResult(ok(null))).toEqual(ok(null));
    expect(slimDraftResult(ok('not-an-object'))).toEqual(ok('not-an-object'));
  });
});
