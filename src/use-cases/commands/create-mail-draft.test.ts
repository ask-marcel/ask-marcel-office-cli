import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './create-mail-draft.ts';

describe('create-mail-draft', () => {
  it('creates a draft in /me/messages with the subject, a Text body, and the comma-separated recipients parsed into Graph address objects', async () => {
    let capturedPath = '';
    let capturedBody: Record<string, unknown> = {};
    const graph = fakeGraphClient({
      post: async (path, body) => {
        capturedPath = path;
        capturedBody = body as Record<string, unknown>;
        return ok({ id: 'msg-1' });
      },
    });
    const result = await execute(graph, { subject: 'Q3 Report', bodyContent: 'Please review.', toRecipients: 'alice@example.com, bob@example.com' });
    expect(result).toEqual(ok({ id: 'msg-1' }));
    expect(capturedPath).toBe('/me/messages');
    expect(capturedBody.subject).toBe('Q3 Report');
    expect(capturedBody.body).toEqual({ contentType: 'Text', content: 'Please review.' });
    expect(capturedBody.toRecipients).toEqual([{ emailAddress: { address: 'alice@example.com' } }, { emailAddress: { address: 'bob@example.com' } }]);
  });

  it('targets the given mail folder and includes the HTML body, cc, bcc, and importance when all optional fields are supplied', async () => {
    let capturedPath = '';
    let capturedBody: Record<string, unknown> = {};
    const graph = fakeGraphClient({
      post: async (path, body) => {
        capturedPath = path;
        capturedBody = body as Record<string, unknown>;
        return ok({ id: 'msg-2' });
      },
    });
    const result = await execute(graph, {
      subject: 'Update',
      bodyContent: '<p>hi</p>',
      bodyContentType: 'HTML',
      toRecipients: 'alice@example.com',
      ccRecipients: 'carol@example.com',
      bccRecipients: 'dave@example.com',
      importance: 'High',
      mailFolderId: 'AAMkFolder',
    });
    expect(result.ok).toBe(true);
    expect(capturedPath).toBe('/me/mailFolders/AAMkFolder/messages');
    expect(capturedBody.body).toEqual({ contentType: 'HTML', content: '<p>hi</p>' });
    expect(capturedBody.ccRecipients).toEqual([{ emailAddress: { address: 'carol@example.com' } }]);
    expect(capturedBody.bccRecipients).toEqual([{ emailAddress: { address: 'dave@example.com' } }]);
    expect(capturedBody.importance).toBe('High');
  });

  it('returns a validation_error when a required field (subject) is missing', async () => {
    const result = await execute(fakeGraphClient(), { bodyContent: 'x', toRecipients: 'a@b.com' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('validation_error');
  });
});
