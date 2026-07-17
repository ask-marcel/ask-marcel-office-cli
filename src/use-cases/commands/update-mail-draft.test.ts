import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute } from './update-mail-draft.ts';

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
        'At least one field must be provided to update (--subject, --body-content, --to-recipients, --cc-recipients, --bcc-recipients, or --importance). Pass an empty string to a recipient flag to clear that list.'
      );
    }
  });
});
