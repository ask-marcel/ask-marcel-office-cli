import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { execute } from './update-mail-draft.ts';

type PatchCall = { path: string; body: Record<string, unknown> };

const capturingGraph = (calls: PatchCall[]): GraphClient => ({
  get: async () => ok({}),
  patch: async (path, body) => {
    calls.push({ path, body: body as Record<string, unknown> });
    return ok({ id: 'm1', isDraft: true });
  },
  post: async () => ok({}),
  getBinary: async () => ok({}),
  getElevated: async () => ok({}),
  teamsChat: async () => ok({}),
  teamsChatIc3: async () => ok({}),
  getBinaryElevated: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
  getCachedTokenInfo: async () => ok({ scopes: [], audience: undefined, expiresAt: undefined, expiresInSeconds: undefined }),
});

describe('update-mail-draft', () => {
  it('patches /me/messages/{id} with only the provided fields (subject, HTML body, recipients, importance)', async () => {
    const calls: PatchCall[] = [];
    const result = await execute(capturingGraph(calls), {
      messageId: 'm1',
      subject: 'Updated',
      bodyContent: '<p>new</p>',
      bodyContentType: 'HTML',
      toRecipients: 'a@example.com',
      ccRecipients: 'c@example.com',
      bccRecipients: 'd@example.com',
      importance: 'Low',
    });
    expect(result).toEqual(ok({ id: 'm1', isDraft: true }));
    expect(calls[0]?.path).toBe('/me/messages/m1');
    expect(calls[0]?.body.subject).toBe('Updated');
    expect(calls[0]?.body.body).toEqual({ contentType: 'HTML', content: '<p>new</p>' });
    expect(calls[0]?.body.toRecipients).toEqual([{ emailAddress: { address: 'a@example.com' } }]);
    expect(calls[0]?.body.ccRecipients).toEqual([{ emailAddress: { address: 'c@example.com' } }]);
    expect(calls[0]?.body.bccRecipients).toEqual([{ emailAddress: { address: 'd@example.com' } }]);
    expect(calls[0]?.body.importance).toBe('Low');
  });

  it('defaults the body contentType to Text when only --body-content is given', async () => {
    const calls: PatchCall[] = [];
    await execute(capturingGraph(calls), { messageId: 'm1', bodyContent: 'plain' });
    expect(calls[0]?.body.body).toEqual({ contentType: 'Text', content: 'plain' });
  });

  it('returns a validation_error when no updatable field is provided (only --message-id)', async () => {
    const result = await execute(capturingGraph([]), { messageId: 'm1' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
    expect(result.error.type === 'validation_error' ? result.error.message : '').toContain('At least one field');
  });

  it('returns a validation_error when --message-id is missing', async () => {
    const result = await execute(capturingGraph([]), { subject: 'S' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });
});
