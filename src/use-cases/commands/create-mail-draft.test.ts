import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { execute } from './create-mail-draft.ts';

type PostCall = { path: string; body: Record<string, unknown> };

const capturingGraph = (calls: PostCall[]): GraphClient => ({
  get: async () => ok({}),
  patch: async () => ok({}),
  post: async (path, body) => {
    calls.push({ path, body: body as Record<string, unknown> });
    return ok({ id: 'draft-1', isDraft: true });
  },
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

describe('create-mail-draft', () => {
  it('posts a Text draft to /me/messages with the subject, body and comma-split recipients', async () => {
    const calls: PostCall[] = [];
    const result = await execute(capturingGraph(calls), { subject: 'Q3 Report', bodyContent: 'Please review.', toRecipients: 'alice@example.com, bob@example.com' });
    expect(result).toEqual(ok({ id: 'draft-1', isDraft: true }));
    expect(calls[0]?.path).toBe('/me/messages');
    expect(calls[0]?.body.subject).toBe('Q3 Report');
    expect(calls[0]?.body.body).toEqual({ contentType: 'Text', content: 'Please review.' });
    expect(calls[0]?.body.toRecipients).toEqual([{ emailAddress: { address: 'alice@example.com' } }, { emailAddress: { address: 'bob@example.com' } }]);
    expect(calls[0]?.body.ccRecipients).toBeUndefined();
    expect(calls[0]?.body.importance).toBeUndefined();
  });

  it('targets the folder-scoped path and includes the HTML body, cc, bcc and importance when supplied', async () => {
    const calls: PostCall[] = [];
    await execute(capturingGraph(calls), {
      subject: 'S',
      bodyContent: '<p>hi</p>',
      bodyContentType: 'HTML',
      toRecipients: 'a@example.com',
      ccRecipients: 'c@example.com',
      bccRecipients: 'd@example.com',
      importance: 'High',
      mailFolderId: 'drafts',
    });
    expect(calls[0]?.path).toBe('/me/mailFolders/drafts/messages');
    expect(calls[0]?.body.body).toEqual({ contentType: 'HTML', content: '<p>hi</p>' });
    expect(calls[0]?.body.ccRecipients).toEqual([{ emailAddress: { address: 'c@example.com' } }]);
    expect(calls[0]?.body.bccRecipients).toEqual([{ emailAddress: { address: 'd@example.com' } }]);
    expect(calls[0]?.body.importance).toBe('High');
  });

  it('returns a validation_error when a required field (to-recipients) is missing', async () => {
    const result = await execute(capturingGraph([]), { subject: 'S', bodyContent: 'B' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('validation_error');
  });
});
