import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { execute as createForwardDraft } from './create-forward-draft.ts';
import { execute as createMailDraft } from './create-mail-draft.ts';
import { execute as createReplyDraft } from './create-reply-draft.ts';
import { execute as updateMailDraft } from './update-mail-draft.ts';

// A test that the slim shape comes back cannot SEE a return site that was
// missed — and one was, on the forward command's HTML path, because a scripted
// edit matched every `return graph.patch(` except that one. So this asserts the
// NEGATIVE across every path each draft command can take: whatever route the
// code takes, a body must never come back. (Lesson 2026-07-16.)

const QUOTED_BODY = { contentType: 'html', content: `<html><body><div>reply</div><div id="divRplyFwdMsg">${'quoted '.repeat(2000)}</div></body></html>` };

// Every Graph answer on every path hands back a FULL message, exactly as Graph
// does — so any unslimmed route shows up as a leaked body.
const echoingGraph = (): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({
    post: async () => ok({ id: 'draft-9', isDraft: true, subject: 'RE: Contoso migration', body: QUOTED_BODY }),
    get: async () => ok({ id: 'draft-9', isDraft: true, body: QUOTED_BODY }),
    patch: async () => ok({ id: 'draft-9', isDraft: true, subject: 'RE: Contoso migration', body: QUOTED_BODY }),
  });

const paths = [
  { label: 'create-reply-draft, text, no subject', run: () => createReplyDraft(echoingGraph(), { replyToMessageId: 'm1', bodyContent: 'x' }) },
  { label: 'create-reply-draft, text, subject override', run: () => createReplyDraft(echoingGraph(), { replyToMessageId: 'm1', bodyContent: 'x', subject: 'S' }) },
  {
    label: 'create-reply-draft, HTML splice',
    run: () => createReplyDraft(echoingGraph(), { replyToMessageId: 'm1', bodyContent: '<p>x</p>', bodyContentType: 'HTML' }),
  },
  {
    label: 'create-forward-draft, text',
    run: () => createForwardDraft(echoingGraph(), { forwardMessageId: 'm1', toRecipients: 'a@contoso.com', bodyContent: 'x' }),
  },
  {
    label: 'create-forward-draft, text, cc override',
    run: () => createForwardDraft(echoingGraph(), { forwardMessageId: 'm1', toRecipients: 'a@contoso.com', bodyContent: 'x', ccRecipients: 'b@contoso.com' }),
  },
  {
    label: 'create-forward-draft, HTML splice',
    run: () => createForwardDraft(echoingGraph(), { forwardMessageId: 'm1', toRecipients: 'a@contoso.com', bodyContent: '<p>x</p>', bodyContentType: 'HTML' }),
  },
  { label: 'update-mail-draft, plain field update', run: () => updateMailDraft(echoingGraph(), { messageId: 'm1', subject: 'S' }) },
  { label: 'update-mail-draft, --comment revise', run: () => updateMailDraft(echoingGraph(), { messageId: 'm1', comment: 'revised' }) },
  { label: 'create-mail-draft, fresh draft', run: () => createMailDraft(echoingGraph(), { subject: 'S', bodyContent: 'x', toRecipients: 'a@contoso.com' }) },
];

describe('never reading a draft’s own body back to the caller who just wrote it', () => {
  it.each(paths)('answers without the body on: $label', async ({ run }) => {
    const result = await run();

    expect(result.ok).toBe(true);
    if (result.ok) {
      const value = result.value as Record<string, unknown>;
      expect(Object.hasOwn(value, 'body')).toBe(false);
      expect(value.id).toBe('draft-9');
      // The quoted history is ~14 KB here and 174 KB in the wild; the answer
      // must not scale with the thread it is attached to.
      expect(JSON.stringify(result.value).length).toBeLessThan(500);
    }
  });
});
