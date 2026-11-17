import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeWriteGraph } from '../../test-helpers/write-graph-fake.ts';
import type { WriteGraph } from '../ports/graph.ts';
import { writeGraphOf } from './command-graph.ts';

/*
 * No command sends mail (product rule). The write view a write command gets
 * refuses, before any request, a change to a path that sends: `send`,
 * `sendMail`, `reply`, `replyAll`, `forward`, in any case, with or without a
 * query string, a trailing slash, the `microsoft.graph.` namespace or a
 * percent-encoded letter, and a JSON batch, which can carry a send in its
 * body. The actions that make a draft stay permitted.
 */

// A write graph that records each request it is asked to send.
const recordingGraph = (sent: string[]): WriteGraph =>
  fakeWriteGraph({
    post: async (path) => {
      sent.push(`POST ${path}`);
      return ok({ id: 'draft-1' });
    },
    patch: async (path) => {
      sent.push(`PATCH ${path}`);
      return ok(undefined);
    },
    delete: async (path) => {
      sent.push(`DELETE ${path}`);
      return ok(undefined);
    },
  });

const SENDING_PATHS = [
  '/me/sendMail',
  '/users/u1/sendMail',
  '/me/messages/m1/send',
  '/me/mailFolders/drafts/messages/m1/send',
  '/me/messages/m1/reply',
  '/me/messages/m1/replyAll',
  '/me/messages/m1/forward',
  '/me/events/e1/forward',
  '/groups/g1/threads/t1/reply',
  '/me/messages/m1/SEND',
  '/me/SendMail?$select=id',
  '/me/messages/m1/send/',
  '/me/messages/m1/microsoft.graph.send',
  '/me/messages/m1/Microsoft.Graph.ReplyAll',
  '/me/messages/m1/%73end',
  '/$batch',
];

const DRAFT_PATHS = [
  '/me/messages',
  '/me/mailFolders/drafts/messages',
  '/me/messages/m1/createReply',
  '/me/messages/m1/createReplyAll',
  '/me/messages/m1/createForward',
  '/me/messages/m1/move',
  '/me/messages/m1/attachments',
  '/me/messages/m1/attachments/createUploadSession',
  '/me/messages/send-later-draft',
  '/me/todo/lists/l1/tasks',
  '/planner/tasks',
];

describe('a write command that asks to send mail', () => {
  it.each(SENDING_PATHS)('is refused before any request when it POSTs to %s', async (path) => {
    const sent: string[] = [];
    const result = await writeGraphOf(recordingGraph(sent)).post(path, {});
    expect(sent).toEqual([]);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error).toEqual({
      type: 'validation_error',
      code: 'send_refused',
      message: `POST ${path} was not sent: it would send mail, and this CLI never sends mail. A draft stays in Drafts until you send it yourself from Outlook.`,
    });
  });

  it('is refused on a PATCH and a DELETE of a sending path too', async () => {
    const sent: string[] = [];
    const view = writeGraphOf(recordingGraph(sent));
    const patched = await view.patch('/me/messages/m1/send', {});
    const deleted = await view.delete('/me/messages/m1/forward');
    expect(sent).toEqual([]);
    expect([patched.ok, deleted.ok]).toEqual([false, false]);
    if (patched.ok || deleted.ok) return;
    expect([patched.error.code, deleted.error.code]).toEqual(['send_refused', 'send_refused']);
  });
});

describe('a write command that makes or changes a draft', () => {
  it.each(DRAFT_PATHS)('is sent when it POSTs to %s', async (path) => {
    const sent: string[] = [];
    expect(await writeGraphOf(recordingGraph(sent)).post(path, {})).toEqual(ok({ id: 'draft-1' }));
    expect(sent).toEqual([`POST ${path}`]);
  });

  it('is sent on a PATCH and a DELETE of a draft', async () => {
    const sent: string[] = [];
    const view = writeGraphOf(recordingGraph(sent));
    await view.patch('/me/messages/d1', { subject: 'Hi' });
    await view.delete('/me/messages/d1');
    expect(sent).toEqual(['PATCH /me/messages/d1', 'DELETE /me/messages/d1']);
  });
});
