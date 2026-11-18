import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { StoredMessage } from '../../test-helpers/mail-store-fake.ts';
import { createMailStoreFake } from '../../test-helpers/mail-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * move-mail-thread moves the messages of one conversation that are in the
 * folder --from, and leaves the messages of the conversation in other folders
 * and the messages of other conversations where they are.
 */

const message = (id: string, folder: string, conversationId: string): StoredMessage => ({ id, parentFolderId: folder, changeKey: `${id}-ck`, conversationId });

const SEED = [message('a1', 'inbox-id', 'c1'), message('a2', 'inbox-id', 'c1'), message('a3', 'sent-id', 'c1'), message('b1', 'inbox-id', 'c2')];

const PARAMS = { conversationId: 'c1', from: 'inbox', to: 'Inbox/Projects' };

describe('move-mail-thread', () => {
  it('shows the messages of the conversation in --from, with their folder, and moves nothing on a dry-run', async () => {
    const mailbox = createMailStoreFake(SEED);
    const answer = await runWrite(mailbox.graph, 'move-mail-thread', PARAMS, { dryRun: true });
    expect(answer['changes']).toEqual([
      { action: 'move-message', target: '/me/messages/a1/move', before: { messageId: 'a1', folderId: 'inbox-id', folder: 'Inbox' } },
      { action: 'move-message', target: '/me/messages/a2/move', before: { messageId: 'a2', folderId: 'inbox-id', folder: 'Inbox' } },
    ]);
    expect(mailbox.calls).toContain("GET /me/mailFolders/inbox-id/messages?$filter=conversationId%20eq%20'c1'&$select=id&$top=1000");
    expect(mailbox.calls.filter((call) => !call.startsWith('GET'))).toEqual([]);
    expect(answer).not.toHaveProperty('skipped');
  });

  it('moves them, and the undo moves them back', async () => {
    const mailbox = createMailStoreFake(SEED);
    const journal = createRunJournalFake();
    expect((await runWrite(mailbox.graph, 'move-mail-thread', PARAMS, { journal }))['done']).toBe(2);
    expect([...mailbox.messages.values()].filter((m) => m.parentFolderId === 'projects-id')).toHaveLength(2);
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone', 'undone']);
    expect([...mailbox.messages.values()].filter((m) => m.parentFolderId === 'inbox-id')).toHaveLength(3);
  });

  it('refuses a refused target before it reads the conversation', async () => {
    const mailbox = createMailStoreFake(SEED);
    const answer = await runWrite(mailbox.graph, 'move-mail-thread', { ...PARAMS, to: 'Junk Email' });
    expect((answer['error'] as { code: string }).code).toBe('folder_refused');
    expect(mailbox.calls.filter((call) => call.includes('/messages'))).toEqual([]);
  });

  it('changes nothing when the conversation has no message in --from', async () => {
    const answer = await runWrite(createMailStoreFake(SEED).graph, 'move-mail-thread', { ...PARAMS, from: 'archive' }, { dryRun: true });
    expect(answer['changes']).toEqual([]);
  });

  it('refuses a conversation with more messages in the folder than one run moves', async () => {
    const mailbox = createMailStoreFake(SEED);
    const graph = {
      ...mailbox.graph,
      get: async (path: string) => (path.includes('$filter') ? ok({ value: [], '@odata.nextLink': 'https://graph/next' }) : mailbox.graph.get(path)),
    };
    const answer = await runWrite(graph, 'move-mail-thread', PARAMS);
    expect((answer['error'] as { message: string }).message).toBe(
      'The conversation has more than 1000 messages in that folder: one run changes at most 1000. Move them with move-mail-message --ids-file.'
    );
  });

  it('refuses an answer it cannot read, and a message id of Graph that is not an id', async () => {
    const mailbox = createMailStoreFake(SEED);
    const answerWith = (value: unknown): typeof mailbox.graph => ({ ...mailbox.graph, get: async (path) => (path.includes('$filter') ? ok(value) : mailbox.graph.get(path)) });
    expect(((await runWrite(answerWith({ items: [] }), 'move-mail-thread', PARAMS))['error'] as { message: string }).message).toBe(
      'Graph gave the messages of the conversation in a shape that this CLI cannot read.'
    );
    expect(((await runWrite(answerWith({ value: [{ id: 'a/b' }] }), 'move-mail-thread', PARAMS))['error'] as { code: string }).code).toBe('invalid_id');
  });

  it('refuses a --from that names no folder', async () => {
    const answer = await runWrite(createMailStoreFake(SEED).graph, 'move-mail-thread', { ...PARAMS, from: 'Nowhere' });
    expect((answer['error'] as { code: string }).code).toBe('folder_not_found');
  });
});
