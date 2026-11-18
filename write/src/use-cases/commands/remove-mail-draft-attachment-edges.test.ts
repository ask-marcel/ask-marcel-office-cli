import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createDraftAttachmentsFake, DRAFTS_ID, textAttachment } from '../../test-helpers/draft-attachments-fake.ts';
import type { StoredDraft } from '../../test-helpers/draft-attachments-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * remove-mail-draft-attachment between the plan and the write (an attachment
 * gone or changed in between, bytes that cannot be read, a failed delete),
 * and its undo, which attaches the file again only to a draft.
 */

const NOTES = textAttachment('a1', 'notes.txt', 'Meeting notes.', { contentType: 'text/plain' });
const draft = (extra: Partial<StoredDraft> = {}): StoredDraft => ({ id: 'd1', isDraft: true, parentFolderId: DRAFTS_ID, attachments: [NOTES], ...extra });
const REMOVE = { messageId: 'd1', attachmentId: 'a1' };
const BEFORE = { messageId: 'd1', attachmentId: 'a1', file: 'notes.txt', size: '184' };
const SELECT = 'GET /me/messages/d1/attachments/a1?$select=id,name,size,contentType,isInline';

describe('remove-mail-draft-attachment between the plan and the write', () => {
  const changedAfterPlan = (mail: ReturnType<typeof createDraftAttachmentsFake>, change: () => void): void => {
    let reads = 0;
    mail.onCall = (call) => {
      reads += call === SELECT ? 1 : 0;
      if (reads === 1 && call.startsWith('GET /me/messages/d1?')) change();
    };
  };

  it('passes over an attachment that is gone, or that changed, after the plan', async () => {
    for (const [change, reason] of [
      [(stored: StoredDraft) => stored.attachments.splice(0, 1), 'The item is no longer there: it was moved or deleted after the plan.'],
      [
        (stored: StoredDraft) => stored.attachments.splice(0, 1, { ...NOTES, name: 'other.txt' }),
        'The item changed after the plan, so this run passes over it. Run the command again to see it as it is now.',
      ],
      [
        (stored: StoredDraft) => stored.attachments.splice(0, 1, { ...NOTES, contentBytes: 'QUJD' }),
        'The item changed after the plan, so this run passes over it. Run the command again to see it as it is now.',
      ],
      [
        (stored: StoredDraft) => Object.assign(stored, { isDraft: false }),
        'The item changed after the plan, so this run passes over it. Run the command again to see it as it is now.',
      ],
    ] as const) {
      const mail = createDraftAttachmentsFake([draft()]);
      changedAfterPlan(mail, () => change(mail.messages.get('d1') as StoredDraft));
      const run = await runWrite(mail.graph, 'remove-mail-draft-attachment', REMOVE);
      expect(run['skipped']).toEqual([{ target: '/me/messages/d1/attachments/a1', before: BEFORE, reason }]);
      expect(mail.calls.filter((call) => call.startsWith('DELETE'))).toEqual([]);
    }
  });

  it('deletes nothing when the bytes of the file cannot be read, and reports a failed delete', async () => {
    const unreadable = createDraftAttachmentsFake([draft()]);
    unreadable.answer.set('GET /me/messages/d1/attachments/a1', { name: 'notes.txt' });
    expect((await runWrite(unreadable.graph, 'remove-mail-draft-attachment', REMOVE))['error']).toMatchObject({
      message: 'Graph gave the bytes of the attachment in a shape that this CLI cannot read.',
    });
    expect(unreadable.calls.filter((call) => call.startsWith('DELETE'))).toEqual([]);
    const refused = createDraftAttachmentsFake([draft()]);
    refused.fail.set('DELETE /me/messages/d1/attachments/a1', { type: 'api_error', status: 403, message: 'Access is denied.' });
    const journal = createRunJournalFake();
    expect((await runWrite(refused.graph, 'remove-mail-draft-attachment', REMOVE, { journal }))['error']).toMatchObject({ message: 'Access is denied.' });
    expect(journal.records.size).toBe(0);
  });
});

describe('the undo of remove-mail-draft-attachment', () => {
  const removed = async (): Promise<{ mail: ReturnType<typeof createDraftAttachmentsFake>; journal: ReturnType<typeof createRunJournalFake> }> => {
    const mail = createDraftAttachmentsFake([draft()]);
    const journal = createRunJournalFake();
    await runWrite(mail.graph, 'remove-mail-draft-attachment', REMOVE, { journal });
    return { mail, journal };
  };

  it('attaches nothing to a message that is no longer a draft, and says missing for a draft that is gone', async () => {
    const { mail, journal } = await removed();
    (mail.messages.get('d1') as StoredDraft).isDraft = false;
    expect(statusesOf(await undoWrite(mail.graph, journal, true))).toEqual([expect.stringMatching(/^changed: /)]);
    mail.messages.delete('d1');
    expect(statusesOf(await undoWrite(mail.graph, journal, true))).toEqual([expect.stringMatching(/^missing: /)]);
    expect(mail.calls.filter((call) => call.startsWith('UPLOAD'))).toEqual([]);
  });

  it('reports a failed upload of the undo, and an upload that gives no id', async () => {
    const { mail, journal } = await removed();
    mail.fail.set('UPLOAD /me/messages/d1/attachments', { type: 'api_error', status: 500, message: 'Upload failed.' });
    expect(statusesOf(await undoWrite(mail.graph, journal, true))).toEqual(['failed: Upload failed.']);
    mail.fail.clear();
    mail.uploadWithoutId = true;
    const undone = await undoWrite(mail.graph, journal, true);
    expect(undone).toMatchObject({ ok: true, value: { changes: [{ status: 'undone' }] } });
    expect(undone.ok && undone.value.changes[0]?.now).toBeUndefined();
  });
});
