import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createDraftAttachmentsFake, DRAFTS_ID, OUTBOX_ID, textAttachment } from '../../test-helpers/draft-attachments-fake.ts';
import type { StoredDraft } from '../../test-helpers/draft-attachments-fake.ts';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * add-mail-draft-attachment between the plan and the write (a draft sent in
 * between, an upload with no id, a failed upload), and its undo, which runs
 * only while the message is a draft that still has the attachment.
 */

const draft = (extra: Partial<StoredDraft> = {}): StoredDraft => ({ id: 'd1', isDraft: true, parentFolderId: DRAFTS_ID, attachments: [], ...extra });

const disk = (): ReturnType<typeof createLocalBytesFake> =>
  createLocalBytesFake({
    'report.pdf': utf8Bytes('%PDF-1.7 report'),
    'empty.txt': new Uint8Array(0),
  });

const ADD = { messageId: 'd1', file: 'report.pdf' };
const BEFORE = { messageId: 'd1', file: 'report.pdf', size: '15', contentType: 'application/pdf' };

describe('add-mail-draft-attachment between the plan and the write', () => {
  it('passes over a draft that Outlook sent after the plan, and uploads nothing', async () => {
    const mail = createDraftAttachmentsFake([draft()]);
    let reads = 0;
    mail.onCall = (call) => {
      reads += call.startsWith('GET /me/messages/d1?') ? 1 : 0;
      if (reads === 2) (mail.messages.get('d1') as StoredDraft).isDraft = false;
    };
    const run = await runWrite(mail.graph, 'add-mail-draft-attachment', ADD, { files: disk() });
    expect(run['skipped']).toEqual([
      {
        target: '/me/messages/d1/attachments',
        before: BEFORE,
        reason: 'The item changed after the plan, so this run passes over it. Run the command again to see it as it is now.',
      },
    ]);
    expect(mail.calls.filter((call) => call.startsWith('UPLOAD'))).toEqual([]);
  });

  it('finds the id of an attachment that the upload gave with no id, by its name among the new attachments', async () => {
    const mail = createDraftAttachmentsFake([draft({ attachments: [textAttachment('a0', 'report.pdf', 'old')] })]);
    mail.uploadWithoutId = true;
    const run = await runWrite(mail.graph, 'add-mail-draft-attachment', ADD, { files: disk() });
    expect((run['changes'] as ReadonlyArray<{ after: unknown }>)[0]?.after).toEqual({ messageId: 'd1', attachmentId: 'att-1' });
  });

  it('reports an upload with no id that it cannot find, and keeps no change', async () => {
    const mail = createDraftAttachmentsFake([draft()]);
    mail.uploadWithoutId = true;
    mail.onCall = (call) => {
      if (call.startsWith('UPLOAD')) mail.answer.set('GET /me/messages/d1/attachments?$select=id,name,size', { value: [] });
    };
    expect((await runWrite(mail.graph, 'add-mail-draft-attachment', ADD, { files: disk() }))['error']).toEqual({
      source: 'graph',
      code: 'attachment_id_unknown',
      message: 'Graph attached the file, but gave no id for it, so the undo cannot remove it. Remove it in Outlook if you do not want it.',
    });
  });

  it('reports a failed upload, and records no change', async () => {
    const mail = createDraftAttachmentsFake([draft()]);
    mail.fail.set('UPLOAD /me/messages/d1/attachments', { type: 'api_error', status: 413, message: 'Too large.' });
    const journal = createRunJournalFake();
    expect((await runWrite(mail.graph, 'add-mail-draft-attachment', ADD, { journal, files: disk() }))['error']).toMatchObject({ message: 'Too large.' });
    expect(journal.records.size).toBe(0);
  });
});

describe('the undo of add-mail-draft-attachment', () => {
  const added = async (): Promise<{ mail: ReturnType<typeof createDraftAttachmentsFake>; journal: ReturnType<typeof createRunJournalFake> }> => {
    const mail = createDraftAttachmentsFake([draft()]);
    const journal = createRunJournalFake();
    await runWrite(mail.graph, 'add-mail-draft-attachment', ADD, { journal, files: disk() });
    return { mail, journal };
  };

  it('keeps the attachment of a message that is no longer a draft, or that is in the Outbox', async () => {
    for (const change of [{ isDraft: false }, { parentFolderId: OUTBOX_ID }]) {
      const { mail, journal } = await added();
      Object.assign(mail.messages.get('d1') as StoredDraft, change);
      expect(statusesOf(await undoWrite(mail.graph, journal, true))).toEqual([expect.stringMatching(/^changed: /)]);
      expect(mail.messages.get('d1')?.attachments.length).toBe(1);
    }
  });

  it('says missing for an attachment or a draft that is gone', async () => {
    const { mail, journal } = await added();
    (mail.messages.get('d1') as StoredDraft).attachments.length = 0;
    expect(statusesOf(await undoWrite(mail.graph, journal, false))).toEqual([expect.stringMatching(/^missing: /)]);
    mail.messages.delete('d1');
    expect(statusesOf(await undoWrite(mail.graph, journal, false))).toEqual([expect.stringMatching(/^missing: /)]);
  });

  it('keeps an attachment that has another name now', async () => {
    const { mail, journal } = await added();
    const attachments = (mail.messages.get('d1') as StoredDraft).attachments;
    attachments[0] = { ...textAttachment('att-1', 'renamed.pdf', 'x') };
    expect(statusesOf(await undoWrite(mail.graph, journal, false))).toEqual([expect.stringMatching(/^changed: /)]);
  });
});
