import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createDraftAttachmentsFake, DRAFTS_ID, OUTBOX_ID, textAttachment } from '../../test-helpers/draft-attachments-fake.ts';
import type { StoredDraft } from '../../test-helpers/draft-attachments-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';
import { commands } from './write-commands.ts';

/*
 * remove-mail-draft-attachment: removes one file attachment of an unsent mail
 * draft. Before the delete, the run reads the bytes of the file, and the run
 * journal keeps them (at most 25 MB), so the undo attaches the file again,
 * with a new id. An attachment that the undo cannot make again (an attached
 * item, a link, an inline picture of the body, a file above 25 MB) is refused.
 */

const NOTES = textAttachment('a1', 'notes.txt', 'Meeting notes.', { contentType: 'text/plain' });
const draft = (extra: Partial<StoredDraft> = {}): StoredDraft => ({ id: 'd1', isDraft: true, parentFolderId: DRAFTS_ID, attachments: [NOTES], ...extra });
const REMOVE = { messageId: 'd1', attachmentId: 'a1' };
const BEFORE = { messageId: 'd1', attachmentId: 'a1', file: 'notes.txt', size: '184' };
const SELECT = 'GET /me/messages/d1/attachments/a1?$select=id,name,size,contentType,isInline';

describe('remove-mail-draft-attachment', () => {
  it('shows the attachment in a dry-run, removes it, and the undo attaches the file again with a new id', async () => {
    const mail = createDraftAttachmentsFake([draft()]);
    const dry = await runWrite(mail.graph, 'remove-mail-draft-attachment', REMOVE, { dryRun: true });
    expect(dry).toEqual({
      dryRun: true,
      command: 'remove-mail-draft-attachment',
      changes: [{ action: 'remove-attachment', target: '/me/messages/d1/attachments/a1', before: BEFORE }],
    });
    expect(mail.calls).toEqual(['GET /me/messages/d1?$select=isDraft,parentFolderId', 'GET /me/mailFolders/outbox?$select=id', SELECT]);
    const journal = createRunJournalFake();
    const run = await runWrite(mail.graph, 'remove-mail-draft-attachment', REMOVE, { journal });
    expect(run).toEqual({
      done: 1,
      skipped: [],
      failed: [],
      runId: expect.any(String),
      changes: [{ action: 'remove-attachment', target: '/me/messages/d1/attachments/a1', before: BEFORE, after: {} }],
    });
    expect(mail.calls.slice(-2)).toEqual(['GET /me/messages/d1/attachments/a1', 'DELETE /me/messages/d1/attachments/a1']);
    expect(mail.messages.get('d1')?.attachments).toEqual([]);
    const undone = await undoWrite(mail.graph, journal, true);
    expect(undone).toMatchObject({ ok: true, value: { changes: [{ status: 'undone', now: { messageId: 'd1', attachmentId: 'att-1' } }] } });
    expect(mail.messages.get('d1')?.attachments).toEqual([{ ...NOTES, id: 'att-1' }]);
  });

  it('reads the bytes that Graph gives wrapped in value, and attaches with no content type when Graph gives none', async () => {
    const mail = createDraftAttachmentsFake([draft()]);
    mail.answer.set('GET /me/messages/d1/attachments/a1', { value: { contentBytes: NOTES.contentBytes } });
    const journal = createRunJournalFake();
    await runWrite(mail.graph, 'remove-mail-draft-attachment', REMOVE, { journal });
    expect(statusesOf(await undoWrite(mail.graph, journal, true))).toEqual(['undone']);
    expect(mail.messages.get('d1')?.attachments).toEqual([{ id: 'att-1', name: 'notes.txt', contentBytes: NOTES.contentBytes }]);
  });
});

describe('the refusals of remove-mail-draft-attachment', () => {
  const refusal = async (seed: StoredDraft, params: Record<string, string> = REMOVE): Promise<{ error: unknown; calls: string[] }> => {
    const mail = createDraftAttachmentsFake([seed]);
    const result = await runWrite(mail.graph, 'remove-mail-draft-attachment', params);
    expect(mail.calls.filter((call) => !call.startsWith('GET'))).toEqual([]);
    return { error: result['error'], calls: mail.calls };
  };

  it('refuses a message that is not a draft, and a draft in the Outbox', async () => {
    expect(await refusal(draft({ isDraft: false }))).toEqual({
      error: expect.objectContaining({ code: 'not_a_draft' }),
      calls: ['GET /me/messages/d1?$select=isDraft,parentFolderId'],
    });
    expect((await refusal(draft({ parentFolderId: OUTBOX_ID }))).error).toMatchObject({ code: 'draft_in_outbox' });
  });

  it('refuses ids that are not Graph ids before any request', async () => {
    expect(await refusal(draft(), { messageId: 'd1', attachmentId: '../send' })).toEqual({
      error: expect.objectContaining({ code: 'invalid_id', message: expect.stringMatching(/^--attachment-id: /) }),
      calls: [],
    });
    expect(await refusal(draft(), { messageId: 'd 1', attachmentId: 'a1' })).toEqual({ error: expect.objectContaining({ code: 'invalid_id' }), calls: [] });
  });

  it('refuses an attachment that the draft does not have', async () => {
    expect((await refusal(draft({ attachments: [] }))).error).toEqual({
      source: 'validation',
      code: 'mail_attachment_not_found',
      message: '--attachment-id: the draft has no attachment with this id. list-mail-attachments gives the id of each attachment of a message.',
    });
  });

  it('refuses an attached item, a link and an inline picture, which the undo cannot make again', async () => {
    for (const type of ['#microsoft.graph.itemAttachment', '#microsoft.graph.referenceAttachment']) {
      expect((await refusal(draft({ attachments: [{ ...NOTES, type }] }))).error).toEqual({
        source: 'validation',
        code: 'not_a_file_attachment',
        message: '--attachment-id: the attachment is an attached item or a link, not a file. The undo cannot make it again, so this command removes files only.',
      });
    }
    expect((await refusal(draft({ attachments: [{ ...NOTES, isInline: true }] }))).error).toEqual({
      source: 'validation',
      code: 'inline_attachment',
      message: '--attachment-id: the attachment is a picture in the body of the mail. The undo cannot put it back in the body, so this command does not remove it.',
    });
  });

  it('takes a file that the list gives as 25000000 bytes, and refuses one above it', async () => {
    const sized = (listed: number): StoredDraft => draft({ attachments: [{ ...NOTES, contentBytes: Buffer.alloc(listed - 170).toString('base64') }] });
    expect((await refusal(sized(25_000_001))).error).toEqual({
      source: 'validation',
      code: 'attachment_too_large_for_undo',
      message:
        '--attachment-id: the attachment is 25000001 bytes. The run journal keeps at most 25000000 bytes of a file for the undo, so this command does not remove it. Remove it in Outlook.',
    });
    const mail = createDraftAttachmentsFake([sized(25_000_000)]);
    expect((await runWrite(mail.graph, 'remove-mail-draft-attachment', REMOVE, { dryRun: true }))['changes']).toEqual([
      { action: 'remove-attachment', target: '/me/messages/d1/attachments/a1', before: { ...BEFORE, size: '25000000' } },
    ]);
  });

  it('refuses for a library caller the same way', async () => {
    const mail = createDraftAttachmentsFake([draft({ isDraft: false })]);
    expect(await commands['remove-mail-draft-attachment']?.execute(mail.graph, REMOVE)).toMatchObject({ ok: false, error: { code: 'not_a_draft' } });
  });
});
