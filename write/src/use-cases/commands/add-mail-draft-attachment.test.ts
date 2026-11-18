import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createDraftAttachmentsFake, DRAFTS_ID, OUTBOX_ID, textAttachment } from '../../test-helpers/draft-attachments-fake.ts';
import type { StoredDraft } from '../../test-helpers/draft-attachments-fake.ts';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * add-mail-draft-attachment: attaches a local file to an unsent mail draft.
 * The plan reads the file and the draft, and refuses a message that is not a
 * draft or that is in the Outbox, before any write. The undo removes the
 * attachment, only while the message is a draft and still has it.
 */

const draft = (extra: Partial<StoredDraft> = {}): StoredDraft => ({ id: 'd1', isDraft: true, parentFolderId: DRAFTS_ID, attachments: [], ...extra });

const disk = (): ReturnType<typeof createLocalBytesFake> =>
  createLocalBytesFake({
    'report.pdf': utf8Bytes('%PDF-1.7 report'),
    'empty.txt': new Uint8Array(0),
  });

const ADD = { messageId: 'd1', file: 'report.pdf' };
const BEFORE = { messageId: 'd1', file: 'report.pdf', size: '15', contentType: 'application/pdf' };

describe('add-mail-draft-attachment', () => {
  it('shows the file in a dry-run, attaches it, and the undo removes it', async () => {
    const mail = createDraftAttachmentsFake([draft()]);
    const dry = await runWrite(mail.graph, 'add-mail-draft-attachment', ADD, { dryRun: true, files: disk() });
    expect(dry).toEqual({ dryRun: true, command: 'add-mail-draft-attachment', changes: [{ action: 'add-attachment', target: '/me/messages/d1/attachments', before: BEFORE }] });
    expect(mail.calls).toEqual([
      'GET /me/messages/d1?$select=isDraft,parentFolderId',
      'GET /me/mailFolders/outbox?$select=id',
      'GET /me/messages/d1/attachments?$select=id,name,size',
    ]);
    const journal = createRunJournalFake();
    const run = await runWrite(mail.graph, 'add-mail-draft-attachment', ADD, { journal, files: disk() });
    expect(run).toEqual({
      done: 1,
      skipped: [],
      failed: [],
      runId: expect.any(String),
      changes: [{ action: 'add-attachment', target: '/me/messages/d1/attachments', before: BEFORE, after: { messageId: 'd1', attachmentId: 'att-1' } }],
    });
    expect(mail.messages.get('d1')?.attachments).toEqual([
      { id: 'att-1', name: 'report.pdf', contentType: 'application/pdf', contentBytes: Buffer.from('%PDF-1.7 report').toString('base64') },
    ]);
    expect(statusesOf(await undoWrite(mail.graph, journal, true))).toEqual(['undone']);
    expect(mail.messages.get('d1')?.attachments).toEqual([]);
  });

  it('gives the attachment the content type of its extension, in any case, and a general type for an extension it does not know', async () => {
    const expected: Record<string, string> = {
      'a.pdf': 'application/pdf',
      'a.txt': 'text/plain',
      'a.csv': 'text/csv',
      'a.md': 'text/markdown',
      'a.html': 'text/html',
      'a.json': 'application/json',
      'photo.Png': 'image/png',
      'a.jpg': 'image/jpeg',
      'a.JPEG': 'image/jpeg',
      'a.gif': 'image/gif',
      'a.zip': 'application/zip',
      'Plan.DOCX': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'a.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'a.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      '.pdf': 'application/pdf',
      'data.xyz': 'application/octet-stream',
      README: 'application/octet-stream',
      pdf: 'application/octet-stream',
    };
    const files = createLocalBytesFake(Object.fromEntries(Object.keys(expected).map((name) => [name, utf8Bytes('x')])));
    const types: Record<string, unknown> = {};
    for (const file of Object.keys(expected)) {
      const mail = createDraftAttachmentsFake([draft()]);
      const dry = await runWrite(mail.graph, 'add-mail-draft-attachment', { messageId: 'd1', file }, { dryRun: true, files });
      types[file] = (dry['changes'] as ReadonlyArray<{ before: { contentType: string } }>)[0]?.before.contentType;
    }
    expect(types).toEqual(expected);
  });

  it('warns in the dry-run of a name that the draft has already, and of a mail above 25 MB', async () => {
    const big = { ...textAttachment('a1', 'REPORT.pdf', 'old'), contentBytes: Buffer.alloc(25_000_000).toString('base64') };
    const mail = createDraftAttachmentsFake([draft({ attachments: [big] })]);
    const dry = await runWrite(mail.graph, 'add-mail-draft-attachment', ADD, { dryRun: true, files: disk() });
    expect((dry['changes'] as ReadonlyArray<{ before: unknown }>)[0]?.before).toEqual({
      ...BEFORE,
      warning: 'The draft has an attachment with this name already. The mail will be 25000185 bytes with its attachments: many mail servers refuse a mail above 25000000 bytes.',
    });
  });

  it('gives no size warning at 25000000 bytes', async () => {
    const mail = createDraftAttachmentsFake([draft({ attachments: [{ ...textAttachment('a1', 'old.bin', ''), contentBytes: Buffer.alloc(24_999_815).toString('base64') }] })]);
    const dry = await runWrite(mail.graph, 'add-mail-draft-attachment', ADD, { dryRun: true, files: disk() });
    expect((dry['changes'] as ReadonlyArray<{ before: unknown }>)[0]?.before).toEqual(BEFORE);
  });
});

describe('the refusals of add-mail-draft-attachment', () => {
  const refusal = async (seed: StoredDraft, params: Record<string, string> = ADD, files: LocalFileReader = disk()): Promise<{ error: unknown; calls: string[] }> => {
    const mail = createDraftAttachmentsFake([seed]);
    const result = await runWrite(mail.graph, 'add-mail-draft-attachment', params, { files });
    return { error: result['error'], calls: mail.calls };
  };

  it('refuses a message that is not a draft, after one read of the message and no write', async () => {
    for (const isDraft of [false, undefined]) {
      const { error, calls } = await refusal({ ...draft(), isDraft });
      expect(error).toEqual({
        source: 'validation',
        code: 'not_a_draft',
        message: '--message-id: the message is not a draft (it was sent or received). This command changes an unsent draft only.',
      });
      expect(calls).toEqual(['GET /me/messages/d1?$select=isDraft,parentFolderId']);
    }
  });

  it('refuses a draft in the Outbox, and writes nothing', async () => {
    const { error, calls } = await refusal(draft({ parentFolderId: OUTBOX_ID }));
    expect(error).toEqual({
      source: 'validation',
      code: 'draft_in_outbox',
      message: '--message-id: the draft is in the Outbox, so Outlook sends it now. This command changes an unsent draft only.',
    });
    expect(calls).toEqual(['GET /me/messages/d1?$select=isDraft,parentFolderId', 'GET /me/mailFolders/outbox?$select=id']);
  });

  it('refuses an id that is not a Graph id, a missing file and an empty file before any request', async () => {
    expect(await refusal(draft(), { messageId: 'd1/send', file: 'report.pdf' })).toEqual({ error: expect.objectContaining({ code: 'invalid_id' }), calls: [] });
    expect(await refusal(draft(), { messageId: 'd1', file: 'nothing.pdf' })).toEqual({ error: expect.objectContaining({ code: 'local_file_not_found' }), calls: [] });
    expect(await refusal(draft(), { messageId: 'd1', file: 'empty.txt' })).toEqual({
      error: { source: 'validation', code: 'attachment_file_empty', message: 'empty.txt is empty. Graph does not attach an empty file.' },
      calls: [],
    });
  });

  it('takes a file of 150000000 bytes and refuses a file of 150000001 bytes', async () => {
    const sized = (size: number): LocalFileReader => ({ stat: async () => ok({ kind: 'file', size }), readBytes: async () => ok(utf8Bytes('x')) });
    expect((await refusal(draft(), ADD, sized(150_000_001))).error).toMatchObject({
      code: 'local_file_too_large',
      message: 'report.pdf is 150000001 bytes, above the limit of 150000000 bytes for this upload.',
    });
    expect((await refusal(draft(), ADD, sized(150_000_000))).error).toBeUndefined();
  });

  it('refuses a message that is not there', async () => {
    expect((await refusal(draft({ id: 'd2' }))).error).toMatchObject({ source: 'graph', message: expect.stringContaining('not found') });
  });
});
