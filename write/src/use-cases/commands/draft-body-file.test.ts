import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createLocalBytesFake, UTF8_BOM, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { createMailboxFake } from '../../test-helpers/mailbox-fake.ts';
import type { Message } from '../../test-helpers/mailbox-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';
import { commands } from './write-commands.ts';

/*
 * `--body-file` on the four draft commands: the body text comes from a local
 * UTF-8 file (a byte order mark is removed), at most 1,000,000 bytes. It
 * stands in for --body-content (create-mail-draft, update-mail-draft) or
 * --comment (create-reply-draft, create-forward-draft). The dry-run reads the
 * file too, and the undo of each command stays as it is.
 */

const SOURCE: Message = { id: 'msg-1', isDraft: false, changeKey: 'msg-1-v1' };
const DRAFT: Message = { id: 'd1', isDraft: true, changeKey: 'd1-v1', subject: 'Plan', body: { contentType: 'Text', content: 'Old text.' } };
const LONG = 'Line one of a long body.\nLine two: café, 東京, emoji 🙂.\n';

const files = (): ReturnType<typeof createLocalBytesFake> =>
  createLocalBytesFake({
    'body.txt': utf8Bytes(LONG, UTF8_BOM),
    'plain.txt': utf8Bytes(LONG),
    'bom-only.txt': Uint8Array.from(UTF8_BOM),
    'empty.txt': new Uint8Array(0),
    'latin1.txt': Uint8Array.from([0x63, 0x61, 0x66, 0xe9]),
    'max.txt': new Uint8Array(1_000_000).fill(0x61),
    'over.txt': new Uint8Array(1_000_001).fill(0x61),
  });

const NEW_MAIL = { subject: 'Report', toRecipients: 'kim.lee@contoso.com' };

describe('--body-file on the draft commands', () => {
  it('makes a new draft with the text of the file, the byte order mark removed, and the undo still moves it to Deleted Items', async () => {
    const mailbox = createMailboxFake();
    const journal = createRunJournalFake();
    const run = await runWrite(mailbox.graph, 'create-mail-draft', { ...NEW_MAIL, bodyFile: 'body.txt' }, { journal, files: files() });
    expect(run).toMatchObject({ id: 'draft-1', runId: expect.any(String) });
    expect(mailbox.messages.get('draft-1')?.['body']).toEqual({ contentType: 'Text', content: LONG });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone']);
    expect(mailbox.messages.has('draft-1')).toBe(false);
  });

  it('reads the file in a dry-run too, and writes nothing', async () => {
    const mailbox = createMailboxFake();
    const disk = files();
    const dry = await runWrite(mailbox.graph, 'create-mail-draft', { ...NEW_MAIL, bodyFile: 'plain.txt' }, { dryRun: true, files: disk });
    expect(dry['changes']).toEqual([{ action: 'create-draft', target: '/me/messages', before: {} }]);
    expect(disk.reads).toEqual(['plain.txt']);
    expect(mailbox.calls).toEqual([]);
    const refused = await runWrite(mailbox.graph, 'create-mail-draft', { ...NEW_MAIL, bodyFile: 'missing.txt' }, { dryRun: true, files: disk });
    expect(refused['error']).toMatchObject({ code: 'local_file_not_found' });
  });

  it('gives a reply and a forward the text of the file above the quote', async () => {
    const mailbox = createMailboxFake([SOURCE]);
    await runWrite(mailbox.graph, 'create-reply-draft', { replyToMessageId: 'msg-1', bodyFile: 'body.txt' }, { files: files() });
    await runWrite(mailbox.graph, 'create-forward-draft', { forwardMessageId: 'msg-1', toRecipients: 'kim.lee@contoso.com', bodyFile: 'plain.txt' }, { files: files() });
    expect(mailbox.calls).toEqual([
      'POST /me/messages/msg-1/createReplyAll',
      'GET /me/messages/draft-1?$select=changeKey',
      'POST /me/messages/msg-1/createForward',
      'GET /me/messages/draft-2?$select=changeKey',
    ]);
    expect(mailbox.messages.get('draft-1')?.['comment']).toBe(LONG);
    expect(mailbox.messages.get('draft-2')?.['comment']).toBe(LONG);
  });

  it('replaces the body of a draft with the text of the file, and the undo writes the body of before back', async () => {
    const mailbox = createMailboxFake([DRAFT]);
    const journal = createRunJournalFake();
    await runWrite(mailbox.graph, 'update-mail-draft', { messageId: 'd1', bodyFile: 'body.txt', bodyContentType: 'HTML' }, { journal, files: files() });
    expect(mailbox.messages.get('d1')?.['body']).toEqual({ contentType: 'HTML', content: LONG });
    expect(statusesOf(await undoWrite(mailbox.graph, journal, true))).toEqual(['undone']);
    expect(mailbox.messages.get('d1')?.['body']).toEqual({ contentType: 'Text', content: 'Old text.' });
  });

  it('takes a file of 1,000,000 bytes and refuses a file of 1,000,001 bytes', async () => {
    const mailbox = createMailboxFake();
    expect(await runWrite(mailbox.graph, 'create-mail-draft', { ...NEW_MAIL, bodyFile: 'max.txt' }, { files: files() })).toMatchObject({ id: 'draft-1' });
    expect((await runWrite(mailbox.graph, 'create-mail-draft', { ...NEW_MAIL, bodyFile: 'over.txt' }, { files: files() }))['error']).toMatchObject({
      code: 'local_file_too_large',
      message: 'over.txt is 1000001 bytes, above the limit of 1000000 bytes for this upload.',
    });
  });
});

describe('the refusals of --body-file', () => {
  const refusal = async (name: string, params: Record<string, string>): Promise<unknown> => {
    const mailbox = createMailboxFake([SOURCE, DRAFT]);
    const result = await runWrite(mailbox.graph, name, params, { files: files() });
    expect(mailbox.calls).toEqual([]);
    return result['error'];
  };

  it('refuses --body-file with the text option of the command, and names both options', async () => {
    expect(await refusal('create-mail-draft', { ...NEW_MAIL, bodyContent: 'x', bodyFile: 'body.txt' })).toMatchObject({
      code: 'body_file_conflict',
      message: '--body-file and --body-content are mutually exclusive: give the body text in one of them.',
    });
    expect(await refusal('create-reply-draft', { replyToMessageId: 'msg-1', comment: 'x', bodyFile: 'body.txt' })).toMatchObject({
      message: '--body-file and --comment are mutually exclusive: give the body text in one of them.',
    });
    expect(await refusal('create-forward-draft', { forwardMessageId: 'msg-1', toRecipients: 'a@b.c', comment: 'x', bodyFile: 'body.txt' })).toMatchObject({
      code: 'body_file_conflict',
    });
    expect(await refusal('update-mail-draft', { messageId: 'd1', bodyContent: 'x', bodyFile: 'body.txt' })).toMatchObject({ code: 'body_file_conflict' });
    expect(await refusal('update-mail-draft', { messageId: 'd1', comment: 'x', bodyFile: 'body.txt' })).toMatchObject({
      message: '--body-file and --comment are mutually exclusive: give the body text in one of them.',
    });
  });

  it('refuses a new draft, a reply and a forward with no body text', async () => {
    expect(await refusal('create-mail-draft', NEW_MAIL)).toMatchObject({ code: 'body_missing', message: 'Give the body text with --body-content or --body-file.' });
    expect(await refusal('create-reply-draft', { replyToMessageId: 'msg-1' })).toMatchObject({
      code: 'body_missing',
      message: 'Give the body text with --comment or --body-file.',
    });
    expect(await refusal('create-forward-draft', { forwardMessageId: 'msg-1', toRecipients: 'a@b.c' })).toMatchObject({ code: 'body_missing' });
  });

  it('refuses a file that is not UTF-8, and a file with no text, and never quotes the file', async () => {
    expect(await refusal('create-mail-draft', { ...NEW_MAIL, bodyFile: 'latin1.txt' })).toEqual({
      source: 'validation',
      code: 'body_file_not_utf8',
      message: 'latin1.txt is not UTF-8 text. Save the body file as UTF-8 (a byte order mark is removed).',
    });
    for (const path of ['empty.txt', 'bom-only.txt']) {
      expect(await refusal('create-reply-draft', { replyToMessageId: 'msg-1', bodyFile: path })).toEqual({
        source: 'validation',
        code: 'body_file_empty',
        message: `${path} has no text. Write the body text in the file.`,
      });
    }
  });

  it('refuses --body-file for a library caller that gives no file reader', async () => {
    const mailbox = createMailboxFake();
    const made = await commands['create-mail-draft']?.execute(mailbox.graph, { ...NEW_MAIL, bodyFile: 'body.txt' });
    expect(made).toMatchObject({ ok: false, error: { code: 'local_files_unavailable' } });
    expect(mailbox.calls).toEqual([]);
  });

  it('lets a library caller give the file reader as the third argument of execute', async () => {
    const mailbox = createMailboxFake([SOURCE]);
    const made = await commands['create-reply-draft']?.execute(mailbox.graph, { replyToMessageId: 'msg-1', bodyFile: 'plain.txt' }, files());
    expect(made).toMatchObject({ ok: true, value: { id: 'draft-1' } });
    expect(mailbox.messages.get('draft-1')?.['comment']).toBe(LONG);
  });
});
