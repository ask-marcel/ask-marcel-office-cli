import { describe, expect, it } from 'bun:test';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createDraftAttachmentsFake, DRAFTS_ID, textAttachment } from '../../test-helpers/draft-attachments-fake.ts';
import type { StoredDraft } from '../../test-helpers/draft-attachments-fake.ts';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';
import { commands } from './write-commands.ts';

/*
 * The faults of the two attachment commands of a mail draft: each Graph read
 * that fails, each answer in a shape that the CLI cannot read, the params that
 * are missing, and the texts of their docs.
 */

const NOTES = textAttachment('a1', 'notes.txt', 'Meeting notes.', { contentType: 'text/plain' });
const draft = (): StoredDraft => ({ id: 'd1', isDraft: true, parentFolderId: DRAFTS_ID, attachments: [NOTES] });
const disk = (): ReturnType<typeof createLocalBytesFake> => createLocalBytesFake({ 'report.pdf': utf8Bytes('%PDF') });
const ADD = { messageId: 'd1', file: 'report.pdf' };
const REMOVE = { messageId: 'd1', attachmentId: 'a1' };
const DENIED: GraphError = { type: 'api_error', status: 403, message: 'Access is denied.' };

const MESSAGE = 'GET /me/messages/d1?$select=isDraft,parentFolderId';
const OUTBOX = 'GET /me/mailFolders/outbox?$select=id';
const LIST = 'GET /me/messages/d1/attachments?$select=id,name,size';
const META = 'GET /me/messages/d1/attachments/a1?$select=id,name,size,contentType,isInline';

const unreadable = (what: string): unknown => ({ source: 'validation', message: `Graph gave ${what} in a shape that this CLI cannot read.` });

describe('the faults of the attachment commands of a mail draft', () => {
  it('stops on each read that fails before the write, and writes nothing', async () => {
    for (const [name, params, call] of [
      ['add-mail-draft-attachment', ADD, MESSAGE],
      ['add-mail-draft-attachment', ADD, OUTBOX],
      ['add-mail-draft-attachment', ADD, LIST],
      ['remove-mail-draft-attachment', REMOVE, MESSAGE],
      ['remove-mail-draft-attachment', REMOVE, META],
    ] as const) {
      const mail = createDraftAttachmentsFake([draft()]);
      mail.fail.set(call, DENIED);
      expect({ call, error: (await runWrite(mail.graph, name, params, { files: disk() }))['error'] }).toEqual({
        call,
        error: expect.objectContaining({ message: 'Access is denied.' }),
      });
      expect(mail.calls.filter((made) => !made.startsWith('GET'))).toEqual([]);
    }
  });

  it('names each answer that it cannot read', async () => {
    for (const [name, params, call, answer, what] of [
      ['add-mail-draft-attachment', ADD, MESSAGE, { isDraft: true }, 'the message'],
      ['add-mail-draft-attachment', ADD, OUTBOX, { name: 'Outbox' }, 'the Outbox folder'],
      ['add-mail-draft-attachment', ADD, LIST, { value: [{ id: 'a/1', name: 'x', size: 1 }] }, 'the attachments of a mail draft'],
      ['remove-mail-draft-attachment', REMOVE, META, { name: 'notes.txt' }, 'an attachment of a mail draft'],
    ] as const) {
      const mail = createDraftAttachmentsFake([draft()]);
      mail.answer.set(call, answer);
      expect((await runWrite(mail.graph, name, params, { files: disk() }))['error']).toEqual(unreadable(what));
    }
  });

  it('refuses missing params and names the option of a bad id', async () => {
    const graph = createDraftAttachmentsFake([draft()]).graph;
    expect(await commands['add-mail-draft-attachment']?.execute(graph, {}, disk())).toEqual({
      ok: false,
      error: { type: 'validation_error', message: '--message-id is missing; --file is missing' },
    });
    expect(await commands['remove-mail-draft-attachment']?.execute(graph, {})).toEqual({
      ok: false,
      error: { type: 'validation_error', message: '--message-id is missing; --attachment-id is missing' },
    });
    for (const [name, params] of [
      ['add-mail-draft-attachment', { messageId: 'd/1', file: 'report.pdf' }],
      ['remove-mail-draft-attachment', { messageId: 'd/1', attachmentId: 'a1' }],
    ] as const) {
      expect((await runWrite(graph, name, params, { files: disk() }))['error']).toMatchObject({
        message: expect.stringMatching(/^--message-id: "d\/1" is not a Graph id/),
      });
    }
  });

  it('gives a library caller the api error of an upload that Graph gave with no id it can find', async () => {
    const mail = createDraftAttachmentsFake([draft()]);
    mail.uploadWithoutId = true;
    mail.onCall = (call) => {
      if (call.startsWith('UPLOAD')) mail.answer.set(LIST, { value: [] });
    };
    expect(await commands['add-mail-draft-attachment']?.execute(mail.graph, ADD, disk())).toMatchObject({
      ok: false,
      error: { type: 'api_error', status: 502, code: 'attachment_id_unknown' },
    });
  });

  it('finds the new attachment by its name, keeps the id that the upload gave, and stops on a failed read of the list', async () => {
    const other = createDraftAttachmentsFake([draft()]);
    other.uploadWithoutId = true;
    other.onCall = (call) => {
      if (call.startsWith('UPLOAD')) other.messages.get('d1')?.attachments.push(textAttachment('x9', 'other.pdf', 'x'));
    };
    expect((await runWrite(other.graph, 'add-mail-draft-attachment', ADD, { files: disk() }))['changes']).toMatchObject([{ after: { attachmentId: 'att-1' } }]);
    const withId = createDraftAttachmentsFake([draft()]);
    withId.onCall = (call) => {
      if (call.startsWith('UPLOAD')) withId.fail.set(LIST, DENIED);
    };
    expect((await runWrite(withId.graph, 'add-mail-draft-attachment', ADD, { files: disk() }))['changes']).toMatchObject([{ after: { attachmentId: 'att-1' } }]);
    const noList = createDraftAttachmentsFake([draft()]);
    noList.uploadWithoutId = true;
    noList.onCall = (call) => {
      if (call.startsWith('UPLOAD')) noList.fail.set(LIST, DENIED);
    };
    expect((await runWrite(noList.graph, 'add-mail-draft-attachment', ADD, { files: disk() }))['error']).toMatchObject({ message: 'Access is denied.' });
  });

  it('deletes nothing when the read of the bytes fails, and keeps a content type that Graph gives as null out of the journal', async () => {
    const failed = createDraftAttachmentsFake([draft()]);
    failed.fail.set('GET /me/messages/d1/attachments/a1', DENIED);
    expect((await runWrite(failed.graph, 'remove-mail-draft-attachment', REMOVE))['error']).toMatchObject({ message: 'Access is denied.' });
    expect(failed.calls.filter((call) => call.startsWith('DELETE'))).toEqual([]);
    const typeless = createDraftAttachmentsFake([draft()]);
    typeless.answer.set('GET /me/messages/d1/attachments/a1', { contentBytes: NOTES.contentBytes, contentType: null });
    const journal = createRunJournalFake();
    await runWrite(typeless.graph, 'remove-mail-draft-attachment', REMOVE, { journal });
    expect(statusesOf(await undoWrite(typeless.graph, journal, true))).toEqual(['undone']);
    expect(typeless.messages.get('d1')?.attachments).toEqual([{ id: 'att-1', name: 'notes.txt', contentBytes: NOTES.contentBytes }]);
  });

  it('keeps an added attachment whose name it cannot read in the undo check', async () => {
    const mail = createDraftAttachmentsFake([{ ...draft(), attachments: [] }]);
    const journal = createRunJournalFake();
    await runWrite(mail.graph, 'add-mail-draft-attachment', ADD, { journal, files: disk() });
    mail.answer.set('GET /me/messages/d1/attachments/att-1?$select=name', { size: 4 });
    expect(statusesOf(await undoWrite(mail.graph, journal, true))).toEqual([expect.stringMatching(/^changed: /)]);
  });
});

describe('the docs of the attachment commands of a mail draft and of --body-file', () => {
  it('say what each undo does', () => {
    expect([commands['add-mail-draft-attachment']?.meta.undo, commands['remove-mail-draft-attachment']?.meta.undo]).toEqual([
      'removes the attachment, only if the message is still an unsent draft out of the Outbox and still has the attachment with its name',
      'attaches the file again from the bytes in the run journal, only if the message is still an unsent draft out of the Outbox. It is a new attachment with a new id',
    ]);
  });

  it('name the option that --body-file stands in for on each draft command', () => {
    const stands = Object.fromEntries(
      ['create-mail-draft', 'update-mail-draft', 'create-reply-draft', 'create-forward-draft'].map((name) => [
        name,
        commands[name]?.meta.options.find((option) => option.name === 'body-file')?.description.match(/in place of (--[a-z-]+)/)?.[1],
      ])
    );
    expect(stands).toEqual({
      'create-mail-draft': '--body-content',
      'update-mail-draft': '--body-content',
      'create-reply-draft': '--comment',
      'create-forward-draft': '--comment',
    });
  });
});
