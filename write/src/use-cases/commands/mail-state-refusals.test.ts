import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createLocalFilesFake } from '@ask-marcel/office-core/test-helpers/local-files-fake.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { StoredMessage } from '../../test-helpers/mail-store-fake.ts';
import { createMailStoreFake } from '../../test-helpers/mail-store-fake.ts';
import { runWrite } from '../../test-helpers/run-write.ts';
import { FIELD_UNDO_HANDLERS } from './mail-field-write.ts';
import { MOVE_UNDO_HANDLERS } from './mail-move-write.ts';
import { commands } from './write-commands.ts';

/*
 * Each refusal of the mail state commands, as the caller gets it: its message,
 * its code and its source (`validation`: a wrong input, which a retry does not
 * cure). The ids and the folders named in a refusal are those of the param.
 */

const SEED: ReadonlyArray<StoredMessage> = [
  { id: 'm1', parentFolderId: 'inbox-id', changeKey: 'm1-ck', isRead: false, conversationId: 'c1' },
  { id: 'd1', parentFolderId: 'drafts-id', changeKey: 'd1-ck', isDraft: true },
];

const refusal = async (name: string, params: Record<string, string>, graph: WriteGraph = createMailStoreFake(SEED).graph): Promise<unknown> =>
  (await runWrite(graph, name, params))['error'];

const NOT_AN_ID = 'is not a Graph id. An id has only letters, digits, "_", "-" and "=", as list-mail-folder-messages and get-mail-message return it.';

describe('the refusals of a wrong param', () => {
  it('name the param of a message id that is not an id, for each command', async () => {
    for (const [name, params] of [
      ['move-mail-message', { messageId: 'm/1', to: 'archive' }],
      ['mark-mail-message', { messageId: 'm/1', read: 'true' }],
      ['set-mail-flag', { messageId: 'm/1', status: 'flagged' }],
      ['set-mail-categories', { messageId: 'm/1', categories: 'Red' }],
      ['delete-mail-draft', { messageId: 'm/1' }],
    ] as const) {
      expect([name, await refusal(name, params)]).toEqual([name, { code: 'invalid_id', source: 'validation', message: `--message-id: "m/1" ${NOT_AN_ID}` }]);
    }
  });

  it('quote at most 40 characters of a bad id', async () => {
    const long = `${'a'.repeat(39)}/${'b'.repeat(60)}`;
    expect(await refusal('mark-mail-message', { messageId: long, read: 'true' })).toMatchObject({ message: `--message-id: "${long.slice(0, 40)}" ${NOT_AN_ID}` });
  });

  it('say which required param is missing, for each command', async () => {
    for (const [name, params] of [
      ['move-mail-message', { messageId: 'm1' }],
      ['move-mail-thread', { conversationId: 'c1', to: 'archive' }],
      ['mark-mail-message', { messageId: 'm1' }],
      ['set-mail-flag', { messageId: 'm1' }],
      ['set-mail-categories', { messageId: 'm1' }],
      ['delete-mail-draft', {}],
    ] as const) {
      expect([name, await refusal(name, params)]).toEqual([name, { source: 'validation', message: expect.stringContaining('--') }]);
    }
  });

  it('name the param of an empty folder', async () => {
    expect(await refusal('move-mail-thread', { conversationId: 'c1', from: 'inbox', to: ' ' })).toEqual({
      code: 'folder_not_found',
      source: 'validation',
      message: '--to: name a mail folder.',
    });
    expect(await refusal('move-mail-thread', { conversationId: 'c1', from: ' ', to: 'archive' })).toEqual({
      code: 'folder_not_found',
      source: 'validation',
      message: '--from: name a mail folder.',
    });
  });

  it('name the folder of a path with no name in it', async () => {
    expect(await refusal('move-mail-message', { messageId: 'm1', to: '/' })).toEqual({
      code: 'folder_not_found',
      source: 'validation',
      message: 'No folder named "" in "Top of Information Store". Run list-mail-folders or list-mail-child-folders to see the folder names, or give the folder id.',
    });
  });

  it('give the refused target and the blind guard as wrong inputs', async () => {
    expect(await refusal('move-mail-thread', { conversationId: 'c1', from: 'inbox', to: 'outbox' })).toMatchObject({ code: 'folder_refused', source: 'validation' });
    const mailbox = createMailStoreFake(SEED);
    for (const name of ['deleteditems', 'recoverableitemsroot', 'junkemail', 'outbox'])
      mailbox.failures.set(`GET /me/mailFolders/${name}`, { type: 'api_error', status: 404, message: 'x' });
    expect(await refusal('move-mail-message', { messageId: 'm1', to: 'archive' }, mailbox.graph)).toMatchObject({ code: 'folder_guard_failed', source: 'validation' });
  });

  it('give an empty batch and a batch too large as wrong inputs', async () => {
    expect(await refusal('mark-mail-message', { messageId: ',', read: 'true' })).toEqual({
      code: 'invalid_batch',
      source: 'validation',
      message: 'Give the messages: --message-id <id> (ids separated by commas) or --ids-file <file> (one id on each line).',
    });
    const files = createLocalFilesFake({ 'ids.txt': 'm1\n  \nx/y\n' });
    expect((await runWrite(createMailStoreFake(SEED).graph, 'mark-mail-message', { idsFile: 'ids.txt', read: 'true' }, { files }))['error']).toEqual({
      code: 'invalid_id',
      source: 'validation',
      message: '--ids-file line 3 is not a Graph id. Give one message id on each line, as list-mail-folder-messages returns it.',
    });
  });

  it('refuse a due date with another status as a wrong input, and take notFlagged', async () => {
    expect(await refusal('set-mail-flag', { messageId: 'm1', status: 'complete', due: '2026-11-20' })).toEqual({
      source: 'validation',
      message: '--due: a due date goes with --status flagged only.',
    });
    expect(await refusal('set-mail-flag', { messageId: 'm1', status: 'flagged', due: 'x2026-11-20' })).toMatchObject({ message: expect.stringContaining('YYYY-MM-DD') });
    const mailbox = createMailStoreFake([{ id: 'm2', parentFolderId: 'inbox-id', changeKey: 'ck', flag: { flagStatus: 'flagged' } }]);
    const run = await runWrite(mailbox.graph, 'set-mail-flag', { messageId: 'm2', status: 'notFlagged' });
    expect((run['changes'] as ReadonlyArray<{ before: { flag: string } }>)[0]?.before.flag).toBe('flagged');
    expect(mailbox.messages.get('m2')?.['flag']).toEqual({ flagStatus: 'notFlagged' });
  });
});

describe('the refusals of an answer of Graph', () => {
  const answering = (graph: WriteGraph, match: string, value: unknown): WriteGraph => ({ ...graph, get: async (path) => (path.includes(match) ? ok(value) : graph.get(path)) });

  it('give a message, a folder, a listing and a conversation that Graph gives in an odd shape as validation errors', async () => {
    const graph = createMailStoreFake(SEED).graph;
    expect(await refusal('mark-mail-message', { messageId: 'm1', read: 'true' }, answering(graph, '/me/messages/m1?', { id: 'm1' }))).toEqual({
      source: 'validation',
      message: 'Graph gave the message m1 in a shape that this CLI cannot read.',
    });
    expect(await refusal('move-mail-message', { messageId: 'm1', to: 'archive' }, answering(graph, '/me/mailFolders/archive?', {}))).toEqual({
      source: 'validation',
      message: 'Graph gave a mail folder in a shape that this CLI cannot read.',
    });
    expect(await refusal('move-mail-message', { messageId: 'm1', to: 'archive' }, answering(graph, '/me/mailFolders/archive?', { id: 'a/b', displayName: 'A' }))).toEqual({
      code: 'invalid_id',
      source: 'validation',
      message: `mail folder: "a/b" ${NOT_AN_ID}`,
    });
    expect(await refusal('move-mail-thread', { conversationId: 'c1', from: 'inbox', to: 'archive' }, answering(graph, '$filter', { items: [] }))).toEqual({
      source: 'validation',
      message: 'Graph gave the messages of the conversation in a shape that this CLI cannot read.',
    });
    expect(await refusal('move-mail-thread', { conversationId: 'c1', from: 'inbox', to: 'archive' }, answering(graph, '$filter', { value: [{ id: 'a/b' }] }))).toEqual({
      code: 'invalid_id',
      source: 'validation',
      message: `message: "a/b" ${NOT_AN_ID}`,
    });
    expect(
      await refusal('move-mail-thread', { conversationId: 'c1', from: 'inbox', to: 'archive' }, answering(graph, '$filter', { value: [], '@odata.nextLink': 'x' }))
    ).toMatchObject({
      code: 'invalid_batch',
      source: 'validation',
    });
  });

  it('give a moved message with no new id, or with a new id that is not an id, as validation errors', async () => {
    const mailbox = createMailStoreFake(SEED);
    expect(await refusal('move-mail-message', { messageId: 'm1', to: 'archive' }, { ...mailbox.graph, post: async () => ok({}) })).toEqual({
      source: 'validation',
      message: 'Graph moved the message but gave no new id.',
    });
    expect(await refusal('move-mail-message', { messageId: 'm1', to: 'archive' }, { ...mailbox.graph, post: async () => ok({ id: 'a/b', parentFolderId: 'archive-id' }) })).toEqual(
      {
        code: 'invalid_id',
        source: 'validation',
        message: `moved message: "a/b" ${NOT_AN_ID}`,
      }
    );
  });

  it('pass on the error of a move of a draft to Deleted Items, and of the listing of a conversation', async () => {
    const mailbox = createMailStoreFake(SEED);
    mailbox.failures.set('POST /me/messages/d1/move', { type: 'api_error', status: 503, message: 'Busy.' });
    mailbox.failures.set('GET /me/mailFolders/inbox-id/messages', { type: 'api_error', status: 503, message: 'Busy too.' });
    expect(await refusal('delete-mail-draft', { messageId: 'd1' }, mailbox.graph)).toMatchObject({ message: 'Busy.' });
    expect(await refusal('move-mail-thread', { conversationId: 'c1', from: 'inbox', to: 'archive' }, mailbox.graph)).toMatchObject({ message: 'Busy too.' });
  });
});

describe('the refusals of the undo of a journal that a person edited', () => {
  const graph = createMailStoreFake(SEED).graph;

  it('are validation errors, for a move and for a state field', async () => {
    const bad = { messageId: '../x', folderId: 'a', backTo: 'b', field: 'isRead', before: false, set: true };
    expect(await MOVE_UNDO_HANDLERS['move-message-back']?.check(graph, bad)).toEqual({
      ok: false,
      error: { type: 'validation_error', message: 'The run journal names an id that is not a Graph id, so this undo does not run.' },
    });
    expect(await FIELD_UNDO_HANDLERS['restore-message-field']?.check(graph, bad)).toEqual({
      ok: false,
      error: {
        type: 'validation_error',
        message: 'The run journal names a message id that is not a Graph id, or a field that no command changes, so this undo does not run.',
      },
    });
  });

  it('read a check answer that is not a record as a change', async () => {
    const odd = { ...graph, get: async () => ok('nothing') };
    expect(await FIELD_UNDO_HANDLERS['restore-message-field']?.check(odd, { messageId: 'm1', field: 'isRead', before: false, set: true })).toEqual(ok('changed'));
  });
});

describe('the reads of the mail state commands', () => {
  it('select the fields that each command changes, and no other', async () => {
    for (const [name, params, select] of [
      ['move-mail-message', { messageId: 'm1', to: 'archive' }, 'parentFolderId,changeKey'],
      ['set-mail-flag', { messageId: 'm1', status: 'flagged' }, 'parentFolderId,changeKey,flag'],
      ['set-mail-categories', { messageId: 'm1', categories: 'Red' }, 'parentFolderId,changeKey,categories'],
      ['delete-mail-draft', { messageId: 'd1' }, 'parentFolderId,changeKey,isDraft'],
    ] as const) {
      const mailbox = createMailStoreFake(SEED);
      await runWrite(mailbox.graph, name, params, { dryRun: true });
      expect([name, mailbox.calls.filter((call) => call.startsWith('GET /me/messages/'))]).toEqual([name, [`GET /me/messages/${params.messageId}?$select=${select}`]]);
    }
  });

  it('read no folder whose id is not an id, and no parent of a folder that has none', async () => {
    const mailbox = createMailStoreFake([{ id: 'm1', parentFolderId: 'a/b', changeKey: 'ck' }]);
    const graph = {
      ...mailbox.graph,
      get: async (path: string) => (path.startsWith('/me/mailFolders/archive?') ? ok({ id: 'archive-id', displayName: 'Archive' }) : mailbox.graph.get(path)),
    };
    await runWrite(graph, 'move-mail-message', { messageId: 'm1', to: 'archive' }, { dryRun: true });
    expect(mailbox.calls.filter((call) => call.startsWith('GET /me/mailFolders/a/') || call.startsWith('GET /me/mailFolders/?'))).toEqual([]);
  });

  it('pass over an entry of a folder listing that is not a folder', async () => {
    const mailbox = createMailStoreFake(SEED);
    const reports = { id: 'reports-id', displayName: 'Reports', parentFolderId: 'root-id' };
    const graph = { ...mailbox.graph, get: async (path: string) => (path.includes('/childFolders') ? ok({ value: ['odd', { id: 'x-id' }, reports] }) : mailbox.graph.get(path)) };
    await runWrite(graph, 'move-mail-message', { messageId: 'm1', to: 'Reports' });
    expect([...mailbox.messages.values()].find((m) => m.id.startsWith('m1'))?.parentFolderId).toBe('reports-id');
  });

  it('read a name with a space as a name at the top, never as an id', async () => {
    const mailbox = createMailStoreFake(SEED);
    await runWrite(mailbox.graph, 'move-mail-message', { messageId: 'm1', to: 'Sent Items' }, { dryRun: true });
    expect(mailbox.calls.filter((call) => call.startsWith('GET /me/mailFolders/Sent'))).toEqual([]);
  });

  it('show the flag of before, and read a flag with no status as not flagged', async () => {
    const mailbox = createMailStoreFake([{ id: 'm2', parentFolderId: 'inbox-id', changeKey: 'ck', flag: {} }]);
    const dry = await runWrite(mailbox.graph, 'set-mail-flag', { messageId: 'm2', status: 'flagged' }, { dryRun: true });
    expect((dry['changes'] as ReadonlyArray<{ before: { flag: string } }>)[0]?.before.flag).toBe('notFlagged');
  });

  it('pass over the empty and blank lines of --ids-file, and add no other id', async () => {
    const files = createLocalFilesFake({ 'ids.txt': '\n  \nm1\n\t\n' });
    const run = await runWrite(createMailStoreFake(SEED).graph, 'mark-mail-message', { idsFile: 'ids.txt', read: 'true' }, { files });
    expect([run['done'], run['skipped']]).toEqual([1, []]);
  });
});

describe('the docs of the mail state commands', () => {
  it('say what each undo does and what each write sends', () => {
    const FIELD = 'gives each message the value of before, only if the message still has the value that the run set';
    const MOVE = 'moves each message back to the folder it was in (the message gets a new id again), only if it is still in the folder that the run put it in';
    const TO = "{ destinationId: '<the id of the folder {to}>' }";
    const docs = Object.fromEntries(
      ['move-mail-message', 'move-mail-thread', 'mark-mail-message', 'set-mail-flag', 'set-mail-categories', 'delete-mail-draft'].map((name) => [
        name,
        [commands[name]?.meta.undo, commands[name]?.meta.bodyTemplate],
      ])
    );
    expect(docs).toEqual({
      'move-mail-message': [MOVE, TO],
      'move-mail-thread': [MOVE, TO],
      'mark-mail-message': [FIELD, '{ isRead: {read} }'],
      'set-mail-flag': [FIELD, expect.stringContaining("flagStatus: '{status}'")],
      'set-mail-categories': [FIELD, "{ categories: ['{categories}', ...] }"],
      'delete-mail-draft': ['moves the draft back to Drafts (it gets a new id again), only if it is still a draft in Deleted Items', "{ destinationId: 'deleteditems' }"],
    });
  });
});
