import { describe, expect, it } from 'bun:test';
import { runRegistryCommand } from '@ask-marcel/office-core/composition/run-registry-command.ts';
import type { RunRegistryCommandFailure } from '@ask-marcel/office-core/composition/run-registry-command.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import type { RunRecorder } from '@ask-marcel/office-core/use-cases/commands/record-run.ts';
import { undoRun } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { UndoReport } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { RunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createMailboxFake } from '../../test-helpers/mailbox-fake.ts';
import type { Message } from '../../test-helpers/mailbox-fake.ts';
import { commands, undoHandlers } from './write-commands.ts';

/*
 * The dry-run and the undo of each draft command, run as the shells run them
 * (runRegistryCommand) on a fake mailbox, with the undo handlers of the package.
 */

const RUN = '20261118T090507Z-0a1b2c3d';
const NOW = new Date('2026-11-18T09:05:07.000Z');

const recorder = (journal: RunJournalFake = createRunJournalFake()): RunRecorder => ({ journal, now: () => NOW, randomSuffix: () => '0a1b2c3d' });

const run = async (
  graph: WriteGraph,
  name: string,
  params: Record<string, string>,
  extra: { readonly dryRun?: boolean; readonly runs?: RunRecorder } = {}
): Promise<Result<unknown, RunRegistryCommandFailure>> => {
  const command = commands[name];
  if (command === undefined) throw new Error(name);
  const deps = {
    graph,
    fs: createFileSystemFake(),
    commandPackage: { registry: commands, binName: 'ask-marcel-office-write' },
    ...(extra.runs === undefined ? {} : { runs: extra.runs }),
  };
  return runRegistryCommand(deps, { name, command, params, surface: 'cli', dryRun: extra.dryRun === true });
};

const undo = async (graph: WriteGraph, journal: RunJournalFake, apply: boolean): Promise<Result<UndoReport, unknown>> =>
  undoRun({ journal, graph, handlers: undoHandlers, now: () => NOW }, RUN, apply);

const SOURCE: Message = { id: 'msg-1', isDraft: false, changeKey: 'msg-1-v1', subject: 'Q3 planning' };

const changesOf = (result: Result<unknown, RunRegistryCommandFailure>): unknown => (result.ok ? (result.value as { changes: unknown }).changes : result.error);

describe('the dry-run of each draft command', () => {
  it('shows the reply it would make from the message, reply-all by default and to the sender with --reply-all false', async () => {
    const mailbox = createMailboxFake([SOURCE]);
    expect(changesOf(await run(mailbox.graph, 'create-reply-draft', { replyToMessageId: 'msg-1', comment: 'Yes.' }, { dryRun: true }))).toEqual([
      { action: 'create-reply-draft', target: '/me/messages/msg-1/createReplyAll', before: { messageId: 'msg-1' } },
    ]);
    expect(changesOf(await run(mailbox.graph, 'create-reply-draft', { replyToMessageId: 'msg-1', comment: 'Yes.', replyAll: 'false' }, { dryRun: true }))).toEqual([
      { action: 'create-reply-draft', target: '/me/messages/msg-1/createReply', before: { messageId: 'msg-1' } },
    ]);
    expect(mailbox.calls).toEqual([]);
  });

  it('shows the forward it would make from the message', async () => {
    const mailbox = createMailboxFake([SOURCE]);
    expect(
      changesOf(await run(mailbox.graph, 'create-forward-draft', { forwardMessageId: 'msg-1', toRecipients: 'kim.lee@contoso.com', comment: 'FYI.' }, { dryRun: true }))
    ).toEqual([{ action: 'create-forward-draft', target: '/me/messages/msg-1/createForward', before: { messageId: 'msg-1' } }]);
  });

  it('shows a draft made in the folder that --mail-folder-id names', async () => {
    const params = { subject: 'Agenda', bodyContent: 'Hi.', toRecipients: 'kim.lee@contoso.com', mailFolderId: 'folder-7' };
    expect(changesOf(await run(createMailboxFake().graph, 'create-mail-draft', params, { dryRun: true }))).toEqual([
      { action: 'create-draft', target: '/me/mailFolders/folder-7/messages', before: {} },
    ]);
  });

  it('reads the fields that an update would write, and shows the draft with its changeKey, with no write', async () => {
    const mailbox = createMailboxFake([{ id: 'draft-9', isDraft: true, changeKey: 'draft-9-v4', subject: 'Old' }]);
    expect(changesOf(await run(mailbox.graph, 'update-mail-draft', { messageId: 'draft-9', subject: 'New' }, { dryRun: true }))).toEqual([
      { action: 'update-draft', target: '/me/messages/draft-9', before: { messageId: 'draft-9', changeKey: 'draft-9-v4' } },
    ]);
    expect(mailbox.calls).toEqual(['GET /me/messages/draft-9?$select=subject,changeKey']);
  });

  it('refuses on a dry-run what the run refuses, before any request', async () => {
    const mailbox = createMailboxFake([SOURCE]);
    const html = { replyToMessageId: 'msg-1', comment: '<div id="divRplyFwdMsg">x</div>', bodyContentType: 'HTML' };
    const refusals = [
      await run(mailbox.graph, 'create-reply-draft', html, { dryRun: true }),
      await run(mailbox.graph, 'create-reply-draft', { comment: 'Yes.' }, { dryRun: true }),
      await run(
        mailbox.graph,
        'create-forward-draft',
        { forwardMessageId: 'msg-1', toRecipients: 'kim.lee@contoso.com', comment: html.comment, bodyContentType: 'HTML' },
        { dryRun: true }
      ),
      await run(mailbox.graph, 'create-forward-draft', { forwardMessageId: 'msg-1', comment: 'FYI.' }, { dryRun: true }),
      await run(mailbox.graph, 'create-mail-draft', { subject: 'Agenda' }, { dryRun: true }),
      await run(mailbox.graph, 'update-mail-draft', { messageId: 'draft-9' }, { dryRun: true }),
      await run(mailbox.graph, 'update-mail-draft', { subject: 'x' }, { dryRun: true }),
    ];
    expect(refusals.map((result) => (result.ok ? 'ran' : result.error.source))).toEqual([
      'validation',
      'validation',
      'validation',
      'validation',
      'validation',
      'validation',
      'validation',
    ]);
    expect(refusals.map((result) => (result.ok ? '' : result.error.message)).filter((message) => message.includes('boundary marker'))).toHaveLength(2);
    expect(mailbox.calls).toEqual([]);
  });

  it('passes on the failure of the read of an update', async () => {
    const result = await run(createMailboxFake().graph, 'update-mail-draft', { messageId: 'draft-9', subject: 'New' }, { dryRun: true });
    expect(result.ok ? '' : result.error.message).toBe('The specified object was not found in the store.');
  });
});

describe('the run and the undo of each draft command', () => {
  it('records a reply and moves it to Deleted Items on undo', async () => {
    const mailbox = createMailboxFake([SOURCE]);
    const journal = createRunJournalFake();
    const made = await run(mailbox.graph, 'create-reply-draft', { replyToMessageId: 'msg-1', comment: 'Yes.' }, { runs: recorder(journal) });
    expect(changesOf(made)).toEqual([
      { action: 'create-reply-draft', target: '/me/messages/msg-1/createReplyAll', before: { messageId: 'msg-1' }, after: { messageId: 'draft-1', changeKey: 'draft-1-v1' } },
    ]);
    expect(mailbox.calls).toEqual(['POST /me/messages/msg-1/createReplyAll', 'GET /me/messages/draft-1?$select=changeKey']);
    const undone = await undo(mailbox.graph, journal, true);
    expect(undone.ok ? undone.value.changes.map((change) => change.status) : []).toEqual(['undone']);
    expect(mailbox.calls.at(-1)).toBe('DELETE /me/messages/draft-1');
  });

  it('records a forward with the changeKey of its last write', async () => {
    const mailbox = createMailboxFake([SOURCE]);
    const made = await run(
      mailbox.graph,
      'create-forward-draft',
      { forwardMessageId: 'msg-1', toRecipients: 'kim.lee@contoso.com', comment: 'FYI.', subject: 'FW: Q3' },
      { runs: recorder() }
    );
    expect(changesOf(made)).toEqual([
      { action: 'create-forward-draft', target: '/me/messages/msg-1/createForward', before: { messageId: 'msg-1' }, after: { messageId: 'draft-1', changeKey: 'draft-1-v2' } },
    ]);
  });

  it('writes back on undo the fields that an update changed, as they were', async () => {
    const mailbox = createMailboxFake([{ id: 'draft-9', isDraft: true, changeKey: 'draft-9-v1', subject: 'Old', importance: 'Low' }]);
    const journal = createRunJournalFake();
    const made = await run(mailbox.graph, 'update-mail-draft', { messageId: 'draft-9', subject: 'New', importance: 'High' }, { runs: recorder(journal) });
    expect(changesOf(made)).toEqual([
      {
        action: 'update-draft',
        target: '/me/messages/draft-9',
        before: { messageId: 'draft-9', changeKey: 'draft-9-v1' },
        after: { messageId: 'draft-9', changeKey: 'draft-9-v2' },
      },
    ]);
    expect(mailbox.messages.get('draft-9')?.subject).toBe('New');
    const undone = await undo(mailbox.graph, journal, true);
    expect(undone.ok ? undone.value.changes.map((change) => change.status) : []).toEqual(['undone']);
    expect([mailbox.messages.get('draft-9')?.subject, mailbox.messages.get('draft-9')?.importance]).toEqual(['Old', 'Low']);
  });

  it('saves the body before a --comment update, which writes the body', async () => {
    const body = { contentType: 'html', content: '<div>old</div><div id="divRplyFwdMsg">quote</div>' };
    const mailbox = createMailboxFake([{ id: 'draft-9', isDraft: true, changeKey: 'draft-9-v1', body }]);
    await run(mailbox.graph, 'update-mail-draft', { messageId: 'draft-9', comment: 'new' }, { dryRun: true });
    expect(mailbox.calls).toEqual(['GET /me/messages/draft-9?$select=body,changeKey']);
  });

  it('records no run for a draft command whose answer names no draft (a 204)', async () => {
    const mailbox = createMailboxFake();
    const empty: WriteGraph = { ...mailbox.graph, post: async () => ok({ ok: true }) };
    const journal = createRunJournalFake();
    expect(await run(empty, 'create-mail-draft', { subject: 'Agenda', bodyContent: 'Hi.', toRecipients: 'kim.lee@contoso.com' }, { runs: recorder(journal) })).toEqual(
      ok({ ok: true, changes: [] })
    );
    expect(journal.records.size).toBe(0);
  });

  it('records an empty changeKey when it cannot read the draft back, so the undo passes over the draft', async () => {
    const mailbox = createMailboxFake();
    const blind: WriteGraph = { ...mailbox.graph, get: async () => err({ type: 'network_error', message: 'The network is down.' }) };
    const journal = createRunJournalFake();
    const made = await run(blind, 'create-mail-draft', { subject: 'Agenda', bodyContent: 'Hi.', toRecipients: 'kim.lee@contoso.com' }, { runs: recorder(journal) });
    expect(changesOf(made)).toEqual([{ action: 'create-draft', target: '/me/messages', before: {}, after: { messageId: 'draft-1', changeKey: '' } }]);
    const preview = await undo(mailbox.graph, journal, false);
    expect(preview.ok ? preview.value.changes.map((change) => change.status) : []).toEqual(['changed']);
  });

  it('saves only the fields that the message has, and nothing when its answer is not a message', async () => {
    const mailbox = createMailboxFake([{ id: 'draft-9', isDraft: true, changeKey: 'draft-9-v1', subject: 'Old' }]);
    const journal = createRunJournalFake();
    await run(mailbox.graph, 'update-mail-draft', { messageId: 'draft-9', subject: 'New', importance: 'High' }, { runs: recorder(journal) });
    const record = journal.records.get(RUN) as { changes: ReadonlyArray<{ undo: { data: { fields: Record<string, unknown> } } }> };
    expect(record.changes.map((change) => Object.keys(change.undo.data.fields))).toEqual([['subject']]);
    const odd: WriteGraph = { ...mailbox.graph, get: async () => ok('not a message') };
    const second = createRunJournalFake();
    await run(odd, 'update-mail-draft', { messageId: 'draft-9', subject: 'Newer' }, { runs: recorder(second) });
    const oddRecord = second.records.get(RUN) as typeof record;
    expect(oddRecord.changes.map((change) => change.undo.data.fields)).toEqual([{}]);
  });

  it('records an empty changeKey before an update when the message answers none', async () => {
    const mailbox = createMailboxFake([{ id: 'draft-9', isDraft: true, changeKey: 'draft-9-v1', subject: 'Old' }]);
    const keyless: WriteGraph = { ...mailbox.graph, get: async () => ok({ subject: 'Old' }) };
    expect(changesOf(await run(keyless, 'update-mail-draft', { messageId: 'draft-9', subject: 'New' }, { dryRun: true }))).toEqual([
      { action: 'update-draft', target: '/me/messages/draft-9', before: { messageId: 'draft-9', changeKey: '' } },
    ]);
  });

  it('passes on the failure of the write of a draft command, and records nothing', async () => {
    const mailbox = createMailboxFake();
    const journal = createRunJournalFake();
    const refused: WriteGraph = { ...mailbox.graph, patch: async () => err({ type: 'api_error', status: 403, message: 'Access is denied.' }) };
    const update = await run(refused, 'update-mail-draft', { messageId: 'draft-9', subject: 'New' }, { runs: recorder(journal) });
    expect(update.ok ? '' : update.error.message).toBe('The specified object was not found in the store.');
    const seeded = createMailboxFake([{ id: 'draft-9', isDraft: true, changeKey: 'draft-9-v1', subject: 'Old' }]);
    const denied = await run({ ...seeded.graph, patch: refused.patch }, 'update-mail-draft', { messageId: 'draft-9', subject: 'New' }, { runs: recorder(journal) });
    expect(denied.ok ? '' : denied.error.message).toBe('Access is denied.');
    const create = await run(
      { ...mailbox.graph, post: refused.patch as WriteGraph['post'] },
      'create-mail-draft',
      { subject: 'A', bodyContent: 'B', toRecipients: 'kim.lee@contoso.com' },
      { runs: recorder(journal) }
    );
    expect(create.ok ? '' : create.error.message).toBe('Access is denied.');
    expect(journal.records.size).toBe(0);
  });
});
