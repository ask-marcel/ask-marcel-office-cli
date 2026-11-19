import { describe, expect, it } from 'bun:test';
import type { RunRecorder } from '@ask-marcel/office-core/use-cases/commands/record-run.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { RunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createMailboxFake } from '../test-helpers/mailbox-fake.ts';
import type { MailboxFake } from '../test-helpers/mailbox-fake.ts';
import { buildWriteCli } from './write-program.ts';

/*
 * The full cycle of a draft command on the write CLI, with a fake mailbox and a
 * fake run journal: the dry-run, the run with its run id, the undo preview, the
 * undo with --apply, the refusal of a second undo, the pass-over of a draft that
 * changed after the run, and the removal of the runs older than 30 days.
 */

const captureStdout = async (run: () => Promise<unknown>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run();
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

type Envelope = { readonly ok: boolean; readonly data?: Record<string, unknown>; readonly error?: string; readonly errorCode?: string };

type Bench = { readonly mailbox: MailboxFake; readonly journal: RunJournalFake; readonly cli: (...args: string[]) => Promise<Envelope> };

const NOW = new Date('2026-11-18T09:05:07.000Z');

const bench = (): Bench => {
  const mailbox = createMailboxFake();
  const journal = createRunJournalFake();
  let suffix = 0;
  const runs: RunRecorder = { journal, now: () => NOW, randomSuffix: () => (++suffix).toString(16).padStart(8, '0') };
  const cli = async (...args: string[]): Promise<Envelope> => {
    const program = buildWriteCli({ graph: mailbox.graph, logger: createLoggerFake(), fs: createFileSystemFake(), runs, onCommandError: () => undefined });
    const out = await captureStdout(async () => program.parseAsync(['node', 'ask-marcel-office-write', '--output', 'json', ...args]).catch(() => undefined));
    return JSON.parse(out) as Envelope;
  };
  return { mailbox, journal, cli };
};

const CREATE = ['create-mail-draft', '--subject', 'Agenda', '--body-content', 'See below.', '--to-recipients', 'robin.chen@contoso.com'];

const RUN_1 = '20261118T090507Z-00000001';

describe('the cycle of a draft command: dry-run, run, undo', () => {
  it('shows the draft it would make on a dry-run, and makes nothing', async () => {
    const { mailbox, journal, cli } = bench();
    expect(await cli(...CREATE, '--dry-run')).toEqual({
      ok: true,
      data: { dryRun: true, command: 'create-mail-draft', changes: [{ action: 'create-draft', target: '/me/messages', before: {} }] },
    });
    expect([mailbox.calls, mailbox.messages.size, journal.records.size]).toEqual([[], 0, 0]);
  });

  it('makes the draft with a run id, shows its undo, does it, and refuses a second undo', async () => {
    const { mailbox, journal, cli } = bench();
    const made = await cli(...CREATE);
    expect(made.data).toMatchObject({
      id: 'draft-1',
      runId: RUN_1,
      changes: [{ action: 'create-draft', target: '/me/messages', before: {}, after: { messageId: 'draft-1', changeKey: 'draft-1-v1' } }],
    });
    expect(journal.records.has(RUN_1)).toBe(true);

    const preview = await cli('undo', RUN_1);
    expect(preview.data).toMatchObject({ runId: RUN_1, command: 'create-mail-draft', applied: false, changes: [{ undo: 'delete-draft', status: 'will-undo' }] });
    expect(mailbox.messages.has('draft-1')).toBe(true);

    const undone = await cli('undo', RUN_1, '--apply');
    expect(undone.data).toMatchObject({ applied: true, changes: [{ status: 'undone' }] });
    expect(mailbox.calls.at(-1)).toBe('DELETE /me/messages/draft-1');
    expect(mailbox.messages.has('draft-1')).toBe(false);

    expect(await cli('undo', RUN_1, '--apply')).toMatchObject({ ok: false, errorCode: 'run_already_undone' });
  });

  it('passes over a draft that changed after the run, and keeps it', async () => {
    const { mailbox, cli } = bench();
    await cli(...CREATE);
    mailbox.edit('draft-1', { subject: 'Agenda (edited in Outlook)' });
    const undone = await cli('undo', RUN_1, '--apply');
    expect(undone.data).toMatchObject({ changes: [{ status: 'changed' }] });
    expect(mailbox.messages.get('draft-1')?.subject).toBe('Agenda (edited in Outlook)');
    expect(mailbox.calls.filter((call) => call.startsWith('DELETE'))).toEqual([]);
  });

  it('passes over a message that is no longer a draft', async () => {
    const { mailbox, cli } = bench();
    await cli(...CREATE);
    mailbox.messages.set('draft-1', { ...mailbox.messages.get('draft-1'), id: 'draft-1', isDraft: false, changeKey: 'draft-1-v1' });
    expect((await cli('undo', RUN_1)).data).toMatchObject({ changes: [{ status: 'changed' }] });
  });

  it('removes the runs older than 30 days at each write, and refuses the undo of an expired run', async () => {
    const { journal, cli } = bench();
    journal.records.set('20261019T090506Z-0000ffff', { version: 1 });
    await cli(...CREATE);
    expect(journal.records.has('20261019T090506Z-0000ffff')).toBe(false);
    expect(await cli('undo', '20261019T090506Z-0000ffff')).toMatchObject({ ok: false, errorCode: 'run_expired' });
  });

  it('says that a run id is not one, or not in the journal', async () => {
    const { cli } = bench();
    expect(await cli('undo', '../token-cache')).toMatchObject({ ok: false, errorCode: 'invalid_run_id' });
    expect(await cli('undo', RUN_1)).toMatchObject({ ok: false, errorCode: 'run_not_found' });
  });

  it('gives --dry-run to the four draft commands, add-mail-draft-attachment, remove-mail-draft-attachment, the six mail state commands, the 14 To Do commands, the three calendar draft commands and the 14 Planner commands only', () => {
    const program = buildWriteCli({ graph: createMailboxFake().graph, logger: createLoggerFake(), fs: createFileSystemFake() });
    const withDryRun = program.commands.filter((command) => command.options.some((option) => option.long === '--dry-run')).map((command) => command.name());
    expect(withDryRun.toSorted((a, b) => a.localeCompare(b))).toEqual([
      'add-mail-draft-attachment',
      'add-planner-checklist-item',
      'add-planner-task-file',
      'add-planner-task-link',
      'add-todo-linked-resource',
      'add-todo-task-attachment',
      'add-todo-task-step',
      'apply-todo-ops',
      'complete-todo-task',
      'create-calendar-draft',
      'create-forward-draft',
      'create-mail-draft',
      'create-planner-bucket',
      'create-planner-task',
      'create-reply-draft',
      'create-todo-list',
      'create-todo-task',
      'delete-calendar-draft',
      'delete-mail-draft',
      'delete-planner-bucket',
      'delete-planner-task',
      'delete-todo-task',
      'delete-todo-task-step',
      'mark-mail-message',
      'move-mail-message',
      'move-mail-thread',
      'move-todo-task',
      'remove-mail-draft-attachment',
      'remove-planner-checklist-item',
      'remove-planner-task-link',
      'remove-todo-task-attachment',
      'set-mail-categories',
      'set-mail-flag',
      'update-calendar-draft',
      'update-mail-draft',
      'update-planner-bucket',
      'update-planner-checklist-item',
      'update-planner-plan',
      'update-planner-task',
      'update-planner-task-details',
      'update-todo-list',
      'update-todo-task',
      'update-todo-task-step',
    ]);
  });
});
