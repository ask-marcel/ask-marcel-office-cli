import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { createRunJournalFake } from '../../test-helpers/run-journal-fake.ts';
import { fakeWriteGraph } from '../../test-helpers/write-graph-fake.ts';
import type { RunRecord } from './record-run.ts';
import type { UndoHandlers } from './undo-run.ts';
import { undoRun } from './undo-run.ts';

/*
 * An undo that gives the item a new id (a mail message moved back to its
 * folder) says the new ids in its outcome, as `now`, so the user can find the
 * item again. An undo whose handler names no new ids has no `now`.
 */

const RUN = '20261118T090507Z-0a1b2c3d';

const moved: RunRecord['changes'][number] = {
  action: 'move-widget',
  target: '/widgets/w1/move',
  before: { widgetId: 'w1' },
  after: { widgetId: 'w2' },
  undo: { kind: 'move-widget-back', data: { widgetId: 'w2' } },
};

const undo = async (answer: unknown): Promise<unknown> => {
  const journal = createRunJournalFake();
  journal.records.set(RUN, { version: 1, runId: RUN, command: 'move-widget', startedAt: '2026-11-18T09:05:07.250Z', changes: [moved] });
  const handlers: UndoHandlers = { 'move-widget-back': { check: async () => ok('ready'), apply: async () => ok(answer) } };
  const report = await undoRun({ journal, graph: fakeWriteGraph(), handlers, now: () => new Date('2026-11-19T10:00:00.000Z') }, RUN, true);
  return report.ok ? report.value.changes : report.error;
};

describe('the new ids of an undo', () => {
  it('are in the outcome when the handler names them', async () => {
    expect(await undo({ now: { widgetId: 'w3' }, id: 'w3' })).toEqual([
      { action: 'move-widget', target: '/widgets/w1/move', after: { widgetId: 'w2' }, undo: 'move-widget-back', status: 'undone', now: { widgetId: 'w3' } },
    ]);
  });

  it('are left out when the handler names none, or names them in another shape', async () => {
    const plain = { action: 'move-widget', target: '/widgets/w1/move', after: { widgetId: 'w2' }, undo: 'move-widget-back', status: 'undone' };
    expect(await undo({ id: 'w2' })).toEqual([plain]);
    expect(await undo({ now: { widgetId: 3 } })).toEqual([plain]);
    expect(await undo(undefined)).toEqual([plain]);
  });
});
