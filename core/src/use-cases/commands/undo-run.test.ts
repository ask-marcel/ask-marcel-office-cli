import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { Result } from '../../domain/result.ts';
import type { RunError } from '../../domain/run-id.ts';
import type { WriteGraph } from '../ports/graph.ts';
import type { RunJournal } from '../ports/run-journal.ts';
import { createRunJournalFake } from '../../test-helpers/run-journal-fake.ts';
import { fakeWriteGraph } from '../../test-helpers/write-graph-fake.ts';
import type { RunRecord } from './record-run.ts';
import type { UndoHandlers, UndoOutcome, UndoReport, UndoStatus } from './undo-run.ts';
import { undoRun } from './undo-run.ts';

/*
 * The undo of a recorded run, on made-up widgets: the handler of `delete-widget`
 * deletes a widget the run made, only if the widget still has the version the
 * run left. The widgets live in a map that the fake graph reads and deletes.
 */

const RUN = '20261118T090507Z-0a1b2c3d';
const NOW = new Date('2026-11-20T10:00:00.000Z');

const made = (id: string, version: string): RunRecord['changes'][number] => ({
  action: 'create-widget',
  target: '/widgets',
  before: {},
  after: { widgetId: id },
  undo: { kind: 'delete-widget', data: { widgetId: id, version } },
});

const record = (changes: RunRecord['changes']): RunRecord => ({ version: 1, runId: RUN, command: 'save-widget', startedAt: '2026-11-18T09:05:07.250Z', changes });

type Widgets = { readonly graph: WriteGraph; readonly items: Map<string, string>; readonly deletes: string[] };

const widgets = (items: Record<string, string>, overrides: Partial<WriteGraph> = {}): Widgets => {
  const store = new Map(Object.entries(items));
  const deletes: string[] = [];
  const graph = fakeWriteGraph({
    get: async (path) => {
      const version = store.get(path.replace('/widgets/', ''));
      return version === undefined ? err({ type: 'api_error', status: 404, message: 'Not found.' }) : ok({ version });
    },
    delete: async (path) => {
      deletes.push(path);
      store.delete(path.replace('/widgets/', ''));
      return ok({ ok: true });
    },
    ...overrides,
  });
  return { graph, items: store, deletes };
};

const HANDLERS: UndoHandlers = {
  'delete-widget': {
    check: async (reads, data) => {
      const read = await reads.get(`/widgets/${String(data['widgetId'])}`);
      if (!read.ok) return read.error.type === 'api_error' && read.error.status === 404 ? ok('missing') : read;
      return ok((read.value as { version: string }).version === data['version'] ? 'ready' : 'changed');
    },
    apply: async (graph, data) => graph.delete(`/widgets/${String(data['widgetId'])}`),
  },
};

const seeded = (changes: RunRecord['changes']): ReturnType<typeof createRunJournalFake> => {
  const journal = createRunJournalFake();
  journal.records.set(RUN, record(changes));
  return journal;
};

const undo = async (journal: RunJournal, graph: WriteGraph, apply: boolean, runId = RUN, now = NOW): Promise<Result<UndoReport, RunError>> =>
  undoRun({ journal, graph, handlers: HANDLERS, now: () => now }, runId, apply);

const outcomeOf = (id: string, status: UndoStatus, reason?: string): UndoOutcome => ({
  action: 'create-widget',
  target: '/widgets',
  after: { widgetId: id },
  undo: 'delete-widget',
  status,
  ...(reason === undefined ? {} : { reason }),
});

const CHANGED = 'The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.';

describe('the undo preview of a run', () => {
  it('shows each change that the undo would reverse, and changes nothing', async () => {
    const journal = seeded([made('w1', 'v1')]);
    const { graph, deletes } = widgets({ w1: 'v1' });
    expect(await undo(journal, graph, false)).toEqual(
      ok({ runId: RUN, command: 'save-widget', startedAt: '2026-11-18T09:05:07.250Z', applied: false, changes: [outcomeOf('w1', 'will-undo')] })
    );
    expect(deletes).toEqual([]);
    expect(journal.records.get(RUN)).toEqual(record([made('w1', 'v1')]));
  });

  it('says which items changed after the run or are gone, so the undo passes over them', async () => {
    const { graph } = widgets({ w1: 'v2' });
    const result = await undo(seeded([made('w1', 'v1'), made('w2', 'v1')]), graph, false);
    expect(result.ok ? result.value.changes : []).toEqual([
      outcomeOf('w1', 'changed', CHANGED),
      outcomeOf('w2', 'missing', 'The item is no longer there, so there is nothing to undo.'),
    ]);
  });
});

describe('the undo of a run with --apply', () => {
  it('reverses each change, last first, and records the undo in the journal', async () => {
    const journal = seeded([made('w1', 'v1'), made('w2', 'v1')]);
    const { graph, deletes } = widgets({ w1: 'v1', w2: 'v1' });
    const result = await undo(journal, graph, true);
    expect(result.ok ? result.value.changes : []).toEqual([outcomeOf('w1', 'undone'), outcomeOf('w2', 'undone')]);
    expect(result.ok ? result.value.applied : false).toBe(true);
    expect(deletes).toEqual(['/widgets/w2', '/widgets/w1']);
    expect(journal.records.get(RUN)).toEqual(
      record([
        { ...made('w1', 'v1'), undoneAt: NOW.toISOString() },
        { ...made('w2', 'v1'), undoneAt: NOW.toISOString() },
      ])
    );
  });

  it('refuses a second undo of a run that is fully undone', async () => {
    const journal = seeded([made('w1', 'v1')]);
    const { graph, deletes } = widgets({ w1: 'v1' });
    await undo(journal, graph, true);
    for (const apply of [false, true]) {
      expect(await undo(journal, graph, apply)).toEqual(err({ code: 'run_already_undone', message: `Run ${RUN} is already fully undone: there is nothing left to undo.` }));
    }
    expect(deletes).toEqual(['/widgets/w1']);
  });

  it('passes over an item that changed after the run, reverses the others, and lets a later undo try the rest again', async () => {
    const journal = seeded([made('w1', 'v1'), made('w2', 'v1')]);
    const { graph, deletes } = widgets({ w1: 'v2', w2: 'v1' });
    const first = await undo(journal, graph, true);
    expect(first.ok ? first.value.changes : []).toEqual([outcomeOf('w1', 'changed', CHANGED), outcomeOf('w2', 'undone')]);
    expect(deletes).toEqual(['/widgets/w2']);
    const second = await undo(journal, graph, false);
    expect(second.ok ? second.value.changes : []).toEqual([outcomeOf('w1', 'changed', CHANGED), outcomeOf('w2', 'already-undone')]);
  });

  it('records nothing when no change was reversed', async () => {
    const journal = seeded([made('w1', 'v1')]);
    await undo(journal, widgets({}).graph, true);
    expect(journal.records.get(RUN)).toEqual(record([made('w1', 'v1')]));
  });

  it('reports the failure of the check or of the reversal of a change, and a kind it has no handler for', async () => {
    const failing = widgets({ w1: 'v1' }, { delete: async () => err({ type: 'api_error', status: 403, message: 'Access is denied.' }) });
    const foreign = { ...made('w2', 'v1'), undo: { kind: 'delete-gadget', data: {} } };
    const result = await undo(seeded([made('w1', 'v1'), foreign]), failing.graph, true);
    expect(result.ok ? result.value.changes : []).toEqual([
      outcomeOf('w1', 'failed', 'Access is denied.'),
      { ...outcomeOf('w2', 'failed', 'This version of the CLI has no undo for "delete-gadget".'), undo: 'delete-gadget' },
    ]);
    const unreadable = widgets({}, { get: async () => err({ type: 'network_error', message: 'The network is down.' }) });
    const checked = await undo(seeded([made('w1', 'v1')]), unreadable.graph, true);
    expect(checked.ok ? checked.value.changes : []).toEqual([outcomeOf('w1', 'failed', 'The network is down.')]);
  });

  it('runs each undo on the write view of the graph, so an undo that would send mail is refused before any request', async () => {
    const { graph, deletes } = widgets({ w1: 'v1' });
    let posted = false;
    const sending: WriteGraph = {
      ...graph,
      post: async () => {
        posted = true;
        return ok({});
      },
    };
    const handlers: UndoHandlers = { 'delete-widget': { check: async () => ok('ready'), apply: async (write) => write.post('/me/messages/w1/send', {}) } };
    const result = await undoRun({ journal: seeded([made('w1', 'v1')]), graph: sending, handlers, now: () => NOW }, RUN, true);
    expect(result.ok ? result.value.changes.map((change) => [change.status, change.reason]) : []).toEqual([
      ['failed', 'POST /me/messages/w1/send was not sent: it would send mail, and this CLI never sends mail. A draft stays in Drafts until you send it yourself from Outlook.'],
    ]);
    expect([posted, deletes]).toEqual([false, []]);
  });

  it('does not take a handler from the prototype of the handler table', async () => {
    const odd = { ...made('w1', 'v1'), undo: { kind: 'toString', data: {} } };
    const result = await undo(seeded([odd]), widgets({}).graph, true);
    expect(result.ok ? result.value.changes.map((change) => change.status) : []).toEqual(['failed']);
  });

  it('still answers the undo when the journal cannot record it, with a warning', async () => {
    const journal = seeded([made('w1', 'v1')]);
    const failingSave: RunJournal = { ...journal, save: async () => err({ type: 'io_failed', message: 'EACCES: permission denied' }) };
    const result = await undo(failingSave, widgets({ w1: 'v1' }).graph, true);
    expect(result.ok ? result.value.warning : undefined).toBe('The undo is done, but the run journal could not record it: EACCES: permission denied');
  });
});

describe('the run ids that the undo refuses', () => {
  it('refuses a run id that is not one, before it reads the journal', async () => {
    const result = await undo(seeded([made('w1', 'v1')]), widgets({}).graph, false, '../token-cache');
    expect(result.ok ? '' : result.error.code).toBe('invalid_run_id');
  });

  it('refuses a run older than 30 days, even while its record is still in the journal, and accepts one of exactly 30 days', async () => {
    const journal = seeded([made('w1', 'v1')]);
    const { graph } = widgets({ w1: 'v1' });
    expect(await undo(journal, graph, false, RUN, new Date('2026-12-18T09:05:08.000Z'))).toEqual(
      err({ code: 'run_expired', message: `Run ${RUN} is older than 30 days: the journal keeps a run 30 days only, so it cannot be undone.` })
    );
    expect((await undo(journal, graph, false, RUN, new Date('2026-12-18T09:05:07.000Z'))).ok).toBe(true);
  });

  it('says that a run is not in the journal', async () => {
    expect(await undo(createRunJournalFake(), widgets({}).graph, false)).toEqual(
      err({ code: 'run_not_found', message: `The run journal has no run ${RUN}. Copy the runId field of the write output: a dry-run has no run id.` })
    );
  });

  it('says that the journal could not be read', async () => {
    const broken: RunJournal = { ...createRunJournalFake(), load: async () => err({ type: 'io_failed', message: 'EIO: i/o error' }) };
    expect(await undo(broken, widgets({}).graph, false)).toEqual(err({ code: 'journal_failed', message: 'The run journal could not be read: EIO: i/o error' }));
  });

  for (const [name, raw] of [
    ['a record of another shape', { version: 2 }],
    ['a record of another run', { ...record([made('w1', 'v1')]), runId: '20261118T090507Z-ffffffff' }],
    ['a record with no change', record([])],
  ] as const) {
    it(`refuses ${name} as unreadable`, async () => {
      const journal = createRunJournalFake();
      journal.records.set(RUN, raw);
      expect(await undo(journal, widgets({}).graph, false)).toEqual(
        err({ code: 'run_unreadable', message: `The journal record of run ${RUN} is not readable, so the run cannot be undone.` })
      );
    });
  }
});
