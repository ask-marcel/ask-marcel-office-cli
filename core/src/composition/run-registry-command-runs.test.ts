import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { err, ok } from '../domain/result.ts';
import type { Result } from '../domain/result.ts';
import type { WriteGraph } from '../infra/write-graph.ts';
import { writeCommandOf } from '../use-cases/commands/command-graph.ts';
import type { Command, WriteCommand } from '../use-cases/commands/command-types.ts';
import { withUnknownParamRejection } from '../use-cases/commands/reject-unknown-params.ts';
import type { AppliedChange, PlanWrite } from '../use-cases/commands/write-plan.ts';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createRunJournalFake } from '../test-helpers/run-journal-fake.ts';
import { fakeWriteGraph } from '../test-helpers/write-graph-fake.ts';
import { runRegistryCommand } from './run-registry-command.ts';
import type { RunRegistryCommandFailure } from './run-registry-command.ts';

/*
 * The dry-run and the run journal of a write with a plan (write-plan.ts), as
 * both shells run it: a dry-run makes no write, and a run that is not a dry-run
 * gets a run id, a record in the journal and the cleanup of the runs older than
 * 30 days.
 */

const MADE: AppliedChange = {
  action: 'create-draft',
  target: '/me/messages',
  before: {},
  after: { messageId: 'draft-1', changeKey: 'ck-1' },
  undo: { kind: 'delete-draft', data: { messageId: 'draft-1', changeKey: 'ck-1' } },
};

// A draft command whose plan reads the Drafts folder, and whose writes post the draft.
const savePlan: PlanWrite = async (reads, params) => {
  const folder = await reads.get('/me/mailFolders/drafts?$select=id');
  if (!folder.ok) return folder;
  return ok({
    changes: [{ action: 'create-draft', target: '/me/messages', before: {} }],
    apply: async (graph) => {
      const posted = await graph.post('/me/messages', { subject: params['subject'] });
      return posted.ok ? ok({ value: { id: 'draft-1', subject: params['subject'] }, changes: [MADE] }) : posted;
    },
  });
};

// `null` builds the command with no plan.
const draftCommand = (plan: PlanWrite | null = savePlan): Command<WriteGraph> => {
  const command: WriteCommand = {
    schema: z.object({ subject: z.string() }),
    execute: async () => ok({ executed: true }),
    meta: { ...fakeCommand().meta, graphMethod: 'POST', effect: 'draft' },
    ...(plan === null ? {} : { plan }),
  };
  return withUnknownParamRejection(writeCommandOf(command));
};

type Call = { readonly member: string; readonly path: string };

// A write graph that logs each call by member and path.
const loggingGraph = (overrides: Partial<WriteGraph> = {}): { readonly graph: WriteGraph; readonly calls: Call[] } => {
  const calls: Call[] = [];
  const base = fakeWriteGraph(overrides);
  const log = (member: string, path: string): void => {
    calls.push({ member, path });
  };
  const graph: WriteGraph = {
    get: async (path, headers) => {
      log('get', path);
      return base.get(path, headers);
    },
    getBinary: async (path) => {
      log('getBinary', path);
      return base.getBinary(path);
    },
    fetchUrl: async (url) => {
      log('fetchUrl', url);
      return base.fetchUrl(url);
    },
    post: async (path, body) => {
      log('post', path);
      return base.post(path, body);
    },
    patch: async (path, body, options) => {
      log('patch', path);
      return base.patch(path, body, options);
    },
    put: async (path, body, type) => {
      log('put', path);
      return base.put(path, body, type);
    },
    delete: async (path, options) => {
      log('delete', path);
      return base.delete(path, options);
    },
    uploadAttachment: async (target, file) => {
      log('uploadAttachment', target.kind);
      return base.uploadAttachment(target, file);
    },
  };
  return { graph, calls };
};

const NOW = new Date('2026-11-18T09:05:07.250Z');

const recorder = (
  journal = createRunJournalFake()
): { readonly journal: ReturnType<typeof createRunJournalFake>; readonly now: () => Date; readonly randomSuffix: () => string } => ({
  journal,
  now: () => NOW,
  randomSuffix: () => '0a1b2c3d',
});

const run = async (
  graph: WriteGraph,
  command: Command<WriteGraph>,
  extra: { readonly dryRun?: boolean; readonly runs?: ReturnType<typeof recorder>; readonly params?: Record<string, string> } = {}
): Promise<Result<unknown, RunRegistryCommandFailure>> =>
  runRegistryCommand(
    { graph, fs: createFileSystemFake(), commandPackage: fakeCommandPackage(), ...(extra.runs === undefined ? {} : { runs: extra.runs }) },
    { name: 'save-widget-draft', command, params: extra.params ?? { subject: 'Agenda' }, surface: 'cli', ...(extra.dryRun === undefined ? {} : { dryRun: extra.dryRun }) }
  );

describe('a dry-run of a write with a plan', () => {
  it('shows each change it would make, after the reads it needs, and sends no write', async () => {
    const { graph, calls } = loggingGraph();
    const runs = recorder();
    const result = await run(graph, draftCommand(), { dryRun: true, runs });
    expect(result).toEqual(ok({ dryRun: true, command: 'save-widget-draft', changes: [{ action: 'create-draft', target: '/me/messages', before: {} }] }));
    expect(calls).toEqual([{ member: 'get', path: '/me/mailFolders/drafts?$select=id' }]);
    expect(runs.journal.records.size).toBe(0);
  });

  it('gives the plan the read members only, so a plan that tries to write fails before any request', async () => {
    const { graph, calls } = loggingGraph();
    let members: ReadonlyArray<string> = [];
    const sneaky: PlanWrite = async (reads) => {
      members = Object.keys(reads).toSorted((a, b) => a.localeCompare(b));
      return ok({ changes: [], apply: async () => ok({ value: {}, changes: [] }) });
    };
    await run(graph, draftCommand(sneaky), { dryRun: true });
    expect(members).toEqual(['fetchUrl', 'get', 'getBinary']);
    expect(calls).toEqual([]);
  });

  it('passes on the failure of a read of the plan', async () => {
    const { graph } = loggingGraph({ get: async () => err({ type: 'api_error', status: 404, message: 'The folder is gone.' }) });
    expect(await run(graph, draftCommand(), { dryRun: true })).toEqual(err({ message: 'The folder is gone.', source: 'graph' }));
  });

  it('refuses an unknown parameter before the plan reads anything', async () => {
    const { graph, calls } = loggingGraph();
    const result = await run(graph, draftCommand(), { dryRun: true, params: { subject: 'Agenda', subjet: 'x' } });
    expect(result.ok ? '' : result.error.code).toBe('unknown_parameter');
    expect(calls).toEqual([]);
  });

  it('refuses a dry-run of a command that has no plan, and runs nothing', async () => {
    const { graph, calls } = loggingGraph();
    expect(await run(graph, draftCommand(null), { dryRun: true })).toEqual(
      err({
        message: '--dry-run: save-widget-draft has no dry-run and no undo. Only the commands whose docs show "Dry-run and undo" have them.',
        code: 'dry_run_unsupported',
        source: 'validation',
      })
    );
    expect(calls).toEqual([]);
  });
});

describe('a run of a write with a plan and a run journal', () => {
  it('makes the changes, answers the run id and the ids before and after, and records the run with its undo', async () => {
    const { graph, calls } = loggingGraph();
    const runs = recorder();
    const result = await run(graph, draftCommand(), { runs });
    expect(result).toEqual(
      ok({
        id: 'draft-1',
        subject: 'Agenda',
        runId: '20261118T090507Z-0a1b2c3d',
        changes: [{ action: 'create-draft', target: '/me/messages', before: {}, after: { messageId: 'draft-1', changeKey: 'ck-1' } }],
      })
    );
    expect(calls.map((call) => call.member)).toEqual(['get', 'post']);
    expect(runs.journal.records.get('20261118T090507Z-0a1b2c3d')).toEqual({
      version: 1,
      runId: '20261118T090507Z-0a1b2c3d',
      command: 'save-widget-draft',
      startedAt: '2026-11-18T09:05:07.250Z',
      changes: [MADE],
    });
  });

  it('removes the runs that started more than 30 days before this one', async () => {
    const runs = recorder();
    runs.journal.records.set('20261019T090506Z-00000000', { version: 1 });
    runs.journal.records.set('20261019T090508Z-00000000', { version: 1 });
    await run(loggingGraph().graph, draftCommand(), { runs });
    expect(runs.journal.cutoffs).toEqual([new Date('2026-10-19T09:05:07.250Z')]);
    expect([...runs.journal.records.keys()].toSorted((a, b) => a.localeCompare(b))).toEqual(['20261019T090508Z-00000000', '20261118T090507Z-0a1b2c3d']);
  });

  it('records nothing when a write of the plan fails', async () => {
    const runs = recorder();
    const result = await run(loggingGraph({ post: async () => err({ type: 'api_error', status: 403, message: 'Access is denied.' }) }).graph, draftCommand(), { runs });
    expect(result).toEqual(err({ message: 'Access is denied.', source: 'graph' }));
    expect([runs.journal.records.size, runs.journal.cutoffs.length]).toEqual([0, 0]);
  });

  it('still answers the change when the journal cannot be written, with a warning that the run has no undo', async () => {
    const runs = recorder(createRunJournalFake({ failSaves: 'EACCES: permission denied' }));
    const result = await run(loggingGraph().graph, draftCommand(), { runs });
    expect(result).toEqual(
      ok({
        id: 'draft-1',
        subject: 'Agenda',
        changes: [{ action: 'create-draft', target: '/me/messages', before: {}, after: { messageId: 'draft-1', changeKey: 'ck-1' } }],
        warning: 'The change is done, but the run journal could not be written, so this run has no undo: EACCES: permission denied',
      })
    );
  });

  it('gives no run id and records nothing when the plan has no change to make', async () => {
    const runs = recorder();
    const nothing: PlanWrite = async () => ok({ changes: [], apply: async () => ok({ value: { moved: 0 }, changes: [] }) });
    expect(await run(loggingGraph().graph, draftCommand(nothing), { runs })).toEqual(ok({ moved: 0, changes: [] }));
    expect([runs.journal.records.size, runs.journal.cutoffs.length]).toEqual([0, 0]);
  });

  for (const answer of [['a', 'b'], null, 'done']) {
    it(`puts the answer ${JSON.stringify(answer)}, which is not an object, under \`value\`, beside the run id`, async () => {
      const runs = recorder();
      const listing: PlanWrite = async () => ok({ changes: [], apply: async () => ok({ value: answer, changes: [MADE] }) });
      expect(await run(loggingGraph().graph, draftCommand(listing), { runs })).toEqual(
        ok({
          value: answer,
          runId: '20261118T090507Z-0a1b2c3d',
          changes: [{ action: 'create-draft', target: '/me/messages', before: {}, after: { messageId: 'draft-1', changeKey: 'ck-1' } }],
        })
      );
    });
  }

  it('records nothing and writes nothing when a read of the plan fails on a run', async () => {
    const runs = recorder();
    const { graph, calls } = loggingGraph({ get: async () => err({ type: 'api_error', status: 404, message: 'The folder is gone.' }) });
    expect(await run(graph, draftCommand(), { runs })).toEqual(err({ message: 'The folder is gone.', source: 'graph' }));
    expect(calls.map((call) => call.member)).toEqual(['get']);
    expect([runs.journal.records.size, runs.journal.cutoffs.length]).toEqual([0, 0]);
  });

  it('runs the plain command, with no run id, when no journal is given (a library caller)', async () => {
    const { graph, calls } = loggingGraph();
    expect(await run(graph, draftCommand())).toEqual(ok({ executed: true }));
    expect(calls).toEqual([]);
  });
});
