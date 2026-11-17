import { err, ok } from '../domain/result.ts';
import { runIdTime } from '../domain/run-id.ts';
import type { RunId } from '../domain/run-id.ts';
import type { RunJournal } from '../use-cases/ports/run-journal.ts';

/**
 * An in-memory run journal. `records` holds each saved record by run id (a
 * test seeds or reads it), `cutoffs` each cutoff the code asked to remove
 * before, and `failSaves` makes each save answer `io_failed` with that text.
 */
export type RunJournalFake = RunJournal & {
  readonly records: Map<string, unknown>;
  readonly cutoffs: Date[];
};

export const createRunJournalFake = (options: { readonly failSaves?: string } = {}): RunJournalFake => {
  const records = new Map<string, unknown>();
  const cutoffs: Date[] = [];
  return {
    records,
    cutoffs,
    save: async (runId, record) => {
      if (options.failSaves !== undefined) return err({ type: 'io_failed', message: options.failSaves });
      records.set(runId, structuredClone(record));
      return ok(undefined);
    },
    load: async (runId) => (records.has(runId) ? ok(structuredClone(records.get(runId))) : err({ type: 'not_found' })),
    removeOlderThan: async (cutoff) => {
      cutoffs.push(cutoff);
      for (const runId of [...records.keys()]) if (runIdTime(runId as RunId) < cutoff) records.delete(runId);
      return ok(undefined);
    },
  };
};
