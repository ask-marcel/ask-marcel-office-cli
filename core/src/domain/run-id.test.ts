import { describe, expect, it } from 'bun:test';
import { newRunId, parseRunId, runIdTime } from './run-id.ts';
import type { RunId } from './run-id.ts';

describe('the id of a write run', () => {
  it('names a run by the UTC second it ran and a random suffix, so the ids sort by time', () => {
    expect(newRunId(new Date('2026-11-18T09:05:07.250Z'), '0a1b2c3d')).toBe(parseRunIdValue('20261118T090507Z-0a1b2c3d'));
  });

  it('gives back the second a run id names, which the 30-day limit reads', () => {
    expect(runIdTime(parseRunIdValue('20261118T090507Z-0a1b2c3d')).toISOString()).toBe('2026-11-18T09:05:07.000Z');
  });

  it('accepts a run id that a write printed', () => {
    expect(parseRunId('20261231T235959Z-ffffffff')).toEqual({ ok: true, value: parseRunIdValue('20261231T235959Z-ffffffff') });
  });

  // The id becomes a file name in the journal folder: a path is refused before it reaches the disk.
  for (const raw of [
    '../token-cache',
    '20261118T090507Z-0a1b2c3d/../../x',
    'x20261118T090507Z-0a1b2c3d',
    '20261118T090507Z-0a1b2c3dx',
    '20261118T090507Z-0A1B2C3D',
    '20261118T090507-0a1b2c3d',
    '',
  ]) {
    it(`refuses ${JSON.stringify(raw)} as a run id with a message that shows the shape of a good one`, () => {
      const parsed = parseRunId(raw);
      expect(parsed.ok).toBe(false);
      if (parsed.ok) return;
      expect(parsed.error.code).toBe('invalid_run_id');
      expect(parsed.error.message).toBe('A run id has the form 20261118T090507Z-0a1b2c3d: the UTC time of the run, then 8 hex digits. Copy the runId field of the write output.');
    });
  }
});

const parseRunIdValue = (raw: string): RunId => {
  const parsed = parseRunId(raw);
  if (!parsed.ok) throw new Error(`not a run id: ${raw}`);
  return parsed.value;
};
