import { describe, expect, it } from 'bun:test';
import { parseIsoDateTime } from './iso-datetime.ts';

const NOW = new Date('2026-05-20T14:30:00Z');
const rejected = (input: string): string => {
  const r = parseIsoDateTime(input, NOW);
  if (r.ok) throw new Error(`${input} was accepted as ${r.value}`);
  return r.error.hint;
};
const accepted = (input: string): string => {
  const r = parseIsoDateTime(input, NOW);
  if (!r.ok) throw new Error(`${input} was rejected`);
  return r.value;
};

describe('the shapes a date input must have', () => {
  it('anchors every shape to the whole input: a prefix or suffix makes it invalid', () => {
    for (const bad of ['x2026-04-01T00:00:00Z', '2026-04-01T00:00:00Zx', 'x2026-04-01', '2026-04-01x', 'x7d', '7dx', '+7dx']) rejected(bad);
  });

  it('refuses a bare date that does not exist, whether the month or the day is impossible', () => {
    for (const bad of ['2026-02-30', '2026-13-01', '2026-04-31', '2026-00-10', '2026-12-32']) rejected(bad);
    expect(accepted('2026-02-28')).toBe('2026-02-28T00:00:00.000Z');
    expect(accepted('2024-02-29')).toBe('2024-02-29T00:00:00.000Z');
  });

  it('ignores surrounding whitespace and letter case', () => {
    expect(accepted('  today  ')).toBe('2026-05-20T00:00:00.000Z');
    expect(accepted('Last-Monday')).toBe('2026-05-11T00:00:00.000Z');
    expect(accepted('NEXT-FRIDAY')).toBe('2026-05-22T00:00:00.000Z');
    expect(accepted('last-wednesday')).toBe('2026-05-13T00:00:00.000Z');
    expect(accepted('next-wednesday')).toBe('2026-05-27T00:00:00.000Z');
  });

  it('lists every accepted shape in the hint, one per line, time zone rule included', () => {
    const hint = rejected('alex');
    for (const phrase of ['Accepted shapes:', 'strict ISO 8601', 'ISO date', 'past offset', 'future offset', 'named:', 'weekday:', 'boundary:', "run's time zone"])
      expect(hint).toContain(phrase);
    expect(hint.split('\n').length).toBeGreaterThan(8);
  });
});
