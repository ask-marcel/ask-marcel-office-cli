import { describe, expect, it } from 'bun:test';
import { isoDateTimeUnsafe, parseIsoDateTime, type IsoDateTime } from './iso-datetime.ts';

// Wednesday, 2026-05-20 at 14:30:00 UTC. ISO weekday 3.
const NOW = new Date('2026-05-20T14:30:00.000Z');

const expectOk = (input: string, expected: string): void => {
  const result = parseIsoDateTime(input, NOW);
  expect(result.ok).toBe(true);
  if (result.ok) expect(result.value).toBe(expected as IsoDateTime);
};

const expectErr = (input: string): void => {
  const result = parseIsoDateTime(input, NOW);
  expect(result.ok).toBe(false);
  if (!result.ok) {
    expect(result.error.type).toBe('invalid_format');
    expect(result.error.input).toBe(input);
    expect(result.error.hint).toContain('strict ISO 8601');
  }
};

describe('parseIsoDateTime — strict ISO passthrough', () => {
  it('an LLM that already computed the ISO string gets it back unchanged (canonical form)', () => {
    expectOk('2026-04-01T00:00:00Z', '2026-04-01T00:00:00Z');
  });

  it('accepts fractional seconds in the ISO form (Graph emits them on delta cursors)', () => {
    expectOk('2026-04-01T00:00:00.123Z', '2026-04-01T00:00:00.123Z');
  });

  it('rejects ISO with a non-UTC offset (Graph expects UTC; offsets would require translation)', () => {
    expectErr('2026-04-01T00:00:00+02:00');
  });
});

describe('parseIsoDateTime — date-only inputs expand to midnight UTC', () => {
  it('`2026-04-01` becomes `2026-04-01T00:00:00.000Z` so the LLM does not need to type the time half', () => {
    expectOk('2026-04-01', '2026-04-01T00:00:00.000Z');
  });

  it('rejects an out-of-range date (Feb 30) instead of silently rolling over to March 2', () => {
    expectErr('2026-02-30');
  });

  it('rejects a malformed date-only string (one component non-numeric)', () => {
    expectErr('2026-XX-01');
  });
});

describe('parseIsoDateTime — relative offsets answer the canonical "last week" / "last month" / "last hour" questions', () => {
  it('`7d` is interpreted as the PAST offset (7 days before now) so `--since 7d` returns the last week', () => {
    expectOk('7d', '2026-05-13T14:30:00.000Z');
  });

  it('`1w` is one week back (same as 7d but spelled with the unit Graph users think in)', () => {
    expectOk('1w', '2026-05-13T14:30:00.000Z');
  });

  it('`2h` is two hours back — useful for narrow "what changed this morning" windows', () => {
    expectOk('2h', '2026-05-20T12:30:00.000Z');
  });

  it('`30m` is thirty minutes back', () => {
    expectOk('30m', '2026-05-20T14:00:00.000Z');
  });

  it('`+7d` is the FUTURE offset for calendar-view end dates ("show me the next week")', () => {
    expectOk('+7d', '2026-05-27T14:30:00.000Z');
  });

  it('`-7d` explicitly signals past (same as the unsigned default, accepted for clarity)', () => {
    expectOk('-7d', '2026-05-13T14:30:00.000Z');
  });

  it('rejects an offset with zero magnitude — `0d` is meaningless for a window endpoint', () => {
    // `0d` parses to 0 days from now — but the regex matches it; behaviour:
    // returns `now` itself. This is intentional, not an error.
    expectOk('0d', '2026-05-20T14:30:00.000Z');
  });

  it('rejects an offset with no unit (LLM mis-typing a bare integer)', () => {
    expectErr('7');
  });

  it('rejects an offset with an unknown unit (years are intentionally not supported — fall back to ISO)', () => {
    expectErr('1y');
  });
});

describe('parseIsoDateTime — named anchors for the most common LLM phrasings', () => {
  it('`now` returns the injected current instant verbatim', () => {
    expectOk('now', '2026-05-20T14:30:00.000Z');
  });

  it("`today` is midnight UTC of the current day (start of the LLM's day-bucket)", () => {
    expectOk('today', '2026-05-20T00:00:00.000Z');
  });

  it('`yesterday` is midnight UTC of the previous day', () => {
    expectOk('yesterday', '2026-05-19T00:00:00.000Z');
  });

  it('`tomorrow` is midnight UTC of the next day (calendar-view end-date use)', () => {
    expectOk('tomorrow', '2026-05-21T00:00:00.000Z');
  });
});

describe('parseIsoDateTime — weekday names', () => {
  // NOW is a Wednesday. The most-recent Monday is 2 days back, the most-recent
  // Wednesday is today, the most-recent Sunday is 3 days back.
  it("`monday` returns the most recent Monday at midnight UTC (this week's Monday)", () => {
    expectOk('monday', '2026-05-18T00:00:00.000Z');
  });

  it('`wednesday` (today) returns today at midnight UTC — covers the on-target boundary case', () => {
    expectOk('wednesday', '2026-05-20T00:00:00.000Z');
  });

  it('`sunday` returns the most recent Sunday at midnight UTC (3 days back since today is Wed)', () => {
    expectOk('sunday', '2026-05-17T00:00:00.000Z');
  });

  it('`tuesday` (yesterday) returns yesterday at midnight UTC', () => {
    expectOk('tuesday', '2026-05-19T00:00:00.000Z');
  });

  it('`friday` returns 5 days back (the most recent Friday before today)', () => {
    expectOk('friday', '2026-05-15T00:00:00.000Z');
  });

  it('`saturday` returns 4 days back', () => {
    expectOk('saturday', '2026-05-16T00:00:00.000Z');
  });

  it('`thursday` returns 6 days back', () => {
    expectOk('thursday', '2026-05-14T00:00:00.000Z');
  });

  it('`last-monday` returns the Monday of LAST week (9 days back from a Wednesday)', () => {
    expectOk('last-monday', '2026-05-11T00:00:00.000Z');
  });

  it('`last-wednesday` (today is Wed) returns one full week back, NOT today', () => {
    expectOk('last-wednesday', '2026-05-13T00:00:00.000Z');
  });

  it('`next-monday` returns the upcoming Monday (5 days forward from a Wednesday)', () => {
    expectOk('next-monday', '2026-05-25T00:00:00.000Z');
  });

  it('`next-wednesday` (today is Wed) returns one full week forward, NOT today', () => {
    expectOk('next-wednesday', '2026-05-27T00:00:00.000Z');
  });

  it('rejects an unknown weekday-like name (typo path)', () => {
    expectErr('mondaay');
  });

  it('rejects `last-notaday` (the prefix matches but the weekday is unknown)', () => {
    expectErr('last-notaday');
  });

  it('rejects `next-notaday` (the prefix matches but the weekday is unknown)', () => {
    expectErr('next-notaday');
  });
});

describe('parseIsoDateTime — boundary anchors', () => {
  it('`start-of-week` returns this Monday at midnight UTC (Mon-Sun convention)', () => {
    expectOk('start-of-week', '2026-05-18T00:00:00.000Z');
  });

  it('`end-of-week` returns Sunday at 23:59:59.999 UTC', () => {
    expectOk('end-of-week', '2026-05-24T23:59:59.999Z');
  });

  it('`start-of-month` returns the 1st of the current month at midnight UTC', () => {
    expectOk('start-of-month', '2026-05-01T00:00:00.000Z');
  });

  it('`end-of-month` returns the last instant of the current month (May has 31 days)', () => {
    expectOk('end-of-month', '2026-05-31T23:59:59.999Z');
  });

  it('`start-of-year` returns Jan 1st of the current year at midnight UTC', () => {
    expectOk('start-of-year', '2026-01-01T00:00:00.000Z');
  });

  it('`end-of-year` returns Dec 31st at 23:59:59.999 UTC of the current year', () => {
    expectOk('end-of-year', '2026-12-31T23:59:59.999Z');
  });
});

describe('parseIsoDateTime — invalid input surfaces a hint that lists every accepted shape', () => {
  it('the empty string fails fast with the full hint list', () => {
    expectErr('');
  });

  it('a whitespace-only string fails (whitespace trimmed, then empty)', () => {
    expectErr('   ');
  });

  it('a non-date free-text input fails (e.g. "alex")', () => {
    expectErr('alex');
  });

  it('rejects a half-quoted relative phrase ("next christmas") rather than guessing', () => {
    expectErr('next christmas');
  });
});

describe('isoDateTimeUnsafe escape hatch — used only at the boundary where the value is already validated', () => {
  it('returns the input as the branded type without parsing', () => {
    const v = isoDateTimeUnsafe('2026-04-01T00:00:00Z');
    expect(v).toBe('2026-04-01T00:00:00Z' as IsoDateTime);
  });
});

describe('parseIsoDateTime — default `now` injection', () => {
  it('uses `new Date()` when no clock is supplied (smoke check — value is within the test window)', () => {
    const before = Date.now();
    const result = parseIsoDateTime('now');
    const after = Date.now();
    expect(result.ok).toBe(true);
    if (result.ok) {
      const parsed = new Date(result.value).getTime();
      expect(parsed).toBeGreaterThanOrEqual(before);
      expect(parsed).toBeLessThanOrEqual(after);
    }
  });
});

describe('parseIsoDateTime — weekday edge cases when "today" sits at week boundaries', () => {
  // Anchor on a Sunday (2026-05-24) so the Mon-based math is exercised at the
  // edge.
  const SUNDAY = new Date('2026-05-24T12:00:00.000Z');

  it("on Sunday, `monday` returns the Monday 6 days earlier (this week's Monday under Mon-Sun convention)", () => {
    const result = parseIsoDateTime('monday', SUNDAY);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('2026-05-18T00:00:00.000Z' as IsoDateTime);
  });

  it('on Sunday, `start-of-week` returns the same Monday', () => {
    const result = parseIsoDateTime('start-of-week', SUNDAY);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('2026-05-18T00:00:00.000Z' as IsoDateTime);
  });

  it('on Sunday, `end-of-week` returns the same Sunday at end-of-day', () => {
    const result = parseIsoDateTime('end-of-week', SUNDAY);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('2026-05-24T23:59:59.999Z' as IsoDateTime);
  });

  // Anchor on a Monday so that the `last-monday` and `next-monday` math is
  // exercised at "today is the target weekday".
  const MONDAY = new Date('2026-05-25T09:00:00.000Z');

  it('on Monday, `monday` returns today (back==0 case)', () => {
    const result = parseIsoDateTime('monday', MONDAY);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('2026-05-25T00:00:00.000Z' as IsoDateTime);
  });

  it('on Monday, `last-monday` returns 7 days back', () => {
    const result = parseIsoDateTime('last-monday', MONDAY);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('2026-05-18T00:00:00.000Z' as IsoDateTime);
  });

  it('on Monday, `next-monday` returns 7 days forward (the back==0 short-circuit)', () => {
    const result = parseIsoDateTime('next-monday', MONDAY);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('2026-06-01T00:00:00.000Z' as IsoDateTime);
  });
});

describe('parseIsoDateTime — end-of-month respects month boundaries', () => {
  it('on Feb 14 2026, end-of-month returns Feb 28 (non-leap year)', () => {
    const feb = new Date('2026-02-14T00:00:00Z');
    const result = parseIsoDateTime('end-of-month', feb);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('2026-02-28T23:59:59.999Z' as IsoDateTime);
  });

  it('on Dec 15 2026, end-of-month returns Dec 31 (last instant of December)', () => {
    const dec = new Date('2026-12-15T00:00:00Z');
    const result = parseIsoDateTime('end-of-month', dec);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('2026-12-31T23:59:59.999Z' as IsoDateTime);
  });
});
