import { describe, expect, it } from 'bun:test';
import { isValidTimeZone, parseIsoDateTime } from './iso-datetime.ts';

// 2026-09-19 20:30 UTC is already Sunday 2026-09-20 04:30 in Shanghai (UTC+8, no DST).
const NOW = new Date('2026-09-19T20:30:00Z');
const at = (input: string, zone: string): string => {
  const r = parseIsoDateTime(input, NOW, zone);
  if (!r.ok) throw new Error(`${input} in ${zone}: ${r.error.type}`);
  return r.value;
};

describe('day boundaries in a time zone', () => {
  it('resolves today, yesterday and tomorrow to the zone midnight, expressed as UTC instants', () => {
    expect(at('today', 'Asia/Shanghai')).toBe('2026-09-19T16:00:00.000Z');
    expect(at('yesterday', 'Asia/Shanghai')).toBe('2026-09-18T16:00:00.000Z');
    expect(at('tomorrow', 'Asia/Shanghai')).toBe('2026-09-20T16:00:00.000Z');
  });

  it('resolves the week, month and year boundaries in the zone', () => {
    expect(at('start-of-week', 'Asia/Shanghai')).toBe('2026-09-13T16:00:00.000Z');
    expect(at('end-of-week', 'Asia/Shanghai')).toBe('2026-09-20T15:59:59.999Z');
    expect(at('start-of-month', 'Asia/Shanghai')).toBe('2026-08-31T16:00:00.000Z');
    expect(at('end-of-month', 'Asia/Shanghai')).toBe('2026-09-30T15:59:59.999Z');
    expect(at('start-of-year', 'Asia/Shanghai')).toBe('2025-12-31T16:00:00.000Z');
    expect(at('end-of-year', 'Asia/Shanghai')).toBe('2026-12-31T15:59:59.999Z');
  });

  it('resolves weekdays from the zone-local calendar', () => {
    expect(at('monday', 'Asia/Shanghai')).toBe('2026-09-13T16:00:00.000Z');
    expect(at('friday', 'Asia/Shanghai')).toBe('2026-09-17T16:00:00.000Z');
    expect(at('sunday', 'Asia/Shanghai')).toBe('2026-09-19T16:00:00.000Z');
    expect(at('last-monday', 'Asia/Shanghai')).toBe('2026-09-06T16:00:00.000Z');
    expect(at('next-monday', 'Asia/Shanghai')).toBe('2026-09-20T16:00:00.000Z');
    expect(at('next-sunday', 'Asia/Shanghai')).toBe('2026-09-26T16:00:00.000Z');
  });

  it('follows a daylight-saving change: the day after the spring-forward night starts an hour earlier in UTC', () => {
    // Amsterdam moves to CEST at 01:00 UTC on 2026-03-29; at 00:30 UTC it is still 01:30 CET.
    const springNight = new Date('2026-03-29T00:30:00Z');
    const r = (input: string): string => {
      const v = parseIsoDateTime(input, springNight, 'Europe/Amsterdam');
      if (!v.ok) throw new Error(input);
      return v.value;
    };
    expect(r('today')).toBe('2026-03-28T23:00:00.000Z');
    expect(r('tomorrow')).toBe('2026-03-29T22:00:00.000Z');
    expect(r('end-of-month')).toBe('2026-03-31T21:59:59.999Z');
  });

  it('keeps UTC as the default zone, and treats a zone of UTC exactly like the default', () => {
    expect(at('today', 'UTC')).toBe('2026-09-19T00:00:00.000Z');
    const d = parseIsoDateTime('today', NOW);
    if (!d.ok) throw new Error('default');
    expect(String(d.value)).toBe('2026-09-19T00:00:00.000Z');
  });

  it('leaves instants, bare dates and offsets alone whatever the zone', () => {
    expect(at('2026-04-01T10:00:00Z', 'Asia/Shanghai')).toBe('2026-04-01T10:00:00Z');
    expect(at('2026-04-01', 'Asia/Shanghai')).toBe('2026-04-01T00:00:00.000Z');
    expect(at('7d', 'Asia/Shanghai')).toBe('2026-09-12T20:30:00.000Z');
    expect(at('now', 'Asia/Shanghai')).toBe('2026-09-19T20:30:00.000Z');
  });

  it('refuses a zone it does not know, naming it, and validates zones on their own', () => {
    const r = parseIsoDateTime('today', NOW, 'Mars/Olympus');
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.error.type).toBe('invalid_zone');
      expect(r.error.input).toBe('Mars/Olympus');
      expect(r.error.hint).toContain('IANA');
    }
    expect(isValidTimeZone('Europe/Amsterdam')).toBe(true);
    expect(isValidTimeZone('UTC')).toBe(true);
    expect(isValidTimeZone('')).toBe(false);
  });
});

describe('which zone names the runtime resolves', () => {
  it('accepts the current IANA names, their older aliases and any letter case, and refuses names that are not zones', () => {
    for (const zone of ['Asia/Kolkata', 'Asia/Calcutta', 'Europe/Kyiv', 'Etc/UTC', 'utc', 'europe/amsterdam']) expect(isValidTimeZone(zone)).toBe(true);
    for (const zone of ['', ' ', 'Mars/Olympus', 'local']) expect(isValidTimeZone(zone)).toBe(false);
  });

  it('resolves a named day in a zone the runtime lists only under an older name', () => {
    // India keeps UTC+5:30 all year: at 20:30 UTC it is already 02:00 on the 20th there.
    expect(at('today', 'Asia/Kolkata')).toBe('2026-09-19T18:30:00.000Z');
    // Kyiv is on summer time (UTC+3): 23:30 on the 19th.
    expect(at('today', 'Europe/Kyiv')).toBe('2026-09-18T21:00:00.000Z');
  });
});
