import { describe, expect, it } from 'bun:test';
import { isoDateTimeUnsafe } from '../../domain/iso-datetime.ts';
import { isoDateTimeField } from './iso-datetime-schema.ts';

describe('isoDateTimeField — Zod fragment for calendar windows', () => {
  it('parses a strict ISO 8601 UTC string to itself (canonical happy path)', () => {
    const r = isoDateTimeField.safeParse('2026-04-01T00:00:00Z');
    expect(r.success).toBe(true);
    if (r.success) expect(r.data).toBe(isoDateTimeUnsafe('2026-04-01T00:00:00Z'));
  });

  it('parses the date-only form to midnight UTC so the LLM can omit the time half', () => {
    const r = isoDateTimeField.safeParse('2026-04-01');
    expect(r.success).toBe(true);
    if (r.success) expect(r.data).toBe(isoDateTimeUnsafe('2026-04-01T00:00:00.000Z'));
  });

  it('parses a relative offset (`7d`) into the corresponding ISO datetime relative to now', () => {
    const r = isoDateTimeField.safeParse('7d');
    expect(r.success).toBe(true);
    if (r.success) expect(r.data).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/);
  });

  it('surfaces a Zod issue with the full domain hint when the input is not a recognised date format', () => {
    const r = isoDateTimeField.safeParse('alex');
    expect(r.success).toBe(false);
    if (!r.success) {
      const msg = r.error.issues[0]?.message ?? '';
      expect(msg).toContain('not a recognised date format ("alex")');
      expect(msg).toContain('strict ISO 8601');
      expect(msg).toContain('past offset');
    }
  });

  it('rejects an empty string with a non-empty refinement before the relative-date transform runs', () => {
    const r = isoDateTimeField.safeParse('');
    expect(r.success).toBe(false);
  });
});
