import { afterEach, describe, expect, it } from 'bun:test';
import { currentDateZone, setDateZone } from './date-zone.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';

describe('the run zone the date fields resolve in', () => {
  afterEach(() => setDateZone('UTC'));

  it('is UTC until the composition root sets it, and then every date field resolves named days there', () => {
    expect(currentDateZone()).toBe('UTC');
    const utc = isoDateTimeField.safeParse('today');
    expect(utc.success).toBe(true);
    if (utc.success) expect(String(utc.data).endsWith('T00:00:00.000Z')).toBe(true);
    setDateZone('Asia/Shanghai');
    expect(currentDateZone()).toBe('Asia/Shanghai');
    const shanghai = isoDateTimeField.safeParse('today');
    expect(shanghai.success).toBe(true);
    if (shanghai.success) expect(String(shanghai.data).endsWith('T16:00:00.000Z')).toBe(true);
    const instant = isoDateTimeField.safeParse('2026-04-01T10:00:00Z');
    if (instant.success) expect(String(instant.data)).toBe('2026-04-01T10:00:00Z');
  });

  it('reports an unknown zone through the field instead of throwing', () => {
    setDateZone('Mars/Olympus');
    const r = isoDateTimeField.safeParse('today');
    expect(r.success).toBe(false);
    if (!r.success) {
      expect(r.error.issues[0]?.message).toContain('IANA');
      expect(r.error.issues[0]?.code).toBe('custom');
    }
    expect(RELATIVE_DATE_DESCRIPTION).toContain('time zone');
  });
});
