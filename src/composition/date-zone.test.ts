import { describe, expect, it } from 'bun:test';
import { resolveDateZone } from './date-zone.ts';

describe('which time zone a run resolves named days in', () => {
  it('takes the flag first, then the environment, then the machine', () => {
    expect(resolveDateZone('Europe/Amsterdam', 'Asia/Shanghai', () => 'America/New_York')).toEqual({ zone: 'Europe/Amsterdam', source: 'flag' });
    expect(resolveDateZone(undefined, 'Asia/Shanghai', () => 'America/New_York')).toEqual({ zone: 'Asia/Shanghai', source: 'env' });
    expect(resolveDateZone(undefined, undefined, () => 'America/New_York')).toEqual({ zone: 'America/New_York', source: 'machine' });
    expect(resolveDateZone('', '', () => 'America/New_York')).toEqual({ zone: 'America/New_York', source: 'machine' });
  });

  it('warns and keeps the machine zone when ASKMARCEL_TZ is not a zone, and falls back to UTC when the machine reports nonsense', () => {
    const bad = resolveDateZone(undefined, 'Mars/Olympus', () => 'Asia/Shanghai');
    expect(bad.zone).toBe('Asia/Shanghai');
    expect(bad.source).toBe('machine');
    expect(bad.warning).toContain('ASKMARCEL_TZ "Mars/Olympus"');
    const stripped = resolveDateZone(undefined, undefined, () => 'Etc/Nowhere');
    expect(stripped).toMatchObject({ zone: 'UTC', source: 'fallback' });
    expect(stripped.warning).toContain('Etc/Nowhere');
    const both = resolveDateZone(undefined, 'Mars/Olympus', () => 'Etc/Nowhere');
    expect(both.zone).toBe('UTC');
    expect(both.warning).toContain('resolve in UTC');
  });
});
