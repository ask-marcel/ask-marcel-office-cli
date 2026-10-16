import { afterEach, describe, expect, it } from 'bun:test';
import { currentDateZone, setDateZone } from './date-zone.ts';

// The zone named days resolve in (core, package split phase 3): the composition
// root sets the zone of the run once, and every date-taking command reads it.

describe('the date zone of a run', () => {
  const before = currentDateZone();
  afterEach(() => setDateZone(before));

  it('is the zone that was set last', () => {
    setDateZone('Asia/Shanghai');
    expect(currentDateZone()).toBe('Asia/Shanghai');
    setDateZone('Europe/Paris');
    expect(currentDateZone()).toBe('Europe/Paris');
  });
});
