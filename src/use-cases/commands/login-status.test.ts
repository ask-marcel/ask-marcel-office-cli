import { describe, expect, it } from 'bun:test';
import { buildLoginSummary } from './login-status.ts';

describe('login summary (slim auth confirmation)', () => {
  it('lists basic plus every currently-available token', () => {
    const s = buildLoginSummary({ elevatedAvailable: true, chatsvcaggAvailable: true, ic3Available: true });
    expect(s.status).toBe('authenticated');
    expect(s.available).toEqual(['basic', 'elevated', 'chatsvcagg', 'ic3']);
  });

  it('omits an unavailable token from the available list', () => {
    const s = buildLoginSummary({ elevatedAvailable: false, chatsvcaggAvailable: true, ic3Available: false });
    expect(s.available).toEqual(['basic', 'chatsvcagg']);
    expect(s.available).not.toContain('elevated');
    expect(s.available).not.toContain('ic3');
  });

  it('always includes basic, since authentication just succeeded', () => {
    const s = buildLoginSummary({ elevatedAvailable: false, chatsvcaggAvailable: false, ic3Available: false });
    expect(s.available).toEqual(['basic']);
  });

  it('points the hint at scopes-check for detail and login --force to refresh', () => {
    const s = buildLoginSummary({ elevatedAvailable: true, chatsvcaggAvailable: true, ic3Available: true });
    expect(s.hint).toContain('scopes-check');
    expect(s.hint).toContain('login --force');
  });
});
