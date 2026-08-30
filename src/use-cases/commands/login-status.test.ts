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

describe('login summary (plain-language tier capabilities)', () => {
  it('describes what every captured token lets the caller read', () => {
    const s = buildLoginSummary({ elevatedAvailable: true, chatsvcaggAvailable: true, ic3Available: true });
    expect(Object.keys(s.unlocked)).toEqual(['basic', 'elevated', 'chatsvcagg', 'ic3']);
    expect(s.unlocked['basic']).toContain('mail');
    expect(s.unlocked['elevated']).toContain('version history');
    expect(s.unlocked['chatsvcagg']).toContain('Teams chat message content');
    expect(s.unlocked['ic3']).toContain('Teams chat history');
    expect(s.missing).toEqual({});
  });

  it('names what a missing token costs and how to get it back', () => {
    const s = buildLoginSummary({ elevatedAvailable: false, chatsvcaggAvailable: true, ic3Available: false });
    expect(Object.keys(s.unlocked)).toEqual(['basic', 'chatsvcagg']);
    expect(Object.keys(s.missing)).toEqual(['elevated', 'ic3']);
    expect(s.missing['elevated']).toContain('version history');
    expect(s.missing['elevated']).toContain('login --force');
    expect(s.missing['ic3']).toContain('login --force');
  });

  it('reports the three optional tiers as missing when only basic was captured', () => {
    const s = buildLoginSummary({ elevatedAvailable: false, chatsvcaggAvailable: false, ic3Available: false });
    expect(Object.keys(s.unlocked)).toEqual(['basic']);
    expect(Object.keys(s.missing)).toEqual(['elevated', 'chatsvcagg', 'ic3']);
  });
});
