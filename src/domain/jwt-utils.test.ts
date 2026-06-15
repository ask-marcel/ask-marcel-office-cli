import { describe, expect, it } from 'bun:test';
import { decodeJwtPayload, isGraphToken, isTokenFresh } from './jwt-utils.ts';

const makeToken = (claims: Record<string, unknown>): string => {
  const header = btoa(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload = btoa(JSON.stringify(claims));
  return `${header}.${payload}.signature`;
};

describe('decodeJwtPayload', () => {
  it('extracts claims from a valid JWT', () => {
    const token = makeToken({ exp: 1700000000, aud: 'graph.microsoft.com', sub: 'user1' });
    const claims = decodeJwtPayload(token);
    expect(claims.exp).toBe(1700000000);
    expect(claims.aud).toBe('graph.microsoft.com');
  });

  it('returns empty object for malformed token', () => {
    expect(decodeJwtPayload('not.a.jwt')).toEqual({});
  });

  it('returns empty object for non-JWT string', () => {
    expect(decodeJwtPayload('hello world')).toEqual({});
  });
});

describe('isTokenFresh', () => {
  it('returns true for token with distant expiry', () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const token = makeToken({ exp: future });
    expect(isTokenFresh(token)).toBe(true);
  });

  it('returns false for expired token', () => {
    const past = Math.floor(Date.now() / 1000) - 100;
    const token = makeToken({ exp: past });
    expect(isTokenFresh(token)).toBe(false);
  });

  it('returns false within buffer window', () => {
    const nearExpiry = Math.floor(Date.now() / 1000) + 10;
    const token = makeToken({ exp: nearExpiry });
    expect(isTokenFresh(token)).toBe(false);
  });
});

describe('isGraphToken', () => {
  it('returns true when aud contains graph.microsoft.com', () => {
    const token = makeToken({ aud: 'https://graph.microsoft.com' });
    expect(isGraphToken(token)).toBe(true);
  });

  it('returns true when aud is an array containing graph.microsoft.com', () => {
    const token = makeToken({ aud: ['api://foo', 'https://graph.microsoft.com'] });
    expect(isGraphToken(token)).toBe(true);
  });

  it('returns true when aud is the Graph resource UUID (00000003-0000-0000-c000-000000000000)', () => {
    const token = makeToken({ aud: '00000003-0000-0000-c000-000000000000' });
    expect(isGraphToken(token)).toBe(true);
  });

  it('returns true when aud is an array containing the Graph resource UUID', () => {
    const token = makeToken({ aud: ['api://foo', '00000003-0000-0000-c000-000000000000'] });
    expect(isGraphToken(token)).toBe(true);
  });

  it('returns false for wrong audience', () => {
    const token = makeToken({ aud: 'management.core.windows.net' });
    expect(isGraphToken(token)).toBe(false);
  });

  it('returns false when the token carries no audience claim at all', () => {
    const token = makeToken({ sub: 'user1' });
    expect(isGraphToken(token)).toBe(false);
  });
});
