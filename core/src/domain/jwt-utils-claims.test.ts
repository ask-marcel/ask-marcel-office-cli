import { describe, expect, it } from 'bun:test';
import { unsignedJwt } from '../test-helpers/jwt.ts';
import { decodeJwtPayload, isGraphToken, isTokenFresh } from './jwt-utils.ts';

/*
 * What core reads from a token without checking its signature (package split
 * phase 3): the claims, whether it is still fresh, and whether Graph is its
 * audience.
 */

// Claims whose base64url form needs `-` and `_`, and every padding length.
const URL_SAFE_CLAIMS = { note: '>>>???', name: 'é' };

describe('the claims of a token', () => {
  it('are read from a base64url payload that uses - and _, whatever its padding', () => {
    for (const filler of ['', 'a', 'ab', 'abc']) {
      expect(decodeJwtPayload(unsignedJwt({ ...URL_SAFE_CLAIMS, filler }))).toEqual({ ...URL_SAFE_CLAIMS, filler });
    }
  });

  it('are none for a token of two or four segments', () => {
    const [header = '', payload = ''] = unsignedJwt({ sub: 'x' }).split('.');
    expect(decodeJwtPayload(`${header}.${payload}`)).toEqual({});
    expect(decodeJwtPayload(`${header}.${payload}.sig.extra`)).toEqual({});
  });
});

describe('a fresh token', () => {
  const now = Date.now() / 1000;

  it('is one that expires after the buffer from now, and not one that expires at its edge', () => {
    expect(isTokenFresh(unsignedJwt({ exp: Math.ceil(now) + 302 }))).toBe(true);
    expect(isTokenFresh(unsignedJwt({ exp: Math.floor(now) + 300 }))).toBe(false);
    expect(isTokenFresh(unsignedJwt({ exp: Math.ceil(now) + 12 }), 10)).toBe(true);
  });

  it('is never one whose expiry is missing or not a number', () => {
    expect(isTokenFresh(unsignedJwt({}))).toBe(false);
    expect(isTokenFresh(unsignedJwt({ exp: `${Math.ceil(now) + 3600}` }))).toBe(false);
  });
});

describe('a Graph token', () => {
  it('names Graph in its audience list, by URL or by resource id', () => {
    expect(isGraphToken(unsignedJwt({ aud: ['https://outlook.office.com', 'https://graph.microsoft.com'] }))).toBe(true);
    expect(isGraphToken(unsignedJwt({ aud: [42, '00000003-0000-0000-c000-000000000000'] }))).toBe(true);
  });

  it('is not one whose audience list names only other services, or holds no text', () => {
    expect(isGraphToken(unsignedJwt({ aud: ['https://outlook.office.com', 'api://00000003-0000-0000-c000-000000000001'] }))).toBe(false);
    expect(isGraphToken(unsignedJwt({ aud: [42, null] }))).toBe(false);
    expect(isGraphToken(unsignedJwt({ aud: 7 }))).toBe(false);
  });
});
