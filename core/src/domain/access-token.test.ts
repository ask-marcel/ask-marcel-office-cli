import { describe, expect, it } from 'bun:test';
import { jwtSegment } from '../test-helpers/jwt.ts';
import { accessToken, accessTokenUnsafe } from './access-token.ts';

const makeJwt = (claims: Record<string, unknown>): string => {
  const header = jwtSegment({ alg: 'RS256' });
  const payload = jwtSegment(claims);
  return `${header}.${payload}.sig`;
};

describe('accessToken brand factory', () => {
  it('accepts a fresh, audience-correct Graph JWT', () => {
    const raw = makeJwt({ exp: Math.floor(Date.now() / 1000) + 3600, aud: 'https://graph.microsoft.com' });
    const result = accessToken(raw);
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(accessTokenUnsafe(raw));
  });

  it('rejects a string that does not look like a JWT', () => {
    const result = accessToken('opaque-token');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('malformed_jwt');
  });

  it('rejects an expired JWT even if the audience is Graph', () => {
    const raw = makeJwt({ exp: Math.floor(Date.now() / 1000) - 3600, aud: 'https://graph.microsoft.com' });
    const result = accessToken(raw);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('expired');
  });

  it('rejects a fresh JWT when the audience is not Graph', () => {
    const raw = makeJwt({ exp: Math.floor(Date.now() / 1000) + 3600, aud: 'https://api.spaces.skype.com' });
    const result = accessToken(raw);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('wrong_audience');
  });

  it('refuses a fresh Graph JWT that is not sent as three base64url segments: padding, a line break from a wrapped paste, a fourth segment', () => {
    const good = makeJwt({ exp: Math.floor(Date.now() / 1000) + 3600, aud: 'https://graph.microsoft.com' });
    const [header, payload] = good.split('.');
    for (const raw of [`${header}=.${payload}.sig`, `${good}\nPART2`, `${good}.extra`]) {
      expect(accessToken(raw)).toEqual({ ok: false, error: { type: 'malformed_jwt' } });
    }
  });
});
