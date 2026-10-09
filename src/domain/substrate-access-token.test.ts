import { describe, expect, it } from 'bun:test';
import { ok, err } from './result.ts';
import { accessTokenUnsafe, isJwtShaped, substrateAccessToken } from './access-token.ts';

const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const jwt = (claims: Record<string, unknown>): string => `${segment({ alg: 'RS256' })}.${segment(claims)}.sig`;
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const CHATSVCAGG = 'https://chatsvcagg.teams.microsoft.com';
const IC3 = 'https://ic3.teams.office.com';

describe('a Teams chat token checked for the service it signs for', () => {
  it('accepts an unexpired token issued for the chat service it is meant for', () => {
    const chat = jwt({ exp: inAnHour(), aud: CHATSVCAGG });
    const media = jwt({ exp: inAnHour(), aud: IC3 });
    expect(substrateAccessToken(chat, 'chatsvcagg')).toEqual(ok(accessTokenUnsafe(chat)));
    expect(substrateAccessToken(media, 'ic3')).toEqual(ok(accessTokenUnsafe(media)));
  });

  it('refuses a token for the other chat service, or for Graph, as the wrong audience', () => {
    expect(substrateAccessToken(jwt({ exp: inAnHour(), aud: IC3 }), 'chatsvcagg')).toEqual(err({ type: 'wrong_audience' }));
    expect(substrateAccessToken(jwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com' }), 'ic3')).toEqual(err({ type: 'wrong_audience' }));
  });

  it('refuses a token that expires within five minutes, and a value that is not a JWT', () => {
    expect(substrateAccessToken(jwt({ exp: Math.floor(Date.now() / 1000) + 60, aud: IC3 }), 'ic3')).toEqual(err({ type: 'expired' }));
    expect(substrateAccessToken('opaque', 'ic3')).toEqual(err({ type: 'malformed_jwt' }));
  });

  it('refuses a chat token that is not sent as three base64url segments, as not a JWT', () => {
    const good = jwt({ exp: inAnHour(), aud: CHATSVCAGG });
    const [header, payload] = good.split('.');
    for (const raw of [`${header}=.${payload}.sig`, `${good}\nPART2`, `${good}.extra`]) {
      expect(substrateAccessToken(raw, 'chatsvcagg')).toEqual(err({ type: 'malformed_jwt' }));
    }
  });
});

describe('a value shaped as a JWT is sent in a header', () => {
  it('is three base64url segments, the last of which may be empty', () => {
    expect(isJwtShaped(jwt({ aud: IC3 }))).toBe(true);
    expect(isJwtShaped('eyJh_b-c.eyJ0.')).toBe(true);
  });

  it('holds no line break, space, padding or other byte, anywhere, and has exactly three segments', () => {
    const good = jwt({ aud: IC3 });
    for (const value of [`${good}\nPART2`, `\n${good}`, `eyJ ${good.slice(3)}`, `${good}=`, `${good}+x`, `${good}.extra`, 'eyJh.eyJ0', '.eyJ0.sig', 'eyJh..sig', '']) {
      expect(isJwtShaped(value)).toBe(false);
    }
  });
});
