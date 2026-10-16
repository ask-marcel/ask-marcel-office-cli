import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { err, ok } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import type { TokenError, TokenSource } from '../use-cases/ports/token-source.ts';
import { NO_SECONDARY_TOKEN_COMMANDS, neededByNote, withNeededBy } from './token-source-needed-by.ts';

/*
 * The second half of a token failure (core, package split phase 3): auth gives
 * the message and the remedy, and the package that asked adds its own commands
 * that need the missing token.
 */

const NEEDED_BY = { elevated: ['get-chat', 'list-chats'], chatsvcagg: ['list-teams-chat-messages'], ic3: ['list-teams-chat-history'] };

const missing: TokenError = { type: 'auth_failed', message: 'The token is not in the cache.', code: 'secondary_token_unavailable' };

const refusing = (error: TokenError): TokenSource => ({
  graphToken: async () => err(error),
  guestToken: async () => err(error),
  substrateToken: async () => err(error),
  substrateRegion: async () => err(error),
});

const PARTNER = tenantIdUnsafe('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04');

describe('a secondary token that is missing', () => {
  it('names the commands of the package that need it, and keeps the code auth gave', async () => {
    const source = withNeededBy(refusing(missing), NEEDED_BY);
    expect(await source.graphToken('elevated')).toEqual(err({ ...missing, message: 'The token is not in the cache. (Commands that need this token: get-chat, list-chats.)' }));
    expect(await source.substrateToken('chatsvcagg')).toEqual(
      err({ ...missing, message: 'The token is not in the cache. (Commands that need this token: list-teams-chat-messages.)' })
    );
    expect(await source.substrateRegion('ic3')).toEqual(err({ ...missing, message: 'The token is not in the cache. (Commands that need this token: list-teams-chat-history.)' }));
  });

  it('adds nothing when the package names no command', async () => {
    expect(await withNeededBy(refusing(missing), NO_SECONDARY_TOKEN_COMMANDS).substrateToken('ic3')).toEqual(err(missing));
  });
});

describe('a failure the list cannot cure', () => {
  it('stays as it is for the basic and the guest token, which every command needs', async () => {
    const source = withNeededBy(refusing(missing), NEEDED_BY);
    expect(await source.graphToken('basic')).toEqual(err(missing));
    expect(await source.guestToken(PARTNER)).toEqual(err(missing));
  });

  it('stays as it is for another code, no code, or a cancelled sign-in', async () => {
    for (const error of [
      { type: 'auth_failed', message: 'No helper.', code: 'token_helper_unavailable' },
      { type: 'auth_failed', message: 'Refused.' },
      { type: 'auth_cancelled' },
    ] satisfies ReadonlyArray<TokenError>) {
      expect(await withNeededBy(refusing(error), NEEDED_BY).substrateToken('ic3')).toEqual(err(error));
    }
  });
});

describe('a token the source has', () => {
  it('passes through with the options the caller gave', async () => {
    const asked: unknown[] = [];
    const source: TokenSource = {
      graphToken: async (tier, options) => {
        asked.push(['graph', tier, options]);
        return ok(accessTokenUnsafe(`${tier}-token`));
      },
      guestToken: async (tenant, options) => {
        asked.push(['guest', tenant, options]);
        return ok(accessTokenUnsafe('guest-token'));
      },
      substrateToken: async (tier, options) => {
        asked.push(['substrate', tier, options]);
        return ok(accessTokenUnsafe(`${tier}-token`));
      },
      substrateRegion: async (tier) => {
        asked.push(['region', tier]);
        return ok('emea' as TeamsRegion);
      },
    };
    const wrapped = withNeededBy(source, NEEDED_BY);
    const rejected = { rejected: accessTokenUnsafe('old') };
    expect(await wrapped.graphToken('elevated', rejected)).toEqual(ok(accessTokenUnsafe('elevated-token')));
    expect(await wrapped.graphToken('basic', rejected)).toEqual(ok(accessTokenUnsafe('basic-token')));
    expect(await wrapped.guestToken(PARTNER, rejected)).toEqual(ok(accessTokenUnsafe('guest-token')));
    expect(await wrapped.substrateToken('ic3', rejected)).toEqual(ok(accessTokenUnsafe('ic3-token')));
    expect(await wrapped.substrateRegion('chatsvcagg')).toEqual(ok('emea' as TeamsRegion));
    expect(asked).toEqual([
      ['graph', 'elevated', rejected],
      ['graph', 'basic', rejected],
      ['guest', PARTNER, rejected],
      ['substrate', 'ic3', rejected],
      ['region', 'chatsvcagg'],
    ]);
  });
});

describe('the note itself', () => {
  it('names the token it is about, or "this token" by default', () => {
    expect(neededByNote(['list-chats'], 'the elevated token')).toBe(' (Commands that need the elevated token: list-chats.)');
    expect(neededByNote(['a', 'b'])).toBe(' (Commands that need this token: a, b.)');
    expect(neededByNote([])).toBe('');
  });
});
