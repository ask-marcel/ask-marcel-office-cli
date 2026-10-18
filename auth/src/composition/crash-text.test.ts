import { describe, expect, it } from 'bun:test';
import { crashText, isCommanderError, isTokenInvocation } from './crash-text.ts';

// The one crash catch of the auth bin (src/main.ts) writes a line to stderr.
// On the token helper path that line must never carry a token.

const TOKEN_RUN = ['node', 'cli.js', 'token', '--tier', 'basic'];
const STATUS_RUN = ['node', 'cli.js', 'status'];
const QUOTED_TOKEN = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.secret';

describe('the crash line of the auth bin', () => {
  it('withholds the message of a crash on the token helper path, so a token that the message quotes never reaches stderr', () => {
    const text = crashText(TOKEN_RUN, new TypeError(`cannot parse ${QUOTED_TOKEN}`));

    expect(text).toBe('the token helper stopped on an unexpected TypeError; its message is withheld.');
    expect(text).not.toContain(QUOTED_TOKEN);
  });

  it('names only a throw when the token helper path crashes on a value that is not an Error', () => {
    const text = crashText(TOKEN_RUN, QUOTED_TOKEN);

    expect(text).toBe('the token helper stopped on an unexpected throw; its message is withheld.');
  });

  it('gives the full message of a crash on the CLI path, where no token is parsed', () => {
    expect(crashText(STATUS_RUN, new Error('the cache file is locked'))).toContain('the cache file is locked');
  });

  it('takes only a first argument of token as the token helper path', () => {
    expect(isTokenInvocation(TOKEN_RUN)).toBe(true);
    expect(isTokenInvocation(STATUS_RUN)).toBe(false);
    expect(isTokenInvocation(['node', 'cli.js', 'docs', 'token'])).toBe(false);
  });

  it('knows a commander parse error, which the CLI already rendered, from any other crash', () => {
    expect(isCommanderError({ code: 'commander.unknownOption' })).toBe(true);
    expect(isCommanderError(new Error('commander.unknownOption'))).toBe(false);
    expect(isCommanderError({ code: 'ENOENT' })).toBe(false);
    expect(isCommanderError({ code: 42 })).toBe(false);
    expect(isCommanderError(null)).toBe(false);
    expect(isCommanderError('commander.unknownOption')).toBe(false);
  });
});
