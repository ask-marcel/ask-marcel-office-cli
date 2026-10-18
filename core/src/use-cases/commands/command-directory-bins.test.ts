import { describe, expect, it } from 'bun:test';
import { fillNames } from './bin-names.ts';
import { PACKAGE_BINS } from './command-directory.ts';

/*
 * The bin of each package of the split (package split phase 4). The auth
 * package, the single package and later read and write fill `{auth-bin}`,
 * `{read-bin}` and `{write-bin}` from this one record.
 */

describe('the bin of each package of the split', () => {
  it('names the auth, read and write bins as each package.json installs them', () => {
    expect(PACKAGE_BINS).toEqual({ auth: 'ask-marcel-office-auth', read: 'ask-marcel-office-read', write: 'ask-marcel-office-write' });
  });

  it('fills a remedy of the auth package with the auth and read bins', () => {
    const names = { binName: PACKAGE_BINS.auth, bins: PACKAGE_BINS };
    expect(fillNames('Run `{bin} login`, then `{read-bin} list-drives`.', names)).toBe('Run `ask-marcel-office-auth login`, then `ask-marcel-office-read list-drives`.');
  });
});
