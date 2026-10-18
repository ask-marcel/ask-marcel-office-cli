import { describe, expect, it } from 'bun:test';
import { fillNames } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import { AUTH_PACKAGE_NAMES } from './auth-package-names.ts';

// The names the auth package fills its texts with: its own bin for `{bin}`
// and `{auth-bin}`, and the bins of read and write for theirs.

describe('the names of the auth package', () => {
  it('names the auth bin as its own bin and the bins of read and write as theirs', () => {
    expect(AUTH_PACKAGE_NAMES).toEqual({
      binName: 'ask-marcel-office-auth',
      bins: { auth: 'ask-marcel-office-auth', read: 'ask-marcel-office-read', write: 'ask-marcel-office-write' },
    });
  });

  it('tells a person to sign in with the auth bin, never with the bin of the single package', () => {
    expect(fillNames('Run `{auth-bin} login`, then `{bin} status`.', AUTH_PACKAGE_NAMES)).toBe('Run `ask-marcel-office-auth login`, then `ask-marcel-office-auth status`.');
  });
});
