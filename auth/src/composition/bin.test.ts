import { describe, expect, it } from 'bun:test';
import { runBin } from './bin.ts';

const run = (args: ReadonlyArray<string>): { readonly code: number; readonly out: string; readonly err: string } => {
  let out = '';
  let err = '';
  const code = runBin(args, {
    out: (text) => {
      out += text;
    },
    err: (text) => {
      err += text;
    },
  });
  return { code, out, err };
};

describe('ask-marcel-office-auth, the workspace skeleton', () => {
  it('answers --version with the version in the package.json of @ask-marcel/office-auth', () => {
    expect(run(['--version'])).toEqual({ code: 0, out: '1.0.0\n', err: '' });
  });

  it('refuses login, because no command has moved in yet', () => {
    const refused = run(['login']);
    expect(refused.code).toBe(1);
    expect(refused.err).toStartWith('@ask-marcel/office-auth 1.0.0 has no commands yet');
  });
});
