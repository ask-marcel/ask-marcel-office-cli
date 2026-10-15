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

describe('ask-marcel-office-read, the workspace skeleton', () => {
  it('answers --version with the version in the package.json of @ask-marcel/office-read', () => {
    expect(run(['--version'])).toEqual({ code: 0, out: '1.0.0\n', err: '' });
  });

  it('refuses list-drives, because no command has moved in yet', () => {
    const refused = run(['list-drives']);
    expect(refused.code).toBe(1);
    expect(refused.err).toStartWith('@ask-marcel/office-read 1.0.0 has no commands yet');
  });
});
