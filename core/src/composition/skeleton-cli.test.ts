import { describe, expect, it } from 'bun:test';
import { runSkeletonCli } from './skeleton-cli.ts';

const AUTH = { name: '@ask-marcel/office-auth', version: '1.0.0' };

const run = (args: ReadonlyArray<string>): { readonly code: number; readonly out: string; readonly err: string } => {
  let out = '';
  let err = '';
  const code = runSkeletonCli(AUTH, args, {
    out: (text) => {
      out += text;
    },
    err: (text) => {
      err += text;
    },
  });
  return { code, out, err };
};

describe('a package of the workspace skeleton', () => {
  it('asked for --version, prints its version on stdout and exits 0', () => {
    expect(run(['--version'])).toEqual({ code: 0, out: '1.0.0\n', err: '' });
  });

  it('asked for a command, says on stderr that it has no commands yet and exits 1', () => {
    expect(run(['login'])).toEqual({
      code: 1,
      out: '',
      err: '@ask-marcel/office-auth 1.0.0 has no commands yet: it is a skeleton of the package split, and its commands move in with a later phase.\n',
    });
  });

  it('called with no argument, answers as for a command', () => {
    expect(run([]).code).toBe(1);
  });
});
