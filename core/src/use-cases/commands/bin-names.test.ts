import { describe, expect, it } from 'bun:test';
import { fillNames, fillNamesIn } from './bin-names.ts';

// The read package of a split, as it renders the names of the three packages.
const READ_PACKAGE = { binName: 'widget-read', bins: { auth: 'widget-auth', read: 'widget-read', write: 'widget-write' } };

describe('the bin placeholders in command references', () => {
  it('names the bin of each package where its placeholder stands, every time it stands there', () => {
    const text =
      'Run `{auth-bin} login`, then `{read-bin} list-drives` or `{write-bin} save-draft`; `{bin} docs <command>` and `{bin} --help` explain both. Again: `{auth-bin} status`.';
    expect(fillNames(text, READ_PACKAGE)).toBe(
      'Run `widget-auth login`, then `widget-read list-drives` or `widget-write save-draft`; `widget-read docs <command>` and `widget-read --help` explain both. Again: `widget-auth status`.'
    );
  });

  it('gives {bin} the bin the package runs as, which need not be one of the three', () => {
    expect(fillNames('`{bin} help-json`', { ...READ_PACKAGE, binName: 'widgets' })).toBe('`widgets help-json`');
  });

  it('leaves a text with no placeholder as it is, braces included', () => {
    expect(fillNames('GET /users/{id}?$select={fields} and {region}', READ_PACKAGE)).toBe('GET /users/{id}?$select={fields} and {region}');
  });

  it('fills every string of a report the CLI built itself, at any depth, and keeps its other values', () => {
    const report = {
      hint: '`{auth-bin} login`',
      basic: { available: false, expiresInSeconds: 12, reason: 'run `{auth-bin} status`' },
      available: ['basic', '{bin}'],
      missing: null,
    };
    expect(fillNamesIn(report, READ_PACKAGE)).toEqual({
      hint: '`widget-auth login`',
      basic: { available: false, expiresInSeconds: 12, reason: 'run `widget-auth status`' },
      available: ['basic', 'widget-read'],
      missing: null,
    });
  });
});
