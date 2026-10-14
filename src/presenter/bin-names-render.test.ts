import { describe, expect, it } from 'bun:test';
import type { RenderContext } from './render-to-string.ts';
import { renderErrorToString, renderToString } from './render-to-string.ts';

// The read package of a split: a remedy names the bin of the package that holds
// the command, so a reader of read is sent to auth for `login`.
const READ_PACKAGE = { binName: 'widget-read', bins: { auth: 'widget-auth', read: 'widget-read', write: 'widget-write' } };

const PLACEHOLDER = /\{(?:(?:auth|read|write)-)?bin\}/;

const listing = (overrides: Partial<RenderContext> = {}): RenderContext => ({
  commandName: 'list-widgets',
  producesBytes: false,
  supportsSelect: true,
  supportsTop: true,
  surface: 'cli',
  ...overrides,
});

describe('the bins a rendered error names', () => {
  it('sends an expired token to the login of the auth package, in the message and in the hint', () => {
    const out = renderErrorToString('Token expired. Run `{auth-bin} login`.', 'text', 'InvalidAuthenticationToken', 'graph', undefined, READ_PACKAGE);
    expect(out).toContain('error: Token expired. Run `widget-auth login`.');
    expect(out).toContain('hint: The cached access token is invalid or expired. Run `widget-auth login` to re-authenticate, then retry. Use `widget-auth status`');
  });

  it('fills the error of the JSON envelope, as it fills the text one', () => {
    const out = JSON.parse(renderErrorToString('Token expired. Run `{auth-bin} login`.', 'json', 'InvalidAuthenticationToken', 'graph', undefined, READ_PACKAGE)) as {
      readonly error: string;
    };
    expect(out.error).toBe('Token expired. Run `widget-auth login`.');
  });

  it('names the bin of this package for the help every package has', () => {
    const out = renderErrorToString("unknown option '--nope'", 'json', 'commander.unknownOption', 'cli', undefined, READ_PACKAGE);
    expect(JSON.parse(out).hint).toBe(
      'Unknown CLI flag. Run `widget-read <command> --help` for the supported flags on that command, or `widget-read help-json --terse --category <name>` (~8 KB) to scan the whole category.'
    );
  });

  it('takes a validation message that already names a command as the remedy, and adds no generic hint', () => {
    const out = renderErrorToString(
      "--user-id 'x' looks like a contact id. Re-run `{read-bin} get-user --user-id` with the mail instead.",
      'text',
      undefined,
      'validation',
      undefined,
      READ_PACKAGE
    );
    expect(out).toBe("error: --user-id 'x' looks like a contact id. Re-run `widget-read get-user --user-id` with the mail instead.\nsource: validation\n");
  });

  it('points a bare validation message at the help of this package', () => {
    const out = renderErrorToString('--message-id is required', 'text', undefined, undefined, undefined, READ_PACKAGE);
    expect(out).toContain('hint: CLI input failed schema validation. Re-read the per-command help (`widget-read <cmd> --help` or `widget-read docs <cmd>`)');
  });
});

describe('the bins a rendered answer names', () => {
  it('prints the next page as a command of the package that holds next-page', () => {
    const out = renderToString({ value: [{ id: 'w1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/widgets?$skip=1' }, 'text', listing(), READ_PACKAGE);
    expect(out).toBe("id: w1\n\n--- next: widget-read next-page --url 'https://graph.microsoft.com/v1.0/widgets?$skip=1'\n");
  });

  it('points an empty --select answer at the docs of this package', () => {
    const out = renderToString({ value: [{ '@odata.etag': 'W/1' }] }, 'text', listing(), READ_PACKAGE);
    expect(out).toContain('selectHint: ');
    expect(out).toContain('in `widget-read docs <command>`.');
  });

  it('points an empty --select JSON answer at the docs of this package', () => {
    const out = JSON.parse(renderToString({ value: [{ '@odata.etag': 'W/1' }] }, 'json', listing(), READ_PACKAGE)) as { readonly selectHint: string };
    expect(out.selectHint).toContain('in `widget-read docs <command>`.');
    expect(out.selectHint).not.toMatch(PLACEHOLDER);
  });

  it('tells a terminal caller of an oversized text answer to redirect the command of this package', () => {
    const big = { value: Array.from({ length: 900 }, (_unused, i) => ({ id: `w${i}`, note: 'x'.repeat(60) })) };
    const sizeHint = renderToString(big, 'text', listing(), READ_PACKAGE).split('\n')[0] ?? '';
    expect(sizeHint).toContain('`widget-read list-widgets ... > ');
    expect(sizeHint).not.toMatch(PLACEHOLDER);
  });

  it('tells a terminal caller of an oversized JSON answer to redirect the command of this package', () => {
    const big = { value: Array.from({ length: 900 }, (_unused, i) => ({ id: `w${i}`, note: 'x'.repeat(60) })) };
    const out = JSON.parse(renderToString(big, 'json', listing(), READ_PACKAGE)) as { readonly sizeHint: string };
    expect(out.sizeHint).toContain('redirect to a file instead: `widget-read list-widgets ... > out.json`.');
  });

  it('leaves a placeholder that Graph data carries as it is: only the CLI text is filled', () => {
    const out = renderToString({ value: [{ id: 'w1', subject: 'Use {bin} and {auth-bin} in the template' }] }, 'text', listing(), READ_PACKAGE);
    expect(out).toBe('id: w1\nsubject: Use {bin} and {auth-bin} in the template\n');
  });
});
