import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../../domain/access-token.ts';
import { tenantIdUnsafe } from '../../domain/tenant-id.ts';
import { tokenFingerprint } from '../../domain/token-fingerprint.ts';
import { parseTokenArgs } from './token-args.ts';

const GUID = '8f2c1a4e-3b6d-4c9a-9e1f-2a7b5c8d0e3f';
const TOKEN = accessTokenUnsafe(`${btoa(JSON.stringify({ alg: 'RS256' }))}.${btoa(JSON.stringify({ exp: 1_900_000_000, aud: 'https://graph.microsoft.com' }))}.sig`);

const argsFor = (argv: ReadonlyArray<string>): ReturnType<typeof parseTokenArgs> => parseTokenArgs(argv);

describe('token helper: what a caller asks for', () => {
  it('a caller asking for the basic tier gets the basic token request', () => {
    expect(argsFor(['--tier', 'basic'])).toEqual({ ok: true, value: { request: { tier: 'basic' } } });
  });

  it('a caller asking for a guest token names the partner tenant', () => {
    expect(argsFor(['--tier', 'guest', '--tenant', GUID])).toEqual({ ok: true, value: { request: { tier: 'guest', tenant: tenantIdUnsafe(GUID) } } });
  });

  it('a caller replaying a refused token passes its fingerprint, in any argument order', async () => {
    const fingerprint = await tokenFingerprint(TOKEN);
    expect(argsFor(['--reject', fingerprint, '--tier', 'chatsvcagg'])).toEqual({ ok: true, value: { request: { tier: 'chatsvcagg' }, rejected: fingerprint } });
  });

  const refused = (argv: ReadonlyArray<string>): { errorCode: string; tier: string | null; message: string; remedy: string } => {
    const args = argsFor(argv);
    if (args.ok) throw new Error('expected the arguments to be refused');
    return args.error;
  };

  it('refuses a call with no tier, naming the five tiers it accepts', () => {
    const line = refused([]);
    expect(line.errorCode).toBe('invalid_arguments');
    expect(line.tier).toBeNull();
    expect(line.message).toContain('basic, elevated, chatsvcagg, ic3, guest');
    expect(line.remedy).toContain('--tier');
  });

  it('refuses a tier it does not know without echoing it', () => {
    const line = refused(['--tier', 'admin-secret']);
    expect(line.errorCode).toBe('invalid_arguments');
    expect(line.tier).toBeNull();
    expect(JSON.stringify(line)).not.toContain('admin-secret');
  });

  it('refuses a guest request that names no tenant', () => {
    const line = refused(['--tier', 'guest']);
    expect(line).toMatchObject({ errorCode: 'invalid_arguments', tier: 'guest' });
    expect(line.message).toContain('needs --tenant');
  });

  it('refuses a tenant on a tier that is not issued by a partner tenant', () => {
    const line = refused(['--tier', 'basic', '--tenant', GUID]);
    expect(line).toMatchObject({ errorCode: 'invalid_arguments', tier: 'basic' });
    expect(line.message).toContain('guest tier only');
  });

  it('refuses a tenant that is not a GUID', () => {
    const line = refused(['--tier', 'guest', '--tenant', 'contoso/../x']);
    expect(line).toMatchObject({ errorCode: 'invalid_arguments', tier: 'guest' });
    expect(line.message).toContain('GUID');
    expect(JSON.stringify(line)).not.toContain('contoso');
  });

  // A caller that pastes the bearer where the fingerprint goes must not see it
  // printed back: stdout carries a token only in a success line.
  it('refuses a token passed as --reject, and never prints it back', () => {
    const line = refused(['--tier', 'basic', '--reject', TOKEN]);
    expect(line).toMatchObject({ errorCode: 'invalid_arguments', tier: 'basic' });
    expect(line.message).toContain('SHA-256');
    expect(JSON.stringify(line)).not.toContain(TOKEN);
  });

  it('refuses a flag it does not know, a flag with no value, and a flag given twice', () => {
    for (const argv of [['--tier', 'basic', '--force', 'yes'], ['--tier', 'basic', 'extra'], ['--tier'], ['--tier', 'basic', '--tier', 'ic3']]) {
      expect(refused(argv)).toMatchObject({ errorCode: 'invalid_arguments', tier: null, message: expect.stringContaining('each followed by its value') });
    }
  });
});
