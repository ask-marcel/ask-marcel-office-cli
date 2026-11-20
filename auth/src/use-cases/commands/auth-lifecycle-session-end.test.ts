import { describe, expect, it } from 'bun:test';
import { AUTH_LIFECYCLE_ENTRIES } from './auth-lifecycle-entries.ts';

// The manifest of `login` declares `--renew`, which starts a new sign-in
// session of about 24 hours, and the response shapes of `login` and `status`
// name the `session` key.

const entry = (name: string): (typeof AUTH_LIFECYCLE_ENTRIES)[number] => {
  const found = AUTH_LIFECYCLE_ENTRIES.find((candidate) => candidate.name === name);
  if (found === undefined) throw new Error(`no lifecycle entry ${name}`);
  return found;
};

describe('the manifest entries of a sign-in session of about 24 hours', () => {
  it('declares login --renew as an optional option that keeps the browser cookies', () => {
    expect(entry('login').options.find((option) => option.name === 'renew')).toStrictEqual({
      name: 'renew',
      key: 'renew',
      required: false,
      description:
        'Start a new sign-in session before the current one ends. Microsoft limits the session to about 24 hours after the browser sign-in, and `status` shows its end as `session.endsAt`. It re-captures every token and a new refresh token in the browser, and keeps the browser cookies, so it usually asks for no password. A plain `login` opens no browser while the session is good. With `--force`, `--force` wins.',
    });
  });

  it('names the session key in the response shapes of login and status', () => {
    for (const name of ['login', 'status']) expect(entry(name).responseShape).toContain('session: { endsAt: string, endsInSeconds: number } | { reason: string }');
  });

  it('says in the status response shape that the end never stops a redemption', () => {
    expect(entry('status').responseShape).toContain('The end is a prediction only: a command still sends the refresh token to Entra.');
  });
});
