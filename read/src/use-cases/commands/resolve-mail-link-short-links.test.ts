import { describe, expect, it } from 'bun:test';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { execute } from './resolve-mail-link.ts';

const resolve = async (url: string): Promise<unknown> => execute(fakeGraphClient(), { url });

describe('resolve-mail-link: the legacy short link `/mail/<id>`', () => {
  it('resolves a short link to the message id it carries', async () => {
    const result = await resolve('https://outlook.office365.com/mail/AAMkAGI2THVS_legacy_short_link');
    expect(result).toEqual({ ok: true, value: { messageId: 'AAMkAGI2THVS_legacy_short_link' } });
  });

  it('resolves a short link whose id has exactly 20 characters, the shortest length it takes as an id', async () => {
    const result = await resolve('https://outlook.office.com/mail/AAMkAGI2THVS_20chars');
    expect(result).toEqual({ ok: true, value: { messageId: 'AAMkAGI2THVS_20chars' } });
  });

  it('refuses a short link on a host that is not Outlook, even with a long id', async () => {
    const result = await resolve('https://example.com/mail/AAMkAGI2THVS_legacy_short_link');
    expect(result).toMatchObject({ ok: false, error: { type: 'validation_error', message: expect.stringContaining('not an Outlook mail link') } });
  });

  it('refuses a link to the mail start page, because it names no message', async () => {
    const result = await resolve('https://outlook.office.com/mail/');
    expect(result).toMatchObject({ ok: false, error: { type: 'validation_error', message: expect.stringContaining('not an Outlook mail link') } });
  });
});
