import { describe, expect, it } from 'bun:test';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { fakeReadGraph } from '@ask-marcel/office-core/test-helpers/read-graph-fake.ts';
import { execute } from './resolve-mail-link.ts';

// The link shapes that resolve-mail-link takes and the shapes that it refuses,
// with the full refusal: its type, its words and its code.

const resolve = async (url: string): Promise<unknown> => execute(fakeReadGraph(), { url });

const NOT_OUTLOOK = err({
  type: 'validation_error',
  message:
    '--url: not an Outlook mail link. Expected shapes: `https://outlook.office.com/owa/?itemid=AAMkA...`, `https://outlook.office.com/mail/inbox/id/AAMkA...`, or `https://outlook.office.com/mail/AAMkA...`. Hosts accepted: outlook.office.com, outlook.office365.com, outlook.live.com.',
});

const ID = 'AAMkAGI2THVS_message_id';

describe('resolve-mail-link: the link shapes that it takes', () => {
  it('takes a link on outlook.live.com, the third Outlook host', async () => {
    expect(await resolve(`https://outlook.live.com/mail/inbox/id/${ID}`)).toEqual(ok({ messageId: ID }));
  });

  it('takes the camel-case `itemId` query of an OWA link', async () => {
    expect(await resolve(`https://outlook.office.com/owa/?itemId=${ID}`)).toEqual(ok({ messageId: ID }));
  });

  it('takes the long id after `/mail/` when an `id` segment is the last segment', async () => {
    expect(await resolve(`https://outlook.office.com/mail/${ID}/id`)).toEqual(ok({ messageId: ID }));
  });

  it('takes a link whose path starts with an empty segment', async () => {
    expect(await resolve(`https://outlook.office.com//mail/inbox/id/${ID}`)).toEqual(ok({ messageId: ID }));
  });
});

describe('resolve-mail-link: the link shapes that it refuses', () => {
  it('refuses an `itemid` query on a path that is not `/owa`', async () => {
    expect(await resolve(`https://outlook.office.com/mail/?itemid=${ID}`)).toEqual(NOT_OUTLOOK);
  });

  it('refuses an OWA link with an empty `itemid`', async () => {
    expect(await resolve('https://outlook.office.com/owa/?itemid=')).toEqual(NOT_OUTLOOK);
  });

  it('refuses a path that does not start with `mail`, with or without an `id` segment', async () => {
    expect(await resolve(`https://outlook.office.com/people/${ID}`)).toEqual(NOT_OUTLOOK);
    expect(await resolve(`https://outlook.office.com/people/inbox/id/${ID}`)).toEqual(NOT_OUTLOOK);
  });

  it('refuses a link to the root of an Outlook host', async () => {
    expect(await resolve('https://outlook.office.com/')).toEqual(NOT_OUTLOOK);
  });

  it('refuses a short link whose second segment has 19 characters', async () => {
    expect(await resolve('https://outlook.office.com/mail/AAMkAGI2THVS_19char')).toEqual(NOT_OUTLOOK);
  });

  it('refuses a calendar link with a validation error and the calendar code', async () => {
    expect(await resolve(`https://outlook.office.com/calendar/item/${ID}`)).toEqual(
      err({ type: 'validation_error', message: '--url looks like an Outlook calendar link, not a mail message link.', code: 'cli_reject_calendar_link_on_mail_resolver' })
    );
  });

  it('refuses a Teams message link with a validation error and the Teams code', async () => {
    expect(await resolve('https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1700000000000')).toEqual(
      err({ type: 'validation_error', message: '--url looks like a Teams message link, not an Outlook mail message link.', code: 'cli_reject_teams_link_on_mail_resolver' })
    );
  });

  it('refuses a host that is not Outlook with the full generic refusal', async () => {
    expect(await resolve(`https://example.com/mail/inbox/id/${ID}`)).toEqual(NOT_OUTLOOK);
  });
});
