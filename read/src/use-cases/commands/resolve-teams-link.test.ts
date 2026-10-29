import { describe, expect, it } from 'bun:test';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { execute } from './resolve-teams-link.ts';

const PREFIX = 'https://teams.microsoft.com/l/message/';
const NOT_TEAMS = err<GraphError>({
  type: 'validation_error',
  message: `--url: not a Teams message link. Expected shape: ${PREFIX}<chat-id>/<message-id>[?tenantId=...&groupId=...]`,
});

const resolve = (url: string): ReturnType<typeof execute> => execute(fakeGraphClient(), { url });

describe('resolving a Teams message link', () => {
  it('gives only the two ids when the link has no query', async () => {
    expect(await resolve(`${PREFIX}19%3Aabc%40thread.v2/1700000000000`)).toEqual(ok({ chatId: '19:abc@thread.v2', messageId: '1700000000000' }));
  });

  it('gives each optional field the link carries, and only those', async () => {
    expect(await resolve(`${PREFIX}c1/m1?tenantId=t1`)).toEqual(ok({ chatId: 'c1', messageId: 'm1', tenantId: 't1' }));
    expect(await resolve(`${PREFIX}c1/m1?groupId=g1&parentMessageId=p1&ctx=chat`)).toEqual(
      ok({ chatId: 'c1', messageId: 'm1', groupId: 'g1', parentMessageId: 'p1', context: 'chat' })
    );
  });

  it('refuses a long link on another host', async () => {
    expect(await resolve('https://example.com/aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa/bbb/ccc')).toEqual(NOT_TEAMS);
  });

  it('refuses a link with an empty chat id or an empty message id', async () => {
    expect(await resolve(`${PREFIX}/m1`)).toEqual(NOT_TEAMS);
    expect(await resolve(`${PREFIX}c1/`)).toEqual(NOT_TEAMS);
  });

  it('points an Outlook mail link to the mail resolver', async () => {
    expect(await resolve('https://outlook.office.com/mail/inbox/id/AAMk')).toEqual(
      err({
        type: 'validation_error',
        message: '--url looks like an Outlook mail message link, not a Teams message link.',
        code: 'cli_reject_mail_link_on_teams_resolver',
      })
    );
  });

  it('points an Outlook calendar link to the calendar resolver', async () => {
    expect(await resolve('https://outlook.office.com/calendar/item/AAMk')).toEqual(
      err({
        type: 'validation_error',
        message: '--url looks like an Outlook calendar item link, not a Teams message link.',
        code: 'cli_reject_calendar_link_on_teams_resolver',
      })
    );
  });

  it('points a OneDrive sharing link to the drive-share resolver', async () => {
    expect(await resolve('https://1drv.ms/w/s!abc')).toEqual(
      err({
        type: 'validation_error',
        message: '--url looks like a OneDrive / SharePoint sharing URL, not a Teams message link.',
        code: 'cli_reject_drive_share_link_on_teams_resolver',
      })
    );
  });
});
