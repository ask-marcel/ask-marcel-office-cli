import { describe, expect, it } from 'bun:test';
import { createMailStoreFake } from '../../test-helpers/mail-store-fake.ts';
import { runWrite } from '../../test-helpers/run-write.ts';

/*
 * A due day of set-mail-flag that has the form YYYY-MM-DD but is not a day of
 * the calendar (a month 13, a day 32) is refused with the words of the option,
 * and no request goes to Graph. Before, such a value threw a RangeError.
 */

describe('set-mail-flag --due', () => {
  it('refuses a month 13 and a day 32 as not a day, and sends no request', async () => {
    const mailbox = createMailStoreFake([{ id: 'm1', parentFolderId: 'inbox-id', changeKey: 'm1-ck', flag: { flagStatus: 'notFlagged' } }]);
    for (const due of ['2026-13-01', '2026-01-32', '2026-00-10']) {
      expect((await runWrite(mailbox.graph, 'set-mail-flag', { messageId: 'm1', status: 'flagged', due }))['error']).toEqual({
        source: 'validation',
        message: '--due must be a day in the form YYYY-MM-DD, for example 2026-11-20',
      });
    }
    expect(mailbox.calls).toEqual([]);
  });
});
