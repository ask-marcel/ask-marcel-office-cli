import { describe, expect, it } from 'bun:test';
import { commands } from './index.ts';

// Guards the drafts-dedup caveat: a future edit must not silently drop the
// warning that a conversationId $filter on the Drafts folder is not a reliable
// "does a draft already exist on this thread" check. The caveat is the human
// half of the fix whose machine half is the find-mail-drafts command.

describe('list-mail-folder-messages documents the Drafts $filter dedup caveat', () => {
  it('warns on the mail-folder-id option that a conversationId filter on Drafts is unreliable for dedup', () => {
    const { meta } = commands['list-mail-folder-messages'];
    const prose = [meta.summary, ...meta.options.map((option) => option.description)].join(' ');
    expect(prose).toContain('not a reliable');
    expect(prose).toContain('conversationId');
  });
});
