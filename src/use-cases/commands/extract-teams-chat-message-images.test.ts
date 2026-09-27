import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const CHAT = '19:abc@thread.v2';
const SHOT = 'https://as-api.asm.skype.com/v1/objects/0-weu-d1-aaa/views/imgo';
const CHART = 'https://ch-prod.asyncgw.teams.microsoft.com/v1/objects/0-weu-d1-bbb/views/imgo?v=1&amp;w=2';
const CONTENT = [
  '<p>very urgent, see screenshot</p>',
  `<img src="${SHOT}" itemtype="http://schema.skype.com/AMSImage" width="250" height="120" alt="image" id="x_0" itemid="0-weu-d1-aaa">`,
  '<img itemscope itemtype="http://schema.skype.com/Emoji" src="https://statics.teams.cdn.office.net/emoticons/cool/default/20_f.png" alt="cool">',
  `<img itemtype="http://schema.skype.com/AMSImage" src="${CHART}" alt="chart">`,
  `<img src="${SHOT}" itemtype="http://schema.skype.com/AMSImage">`,
  '<img itemtype="http://schema.skype.com/AMSImage" alt="broken">',
].join('');

const run = async (graph: ReturnType<typeof fakeGraphClient>): Promise<Result<unknown, GraphError>> => {
  const command = commands['extract-teams-chat-message-images'];
  if (!command) throw new Error('extract-teams-chat-message-images is not registered');
  return command.execute(graph, { chatId: CHAT, messageId: '1727000000000' });
};

describe('extract-teams-chat-message-images', () => {
  it('is marked experimental, refuses a call without a message id, and answers media for the --output-dir flag', async () => {
    const command = commands['extract-teams-chat-message-images'];
    expect(command?.meta.stability).toBe('experimental');
    expect(command?.meta.producesMedia).toBe(true);
    const refused = await command?.execute(fakeGraphClient(), { chatId: CHAT });
    expect(refused?.ok).toBe(false);
    if (refused !== undefined && !refused.ok) expect(refused.error.type).toBe('validation_error');
  });

  it('fetches every pasted image of a message once, as a media list the --output-dir flag can land, and leaves emoji alone', async () => {
    const reads: string[] = [];
    const media: string[] = [];
    const graph = fakeGraphClient({
      teamsChatIc3: async (path: string) => {
        reads.push(path);
        return ok({ id: '1727000000000', content: CONTENT });
      },
      teamsChatMedia: async (url: string) => {
        media.push(url);
        return ok(
          url.startsWith('https://as-api') ? { contentType: 'image/jpeg', size: 3, base64: 'AAAA' } : { contentType: 'image/png; charset=binary', size: 2, base64: 'BBB=' }
        );
      },
    });
    const result = await run(graph);
    if (!result.ok) throw new Error(result.error.message);
    expect(result.value).toEqual({
      count: 2,
      media: [
        { path: 'image-1.jpg', contentType: 'image/jpeg', sizeBytes: 3, base64: 'AAAA' },
        { path: 'image-2.png', contentType: 'image/png; charset=binary', sizeBytes: 2, base64: 'BBB=' },
      ],
    });
    expect(reads).toEqual(['/v1/users/ME/conversations/19%3Aabc%40thread.v2/messages/1727000000000']);
    expect(media).toEqual([SHOT, 'https://ch-prod.asyncgw.teams.microsoft.com/v1/objects/0-weu-d1-bbb/views/imgo?v=1&w=2']);
  });

  it('says so when a message holds no pasted image, names an unknown image type as .bin, and passes on a failed read', async () => {
    const none = await run(fakeGraphClient({ teamsChatIc3: async () => ok({ content: '<p>no image</p>' }) }));
    expect(none).toEqual(ok({ count: 0, media: [], note: 'This message holds no pasted image.' }));
    const empty = await run(fakeGraphClient({ teamsChatIc3: async () => ok({ id: '5', messagetype: 'Event/Call' }) }));
    expect(empty).toEqual(ok({ count: 0, media: [], note: 'This message holds no pasted image.' }));
    for (const [contentType, extension] of [
      ['image/gif', 'gif'],
      ['image/webp', 'webp'],
    ]) {
      const typed = await run(
        fakeGraphClient({
          teamsChatIc3: async () => ok({ content: `<img itemtype="http://schema.skype.com/AMSImage" src="${SHOT}">` }),
          teamsChatMedia: async () => ok({ contentType, size: 1, base64: 'AA==' }),
        })
      );
      expect(typed.ok && (typed.value as { media: ReadonlyArray<{ path: string }> }).media[0]?.path).toBe(`image-1.${extension}`);
    }
    const odd = await run(
      fakeGraphClient({
        teamsChatIc3: async () => ok({ content: `<img itemtype="http://schema.skype.com/AMSImage" src="${SHOT}">` }),
        teamsChatMedia: async () => ok({ contentType: 'image/heic', size: 1, base64: 'AA==' }),
      })
    );
    expect(odd.ok && (odd.value as { media: ReadonlyArray<{ path: string }> }).media[0]?.path).toBe('image-1.bin');
    const gone = err({ type: 'api_error' as const, status: 404, message: 'NotFound: gone' });
    for (const graph of [
      fakeGraphClient({ teamsChatIc3: async () => gone }),
      fakeGraphClient({ teamsChatIc3: async () => ok({ content: CONTENT }), teamsChatMedia: async () => gone }),
    ]) {
      const failed = await run(graph);
      expect(failed.ok).toBe(false);
      if (!failed.ok) expect(failed.error.message).toBe('NotFound: gone');
    }
  });
});
