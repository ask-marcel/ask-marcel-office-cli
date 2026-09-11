import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import type { ChannelMessage } from './channel-message-html.ts';
import { fetchHostedImages, hostedImageSources } from './channel-message-images.ts';

const HOSTED = 'https://graph.microsoft.com/v1.0/teams/t1/channels/c1/messages/1/hostedContents/h1/$value';

const withBody = (html: string, replies: ReadonlyArray<ChannelMessage> = []): ChannelMessage => ({
  id: '1',
  messageType: 'message',
  body: { contentType: 'html', content: html },
  replies,
});

describe('finding and fetching the hosted images of a page', () => {
  it('lists each distinct hosted image once, across posts and their replies, whatever the quote style', () => {
    const other = HOSTED.replace('h1', 'h2');
    const srcs = hostedImageSources([
      withBody(`<img src="${HOSTED}"><img src='${other}'><img src="https://example.com/x.png"><img alt="no src">`, [withBody(`<img src="${HOSTED}">`)]),
    ]);
    expect(srcs).toEqual([HOSTED, other]);
  });

  it('embeds only images at or under 2 MB that Graph returns, skipping failures, non-images and non-Graph URLs', async () => {
    const paths: string[] = [];
    const graph = fakeGraphClient({
      getBinary: async (path) => {
        paths.push(path);
        if (path.endsWith('h1/$value')) return ok({ contentType: 'image/png', size: 4, base64: 'AAAA' });
        if (path.endsWith('h2/$value')) return ok({ contentType: 'text/html', size: 4, base64: 'PGI+' });
        if (path.endsWith('h3/$value')) return ok({ contentType: 'image/png', size: 2_000_001, base64: 'BBBB' });
        if (path.endsWith('h5/$value')) return ok({ contentType: 'image/png', size: 4 });
        return err({ type: 'api_error', status: 404, message: 'gone' });
      },
    });
    const images = await fetchHostedImages(graph, [
      HOSTED,
      HOSTED.replace('h1', 'h2'),
      HOSTED.replace('h1', 'h3'),
      HOSTED.replace('h1', 'h4'),
      HOSTED.replace('h1', 'h5'),
      'https://example.com/x.png',
    ]);
    expect(paths).toEqual(['h1', 'h2', 'h3', 'h4', 'h5'].map((h) => `/teams/t1/channels/c1/messages/1/hostedContents/${h}/$value`));
    expect([...images.entries()]).toEqual([[HOSTED, 'data:image/png;base64,AAAA']]);
  });

  it('embeds an image exactly at the 2 MB limit, and reads the beta origin too', async () => {
    const graph = fakeGraphClient({ getBinary: async () => ok({ contentType: 'image/jpeg', size: 2_000_000, base64: 'CCCC' }) });
    const beta = HOSTED.replace('/v1.0/', '/beta/');
    const images = await fetchHostedImages(graph, [HOSTED, beta]);
    expect(images.get(HOSTED)).toBe('data:image/jpeg;base64,CCCC');
    expect(images.get(beta)).toBe('data:image/jpeg;base64,CCCC');
  });
});

describe('the URL and body shapes the image scan must survive', () => {
  it('reads a spaced, single-quoted src, ignores a message without a body or replies, and finds an image in a nested reply', () => {
    const nested = HOSTED.replace('h1', 'h9');
    const bare: ChannelMessage = { id: 'b', messageType: 'message' };
    const srcs = hostedImageSources([bare, { ...withBody(`<img src = '${HOSTED}'>`), replies: undefined }, withBody('<p>no image</p>', [withBody(`<img src="${nested}">`)])]);
    expect(srcs).toEqual([HOSTED, nested]);
  });

  it('refuses a URL that only contains the Graph origin somewhere inside it', async () => {
    const paths: string[] = [];
    const graph = fakeGraphClient({
      getBinary: async (path) => {
        paths.push(path);
        return ok({ contentType: 'image/png', size: 4, base64: 'AAAA' });
      },
    });
    const images = await fetchHostedImages(graph, [`https://evil.example/${HOSTED}`]);
    expect(paths).toEqual([]);
    expect(images.size).toBe(0);
  });
});
