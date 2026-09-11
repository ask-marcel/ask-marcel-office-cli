import { describe, expect, it } from 'bun:test';
import { bodyMarkdown, isChannelPost, type ChannelMessage } from './channel-message-html.ts';

const NO_IMAGES: ReadonlyMap<string, string> = new Map();
const HOSTED = 'https://graph.microsoft.com/v1.0/teams/t1/channels/c1/messages/1/hostedContents/h1/$value';

const post = (over: Partial<ChannelMessage> = {}): ChannelMessage => ({
  id: '1700000000000',
  replyToId: null,
  messageType: 'message',
  createdDateTime: '2026-09-08T14:03:11Z',
  from: { user: { id: 'u1', displayName: 'Alex Kim' } },
  body: { contentType: 'html', content: '<p>Budget review is <strong>Monday</strong></p>' },
  attachments: [],
  mentions: [],
  reactions: [],
  ...over,
});

const body = (m: ChannelMessage, images: ReadonlyMap<string, string> = NO_IMAGES, inline = false): string => bodyMarkdown(m, images, inline);

describe('telling a post from a system event', () => {
  it('a post has messageType message; a membership event is unknownFutureValue with an eventDetail', () => {
    expect(isChannelPost(post())).toBe(true);
    expect(isChannelPost(post({ messageType: 'unknownFutureValue', eventDetail: { '@odata.type': '#microsoft.graph.membersAddedEventMessageDetail' } }))).toBe(false);
    expect(isChannelPost(post({ messageType: 'systemEventMessage' }))).toBe(false);
  });
});

describe('turning a Teams message body into markdown', () => {
  it('converts an HTML body and trims it', () => {
    expect(body(post())).toBe('Budget review is **Monday**');
  });

  it('keeps a plain-text body verbatim and says when a message was deleted', () => {
    expect(body(post({ body: { contentType: 'text', content: 'line one\nline two' } }))).toBe('line one\nline two');
    expect(body(post({ deletedDateTime: '2026-09-09T08:00:00Z', body: { contentType: 'html', content: '' } }))).toBe('_message deleted_');
    expect(body(post({ body: { contentType: 'text' } }))).toBe('');
  });

  it('flattens @mentions to their text and resolves attachment placeholders from the attachments list', () => {
    const md = body(
      post({
        body: {
          contentType: 'html',
          content:
            '<p><at id="0">Robin Chen</at> see <attachment id="a1"></attachment> and <attachment id="a2"></attachment> <attachment id="a3"></attachment> <attachment id="zz"></attachment> <emoji id="smile" alt="😀" title="smile"></emoji></p>',
        },
        mentions: [{ id: 0, mentionText: 'Robin Chen' }],
        attachments: [
          { id: 'a1', contentType: 'reference', contentUrl: 'https://contoso.sharepoint.com/sites/ops/deck.pptx', name: 'deck.pptx' },
          { id: 'a2', contentType: 'meetingReference', name: 'Budget sync', content: '{}' },
          { id: 'a3', contentType: 'tabReference', name: 'Planner' },
        ],
      })
    );
    expect(md).toBe('@Robin Chen see [deck.pptx](https://contoso.sharepoint.com/sites/ops/deck.pptx) and [meeting: Budget sync] [tab: Planner] [attachment] 😀');
  });

  it('separates an attachment glued to the text before it, as Teams writes a scheduled meeting, and names an unnamed one by its type', () => {
    const md = body(
      post({
        body: { contentType: 'html', content: '<p>Scheduled a meeting<attachment id="m1"></attachment><attachment id="r1"/></p>' },
        attachments: [
          { id: 'm1', contentType: 'meetingReference', name: 'Budget sync' },
          { id: 'r1', contentType: 'reference', contentUrl: null, name: null },
        ],
      })
    );
    expect(md).toBe('Scheduled a meeting [meeting: Budget sync] [attachment: reference]');
  });

  it('resolves a mention whose tag carries no text through the mentions list, and falls back to a bare @mention', () => {
    const md = body(post({ body: { contentType: 'html', content: '<p><at id="3"></at> and <at id="9"></at></p>' }, mentions: [{ id: 3, mentionText: 'Jordan Avery' }] }));
    expect(md).toBe('@Jordan Avery and @mention');
  });

  it('summarises an adaptive card by the text it carries, and names any other attachment by its name', () => {
    const card = JSON.stringify({
      type: 'AdaptiveCard',
      body: [
        { type: 'TextBlock', text: 'Deploy approved' },
        { type: 'Container', items: [{ type: 'TextBlock', title: 'Owner', text: 'Alex' }] },
      ],
    });
    const md = body(
      post({
        body: {
          contentType: 'html',
          content: '<attachment id="c1"></attachment> <attachment id="c2"></attachment> <attachment id="c3"></attachment> <attachment id="c4"></attachment>',
        },
        attachments: [
          { id: 'c1', contentType: 'application/vnd.microsoft.card.adaptive', content: card },
          { id: 'c2', contentType: 'application/vnd.microsoft.card.adaptive', content: '{not json' },
          { id: 'c3', contentType: 'image/png', name: 'chart.png' },
          { id: 'c4', contentType: 'application/vnd.microsoft.card.hero', content: '{"body":[{"type":"Image","url":"x"}]}' },
        ],
      })
    );
    expect(md).toBe('[card: Deploy approved / Owner / Alex] [card: unreadable] [attachment: chart.png] [card]');
  });

  it('caps a card summary at 300 characters', () => {
    const card = JSON.stringify({ body: [{ text: 'x'.repeat(400) }] });
    const md = body(
      post({
        body: { contentType: 'html', content: '<attachment id="c1"></attachment>' },
        attachments: [{ id: 'c1', contentType: 'application/vnd.microsoft.card.adaptive', content: card }],
      })
    );
    expect(md).toBe(`[card: ${'x'.repeat(300)}]`);
  });

  it('replaces a hosted image with a placeholder naming the flag when embedding is off, with the fetched data URI when on, and leaves other images alone', () => {
    const html = { contentType: 'html', content: `<p>chart</p><img src="${HOSTED}" alt="chart"><img src="https://example.com/x.png" alt="ext">` };
    expect(body(post({ body: html }))).toBe('chart\n\n[image: hosted in Teams, pass --inline-images true to embed]![ext](https://example.com/x.png)');
    expect(body(post({ body: html }), new Map([[HOSTED, 'data:image/png;base64,AAAA']]), true)).toBe(
      'chart\n\n![chart](data:image/png;base64,AAAA)![ext](https://example.com/x.png)'
    );
    expect(body(post({ body: html }), NO_IMAGES, true)).toContain('[image: could not be embedded]');
  });

  it('drops the system-event marker tag', () => {
    expect(body(post({ body: { contentType: 'html', content: '<systemEventMessage/>' } }))).toBe('');
  });
});
