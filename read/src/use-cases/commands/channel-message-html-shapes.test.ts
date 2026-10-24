import { describe, expect, it } from 'bun:test';
import { bodyMarkdown, type ChannelMessage } from './channel-message-html.ts';

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

describe('the attribute and tag spellings a Teams body can carry', () => {
  it('reads single-quoted attributes with spaces around the equals sign', () => {
    const md = body(
      post({ body: { contentType: 'html', content: `<p><at id = '3'></at><img src = '${HOSTED}' alt = 'chart'></p>` }, mentions: [{ id: 3, mentionText: 'Jordan Avery' }] }),
      new Map([[HOSTED, 'data:image/png;base64,AAAA']]),
      true
    );
    expect(md).toBe('@Jordan Avery![chart](data:image/png;base64,AAAA)');
  });

  it('drops the inner content of attachment and emoji tags, keeps an emoji with no alt silent, and clears every system-event tag form', () => {
    const md = body(
      post({
        body: {
          contentType: 'html',
          content: '<p>a<attachment id="a1">inner text\nmore</attachment><emoji alt="😀">x y</emoji><emoji id="q"></emoji></p><systemEventMessage foo="1"/><systemEventMessage>',
        },
        attachments: [{ id: 'a1', contentType: 'tabReference', name: 'Planner', contentUrl: 'https://contoso.sharepoint.com/tab' }],
      })
    );
    expect(md).toBe('a [tab: Planner]😀');
  });

  it('restores more than ten parked items in one body', () => {
    const tags = Array.from({ length: 11 }, (_v, i) => `<at id="${i}">P${i}</at>`).join(' ');
    const md = body(post({ body: { contentType: 'html', content: `<p>${tags}</p>` } }));
    expect(md).toBe(Array.from({ length: 11 }, (_v, i) => `@P${i}`).join(' '));
    expect(md).not.toContain('AMSLOT');
  });

  it('gives a mention the text in its tag over the mentions list, trimmed, and a bare @mention when neither names it', () => {
    const md = body(post({ body: { contentType: 'html', content: '<p><at id="0"> Robin Chen </at> <at></at></p>' }, mentions: [{ id: 0, mentionText: 'Robin' }] }));
    expect(md).toBe('@Robin Chen @mention');
    expect(body(post({ body: { contentType: 'html', content: '<at id="1"></at>' }, mentions: undefined }))).toBe('@mention');
  });

  it('reads card text trimmed, walks past null values and an array root, and marks a card without content or with broken JSON', () => {
    const card = JSON.stringify([{ text: '  Deploy approved  ', speak: null }, { items: [{ title: 'Owner' }] }]);
    const md = body(
      post({
        body: { contentType: 'html', content: '<attachment id="c1"></attachment><attachment id="c2"></attachment><attachment id="c3"></attachment>' },
        attachments: [
          { id: 'c1', contentType: 'application/vnd.microsoft.card.adaptive', content: card },
          { id: 'c2', contentType: 'application/vnd.microsoft.card.adaptive' },
          { id: 'c3', contentType: 'application/vnd.microsoft.card.adaptive', content: '{"a":' },
        ],
      })
    );
    expect(md).toBe('[card: Deploy approved / Owner] [card] [card: unreadable]');
  });

  it('names a nameless shared file "attachment", a nameless typeless one "unknown", and never links a tab even when it carries a URL', () => {
    const md = body(
      post({
        body: { contentType: 'html', content: '<attachment id="r1"></attachment><attachment id="x1"></attachment><attachment id="t1"></attachment>' },
        attachments: [
          { id: 'r1', contentType: 'reference', contentUrl: 'https://contoso.sharepoint.com/sites/ops/deck.pptx' },
          { id: 'x1' },
          { id: 't1', contentType: 'tabReference', name: 'Planner', contentUrl: 'https://contoso.sharepoint.com/tab' },
        ],
      })
    );
    expect(md).toBe('[attachment](https://contoso.sharepoint.com/sites/ops/deck.pptx) [attachment: unknown] [tab: Planner]');
  });

  it('embeds a hosted image with no alt as an image with an empty alt, and renders nothing for a message without a body', () => {
    expect(body(post({ body: { contentType: 'html', content: `<img src="${HOSTED}">` } }), new Map([[HOSTED, 'data:image/png;base64,AAAA']]), true)).toBe(
      '![](data:image/png;base64,AAAA)'
    );
    expect(body(post({ body: undefined }))).toBe('');
  });
});

describe('the leftovers a body can carry', () => {
  it('drops an image without a src (as turndown does), keeps a slot-looking word the user typed, and resolves nothing when the attachments field is absent', () => {
    expect(body(post({ body: { contentType: 'html', content: '<p>see <img alt="x"> AMSLOT7X <attachment id="a1"></attachment></p>' }, attachments: undefined }))).toBe(
      'see  AMSLOT7X [attachment]'
    );
  });
});
