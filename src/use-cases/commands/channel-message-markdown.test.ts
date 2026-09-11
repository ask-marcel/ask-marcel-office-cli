import { describe, expect, it } from 'bun:test';
import type { ChannelMessage } from './channel-message-html.ts';
import { renderThread, renderTranscript } from './channel-message-markdown.ts';

const alex = { user: { id: 'u1', displayName: 'Alex Kim' } };
const robin = { user: { id: 'u2', displayName: 'Robin Chen' } };
const NO_IMAGES: ReadonlyMap<string, string> = new Map();

const post = (over: Partial<ChannelMessage> = {}): ChannelMessage => ({
  id: '1700000000000',
  replyToId: null,
  messageType: 'message',
  createdDateTime: '2026-09-08T14:03:11Z',
  from: alex,
  body: { contentType: 'html', content: '<p>Budget review is <strong>Monday</strong></p>' },
  attachments: [],
  mentions: [],
  reactions: [],
  ...over,
});

const reply = (over: Partial<ChannelMessage> = {}): ChannelMessage =>
  post({
    id: '1700000000001',
    replyToId: '1700000000000',
    createdDateTime: '2026-09-08T14:10:00Z',
    from: robin,
    body: { contentType: 'html', content: '<p>Works for me</p>' },
    ...over,
  });

const text = (r: string): string => r;

describe('rendering one post with its replies as a thread', () => {
  it('heads the post with its date, time and author, converts the body, and nests each reply as a quoted block in order', () => {
    const md = text(renderThread(post(), [reply({ createdDateTime: '2026-09-08T15:00:00Z', body: { contentType: 'html', content: '<p>Second</p>' } }), reply()], NO_IMAGES, false));
    expect(md).toBe(
      [
        '### 2026-09-08 14:03 · Alex Kim',
        '',
        'Budget review is **Monday**',
        '',
        '> **Robin Chen · 2026-09-08 14:10**',
        '> Works for me',
        '',
        '> **Robin Chen · 2026-09-08 15:00**',
        '> Second',
      ].join('\n')
    );
  });

  it('carries a subject as a bold title, high importance in the header, an edit marker, and a reactions line ordered by count', () => {
    const md = text(
      renderThread(
        post({
          subject: ' Q4 budget ',
          importance: 'high',
          lastEditedDateTime: '2026-09-08T14:30:00Z',
          reactions: [{ reactionType: '👍' }, { reactionType: '❤️' }, { reactionType: '👍' }, {}],
        }),
        [],
        NO_IMAGES,
        false
      )
    );
    expect(md).toBe(
      ['### 2026-09-08 14:03 · Alex Kim · high importance', '', '**Q4 budget**', '', 'Budget review is **Monday**', '', '_(edited)_', '', '_reactions: 👍 2, ❤️ 1_'].join('\n')
    );
  });

  it('shows a deleted post as deleted, without an edit marker even when it was edited before', () => {
    const md = text(
      renderThread(post({ deletedDateTime: '2026-09-09T08:00:00Z', lastEditedDateTime: '2026-09-08T14:30:00Z', body: { contentType: 'html', content: '' } }), [], NO_IMAGES, false)
    );
    expect(md).toBe(['### 2026-09-08 14:03 · Alex Kim', '', '_message deleted_'].join('\n'));
  });

  it('keeps a plain-text body verbatim, quoting it line by line inside a reply, markers included', () => {
    const md = text(
      renderThread(
        post({ body: { contentType: 'text', content: 'line one\nline two' } }),
        [reply({ body: { contentType: 'text', content: 'a\n\nb' }, lastEditedDateTime: '2026-09-08T14:12:00Z' })],
        NO_IMAGES,
        false
      )
    );
    expect(md).toBe(['### 2026-09-08 14:03 · Alex Kim', '', 'line one\nline two', '', '> **Robin Chen · 2026-09-08 14:10**', '> a', '>', '> b', '>', '> _(edited)_'].join('\n'));
  });

  it('falls back to the application, then to the user id, then to an unknown sender, and keeps an unparseable date as is', () => {
    const bot = text(renderThread(post({ from: { application: { displayName: 'Planner' } }, createdDateTime: 'yesterday' }), [], NO_IMAGES, false));
    expect(bot.startsWith('### yesterday · Planner')).toBe(true);
    const idOnly = text(renderThread(post({ from: { user: { id: 'u9', displayName: null } } }), [], NO_IMAGES, false));
    expect(idOnly.startsWith('### 2026-09-08 14:03 · user u9')).toBe(true);
    const nobody = text(renderThread(post({ from: null, createdDateTime: undefined }), [], NO_IMAGES, false));
    expect(nobody.startsWith('### unknown time · unknown sender')).toBe(true);
  });

  it('orders replies without a date first, as an empty date sorts before any instant', () => {
    const md = text(renderThread(post(), [reply(), reply({ createdDateTime: undefined, body: { contentType: 'html', content: '<p>Undated</p>' } })], NO_IMAGES, false));
    expect(md.indexOf('Undated')).toBeLessThan(md.indexOf('Works for me'));
  });
});

describe('rendering a channel page as a transcript', () => {
  it('orders posts oldest first, skips system events and counts them, and names the channel', () => {
    const result = renderTranscript(
      [
        post({ id: '3', createdDateTime: '2026-09-09T09:00:00Z', body: { contentType: 'html', content: '<p>Third</p>' } }),
        post({ id: 'e', messageType: 'unknownFutureValue', eventDetail: { '@odata.type': '#microsoft.graph.membersAddedEventMessageDetail' } }),
        post({ id: '1', createdDateTime: '2026-09-08T09:00:00Z', body: { contentType: 'html', content: '<p>First</p>' }, replies: [reply()] }),
      ],
      { channelName: 'General', images: NO_IMAGES, inline: false }
    );
    expect(result.posts).toBe(2);
    expect(result.systemEventsOmitted).toBe(1);
    expect(result.text).toBe(
      [
        '# Transcript of General',
        '',
        '_Times are UTC, oldest post first._',
        '',
        '### 2026-09-08 09:00 · Alex Kim',
        '',
        'First',
        '',
        '> **Robin Chen · 2026-09-08 14:10**',
        '> Works for me',
        '',
        '### 2026-09-09 09:00 · Alex Kim',
        '',
        'Third',
      ].join('\n')
    );
  });

  it('says so when the range holds no post, with a generic title when the channel has no name', () => {
    const result = renderTranscript([post({ messageType: 'systemEventMessage' })], { images: NO_IMAGES, inline: false });
    expect(result.text).toBe(['# Channel transcript', '', '_Times are UTC, oldest post first._', '', '_No posts in this range._'].join('\n'));
    expect(result.posts).toBe(0);
    expect(result.systemEventsOmitted).toBe(1);
  });

  it('treats a blank channel name as no name', () => {
    const result = renderTranscript([post()], { channelName: '  ', images: NO_IMAGES, inline: false });
    expect(result.text.startsWith('# Channel transcript\n')).toBe(true);
  });
});

describe('the sender, date and reaction shapes a thread must survive', () => {
  it('keeps a date that does not start with an instant, and names a post with neither user nor application an unknown sender', () => {
    const md = text(renderThread(post({ createdDateTime: 'about 2026-09-08T14:03:11Z', from: { user: null, application: null } }), [], NO_IMAGES, false));
    expect(md.startsWith('### about 2026-09-08T14:03:11Z · unknown sender')).toBe(true);
  });

  it('orders reactions by count whatever their order of arrival, and renders none when the field is absent', () => {
    const md = text(renderThread(post({ reactions: [{ reactionType: '❤️' }, { reactionType: '👍' }, { reactionType: '👍' }] }), [], NO_IMAGES, false));
    expect(md.endsWith('_reactions: 👍 2, ❤️ 1_')).toBe(true);
    expect(text(renderThread(post({ reactions: undefined }), [], NO_IMAGES, false))).toBe(['### 2026-09-08 14:03 · Alex Kim', '', 'Budget review is **Monday**'].join('\n'));
  });

  it('sorts an undated reply first whatever its position among dated ones', () => {
    const md = text(
      renderThread(
        post(),
        [
          reply({ createdDateTime: '2026-09-08T16:00:00Z', body: { contentType: 'text', content: 'late' } }),
          reply({ createdDateTime: undefined, body: { contentType: 'text', content: 'undated' } }),
          reply({ body: { contentType: 'text', content: 'early' } }),
        ],
        NO_IMAGES,
        false
      )
    );
    expect(md.indexOf('undated')).toBeLessThan(md.indexOf('early'));
    expect(md.indexOf('early')).toBeLessThan(md.indexOf('late'));
  });
});
