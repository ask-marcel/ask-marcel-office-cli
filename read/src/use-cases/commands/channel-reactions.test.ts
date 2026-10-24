import { describe, expect, it } from 'bun:test';
import type { ChannelMessage } from './channel-message-html.ts';
import { renderThread, renderTranscript } from './channel-message-markdown.ts';

const NO_IMAGES: ReadonlyMap<string, string> = new Map();

const post = (over: Partial<ChannelMessage> = {}): ChannelMessage => ({
  id: '1700000000000',
  replyToId: null,
  messageType: 'message',
  createdDateTime: '2025-12-02T09:00:00Z',
  from: { user: { id: 'u1', displayName: 'Alex Kim' } },
  body: { contentType: 'html', content: '<p>Year-end party</p>' },
  ...over,
});

// Graph gives a reactor's id and, usually, no name: the name comes from the
// posts in the same read when that person wrote one.
const REACTIONS: ChannelMessage['reactions'] = [
  { reactionType: '❤️', createdDateTime: '2026-09-23T08:02:00Z', user: { user: { id: 'u9', displayName: null } } },
  { reactionType: '👍', createdDateTime: '2026-09-22T10:14:30Z', user: { user: { id: 'u2', displayName: null } } },
  { reactionType: '👍', createdDateTime: '2026-09-22T11:00:00Z', user: { user: { id: 'u3', displayName: 'Jordan Avery' } } },
  { reactionType: '😮', createdDateTime: '2026-09-22T12:00:00Z', user: null },
];

describe('reactions on a channel post', () => {
  it('keep their counts and gain who reacted and when, oldest first, named from the posts in the same read', () => {
    const replyByRobin = post({
      id: '1700000000001',
      replyToId: '1700000000000',
      createdDateTime: '2026-07-01T09:00:00Z',
      from: { user: { id: 'u2', displayName: 'Robin Chen' } },
    });
    const md = renderThread(post({ reactions: REACTIONS }), [replyByRobin], NO_IMAGES, false);
    expect(md).toContain(
      '_reactions: 👍 2, ❤️ 1, 😮 1_\n\n_reacted: 👍 Robin Chen 2026-09-22 10:14 · 👍 Jordan Avery 2026-09-22 11:00 · 😮 someone 2026-09-22 12:00 · ❤️ user u9 2026-09-23 08:02_'
    );
  });

  it('name reactors from every post of a transcript, and add nothing when Graph gives no reaction time', () => {
    const byRobin = post({ id: '1700000000002', createdDateTime: '2026-01-05T09:00:00Z', from: { user: { id: 'u2', displayName: 'Robin Chen' } } });
    const { text } = renderTranscript([post({ reactions: [REACTIONS[1] ?? {}] }), byRobin], { images: NO_IMAGES, inline: false });
    expect(text).toContain('_reacted: 👍 Robin Chen 2026-09-22 10:14_');
    const robinReplied = post({
      reactions: [REACTIONS[1] ?? {}],
      replies: [post({ id: '1700000000003', replyToId: '1700000000000', from: { user: { id: 'u2', displayName: 'Robin Chen' } } })],
    });
    expect(renderTranscript([robinReplied], { images: NO_IMAGES, inline: false }).text).toContain('_reacted: 👍 Robin Chen 2026-09-22 10:14_');
    const undated = renderThread(post({ reactions: [{ reactionType: '👍' }] }), [], NO_IMAGES, false);
    expect(undated).toContain('_reactions: 👍 1_');
    expect(undated).not.toContain('_reacted:');
  });
});
