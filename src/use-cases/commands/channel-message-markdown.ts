import { bodyMarkdown, isChannelPost, nonEmpty, type ChannelMessage } from './channel-message-html.ts';

/**
 * A Teams channel post with its replies as a markdown thread, or a page of
 * posts as a dated transcript, oldest first. Bodies come from
 * `channel-message-html.ts`; this module owns the headings, the quoted
 * replies and the markers (subject, importance, edit, deletion, reactions).
 */

type TranscriptOptions = { readonly channelName?: string; readonly images: ReadonlyMap<string, string>; readonly inline: boolean };
type Transcript = { readonly text: string; readonly posts: number; readonly systemEventsOmitted: number };

const ISO_MINUTE = /^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})/;

const authorOf = (m: ChannelMessage): string => {
  if (nonEmpty(m.from?.user?.displayName)) return m.from.user.displayName;
  if (nonEmpty(m.from?.application?.displayName)) return m.from.application.displayName;
  if (nonEmpty(m.from?.user?.id)) return `user ${m.from.user.id}`;
  return 'unknown sender';
};

const whenOf = (iso: string | undefined): string => {
  if (iso === undefined) return 'unknown time';
  const m = ISO_MINUTE.exec(iso);
  return m === null ? iso : `${m[1]} ${m[2]}`;
};

const reactionsLine = (m: ChannelMessage): string | undefined => {
  const counts = new Map<string, number>();
  if (m.reactions !== undefined) for (const r of m.reactions) if (nonEmpty(r.reactionType)) counts.set(r.reactionType, (counts.get(r.reactionType) ?? 0) + 1);
  if (counts.size === 0) return undefined;
  const ordered = [...counts.entries()].toSorted((a, b) => b[1] - a[1]).map(([type, n]) => `${type} ${n}`);
  return `_reactions: ${ordered.join(', ')}_`;
};

/** The body plus its trailing markers, the lines a post and a reply share. */
const bodyLines = (m: ChannelMessage, images: ReadonlyMap<string, string>, inline: boolean): ReadonlyArray<string> => {
  const lines = [bodyMarkdown(m, images, inline)];
  if (!nonEmpty(m.deletedDateTime) && nonEmpty(m.lastEditedDateTime)) lines.push('', '_(edited)_');
  const reactions = reactionsLine(m);
  if (reactions !== undefined) lines.push('', reactions);
  return lines;
};

const postBlock = (m: ChannelMessage, images: ReadonlyMap<string, string>, inline: boolean): string => {
  const header = `### ${whenOf(m.createdDateTime)} · ${authorOf(m)}${m.importance === 'high' ? ' · high importance' : ''}`;
  const title = nonEmpty(m.subject) ? [`**${m.subject.trim()}**`, ''] : [];
  return [header, '', ...title, ...bodyLines(m, images, inline)].join('\n');
};

const replyBlock = (m: ChannelMessage, images: ReadonlyMap<string, string>, inline: boolean): string => {
  const quoted = bodyLines(m, images, inline)
    .flatMap((l) => l.split('\n'))
    .map((l) => (l === '' ? '>' : `> ${l}`));
  return [`> **${authorOf(m)} · ${whenOf(m.createdDateTime)}**`, ...quoted].join('\n');
};

const byCreated = (a: ChannelMessage, b: ChannelMessage): number => (a.createdDateTime ?? '').localeCompare(b.createdDateTime ?? '');

/** One root post followed by its replies, oldest reply first. */
const renderThread = (root: ChannelMessage, replies: ReadonlyArray<ChannelMessage>, images: ReadonlyMap<string, string>, inline: boolean): string =>
  [postBlock(root, images, inline), ...replies.toSorted(byCreated).map((r) => replyBlock(r, images, inline))].join('\n\n');

/** A page of root posts (each with its expanded replies) as one dated transcript, oldest post first; system events are counted, not shown. */
const renderTranscript = (messages: ReadonlyArray<ChannelMessage>, options: TranscriptOptions): Transcript => {
  const posts = messages.filter(isChannelPost).toSorted(byCreated);
  const systemEventsOmitted = messages.length - posts.length;
  const header = [nonEmpty(options.channelName) ? `# Transcript of ${options.channelName}` : '# Channel transcript', '', '_Times are UTC, oldest post first._'].join('\n');
  const blocks = posts.length === 0 ? ['_No posts in this range._'] : posts.map((p) => renderThread(p, p.replies ?? [], options.images, options.inline));
  return { text: [header, ...blocks].join('\n\n'), posts: posts.length, systemEventsOmitted };
};

export { renderThread, renderTranscript };
export type { Transcript };
