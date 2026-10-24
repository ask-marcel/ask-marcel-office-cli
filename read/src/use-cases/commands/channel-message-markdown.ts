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

type Reaction = NonNullable<ChannelMessage['reactions']>[number];
type People = ReadonlyMap<string, string>;

// Graph names a reactor by id, usually without a display name; the people who
// wrote a post or a reply in the same read supply the names they carry.
const peopleOf = (messages: ReadonlyArray<ChannelMessage>): People => {
  const people = new Map<string, string>();
  for (const m of messages.flatMap((p) => (p.replies === undefined ? [p] : [p, ...p.replies]))) {
    const user = m.from?.user;
    if (nonEmpty(user?.id) && nonEmpty(user.displayName)) people.set(user.id, user.displayName);
  }
  return people;
};

const reactorOf = (r: Reaction, people: People): string => {
  const user = r.user?.user;
  if (nonEmpty(user?.displayName)) return user.displayName;
  if (nonEmpty(user?.id)) return people.get(user.id) ?? `user ${user.id}`;
  return 'someone';
};

// Who reacted and when, oldest first: a reaction is the one trace a quiet
// channel shows of being read, and it bumps no post a reader would open.
const reactedLine = (m: ChannelMessage, people: People): string | undefined => {
  if (m.reactions === undefined) return undefined;
  const dated = m.reactions
    .filter((r): r is Reaction & { reactionType: string; createdDateTime: string } => nonEmpty(r.reactionType) && nonEmpty(r.createdDateTime))
    .toSorted((a, b) => a.createdDateTime.localeCompare(b.createdDateTime));
  if (dated.length === 0) return undefined;
  return `_reacted: ${dated.map((r) => [r.reactionType, reactorOf(r, people), whenOf(r.createdDateTime)].join(' ')).join(' · ')}_`;
};

/** The body plus its trailing markers, the lines a post and a reply share. */
const bodyLines = (m: ChannelMessage, images: ReadonlyMap<string, string>, inline: boolean, people: People): ReadonlyArray<string> => {
  const lines = [bodyMarkdown(m, images, inline)];
  if (!nonEmpty(m.deletedDateTime) && nonEmpty(m.lastEditedDateTime)) lines.push('', '_(edited)_');
  const reactions = reactionsLine(m);
  if (reactions !== undefined) lines.push('', reactions);
  const reacted = reactedLine(m, people);
  if (reacted !== undefined) lines.push('', reacted);
  return lines;
};

const postBlock = (m: ChannelMessage, images: ReadonlyMap<string, string>, inline: boolean, people: People): string => {
  const header = `### ${whenOf(m.createdDateTime)} · ${authorOf(m)}${m.importance === 'high' ? ' · high importance' : ''}`;
  const link = nonEmpty(m.webUrl) ? [`link: ${m.webUrl}`] : [];
  const title = nonEmpty(m.subject) ? [`**${m.subject.trim()}**`, ''] : [];
  return [header, ...link, '', ...title, ...bodyLines(m, images, inline, people)].join('\n');
};

const replyBlock = (m: ChannelMessage, images: ReadonlyMap<string, string>, inline: boolean, people: People): string => {
  const quoted = bodyLines(m, images, inline, people)
    .flatMap((l) => l.split('\n'))
    .map((l) => (l === '' ? '>' : `> ${l}`));
  const link = nonEmpty(m.webUrl) ? [`> link: ${m.webUrl}`] : [];
  return [`> **${authorOf(m)} · ${whenOf(m.createdDateTime)}**`, ...link, ...quoted].join('\n');
};

const byCreated = (a: ChannelMessage, b: ChannelMessage): number => (a.createdDateTime ?? '').localeCompare(b.createdDateTime ?? '');

/** One root post followed by its replies, oldest reply first. */
const renderThread = (
  root: ChannelMessage,
  replies: ReadonlyArray<ChannelMessage>,
  images: ReadonlyMap<string, string>,
  inline: boolean,
  people: People = peopleOf([root, ...replies])
): string => [postBlock(root, images, inline, people), ...replies.toSorted(byCreated).map((r) => replyBlock(r, images, inline, people))].join('\n\n');

/** A page of root posts (each with its expanded replies) as one dated transcript, oldest post first; system events are counted, not shown. */
const renderTranscript = (messages: ReadonlyArray<ChannelMessage>, options: TranscriptOptions): Transcript => {
  const posts = messages.filter(isChannelPost).toSorted(byCreated);
  const systemEventsOmitted = messages.length - posts.length;
  const header = [nonEmpty(options.channelName) ? `# Transcript of ${options.channelName}` : '# Channel transcript', '', '_Times are UTC, oldest post first._'].join('\n');
  const people = peopleOf(posts);
  const blocks = posts.length === 0 ? ['_No posts in this range._'] : posts.map((p) => renderThread(p, p.replies ?? [], options.images, options.inline, people));
  return { text: [header, ...blocks].join('\n\n'), posts: posts.length, systemEventsOmitted };
};

export { renderThread, renderTranscript };
export type { Transcript };
