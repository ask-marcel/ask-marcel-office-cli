import { parseJson } from '../../domain/json.ts';
import { htmlToMarkdown } from '../../infra/turndown-adapter.ts';

/**
 * The body of a Teams channel message (Graph `chatMessage`) as markdown.
 * Teams HTML carries its own tags on top of the usual ones: `<at id>` for a
 * mention, `<attachment id>` as the placeholder of an entry in
 * `attachments[]`, `<emoji alt>`, and `<img src=".../hostedContents/{id}/$value">`
 * for a pasted image only an authenticated Graph call can read. Each is
 * resolved before the HTML reaches turndown; the markdown it becomes is parked
 * in an alphanumeric slot turndown cannot escape, then restored. The empty
 * `<systemEventMessage/>` of a membership event needs nothing: turndown drops
 * an unknown empty tag, and events are filtered before rendering anyway.
 */

type ChannelAttachment = {
  readonly id?: string | null;
  readonly contentType?: string | null;
  readonly contentUrl?: string | null;
  readonly name?: string | null;
  readonly content?: string | null;
};

type ChannelMessage = {
  readonly id?: string;
  readonly replyToId?: string | null;
  readonly messageType?: string;
  readonly createdDateTime?: string;
  readonly lastEditedDateTime?: string | null;
  readonly deletedDateTime?: string | null;
  readonly subject?: string | null;
  readonly importance?: string;
  readonly from?: {
    readonly user?: { readonly id?: string; readonly displayName?: string | null } | null;
    readonly application?: { readonly displayName?: string | null } | null;
  } | null;
  readonly body?: { readonly contentType?: string; readonly content?: string };
  readonly attachments?: ReadonlyArray<ChannelAttachment>;
  readonly mentions?: ReadonlyArray<{ readonly id?: number; readonly mentionText?: string | null }>;
  readonly reactions?: ReadonlyArray<{ readonly reactionType?: string }>;
  readonly eventDetail?: { readonly '@odata.type'?: string };
  readonly replies?: ReadonlyArray<ChannelMessage>;
};

const CARD_TEXT_LIMIT = 300;
const IMAGE_OFF_PLACEHOLDER = '[image: hosted in Teams, pass --inline-images true to embed]';
const IMAGE_MISSING_PLACEHOLDER = '[image: could not be embedded]';

const ID_ATTR = /\bid\s*=\s*(?:"([^"]*)"|'([^']*)')/i;
const SRC_ATTR = /\bsrc\s*=\s*(?:"([^"]*)"|'([^']*)')/i;
const ALT_ATTR = /\balt\s*=\s*(?:"([^"]*)"|'([^']*)')/i;
const IMG_TAG = /<img\b([^>]*)>/gi;
const AT_TAG = /<at\b[^>]*>([\s\S]*?)<\/at>/gi;
const ATTACHMENT_TAG = /<attachment\b([^>]*?)(?:\/>|>[\s\S]*?<\/attachment>)/gi;
const EMOJI_TAG = /<emoji\b([^>]*?)(?:\/>|>[\s\S]*?<\/emoji>)/gi;
const SLOT = /AMSLOT(\d+)X/g;

const nonEmpty = (s: string | null | undefined): s is string => typeof s === 'string' && s.trim() !== '';
const attr = (pattern: RegExp, attrs: string): string | undefined => {
  const m = pattern.exec(attrs);
  return m === null ? undefined : (m[1] ?? m[2]);
};

/** A post is `messageType: message`; membership and channel events come back as `unknownFutureValue` (or `systemEventMessage`) with an `eventDetail`. */
const isChannelPost = (m: ChannelMessage): boolean => m.messageType === 'message';

// Arrays are objects whose entries are their items, so one walk covers both.
const collectCardText = (node: unknown, out: string[]): void => {
  if (node === null || typeof node !== 'object') return;
  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if ((key === 'title' || key === 'text') && nonEmpty(value as string | null | undefined)) out.push((value as string).trim());
    else collectCardText(value, out);
  }
};

const cardMarkdown = (a: ChannelAttachment): string => {
  if (!nonEmpty(a.content)) return '[card]';
  const parsed = parseJson(a.content);
  if (!parsed.ok) return '[card: unreadable]';
  const texts: string[] = [];
  collectCardText(parsed.value, texts);
  return texts.length === 0 ? '[card]' : `[card: ${texts.join(' / ').slice(0, CARD_TEXT_LIMIT)}]`;
};

const attachmentMarkdown = (a: ChannelAttachment | undefined): string => {
  if (a === undefined) return '[attachment]';
  const name = nonEmpty(a.name) ? a.name : 'attachment';
  const type = nonEmpty(a.contentType) ? a.contentType : 'unknown';
  if (type === 'reference' && nonEmpty(a.contentUrl)) return `[${name}](${a.contentUrl})`;
  if (type === 'meetingReference') return `[meeting: ${name}]`;
  if (type === 'tabReference') return `[tab: ${name}]`;
  if (type.startsWith('application/vnd.microsoft.card')) return cardMarkdown(a);
  return `[attachment: ${nonEmpty(a.name) ? a.name : type}]`;
};

const mentionText = (m: ChannelMessage, attrs: string, inner: string): string => {
  if (nonEmpty(inner)) return `@${inner.trim()}`;
  const id = Number(attr(ID_ATTR, attrs));
  const found = m.mentions?.find((x) => x.id === id);
  return nonEmpty(found?.mentionText) ? `@${found.mentionText}` : '@mention';
};

const imageHtml = (attrs: string, images: ReadonlyMap<string, string>, inline: boolean, park: (md: string) => string): string => {
  const src = attr(SRC_ATTR, attrs);
  if (src === undefined || !src.includes('/hostedContents/')) return `<img${attrs}>`;
  if (!inline) return park(IMAGE_OFF_PLACEHOLDER);
  const dataUri = images.get(src);
  if (dataUri === undefined) return park(IMAGE_MISSING_PLACEHOLDER);
  const alt = attr(ALT_ATTR, attrs) ?? '';
  return `<img src="${dataUri}" alt="${alt}">`;
};

/**
 * Teams-specific tags become markdown parked in slots; everything else stays
 * HTML for turndown. An attachment placeholder gets a leading space because
 * Teams glues the tag to the text before it (`Scheduled a meeting<attachment`),
 * and HTML whitespace collapsing absorbs the double space when there was one.
 */
const prepareHtml = (m: ChannelMessage, images: ReadonlyMap<string, string>, inline: boolean): { readonly html: string; readonly slots: ReadonlyArray<string> } => {
  const slots: string[] = [];
  const park = (md: string): string => {
    slots.push(md);
    return `AMSLOT${slots.length - 1}X`;
  };
  const byId = new Map<string, ChannelAttachment>();
  const attachments = m.attachments;
  if (attachments !== undefined) for (const a of attachments) if (nonEmpty(a.id)) byId.set(a.id, a);
  const html = (m.body?.content ?? '')
    .replace(AT_TAG, (_tag, inner: string) => park(mentionText(m, _tag, inner)))
    .replace(ATTACHMENT_TAG, (_tag, attrs: string) => {
      const id = attr(ID_ATTR, attrs);
      return ` ${park(attachmentMarkdown(id === undefined ? undefined : byId.get(id)))}`;
    })
    .replace(EMOJI_TAG, (_tag, attrs: string) => park(attr(ALT_ATTR, attrs) ?? ''))
    .replace(IMG_TAG, (_tag, attrs: string) => imageHtml(attrs, images, inline, park));
  return { html, slots };
};

// A slot with no parked markdown is the user's own text, kept as written.
const restoreSlots = (md: string, slots: ReadonlyArray<string>): string => md.replace(SLOT, (whole, index: string) => slots[Number(index)] ?? whole);

/**
 * The body as markdown: a deleted message says so, a plain-text body is kept
 * verbatim, an HTML body goes through the Teams tags and turndown. Turndown
 * never fails here (it degrades to stripped text on its own), so the
 * conversion is unwrapped rather than threaded as a Result.
 */
const bodyMarkdown = (m: ChannelMessage, images: ReadonlyMap<string, string>, inline: boolean): string => {
  if (nonEmpty(m.deletedDateTime)) return '_message deleted_';
  if (m.body?.contentType === 'text') return m.body.content ?? '';
  const prepared = prepareHtml(m, images, inline);
  const converted = htmlToMarkdown(prepared.html);
  return restoreSlots(converted.ok ? converted.value : '', prepared.slots);
};

export { bodyMarkdown, isChannelPost, nonEmpty };
export type { ChannelAttachment, ChannelMessage };
