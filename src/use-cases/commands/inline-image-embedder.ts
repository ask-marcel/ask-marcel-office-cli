/**
 * Email body HTML often references inline attachments via `cid:<id>`
 * URLs (e.g. `<img src="cid:logo123">`). Those `cid:` refs are
 * meaningless outside the original mail viewer. To make markdown /
 * HTML output self-contained we replace each `cid:<id>` with a base64
 * `data:` URI sourced from the matching inline attachment.
 *
 * Hardening #1: only `image/*` content-types are embedded. Anything
 * else (`text/html`, `application/javascript`, ...) would let an
 * attacker turn an attachment into an executable payload via the
 * data URI; skip and leave the original `cid:` ref in place.
 */

type InlineAttachment = {
  readonly contentId: string;
  readonly contentType: string;
  readonly contentBytes: string;
};

const embedInlineImages = (html: string, attachments: ReadonlyArray<InlineAttachment>): string => {
  let out = html;
  for (const a of attachments) {
    if (a.contentId === '') continue;
    if (!a.contentType.toLowerCase().startsWith('image/')) continue;
    const dataUri = `data:${a.contentType};base64,${a.contentBytes}`;
    // String.replaceAll(searchValue: string, replaceValue) does a literal
    // global replace — no regex, so contentId metacharacters (`.`, `+`,
    // etc.) are treated literally with no escape step needed.
    out = out.replaceAll(`cid:${a.contentId}`, dataUri);
  }
  return out;
};

const IMG_TAG = /<img\b[^>]*>/gi;
const SRC_ATTR = /\bsrc\s*=\s*(?:"([^"]*)"|'([^']*)')/i;
const ALT_ATTR = /\balt\s*=\s*(?:"([^"]*)"|'([^']*)')/i;

// A `cid:` ref that survives the embed pass (embedding disabled, an
// individual fetch failed, a non-image inline attachment, or a cid with no
// matching attachment) would render as a broken `![…](cid:…)` markdown link.
// Swap the whole <img> tag for a readable placeholder instead, labelled by
// the attachment name when known, else the img alt text, else the cid's
// filename-ish prefix before the `@`.
const replaceUnresolvedCidImages = (html: string, labelByContentId: ReadonlyMap<string, string>): string =>
  html.replace(IMG_TAG, (tag) => {
    const src = SRC_ATTR.exec(tag);
    const srcValue = src?.[1] ?? src?.[2];
    if (srcValue === undefined || !srcValue.startsWith('cid:')) return tag;
    const cid = srcValue.slice('cid:'.length);
    const alt = ALT_ATTR.exec(tag);
    const altText = alt?.[1] ?? alt?.[2];
    const cidPrefix = cid.split('@')[0];
    const label = labelByContentId.get(cid) ?? (altText !== undefined && altText !== '' ? altText : undefined) ?? (cidPrefix !== undefined && cidPrefix !== '' ? cidPrefix : cid);
    return `[inline image: ${label}]`;
  });

export { embedInlineImages, replaceUnresolvedCidImages };
export type { InlineAttachment };
