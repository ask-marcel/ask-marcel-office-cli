import type { ReadGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { nonEmpty } from './channel-message-html.ts';
import type { ChannelMessage } from './channel-message-html.ts';

/**
 * Pasted images in a channel message are `<img src="https://graph.microsoft.com/v1.0/teams/.../hostedContents/{id}/$value">`,
 * readable only with the bearer. The renderer embeds them on request, the
 * way `convert-mail-to-markdown` embeds `cid:` images: image/* only, 2 MB or
 * less, anything else left to a placeholder.
 */

const HOSTED_IMAGE_SIZE_LIMIT_BYTES = 2_000_000;
const GRAPH_ORIGIN = /^https:\/\/graph\.microsoft\.com\/(?:v1\.0|beta)/;
const IMG_TAG = /<img\b([^>]*)>/gi;
const SRC_ATTR = /\bsrc\s*=\s*(?:"([^"]*)"|'([^']*)')/i;

/** Every distinct hosted-image URL across the posts and their inlined replies, in order of appearance. */
// The `src` of each `<img>` tag of one body, in order.
const imageSources = (html: string): ReadonlyArray<string> =>
  Array.from(html.matchAll(IMG_TAG), (tag) => {
    const match = tag[1] === undefined ? null : SRC_ATTR.exec(tag[1]);
    return match === null ? undefined : (match[1] ?? match[2]);
  }).filter((src): src is string => src !== undefined);

const hostedImageSources = (messages: ReadonlyArray<ChannelMessage>): ReadonlyArray<string> => {
  const found: string[] = [];
  const visit = (m: ChannelMessage): void => {
    for (const src of imageSources(m.body?.content ?? '')) {
      if (src.includes('/hostedContents/') && !found.includes(src)) found.push(src);
    }
    if (m.replies !== undefined) for (const r of m.replies) visit(r);
  };
  for (const m of messages) visit(m);
  return found;
};

/** Fetches each hosted image through Graph; only `image/*` bodies at or under 2 MB become data URIs. */
const fetchHostedImages = async (graph: ReadGraph, sources: ReadonlyArray<string>): Promise<ReadonlyMap<string, string>> => {
  const images = new Map<string, string>();
  for (const src of sources) {
    if (!GRAPH_ORIGIN.test(src)) continue;
    const fetched = await graph.getBinary(src.replace(GRAPH_ORIGIN, ''));
    if (!fetched.ok) continue;
    const body = fetched.value as { readonly contentType?: string; readonly size?: number; readonly base64?: string };
    if (!nonEmpty(body.contentType) || !body.contentType.toLowerCase().startsWith('image/')) continue;
    if ((body.size ?? 0) > HOSTED_IMAGE_SIZE_LIMIT_BYTES || !nonEmpty(body.base64)) continue;
    images.set(src, `data:${body.contentType};base64,${body.base64}`);
  }
  return images;
};

export { fetchHostedImages, hostedImageSources };
