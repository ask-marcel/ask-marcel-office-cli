import type { GraphClient } from '../../infra/graph-client.ts';

/**
 * OneNote page HTML references images as absolute Graph resource URLs:
 *   `<img src="https://graph.microsoft.com/v1.0/users/…/onenote/resources/{id}/$value" …>`
 * Unlike mail (which uses `cid:` refs), these are live Graph endpoints, so the
 * markdown isn't self-contained until each one is fetched and inlined as a
 * base64 `data:` URI.
 *
 * Per-image isolation (mirrors the mail inline-image embedder): any failure —
 * fetch error, a non-image content-type, an oversize image, or a 302 redirect
 * that comes back without inline bytes — leaves the original URL in place. One
 * bad resource never sinks the whole page.
 *
 * File attachments are embedded by OneNote as
 *   `<object data="…/onenote/resources/{id}/$value" data-attachment="name" type="mime">`.
 * turndown drops `<object>`, so those would vanish; instead each is rewritten to
 * a visible `[OneNote attachment: name (mime)]` annotation (bytes are NOT fetched
 * — the resource endpoint is auth-gated, so surfacing the file's existence is the
 * useful signal for an LLM). This runs before image embedding.
 */

// Bare-URL match (no capture group) so `String.match(/g)` yields a clean
// `string[]` — no `string | undefined` group narrowing. Only `/v1.0` (the
// version `graph.getBinary` itself targets) is matched.
const RESOURCE_URL = /https:\/\/graph\.microsoft\.com\/v1\.0\/[^"\s]*\/onenote\/resources\/[^"\s]*\/\$value/gi;
const OBJECT_EL = /<object\b([^>]*\bdata="https:\/\/graph\.microsoft\.com\/v1\.0\/[^"]*\/onenote\/resources\/[^"]*\/\$value"[^>]*)>(?:\s*<\/object>)?/gi;
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const IMAGE_SIZE_LIMIT_BYTES = 2_000_000;

// Literal patterns (not `new RegExp(name)`) so the attribute name is never an
// injection vector — only these two attributes are ever read off an `<object>`.
// Non-global, so `.exec` is stateless and the shared instances are reuse-safe.
const DATA_ATTACHMENT_ATTR = /\bdata-attachment="([^"]*)"/;
const TYPE_ATTR = /\btype="([^"]*)"/;
const attrValue = (attrs: string, pattern: RegExp): string | undefined => pattern.exec(attrs)?.[1];

const annotateOnenoteObjects = (html: string): string =>
  html.replace(OBJECT_EL, (_match, attrs: string) => {
    const name = attrValue(attrs, DATA_ATTACHMENT_ATTR) ?? 'file';
    const type = attrValue(attrs, TYPE_ATTR);
    const suffix = type === undefined ? '' : ` (${type})`;
    return `<p>[OneNote attachment: ${name}${suffix}]</p>`;
  });

const embedOnenoteResources = async (graph: GraphClient, html: string): Promise<string> => {
  // Rewrite <object> file attachments first so their resource URLs aren't then
  // treated as candidate <img> sources.
  const annotated = annotateOnenoteObjects(html);
  const urls = Array.from(new Set(annotated.match(RESOURCE_URL) ?? []));
  let out = annotated;
  for (const url of urls) {
    // The regex guarantees the `GRAPH_BASE` prefix, so the slice yields the
    // request path `getBinary` expects (it re-prepends the base internally).
    const fetched = await graph.getBinary(url.slice(GRAPH_BASE.length));
    if (!fetched.ok) continue;
    const value = fetched.value as { readonly contentType?: string; readonly size?: number; readonly base64?: string };
    if (typeof value.base64 !== 'string' || typeof value.contentType !== 'string') continue;
    if (!value.contentType.toLowerCase().startsWith('image/')) continue;
    if ((value.size ?? 0) > IMAGE_SIZE_LIMIT_BYTES) continue;
    out = out.replaceAll(url, `data:${value.contentType};base64,${value.base64}`);
  }
  return out;
};

export { embedOnenoteResources };
