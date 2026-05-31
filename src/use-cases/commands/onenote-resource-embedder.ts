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
 * bad resource never sinks the whole page. Only `<img src>` resources are
 * embedded here; `<object>` file attachments are out of scope (noted follow-up).
 */

// Bare-URL match (no capture group) so `String.match(/g)` yields a clean
// `string[]` — no `string | undefined` group narrowing. Only `/v1.0` (the
// version `graph.getBinary` itself targets) is matched.
const RESOURCE_URL = /https:\/\/graph\.microsoft\.com\/v1\.0\/[^"\s]*\/onenote\/resources\/[^"\s]*\/\$value/gi;
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const IMAGE_SIZE_LIMIT_BYTES = 2_000_000;

const embedOnenoteResources = async (graph: GraphClient, html: string): Promise<string> => {
  const urls = Array.from(new Set(html.match(RESOURCE_URL) ?? []));
  let out = html;
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
