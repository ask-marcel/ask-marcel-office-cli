import type { GraphClient } from '../../infra/graph-client.ts';

/**
 * Shared helpers for the SharePoint-link-extraction commands
 * (`extract-sharepoint-links-in-mail`, `extract-sharepoint-links-in-documents`).
 *
 * `extractSharepointUrls(text)` finds every `https://*.sharepoint.com/...`
 * URL inside a string — an HTML mail body or the joined `Target` attributes
 * of an OOXML package's external relationships — and returns the
 * deduplicated list (both `<a href=...>` and bare-text occurrences).
 *
 * `buildShareToken(url)` encodes a URL for Graph's `/shares/{token}`
 * resolver per [shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get):
 * `u!` + base64url(url) with no padding.
 *
 * `resolveSharepointUrls(graph, urls)` fans out one `/shares/{token}/driveItem`
 * resolve per URL (capped at `MAX_LINKS`, per-link errors captured in the
 * entry rather than failing the whole call) — the orchestration shared by
 * both commands.
 */

const SP_URL_PATTERN = /https:\/\/[\w-]+(?:\.[\w-]+)*\.sharepoint\.com[^\s"'<>)]*/gi;

const MAX_LINKS = 25; // Hardening #4: cap fan-out

type ResolvedLink = {
  readonly url: string;
  readonly driveId?: string;
  readonly itemId?: string;
  readonly name?: string;
  readonly webUrl?: string;
  readonly error?: string;
};

type ResolvedLinks = {
  readonly links: ReadonlyArray<ResolvedLink>;
  readonly truncated: boolean;
  readonly skippedCount: number;
};

const stripFragment = (url: string): string => {
  const hash = url.indexOf('#');
  return hash === -1 ? url : url.slice(0, hash);
};

const extractSharepointUrls = (htmlBody: string): ReadonlyArray<string> => {
  const matches = htmlBody.match(SP_URL_PATTERN) ?? [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (const raw of matches) {
    const cleaned = stripFragment(raw);
    if (seen.has(cleaned)) continue;
    seen.add(cleaned);
    out.push(cleaned);
  }
  return out;
};

const buildShareToken = (url: string): string => {
  const b64 = btoa(url).replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '');
  return `u!${b64}`;
};

const resolveOne = async (graph: GraphClient, url: string): Promise<ResolvedLink> => {
  const token = buildShareToken(url);
  const result = await graph.get(`/shares/${token}/driveItem`);
  if (!result.ok) return { url, error: result.error.type === 'api_error' ? result.error.message : `${result.error.type}: ${result.error.message}` };
  const item = result.value as { id?: string; name?: string; webUrl?: string; parentReference?: { driveId?: string } };
  return {
    url,
    driveId: item.parentReference?.driveId,
    itemId: item.id,
    name: item.name,
    webUrl: item.webUrl,
  };
};

const resolveSharepointUrls = async (graph: GraphClient, urls: ReadonlyArray<string>): Promise<ResolvedLinks> => {
  const truncated = urls.length > MAX_LINKS;
  const skippedCount = truncated ? urls.length - MAX_LINKS : 0;
  const kept = truncated ? urls.slice(0, MAX_LINKS) : urls;
  const links = await Promise.all(kept.map((u) => resolveOne(graph, u)));
  return { links, truncated, skippedCount };
};

export { buildShareToken, extractSharepointUrls, resolveSharepointUrls };
export type { ResolvedLink, ResolvedLinks };
