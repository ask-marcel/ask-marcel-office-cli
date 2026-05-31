import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import type { CommandMeta } from './command-types.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { extractExternalRels } from './ooxml-metadata.ts';
import { extractSharepointUrls, resolveSharepointUrls } from './sharepoint-link-extractor.ts';
import type { ResolvedLink } from './sharepoint-link-extractor.ts';

/**
 * The document-side sibling of `extract-sharepoint-links-in-mail`: pulls a
 * docx/xlsx/pptx from a OneDrive / SharePoint item and finds every
 * `*.sharepoint.com` URL embedded in it, resolving each to its driveItem so an
 * agent can follow references out of a document the same way it follows them
 * out of an email. Read-only — no conversion happens here.
 *
 * In OOXML, external hyperlinks live in the `_rels/*.rels` parts as
 * `<Relationship Target="…" TargetMode="External"/>`. `extractExternalRels`
 * (shared with the metadata extractors) already surfaces exactly those, so this
 * command is pure composition: fetch bytes → open the package → collect external
 * targets → keep the SharePoint ones → resolve via `/shares/{token}/driveItem`.
 */

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1) });

type DocumentLinkSummary = {
  readonly driveId: string;
  readonly itemId: string;
  readonly links: ReadonlyArray<ResolvedLink>;
  readonly truncated: boolean;
  readonly skippedCount: number;
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<DocumentLinkSummary, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;

  // 1. Pull the raw document bytes.
  const bytes = await fetchRawBytes(graph, `/drives/${driveId}/items/${itemId}/content`);
  if (!bytes.ok) return bytes;

  // 2. Open as an OOXML package. A non-OOXML input (pdf, image, anything that
  //    isn't a ZIP) can't carry relationship-based hyperlinks → friendly 400.
  const zip = await openOoxmlZip(bytes.value);
  if (!zip.ok) {
    return err({
      type: 'api_error',
      status: 400,
      message:
        'not an Office Open XML document — only .docx/.xlsx/.pptx (and their macro-enabled / template variants) carry relationship-based external hyperlinks. OpenDocument (.odt/.ods/.odp) stores links inside content.xml and is not yet supported here.',
    });
  }

  // 3. Collect external relationship targets, keep the SharePoint ones, resolve each.
  const targets = extractExternalRels(zip.value).map((rel) => rel.target);
  const { links, truncated, skippedCount } = await resolveSharepointUrls(graph, extractSharepointUrls(targets.join('\n')));

  return ok({ driveId, itemId, links, truncated, skippedCount });
};

const meta: CommandMeta = {
  summary:
    'Find every `*.sharepoint.com` URL embedded in a Word / Excel / PowerPoint document on OneDrive or SharePoint and resolve each one to its driveItem (driveId, itemId, name, webUrl) so the agent can feed those into `download-drive-item-as-pdf` / `-as-markdown` etc. The document sibling of `extract-sharepoint-links-in-mail`. Reads external hyperlinks from the OOXML package’s relationship parts (`_rels/*.rels`, `TargetMode="External"`), so it catches links wherever they live (body text, headers/footers, cell formulas, slide shapes). Read-only — no conversion happens here. Capped at 25 unique URLs per call (returns `truncated: true` and `skippedCount` when there are more); duplicates are deduplicated; per-link errors are captured inside each entry instead of failing the whole call. Non-OOXML inputs (pdf/images) and OpenDocument (.odt/.ods/.odp) return an api_error.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description:
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the .docx/.xlsx/.pptx file. Returned by `list-folder-files` or `search-onedrive-files`.' },
  ],
  example: "ask-marcel extract-sharepoint-links-in-documents --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ driveId, itemId, links: [{ url, driveId, itemId, name, webUrl } | { url, error }], truncated, skippedCount }` — one entry per unique SharePoint URL found in the document’s external relationships.',
};

export { execute, meta, schema };
export type { DocumentLinkSummary };
