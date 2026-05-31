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

  // 2. Open as a zip package. A non-zip input (pdf, image, anything that isn't a
  //    ZIP) can't carry extractable hyperlinks → friendly 400.
  const zip = await openOoxmlZip(bytes.value);
  if (!zip.ok) {
    return err({
      type: 'api_error',
      status: 400,
      message:
        'not a zip-based Office document — only OOXML (.docx/.xlsx/.pptx and their macro-enabled / template variants) and OpenDocument (.odt/.ods/.odp) carry extractable hyperlinks. pdf, images, and plain binaries have none.',
    });
  }

  // 3. Collect candidate URLs. OOXML keeps external hyperlinks in the package's
  //    relationship parts (`_rels/*.rels`, TargetMode="External"); OpenDocument
  //    keeps them inline in content.xml / styles.xml as `xlink:href` attributes
  //    (the SharePoint regex pulls them straight out of the raw XML). Then keep
  //    the SharePoint ones and resolve each.
  const isOdf = (zip.value.read('mimetype') ?? '').startsWith('application/vnd.oasis.opendocument');
  const haystack = isOdf
    ? ['content.xml', 'styles.xml'].map((part) => zip.value.read(part) ?? '').join('\n')
    : extractExternalRels(zip.value)
        .map((rel) => rel.target)
        .join('\n');
  const { links, truncated, skippedCount } = await resolveSharepointUrls(graph, extractSharepointUrls(haystack));

  return ok({ driveId, itemId, links, truncated, skippedCount });
};

const meta: CommandMeta = {
  summary:
    'Find every `*.sharepoint.com` URL embedded in a Word / Excel / PowerPoint or OpenDocument file on OneDrive or SharePoint and resolve each one to its driveItem (driveId, itemId, name, webUrl) so the agent can feed those into `download-drive-item-as-pdf` / `-as-markdown` etc. The document sibling of `extract-sharepoint-links-in-mail`. For OOXML (.docx/.xlsx/.pptx) it reads external hyperlinks from the package’s relationship parts (`_rels/*.rels`, `TargetMode="External"`); for OpenDocument (.odt/.ods/.odp) it reads the inline `xlink:href` links in content.xml / styles.xml — either way it catches links wherever they live (body text, headers/footers, cell formulas, slide shapes). Read-only — no conversion happens here. Capped at 25 unique URLs per call (returns `truncated: true` and `skippedCount` when there are more); duplicates are deduplicated; per-link errors are captured inside each entry instead of failing the whole call. Non-zip inputs (pdf/images) return an api_error.',
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
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID of the .docx/.xlsx/.pptx or .odt/.ods/.odp file. Returned by `list-folder-files` or `search-onedrive-files`.',
    },
  ],
  example: "ask-marcel extract-sharepoint-links-in-documents --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ driveId, itemId, links: [{ url, driveId, itemId, name, webUrl } | { url, error }], truncated, skippedCount }` — one entry per unique SharePoint URL found in the document’s external relationships.',
};

export { execute, meta, schema };
export type { DocumentLinkSummary };
