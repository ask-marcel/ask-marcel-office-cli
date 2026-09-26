import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { commentsFromBytes } from './document-comments.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';
import { DRIVE_ID_DESCRIPTION } from './option-descriptions.ts';
import { TENANT_ID_OPTION, brandTenantId, tenantIdShape } from './tenant-option.ts';

const schema = z.object({ driveId: z.string().min(1), itemId: z.string().min(1), ...tenantIdShape });

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const tenant = parsed.data.tenantId === undefined ? undefined : brandTenantId(parsed.data.tenantId);
  if (tenant !== undefined && !tenant.ok) return tenant;
  const tenantId = tenant?.ok === true ? tenant.value : undefined;

  // The name decides the reader, and a folder has no comments to read.
  const metaPath = `/drives/${driveId}/items/${itemId}`;
  const meta = tenantId === undefined ? await graph.get(metaPath) : await graph.getGuest(metaPath, tenantId);
  if (!meta.ok) return meta;
  const item = meta.value as { name?: string; folder?: unknown };
  const name = item.name ?? '';
  if (item.folder !== undefined && item.folder !== null) {
    return err({
      type: 'api_error',
      status: 400,
      message: `item '${name}' is a folder, not a file — use \`list-folder-files --drive-id ${driveId} --item-id ${itemId}\` to enumerate its children, then pick a document from inside it.`,
    });
  }

  const bytes = await fetchRawBytes(graph, `/drives/${driveId}/items/${itemId}/content`, { tenantId });
  if (!bytes.ok) return bytes;
  const comments = await commentsFromBytes(bytes.value, name);
  if (!comments.ok) return comments;
  return ok({ name, ...comments.value });
};

const meta: CommandMeta = {
  summary:
    'List the comments in a Word, Excel or PowerPoint file on OneDrive / SharePoint as one flat list: who wrote each and when, where it sits (the commented text in a document, the cell as `Sheet!A1` in a workbook, the slide in a deck), what it says, and the people it @-mentions. Covers Word comments and replies, Excel notes and threaded comments, and PowerPoint legacy and modern comments, macro-enabled and template variants included. The legacy copy Excel keeps of each threaded comment is dropped, so every comment appears once. For the comments together with the rest of the side-channel metadata, `download-drive-item-as-markdown` has an include-metadata switch.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/content',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-get-content',
  options: [
    { name: 'drive-id', key: 'driveId', required: true, description: DRIVE_ID_DESCRIPTION },
    {
      name: 'item-id',
      key: 'itemId',
      required: true,
      description: 'driveItem ID of the docx / xlsx / pptx file. Returned by `list-folder-files`, `search-onedrive-files` or `microsoft-search-query`.',
    },
    TENANT_ID_OPTION,
  ],
  example: "ask-marcel-office list-document-comments --drive-id 'b!1234' --item-id '01ABC'",
  responseShape:
    '`{ name, format, count, comments: [{ author, date?, anchor?, text, mentions }] }`. `format` is `docx`, `xlsx` or `pptx`. `anchor` is the commented text (docx), the cell as `Sheet!A1` with non-plain sheet names quoted (xlsx), or `slide N` (pptx); it is absent when the file does not record it. `date` is absent on Excel notes, which keep none. `mentions` lists the names written after an `@` that the file knows (its people and its commenters). A file of any other type returns a 415.',
};

export { execute, meta, schema };
