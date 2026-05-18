import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { officeToMarkdown } from './office-to-markdown.ts';
import { normalizeVersionId } from './version-id.ts';

const schema = z.object({
  driveId: z.string().min(1),
  itemId: z.string().min(1),
  versionId: z.string().min(1),
});

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { driveId, itemId } = parsed.data;
  const versionId = normalizeVersionId(parsed.data.versionId);

  const meta = await graph.get(`/drives/${driveId}/items/${itemId}`);
  if (!meta.ok) return meta;
  const name = (meta.value as { name?: string }).name ?? '';

  // Historical-version stream content requires an "elevated" Graph
  // token (M365ChatClient appid). The Teams web client token returns
  // 403 logicalPermissionAccessDenied on the same endpoint.
  return officeToMarkdown(graph, `/drives/${driveId}/items/${itemId}/versions/${versionId}/content`, name, { elevated: true });
};

const meta: CommandMeta = {
  summary:
    'Download a *historical version* of a OneDrive / SharePoint file converted to markdown. Same local conversion pipeline as `download-drive-item-as-markdown`: docx via mammoth, xlsx via sheetjs (markdown tables per sheet), csv as a markdown table, plus plain-text passthrough. Uses an elevated Graph token (captured at login from m365.cloud.microsoft / M365ChatClient) for the bytes-fetch, since the Teams web client token cannot fetch historical-version stream content (returns 403 logicalPermissionAccessDenied). For pptx use `download-drive-item-version-as-pdf`. Loop/Fluid/Whiteboard use Graph `?format=html` (the four inputs Microsoft documents).',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/drives/{drive-id}/items/{item-id}/versions/{version-id}/content?format=html',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitemversion-get-content',
  options: [
    {
      name: 'drive-id',
      key: 'driveId',
      required: true,
      description:
        'Microsoft Graph drive ID. Use `ask-marcel list-drives` for the personal OneDrive, ' +
        'or `ask-marcel list-sharepoint-site-drives --site-id <id>` for a SharePoint document library.',
    },
    { name: 'item-id', key: 'itemId', required: true, description: 'driveItem ID of the file. Returned by `list-folder-files` or `search-onedrive-files`.' },
    {
      name: 'version-id',
      key: 'versionId',
      required: true,
      description:
        'driveItemVersion ID. Returned by `ask-marcel list-drive-item-versions`. ' +
        'Pick a non-current version (the first entry is the live file and Graph rejects this endpoint for it).',
    },
  ],
  example: "ask-marcel download-drive-item-version-as-markdown --drive-id 'b!1234' --item-id '01ABC' --version-id '4.0'",
  responseShape: '`{ contentType: "text/markdown", size: <chars>, text: "..." }` for the converted case; raw-bytes envelope for plain-text source extensions.',
  needsElevatedToken: true,
  producesBytes: true,
};

export { execute, meta, schema };
