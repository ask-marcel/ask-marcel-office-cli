import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { buildShareToken } from './sharepoint-link-extractor.ts';

// Encoder for Microsoft Graph's `/shares/{token}` resolver. Takes any
// OneDrive / SharePoint sharing URL and emits the `u!<base64url>` share
// token that Graph's [shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get)
// endpoint accepts. Pure encoding — no HTTP — the actual `/shares/{token}/driveItem`
// fetch is left to a follow-up call (the LLM may want different downstream
// commands depending on the resolved driveItem: `download-onedrive-file-content`,
// `convert-mail-attachment-to-markdown`, etc.).
//
// Accepted host shapes:
//   - `*.sharepoint.com`           — tenant SharePoint share URLs (`:b:/s/site/...`)
//   - `*-my.sharepoint.com`        — personal OneDrive share URLs
//   - `1drv.ms`                    — Microsoft's short-link shortener
//
// Internal helper `buildShareToken` lives in `sharepoint-link-extractor.ts`
// because three other commands (extract-sharepoint-links-in-mail,
// convert-mail-attachment-to-markdown, convert-mail-attachment-to-pdf) also
// use it. Don't inline-duplicate it here.
const schema = z.object({
  url: z.url(),
});

const ACCEPTED_HOST_PATTERNS: ReadonlyArray<RegExp> = [
  /\.sharepoint\.com$/i, // tenant + personal (`*-my.sharepoint.com` ends with `.sharepoint.com`)
  /^1drv\.ms$/i,
];

type Resolved = {
  readonly shareToken: string;
  readonly graphPath: string;
  readonly originalUrl: string;
};

const isAcceptedHost = (hostname: string): boolean => ACCEPTED_HOST_PATTERNS.some((re) => re.test(hostname));

const parse = (raw: string): Resolved | null => {
  // No try/catch — Zod's `.url()` refinement on the input schema already
  // validated the URL format before this parser runs. Also satisfies the
  // atelier rule restricting try/catch to `src/infra/**`.
  const url = new URL(raw);
  if (!isAcceptedHost(url.hostname)) return null;
  const shareToken = buildShareToken(raw);
  return {
    shareToken,
    graphPath: `/shares/${shareToken}/driveItem`,
    originalUrl: raw,
  };
};

const execute: Command['execute'] = async (_graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const resolved = parse(parsed.data.url);
  if (resolved === null) {
    return err({
      type: 'validation_error',
      message:
        '--url: not a recognised OneDrive / SharePoint sharing URL. Accepted hosts: `*.sharepoint.com` (tenant + personal OneDrive — `*-my.sharepoint.com` is a subdomain that matches), `1drv.ms` (Microsoft short link).',
    });
  }
  return ok(resolved);
};

const meta: CommandMeta = {
  summary:
    "Encode a OneDrive / SharePoint sharing URL into the Graph `/shares/{token}` share token (`u!<base64url>` per [shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get)). Pure transformation — no Graph call. Pipe the returned `graphPath` (`/shares/{token}/driveItem`) into a sibling lookup (`get-drive-item`, `download-onedrive-file-content`, `convert-mail-attachment-to-pdf`, etc.) once the file has been resolved to a `driveItem`. Accepts any `*.sharepoint.com` URL (tenant + `*-my.sharepoint.com` personal OneDrive) and Microsoft's short-link host `1drv.ms`.",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '{url}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/shares-get',
  options: [
    {
      name: 'url',
      key: 'url',
      required: true,
      description:
        'A OneDrive / SharePoint sharing URL — the address from the "Copy link" / "Share" action in the OneDrive or SharePoint UI. Examples: `https://contoso.sharepoint.com/:b:/s/sitename/EaB1cD...`, `https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/file.pdf`, `https://1drv.ms/b/s!AbCdEfGh...`. The CLI does not follow the redirect on `1drv.ms` links — the short URL itself is encoded as the share token (Graph resolves it on the server side).',
    },
  ],
  example: "ask-marcel resolve-drive-share-link --url 'https://contoso.sharepoint.com/:b:/s/team/EaB1cD2eF...?e=abc'",
  responseShape:
    '`{ shareToken: string, graphPath: string, originalUrl: string }`. `shareToken` is the `u!<base64url>` form. `graphPath` is the ready-to-use `/shares/{token}/driveItem` URL — pass it to `ask-marcel next-page --url <link>` for a one-shot driveItem fetch, or feed the `shareToken` into any future `/shares/{token}/...` endpoint. `originalUrl` is echoed back for round-trip confirmation.',
};

export { execute, meta, schema };
