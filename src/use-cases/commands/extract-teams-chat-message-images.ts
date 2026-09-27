import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({ chatId: z.string().min(1), messageId: z.string().min(1) });

// A pasted image is an <img> of item type AMSImage whose src points at Teams'
// media service; emoji and stickers carry other item types and stay put.
const AMS_IMAGE = /<img\b[^>]*\bitemtype="http:\/\/schema\.skype\.com\/AMSImage"[^>]*>/gi;
const SRC = /\bsrc="([^"]+)"/i;

const pastedImageUrls = (content: string): ReadonlyArray<string> => [
  ...new Set(
    [...content.matchAll(AMS_IMAGE)].flatMap(([tag]) => {
      const src = SRC.exec(tag)?.[1];
      return src === undefined ? [] : [src.replaceAll('&amp;', '&')];
    })
  ),
];

const EXTENSIONS: Readonly<Record<string, string>> = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp' };

type Blob = { readonly contentType: string; readonly size: number; readonly base64: string };

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { chatId, messageId } = parsed.data;
  const message = await graph.teamsChatIc3(`/v1/users/ME/conversations/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`);
  if (!message.ok) return message;
  const content = (message.value as { readonly content?: unknown }).content;
  const urls = typeof content === 'string' ? pastedImageUrls(content) : [];
  if (urls.length === 0) return ok({ count: 0, media: [], note: 'This message holds no pasted image.' });
  const blobs: Blob[] = [];
  for (const image of await Promise.all(urls.map((url) => graph.teamsChatMedia(url)))) {
    if (!image.ok) return image;
    blobs.push(image.value as Blob);
  }
  const media = blobs.map((blob, index) => {
    const extension = EXTENSIONS[blob.contentType.split(';')[0]] ?? 'bin';
    return { path: `image-${index + 1}.${extension}`, contentType: blob.contentType, sizeBytes: blob.size, base64: blob.base64 };
  });
  return ok({ count: media.length, media });
};

const meta: CommandMeta = {
  summary:
    "Fetch the images pasted into one Teams chat message (screenshots, charts), which `list-teams-chat-messages` shows only as an empty `<img>`: the message is read through the IC3 substrate and each pasted image downloaded from Teams' media service with the same token, emoji and stickers left out. Answers them as a media list, ready to save as files for a vision model. Take the ids from `list-teams-chat-messages` or `list-teams-chat-history`. **Best-effort, may break on Microsoft client updates** — the substrate and the media service are not in the public Microsoft Graph API.",
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: "/v1/users/ME/conversations/{chat-id}/messages/{message-id} (IC3), then each pasted image from Teams' media service",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chatmessage-list-hostedcontents',
  options: [
    {
      name: 'chat-id',
      key: 'chatId',
      required: true,
      description: 'Teams chat ID (`19:...@thread.v2` or `19:..._...@unq.gbl.spaces`), from `list-teams-chats-with-messages` or `find-chats-with-user`.',
    },
    { name: 'message-id', key: 'messageId', required: true, description: 'The message `id`, from `list-teams-chat-messages` or `list-teams-chat-history`.' },
  ],
  example: "ask-marcel-office extract-teams-chat-message-images --chat-id '19:abc@thread.v2' --message-id '1727000000000'",
  responseShape:
    '`{ count, media: [{ path, contentType, sizeBytes, base64 }] }` — `path` is `image-<n>.<ext>`, in the order the images appear. Pair with the global `--output-dir <dir>` to write each image there (the response then carries `savedTo` in place of `base64`). `count: 0` with a `note` means the message holds no pasted image.',
  producesMedia: true,
  needsSubstrateToken: 'ic3',
  stability: 'experimental',
};

export { execute, meta, schema };
