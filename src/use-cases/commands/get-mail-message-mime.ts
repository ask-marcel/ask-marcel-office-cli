import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { inlineBinary } from './fetch-raw-bytes.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({ messageId: z.string().min(1) });

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  return inlineBinary(graph, `/me/messages/${parsed.data.messageId}/$value`);
};

const meta: CommandMeta = {
  summary:
    'Return the raw RFC 5322 MIME source of a single Outlook message — full headers, every attachment encoded inline. Useful for archiving, full-fidelity forensic inspection, or feeding into a tool that reads MIME directly. For human-readable content prefer `get-mail-message` or `convert-mail-to-markdown`.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}/$value',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      description: 'Outlook message ID. Returned by `list-mail-messages` or `search-mail-messages`.',
    },
  ],
  example: "ask-marcel get-mail-message-mime --message-id 'AAMkAD...'",
  responseShape:
    '`{ contentType: "message/rfc822" or similar, size, base64 }` — the raw MIME envelope, inlined. Pair with the global `--output-path <path>` flag to land the .eml on disk and replace `base64` with `savedTo` for messages with large attachments.',
};

export { execute, meta, schema };
