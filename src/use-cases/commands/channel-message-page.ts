import { err } from '../../domain/result.ts';
import type { Command, CommandOptionMeta } from './command-types.ts';

/**
 * Graph caps a page of channel messages (and of a message's replies) at 50:
 * `$top=51` is a 400 naming the limit (probed live 2026-09-09). The generic
 * `--top` flag promises up to 1,000, so these commands carry their own flag
 * wording and refuse a larger page before the round trip.
 */
const CHANNEL_MESSAGES_TOP_CAP = 50;

const CHANNEL_MESSAGES_TOP_OPTION: CommandOptionMeta = {
  name: 'top',
  key: 'top',
  required: false,
  description:
    'OData $top: page size, 1 to 50 (Graph caps a page of channel messages at 50 and answers 400 above it; the CLI refuses larger values up front). Without it Graph picks its own page size, around 20. Older messages continue through the `next:` footer with `next-page`.',
};

/** Wraps an execute so a `--top` above the cap is refused before Graph is called. */
const withChannelMessagesTopCap =
  (inner: Command['execute']): Command['execute'] =>
  async (graph, params) => {
    if (Number(params['top']) > CHANNEL_MESSAGES_TOP_CAP) {
      return err({
        type: 'validation_error',
        message: `--top must be at most ${CHANNEL_MESSAGES_TOP_CAP}: Graph caps a page of channel messages at ${CHANNEL_MESSAGES_TOP_CAP} and answers 400 above it. Continue through the \`next:\` footer with \`next-page\`.`,
      });
    }
    return inner(graph, params);
  };

export { CHANNEL_MESSAGES_TOP_CAP, CHANNEL_MESSAGES_TOP_OPTION, withChannelMessagesTopCap };
