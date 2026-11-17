import type { GraphError } from './graph-error.ts';
import type { Result } from './result.ts';
import { err } from './result.ts';

/*
 * The no-send check of the write graph (product rule: no command sends mail).
 * A POST to one of these Graph actions sends mail at once, and a JSON batch
 * can carry such a POST in its body. The actions that make a draft
 * (`createReply`, `createReplyAll`, `createForward`) are not on the list.
 */

// The last path segment of each sending action, in lower case.
const SENDING_ACTIONS: ReadonlySet<string> = new Set(['send', 'sendmail', 'reply', 'replyall', 'forward', '$batch']);

// A segment with a percent-encoded letter (`%73end`) is compared decoded; a
// malformed escape cannot name an action, so it is compared as it is.
const decoded = (segment: string): string => {
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
};

// The action a path ends in: no query, no trailing slash, no `microsoft.graph.` namespace, lower case.
const actionOf = (path: string): string => {
  const segments = (path.split(/[?#]/, 1)[0] ?? '').split('/').filter((segment) => segment !== '');
  return decoded(segments.at(-1) ?? '')
    .trim()
    .toLowerCase()
    .replace(/^microsoft\.graph\./, '');
};

/** Whether a change to `path` would send mail. */
const sendsMail = (path: string): boolean => SENDING_ACTIONS.has(actionOf(path));

const sendRefused = (method: string, path: string): Result<never, GraphError> =>
  err({
    type: 'validation_error',
    code: 'send_refused',
    message: `${method} ${path} was not sent: it would send mail, and this CLI never sends mail. A draft stays in Drafts until you send it yourself from Outlook.`,
  });

/** Runs `change` only when `path` sends no mail; refuses it otherwise, before any request. */
const unlessSendsMail = async (method: string, path: string, change: () => Promise<Result<unknown, GraphError>>): Promise<Result<unknown, GraphError>> =>
  sendsMail(path) ? sendRefused(method, path) : change();

export { sendsMail, unlessSendsMail };
