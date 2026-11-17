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

// The path as the request sends it. The URL parser resolves dot segments
// (`x/..`, `%2e%2e`) and drops the query, as fetch does with the Graph root in
// front, so a path cannot reach a send through `..`. The root makes any path parse.
const sentPath = (path: string): string => new URL(`https://graph.invalid/v1.0${path}`).pathname;

// Each percent-encoded byte decoded on its own (`%73end` is `send`): a sending
// action is ASCII, and a malformed escape stays as it is.
const decoded = (segment: string): string => segment.replaceAll(/%([\da-f]{2})/gi, (_escape, hex: string) => String.fromCodePoint(Number.parseInt(hex, 16)));

// The action a path ends in: no trailing slash, no `microsoft.graph.` namespace, lower case.
const actionOf = (path: string): string => {
  const segments = sentPath(path)
    .split('/')
    .filter((segment) => segment !== '');
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
const unlessSendsMail = async <T>(method: string, path: string, change: () => Promise<Result<T, GraphError>>): Promise<Result<T, GraphError>> =>
  sendsMail(path) ? sendRefused(method, path) : change();

export { sendsMail, unlessSendsMail };
