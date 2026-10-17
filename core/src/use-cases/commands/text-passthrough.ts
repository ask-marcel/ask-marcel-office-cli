/**
 * Microsoft Graph's `?format=pdf` / `?format=html` conversion accepts a
 * fixed set of source extensions. For anything outside that whitelist
 * Graph (or its downstream Office Online sandbox) returns a 4xx, which
 * is a cryptic failure mode for an LLM agent.
 *
 * The conversion commands short-circuit on the extensions below and
 * return the raw bytes instead of attempting (and wasting) a Graph
 * round-trip:
 *
 *   - PLAIN_TEXT_EXTENSIONS: extensions that are themselves plain text
 *     (txt/md/json/yaml/etc.) — turning these through `?format=pdf` or
 *     `?format=html` is meaningless; just return the bytes.
 *
 *   - `pdf`: the `*-as-pdf` commands ask Graph to convert TO pdf, but
 *     `pdf` is NOT in Graph's `format=pdf` source-input list (verified
 *     2026-05 against svc.ms; CDN responds 406 InputFormatNotSupported).
 *     The user wanted a PDF and the source IS a PDF, so return its
 *     bytes directly.
 */

const PLAIN_TEXT_EXTENSIONS: ReadonlySet<string> = new Set(['txt', 'md', 'markdown', 'html', 'htm', 'json', 'xml', 'log', 'yaml', 'yml', 'css', 'js', 'ts', 'sh', 'sql']);

const extensionOf = (name: string): string => {
  const dot = name.lastIndexOf('.');
  if (dot === -1) return '';
  return name.slice(dot + 1).toLowerCase();
};

const isPlainTextFilename = (name: string): boolean => PLAIN_TEXT_EXTENSIONS.has(extensionOf(name));

const isPdfSource = (name: string): boolean => extensionOf(name) === 'pdf';

/**
 * Content-sniff for text — decode bytes STRICTLY as UTF-8. Returns the decoded
 * string when the bytes are valid UTF-8, `undefined` when they are not (i.e.
 * binary). The conversion / content commands use this instead of trusting a
 * file extension: it detects any text file regardless of name AND can never
 * mangle binary into `�` (the non-fatal decoder's silent corruption — the bug
 * a binary file named `.txt` triggered). Non-UTF-8-encoded text (UTF-16, etc.)
 * is reported as binary, which is faithful (base64) rather than lossy.
 *
 * `{ fatal: true }` makes `decode` throw on the first invalid sequence; the
 * try/catch is a pure-domain fallback around a native synchronous thrower.
 */
const decodeUtf8Text = (bytes: Uint8Array): string | undefined => {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return undefined;
  }
};

export { decodeUtf8Text, extensionOf, isPdfSource, isPlainTextFilename };
