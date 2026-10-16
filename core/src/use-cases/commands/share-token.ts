import { bytesToBase64 } from '../../domain/utilities/base64.ts';

/**
 * `buildShareToken(url)` encodes a URL for Graph's `/shares/{token}`
 * resolver per [shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get):
 * `u!` + base64url of the URL's UTF-8 bytes, with no padding. Encoding the
 * UTF-8 bytes (not the raw string) is what lets non-ASCII paths resolve — a
 * OneDrive URL with an accented or CJK filename otherwise mints a wrong token
 * (or throws, above U+00FF). Core (package split): the link resolvers of read
 * and the PDF converters of write both open a shared file this way.
 */
const buildShareToken = (url: string): string => {
  const b64 = bytesToBase64(new TextEncoder().encode(url)).replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '');
  return `u!${b64}`;
};

export { buildShareToken };
