export const decodeJwtPayload = (token: string): Record<string, unknown> => {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return {};
    const b64 = parts[1].replaceAll('-', '+').replaceAll('_', '/');
    const padded = b64.padEnd(b64.length + ((4 - (b64.length % 4)) % 4), '=');
    return JSON.parse(Buffer.from(padded, 'base64').toString('utf-8')) as Record<string, unknown>;
  } catch {
    return {};
  }
};

export const isTokenFresh = (token: string, bufferSeconds = 300): boolean => {
  const claims = decodeJwtPayload(token);
  const exp = claims.exp as number | undefined;
  if (typeof exp !== 'number') return false;
  return Date.now() / 1000 < exp - bufferSeconds;
};

const GRAPH_AUDIENCE_UUID = '00000003-0000-0000-c000-000000000000';
const GRAPH_AUDIENCE_URL = 'graph.microsoft.com';

export const isGraphToken = (token: string): boolean => {
  const claims = decodeJwtPayload(token);
  const aud = claims.aud;
  if (typeof aud === 'string') {
    return aud === GRAPH_AUDIENCE_UUID || aud.includes(GRAPH_AUDIENCE_URL);
  }
  if (Array.isArray(aud)) {
    return aud.some((a) => typeof a === 'string' && (a === GRAPH_AUDIENCE_UUID || a.includes(GRAPH_AUDIENCE_URL)));
  }
  return false;
};
