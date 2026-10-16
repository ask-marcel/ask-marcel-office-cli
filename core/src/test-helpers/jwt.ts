// A test token shaped as Entra ID sends one: base64url JSON segments and a
// signature nothing here checks. The AccessToken brands refuse any other
// shape, so padded `btoa` output is not a token.
export const jwtSegment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');

export const unsignedJwt = (claims: Record<string, unknown>): string => `${jwtSegment({ alg: 'RS256' })}.${jwtSegment(claims)}.sig`;
