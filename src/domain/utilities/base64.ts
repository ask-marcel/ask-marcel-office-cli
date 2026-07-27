/**
 * Encode raw bytes to standard base64.
 *
 * Built byte-by-byte into a Latin-1 binary string so every value stays in
 * `btoa`'s 0x00-0xFF input range — `btoa` throws on any code unit above 0xFF,
 * and the char-at-a-time loop avoids the call-stack blow-up of
 * `String.fromCharCode(...bytes)` on large inputs.
 *
 * Callers that need base64url apply the `+`/`/`/`=` transform themselves.
 */
export const bytesToBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};
