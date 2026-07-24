/*
 * Flags a delta endpoint used to advertise and no longer does.
 *
 * Delta endpoints are not ordinary collections: Microsoft honours a different
 * OData subset on each one, and the ones they do not honour are not rejected,
 * they are silently dropped. Advertising such a flag hands the caller a lever
 * that visibly does nothing, and every one of these was advertised here until
 * 2026-07-23, so callers have had time to learn them.
 *
 * Zod cannot carry this refusal: `z.object()` strips unknown keys instead of
 * failing, and `.strict()` is unavailable on any command with a flag alias,
 * because the alias normalizer leaves both the alias key and the canonical key
 * in the params bag.
 *
 * Each entry says what the live probe showed, so the message explains rather
 * than just refuses.
 */

type WithdrawnFlag = { readonly key: string; readonly reason: string };

const withdrawnFlagRefusal = (commandName: string, withdrawn: ReadonlyArray<WithdrawnFlag>, params: Record<string, string>): string | undefined => {
  const hit = withdrawn.find((flag) => Object.hasOwn(params, flag.key));
  if (hit === undefined) return undefined;
  return `--${hit.key} is not supported on ${commandName}. ${hit.reason}`;
};

export { withdrawnFlagRefusal };
export type { WithdrawnFlag };
