/*
 * Format an unknown thrown value into a human-readable string.
 *
 * Replaces `String(err)` in catch blocks. `String(obj)` returns
 * "[object Object]" for plain-object throws and loses the message
 * entirely (SonarJS S6551).
 *
 * Use this in EVERY `catch (error)` block in src/infra/** and in any
 * pure-domain native-API fallback. Safe on any input.
 *
 * See skills/atelier/references/workflow.md (SonarJS table, S6551).
 */

export const formatError = (err: unknown): string => {
  if (err instanceof Error) return err.message;
  if (typeof err === 'string') return err;
  // Numbers only. String(NaN) is 'NaN' and String(Infinity) is 'Infinity', which the
  // JSON.stringify fallback would flatten to 'null', so the number branch earns its
  // keep. Booleans are deliberately NOT special-cased: JSON.stringify renders them
  // identically to String, making a boolean branch dead code (and an unkillable
  // equivalent mutant that would cap this file below the 90% mutation gate).
  if (typeof err === 'number') return String(err);
  try {
    // JSON.stringify returns undefined, not a string, for undefined, a function and a
    // symbol (a `throw undefined` or `Promise.reject()` logged "error: undefined" by luck
    // until 2026-09-27); String covers those three.
    return JSON.stringify(err) ?? String(err);
  } catch {
    return '[unstringifiable error]';
  }
};
