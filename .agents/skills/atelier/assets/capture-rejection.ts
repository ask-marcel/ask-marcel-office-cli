/*
 * Capture a promise rejection and return the caught Error.
 *
 * Replaces `await expect(p).rejects.toThrow(...)`, which trips SonarJS S4123
 * ("unexpected await of a non-Promise value") because the matcher chain is
 * not a real Thenable. This helper reads more clearly anyway:
 *
 *   const err = await captureRejection(doSomethingThatThrows());
 *   expect(err.message).toBe('expected message');
 *
 * Throws:
 *   - if the promise resolved (so the test fails loudly instead of silently passing)
 *   - if the rejection value is not an Error (in atelier codebases, all
 *     rejections must be Errors — enforced by @typescript-eslint/prefer-promise-reject-errors)
 *
 * See skills/atelier/references/workflow.md (SonarJS table, S4123) and
 * skills/atelier/references/result-type.md (Testing Result-returning code).
 */

const formatNonError = (value: unknown): string => {
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  try {
    return JSON.stringify(value);
  } catch {
    return '[unstringifiable value]';
  }
};

export const captureRejection = async (promise: Promise<unknown>): Promise<Error> => {
  try {
    await promise;
  } catch (e) {
    if (e instanceof Error) return e;
    throw new Error(`captureRejection: rejected with non-Error value: ${formatNonError(e)}`);
  }
  throw new Error('captureRejection: expected promise to reject, but it resolved');
};
