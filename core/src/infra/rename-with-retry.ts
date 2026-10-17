/*
 * A rename that waits out a short-lived refusal. On Windows, renaming over a
 * file another process holds open (a concurrent reader of the token cache, an
 * antivirus scan) fails with EPERM, EBUSY or EACCES for a moment; elsewhere the
 * rename is atomic and these codes mean a real permission problem, which the
 * retries only delay. Any other error is thrown at once.
 */

const RETRYABLE_CODES: ReadonlySet<string> = new Set(['EPERM', 'EBUSY', 'EACCES']);

export const RENAME_RETRY_DELAYS_MS: ReadonlyArray<number> = [50, 100, 200, 400];

const isRetryable = (e: unknown): boolean => e instanceof Error && RETRYABLE_CODES.has((e as { code?: string }).code ?? '');

const wait = async (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

export const renameWithRetry = async (
  rename: (from: string, to: string) => Promise<void>,
  from: string,
  to: string,
  delaysMs: ReadonlyArray<number> = RENAME_RETRY_DELAYS_MS
): Promise<void> => {
  for (const delay of delaysMs) {
    try {
      return await rename(from, to);
    } catch (e) {
      if (!isRetryable(e)) throw e;
      await wait(delay);
    }
  }
  return rename(from, to);
};
