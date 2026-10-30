/**
 * Past this many added plus removed lines (a changed line counts once each way)
 * a line diff is about as long as both texts, so the diff stops early and the
 * diff commands say so instead (src/infra/line-diff-adapter.ts).
 */
export const MAX_EDITS = 1_000;
