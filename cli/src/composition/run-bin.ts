import { startErrorLine } from '../presenter/start-error-line.ts';
import type { StartPackageBin } from '../use-cases/start-package-bin.ts';

export type RunBinDeps = {
  readonly start: StartPackageBin;
  readonly writeError: (line: string) => void;
  readonly fail: () => void;
};

// A bin of the facade: starts the bin of its package, and on a failure writes one line and fails.
export const runBin = async (deps: RunBinDeps, packageName: string): Promise<void> => {
  const started = await deps.start(packageName);
  if (started.ok) return;
  deps.writeError(startErrorLine(started.error));
  deps.fail();
};
