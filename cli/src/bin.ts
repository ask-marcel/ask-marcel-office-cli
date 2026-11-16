import { formatError } from '@ask-marcel/office-core/domain/utilities/format-error.ts';
import { createBunFileSystem } from '@ask-marcel/office-core/infra/filesystem-bun.ts';
import { createNodeFileSystem } from '@ask-marcel/office-core/infra/filesystem-node.ts';
import { runBin } from './composition/run-bin.ts';
import { createModuleLoader } from './infra/module-loader.ts';
import { createPackageLocator } from './infra/package-locator.ts';
import { createStartPackageBin } from './use-cases/start-package-bin.ts';

/*
 * The shared body of the three bins of ask-marcel-office-cli (ADR 0005). Each
 * bin (src/auth.ts, src/read.ts, src/write.ts) names its package and calls
 * this. The bin of the package then runs in this process: it reads
 * process.argv, writes its own output and sets its own exit code. The facade
 * writes only to stderr, because stdout can be the MCP frame channel of that
 * bin.
 */
export const facadeBin = async (packageName: string): Promise<void> => {
  const fs = typeof globalThis.Bun === 'undefined' ? createNodeFileSystem() : createBunFileSystem();
  const start = createStartPackageBin({ locator: createPackageLocator(), fs, loader: createModuleLoader() });
  const fail = (): void => {
    process.exitCode = 1;
  };
  // The one crash catch of this entry (rule 17).
  try {
    await runBin({ start, writeError: (line) => process.stderr.write(line), fail }, packageName);
  } catch (e) {
    process.stderr.write(`[crash] ${formatError(e)}\n`);
    fail();
  }
};
