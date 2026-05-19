import updateNotifier from 'update-notifier';
import pkg from '../package.json' with { type: 'json' };
import { buildDeps } from './composition/build-deps.ts';
import { buildCli } from './composition/cli.ts';
import { formatError } from './domain/utilities/format-error.ts';

const ONE_WEEK_MS = 1000 * 60 * 60 * 24 * 7;

const main = async (): Promise<void> => {
  updateNotifier({ pkg: { name: pkg.name, version: pkg.version }, updateCheckInterval: ONE_WEEK_MS }).notify({ defer: false });
  const deps = buildDeps();
  const cli = buildCli({
    auth: deps.auth,
    graph: deps.graph,
    logger: deps.logger,
    processRunner: deps.processRunner,
    fs: deps.fs,
    version: pkg.version,
    onCommandError: () => {
      process.exitCode = 1;
    },
  });
  await cli.parseAsync();
};

const isCommanderError = (e: unknown): boolean =>
  e !== null && typeof e === 'object' && 'code' in e && typeof e.code === 'string' && (e as { code: string }).code.startsWith('commander.');

try {
  await main();
} catch (e) {
  // Commander parser errors (unknown option, missing required, etc.) are
  // already rendered as a JSON envelope on stdout by exitOverride in
  // buildCli. Don't double-print them to stderr — that would re-introduce
  // the two-channel error contract the audit flagged. Truly uncaught errors
  // (assertion failures, OOM, etc.) still hit the [crash] line.
  if (isCommanderError(e)) {
    process.exit(1);
  }
  process.stderr.write(`[crash] ${formatError(e)}\n`);
  process.exit(1);
}
