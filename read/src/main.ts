/*
 * The bin of @ask-marcel/office-read (dist/cli.js). It only wires the
 * process to the composition root (composition/bin.ts).
 */

import { runBin } from './composition/bin.ts';

process.exitCode = runBin(process.argv.slice(2), {
  out: (text) => process.stdout.write(text),
  err: (text) => process.stderr.write(text),
});
