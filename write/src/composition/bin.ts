/*
 * The composition root of the write bin, still a skeleton of the package
 * split: the package's own name and version, given to core's skeleton CLI
 * until the write commands move in (phase 5).
 */

import type { SkeletonOutput } from '@ask-marcel/office-core/composition/skeleton-cli.ts';
import { runSkeletonCli } from '@ask-marcel/office-core/composition/skeleton-cli.ts';
import pkg from '../../package.json' with { type: 'json' };

export const runBin = (args: ReadonlyArray<string>, output: SkeletonOutput): number => runSkeletonCli(pkg, args, output);
