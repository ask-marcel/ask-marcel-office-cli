/*
 * The local file reader of the write bin: the file system of this runtime. A
 * write command reads through it a local file that a param names (`--ids-file`
 * today, and later `--file` and `--body-file`), with the size limit of that
 * command (core/src/use-cases/commands/read-local-file.ts). The CLI and the MCP
 * server of the write bin both get it.
 */
import { createBunFileSystem } from '@ask-marcel/office-core/infra/filesystem-bun.ts';
import { createNodeFileSystem } from '@ask-marcel/office-core/infra/filesystem-node.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';

export const buildWriteFiles = (): LocalFileReader => (typeof globalThis.Bun === 'undefined' ? createNodeFileSystem() : createBunFileSystem());
