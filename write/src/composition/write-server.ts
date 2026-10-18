/*
 * The MCP server of the write package: the core gateway (`mcp-shell.ts`) given
 * the write package (`write-package.ts`), on the write graph alone. It has
 * three tools: list-commands, get-command-docs and run-write-command. The bin
 * imports this module dynamically, so a plain command never loads the SDK.
 */
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { McpShellDeps } from '@ask-marcel/office-core/composition/mcp-shell.ts';
import { buildMcpShell } from '@ask-marcel/office-core/composition/mcp-shell.ts';
import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import { commandPackage } from './write-package.ts';

type BuildWriteMcpServerDeps = Omit<McpShellDeps<WriteGraph>, 'commandPackage' | 'addTools'>;

const buildWriteMcpServer = (deps: BuildWriteMcpServerDeps): McpServer => buildMcpShell({ ...deps, commandPackage });

export { buildWriteMcpServer };
export type { BuildWriteMcpServerDeps };
