/*
 * The read package's MCP server: the core gateway (`mcp-shell.ts`) given the
 * read package (`single-package.ts`). It has no sign-in tools: `login` and
 * `status` are tools of the auth package's own server (`auth-server.ts`).
 * main.ts imports this module dynamically, so a plain command never loads the SDK.
 */
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { ReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';
import type { McpShellDeps } from '@ask-marcel/office-core/composition/mcp-shell.ts';
import { buildMcpShell } from '@ask-marcel/office-core/composition/mcp-shell.ts';
import { commandPackage } from './single-package.ts';

type BuildMcpServerDeps = Omit<McpShellDeps<ReadGraph>, 'commandPackage' | 'addTools'>;

const buildMcpServer = (deps: BuildMcpServerDeps): McpServer => buildMcpShell({ ...deps, commandPackage });

export { buildMcpServer };
export type { BuildMcpServerDeps };
