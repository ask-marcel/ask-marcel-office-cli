/*
 * The single package's MCP server: the core gateway (`mcp-shell.ts`) given the
 * single package (`single-package.ts`) and the auth tools (`auth-mcp.ts`). Its
 * deps are the ones `main.ts` and the real-registry suites have always passed.
 * main.ts imports this module dynamically, so a plain command never loads the SDK.
 */
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { AuthToolsDeps } from './auth-mcp.ts';
import { addAuthTools } from './auth-mcp.ts';
import type { McpShellDeps } from './mcp-shell.ts';
import { buildMcpShell } from './mcp-shell.ts';
import { commandPackage } from './single-package.ts';

type BuildMcpServerDeps = Omit<McpShellDeps, 'commandPackage' | 'addTools'> & AuthToolsDeps;

const buildMcpServer = (deps: BuildMcpServerDeps): McpServer => buildMcpShell({ ...deps, commandPackage, addTools: (server) => addAuthTools(server, deps) });

export { buildMcpServer };
export type { BuildMcpServerDeps };
