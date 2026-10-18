/*
 * The MCP server of the auth package: its own server, not the gateway of the
 * Graph commands (`mcp-shell.ts`), with the `login` and `status` tools of
 * `auth-mcp.ts`. It takes no graph and no file system. No tool offers `token`,
 * and no tool result holds a token (LESSONS 2026-10-15).
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { AuthToolsDeps } from './auth-mcp.ts';
import { addAuthTools } from './auth-mcp.ts';
import { authPackage } from './auth-package.ts';

type BuildAuthMcpServerDeps = Omit<AuthToolsDeps, 'names'> & { readonly version: string };

const buildAuthMcpServer = (deps: BuildAuthMcpServerDeps): McpServer => {
  const server = new McpServer({ name: authPackage.binName, version: deps.version });
  addAuthTools(server, { ...deps, names: authPackage });
  return server;
};

export { buildAuthMcpServer };
export type { BuildAuthMcpServerDeps };
