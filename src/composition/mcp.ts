/*
 * The single package's MCP server, the module `main.ts` (dynamically) and the
 * real-registry suites import. The gateway is in `mcp-shell.ts`, which
 * package-split step 13 turns into the core gateway.
 */
export { buildMcpServer } from './mcp-shell.ts';
export type { BuildMcpServerDeps } from './mcp-shell.ts';
