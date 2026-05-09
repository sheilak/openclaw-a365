import { AsyncLocalStorage } from "node:async_hooks";
import { Type, type TSchema } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import type { TurnContext, Authorization } from "@microsoft/agents-hosting";
import {
  McpToolServerConfigurationService,
  defaultToolingConfigurationProvider,
  resolveTokenScopeForServer,
  type MCPServerConfig,
  type McpClientTool,
} from "@microsoft/agents-a365-tooling";
import { AgenticAuthenticationService } from "@microsoft/agents-a365-runtime";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { ExecuteToolScope } from "@microsoft/opentelemetry";
import type { A365Config } from "./types.js";
import { getA365Runtime } from "./runtime.js";

type ToolResult = AgentToolResult<unknown>;

/**
 * Per-turn context required to invoke MCP tools. Stored in AsyncLocalStorage
 * so the synchronous AgentTool.execute closure can read it during dispatch
 * and mint a fresh agentic-user token via OBO for each call.
 */
export type MCPToolContext = {
  turnContext: TurnContext;
  authorization: Authorization;
  authHandlerName: string;
};

const mcpContextStorage = new AsyncLocalStorage<MCPToolContext>();

export function runWithMcpToolContext<T>(ctx: MCPToolContext, fn: () => T): T {
  return mcpContextStorage.run(ctx, fn);
}

export function getMcpToolContext(): MCPToolContext | undefined {
  return mcpContextStorage.getStore();
}

function getLogger() {
  try {
    return getA365Runtime().logging.getChildLogger({ name: "a365-mcp" });
  } catch {
    return { debug: () => {}, info: () => {}, warn: () => {}, error: () => {} };
  }
}

/**
 * Wrap an AgentTool so every call is enclosed in an A365 ExecuteToolScope.
 * Centralises observability so tool implementations stay focused on the call.
 */
export function wrapToolWithA365Instrumentation(
  tool: AgentTool<TSchema, unknown>,
  cfg?: A365Config,
): AgentTool<TSchema, unknown> {
  const inner = tool.execute;
  return {
    ...tool,
    execute: async (toolCallId: string, params: unknown, signal, onUpdate) => {
      const log = getLogger();
      log.info(`MCP tool call: ${tool.name}`, {
        toolCallId,
        args: params,
      });
      const scope = ExecuteToolScope.start(
        {},
        {
          toolName: tool.name,
          toolCallId,
          arguments: (params ?? undefined) as Record<string, unknown> | undefined,
        },
        {
          agentId: cfg?.graph?.aaInstanceId || cfg?.agentIdentity || "",
          agentEmail: cfg?.agentIdentity,
          tenantId: cfg?.tenantId,
        },
      );
      try {
        return await scope.withActiveSpanAsync(async () => {
          const result = await inner(toolCallId, params as never, signal, onUpdate);
          try {
            scope.recordResponse(result as Record<string, unknown>);
          } catch {
            // best-effort
          }
          const isErr = (result as { isError?: boolean })?.isError === true;
          if (isErr) {
            log.warn(`MCP tool returned isError: ${tool.name}`, { result });
          } else {
            log.info(`MCP tool ok: ${tool.name}`);
          }
          return result;
        });
      } catch (err) {
        log.error(
          `MCP tool threw: ${tool.name}: ${err instanceof Error ? err.message : String(err)}`,
        );
        scope.recordError(err as Error);
        throw err;
      } finally {
        scope.dispose();
      }
    },
  };
}

/**
 * Module-level cache populated by `primeMcpToolsForTurn` on the first turn.
 * Returned verbatim by `getCachedMcpTools` to satisfy openclaw's synchronous
 * `agentTools` factory. Empty until the first message handler runs discovery.
 */
let cachedTools: AgentTool<TSchema, unknown>[] = [];
let cachedScopesByServer: Map<string, string> = new Map();
let primingPromise: Promise<void> | undefined;

export function getCachedMcpTools(): AgentTool<TSchema, unknown>[] {
  return cachedTools;
}

/**
 * Discover MCP servers + tools once. Subsequent calls are no-ops while the
 * cache is populated. Safe to call on every turn.
 */
export async function primeMcpToolsForTurn(
  cfg: A365Config | undefined,
  turnContext: TurnContext,
  authorization: Authorization,
  authHandlerName: string,
): Promise<void> {
  if (cachedTools.length > 0) return;
  if (primingPromise) return primingPromise;

  primingPromise = (async () => {
    const log = getLogger();
    try {
      const { tools, scopesByServer } = await discoverMcpTools(
        cfg,
        turnContext,
        authorization,
        authHandlerName,
      );
      cachedTools = tools;
      cachedScopesByServer = scopesByServer;
      const toolNames = tools.map((t) => t.name);
      log.info(
        `MCP tools discovered: count=${tools.length} servers=${scopesByServer.size} names=${toolNames.join(",") || "(none)"}`,
      );
      // If discovery returned no tools (e.g. consent not yet granted), allow a future
      // turn to retry — otherwise the resolved primingPromise would short-circuit forever.
      if (tools.length === 0) {
        primingPromise = undefined;
      }
    } catch (err) {
      log.error(`MCP discovery failed: ${err instanceof Error ? err.message : String(err)}`);
      // Leave cache empty so a later turn retries.
      primingPromise = undefined;
      throw err;
    }
  })();

  return primingPromise;
}

async function discoverMcpTools(
  cfg: A365Config | undefined,
  turnContext: TurnContext,
  authorization: Authorization,
  authHandlerName: string,
): Promise<{
  tools: AgentTool<TSchema, unknown>[];
  scopesByServer: Map<string, string>;
}> {
  const toolingCfg = defaultToolingConfigurationProvider.getConfiguration();
  const sharedScope = toolingCfg.mcpPlatformAuthenticationScope;
  const configService = new McpToolServerConfigurationService();

  const servers = await configService.listToolServers(
    turnContext,
    authorization,
    authHandlerName,
  );

  const tools: AgentTool<TSchema, unknown>[] = [];
  const scopesByServer = new Map<string, string>();

  for (const server of servers) {
    const scope = resolveTokenScopeForServer(server, sharedScope);
    scopesByServer.set(server.mcpServerName, scope);

    // Dev-mode listToolServers does NOT attach Authorization (it only reads BEARER_TOKEN_*
    // env vars). Mint an OBO token here so the SDK's listTools call is authenticated.
    if (!server.headers || !("Authorization" in server.headers)) {
      try {
        const token = await AgenticAuthenticationService.GetAgenticUserToken(
          authorization,
          authHandlerName,
          turnContext,
          [scope],
        );
        server.headers = { ...(server.headers ?? {}), Authorization: `Bearer ${token}` };
      } catch (err) {
        getLogger().warn(
          `OBO failed for discovery of '${server.mcpServerName}' (scope ${scope}): ${err instanceof Error ? err.message : String(err)}`,
        );
        continue;
      }
    }

    let mcpTools: McpClientTool[];
    try {
      mcpTools = await configService.getMcpClientTools(server.mcpServerName, server);
    } catch (err) {
      getLogger().warn(
        `getMcpClientTools failed for '${server.mcpServerName}': ${err instanceof Error ? err.message : String(err)}`,
      );
      continue;
    }

    for (const t of mcpTools) {
      tools.push(
        wrapToolWithA365Instrumentation(buildAgentToolFromMcp(server, t), cfg),
      );
    }
  }

  return { tools, scopesByServer };
}

/**
 * Build an AgentTool that, on execute, opens an MCP client to the discovered
 * server URL with a freshly-minted per-audience agentic-user token.
 */
function buildAgentToolFromMcp(
  server: MCPServerConfig,
  mcpTool: McpClientTool,
): AgentTool<TSchema, unknown> {
  const qualifiedName = `mcp__${server.mcpServerName}__${mcpTool.name}`;
  // pi-agent-core treats `parameters` as a JSON-Schema-shaped TSchema; pass the
  // MCP inputSchema through as an unsafe TypeBox schema so we don't have to
  // translate JSON Schema → TypeBox per tool.
  const parameters = Type.Unsafe<unknown>(
    (mcpTool.inputSchema ?? { type: "object" }) as Record<string, unknown>,
  );

  return {
    name: qualifiedName,
    label: mcpTool.name,
    description: mcpTool.description ?? `MCP tool ${mcpTool.name} (server: ${server.mcpServerName})`,
    parameters,
    execute: async (_toolCallId: string, params: unknown): Promise<ToolResult> => {
      const ctx = getMcpToolContext();
      if (!ctx) {
        return {
          isError: true,
          content: [
            { type: "text", text: "MCP tool invoked outside of a turn context (no auth available)." },
          ],
        } as unknown as ToolResult;
      }

      const sharedScope =
        defaultToolingConfigurationProvider.getConfiguration().mcpPlatformAuthenticationScope;
      const scope = cachedScopesByServer.get(server.mcpServerName) ?? resolveTokenScopeForServer(server, sharedScope);

      let token: string;
      try {
        token = await AgenticAuthenticationService.GetAgenticUserToken(
          ctx.authorization,
          ctx.authHandlerName,
          ctx.turnContext,
          [scope],
        );
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `Failed to acquire agentic-user token for MCP server '${server.mcpServerName}' (scope ${scope}): ${err instanceof Error ? err.message : String(err)}`,
            },
          ],
        } as unknown as ToolResult;
      }

      const headers = {
        ...(server.headers ?? {}),
        Authorization: `Bearer ${token}`,
      };

      const transport = new StreamableHTTPClientTransport(new URL(server.url), {
        requestInit: { headers },
      });
      const client = new Client({ name: `${server.mcpServerName} Client`, version: "1.0" });

      try {
        await client.connect(transport);
        const result = await client.callTool({
          name: mcpTool.name,
          arguments: (params ?? {}) as Record<string, unknown>,
        });
        // MCP CallToolResult shape lines up with AgentToolResult: { content, isError? }.
        return result as unknown as ToolResult;
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `MCP tool '${qualifiedName}' failed: ${err instanceof Error ? err.message : String(err)}`,
            },
          ],
        } as unknown as ToolResult;
      } finally {
        try {
          await client.close();
        } catch {
          // best-effort
        }
      }
    },
  };
}
