import type { OpenClawConfig, RuntimeEnv } from "openclaw/plugin-sdk";
import type { A365Config, A365MessageMetadata } from "./types.js";
import { getA365Runtime } from "./runtime.js";
import { runWithMcpToolContext, primeMcpToolsForTurn } from "./mcp-tools.js";
import { resolveA365Credentials } from "./token.js";
import { saveConversationReference } from "./conversation-store.js";
import { setAdapter, setBlueprintClientId } from "./adapter-store.js";
import { InvokeAgentScope, OutputScope, BaggageMiddleware } from "@microsoft/opentelemetry";
import type { A365SpanDetails, OutputResponse } from "@microsoft/opentelemetry";
import { defaultObservabilityConfigurationProvider } from "@microsoft/agents-a365-observability";
import { defaultToolingConfigurationProvider } from "@microsoft/agents-a365-tooling";
import { preloadObservabilityToken } from "./observability.js";

const getAuthScopes = (): string[] => {
  const obs =
    defaultObservabilityConfigurationProvider.getConfiguration().observabilityAuthenticationScopes;
  const mcp = defaultToolingConfigurationProvider.getConfiguration().mcpPlatformAuthenticationScope;
  // De-dup just in case the providers ever overlap.
  return Array.from(new Set([...obs, mcp]));
};

const getObservabilityScopes = (): string[] => [
  ...defaultObservabilityConfigurationProvider.getConfiguration().observabilityAuthenticationScopes,
];

export type MonitorA365Opts = {
  cfg: OpenClawConfig;
  runtime?: RuntimeEnv;
  abortSignal?: AbortSignal;
};

export type MonitorA365Result = {
  app: unknown;
  shutdown: () => Promise<void>;
};

/**
 * Activity shape for metadata extraction.
 */
export type ActivityForMetadata = {
  from?: { id: string; name?: string; aadObjectId?: string };
  recipient?: { id: string; name?: string };
  conversation?: { id: string; isGroup?: boolean; tenantId?: string };
  serviceUrl?: string;
  id?: string;
  channelId?: string;
  locale?: string;
  channelData?: {
    tenant?: { id: string };
    team?: { id: string; name?: string };
    channel?: { id: string; name?: string };
  };
};

/**
 * Extract message metadata from an Agents SDK activity.
 */
export function extractMessageMetadata(activity: ActivityForMetadata): A365MessageMetadata {
  return {
    userId: activity.from?.id || "",
    userEmail: activity.from?.aadObjectId || activity.from?.id,
    userName: activity.from?.name,
    userAadId: activity.from?.aadObjectId,
    conversationId: activity.conversation?.id || "",
    isGroup: activity.conversation?.isGroup || false,
    tenantId: activity.conversation?.tenantId || activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl || "",
    activityId: activity.id,
    channelId: activity.channelId,
    teamId: activity.channelData?.team?.id,
    teamName: activity.channelData?.team?.name,
    channelName: activity.channelData?.channel?.name,
  };
}

/**
 * Build a StoredConversationReference from an activity for proactive messaging.
 */
export function buildConversationReference(activity: ActivityForMetadata): StoredConversationReference {
  return {
    conversationId: activity.conversation?.id || "",
    serviceUrl: activity.serviceUrl || "",
    channelId: activity.channelId || "msteams",
    botId: activity.recipient?.id || "",
    botName: activity.recipient?.name,
    userId: activity.from?.id || "",
    userName: activity.from?.name,
    userAadId: activity.from?.aadObjectId,
    tenantId: activity.conversation?.tenantId || activity.channelData?.tenant?.id,
    isGroup: activity.conversation?.isGroup || false,
    locale: activity.locale,
    updatedAt: Date.now(),
  };
}

/**
 * Start the A365 Microsoft Agents provider.
 */
export async function monitorA365Provider(opts: MonitorA365Opts): Promise<MonitorA365Result> {
  const core = getA365Runtime();
  const log = core.logging.getChildLogger({ name: "a365" });
  const cfg = opts.cfg;
  const a365Cfg = cfg.channels?.a365 as A365Config | undefined;

  if (!a365Cfg?.enabled) {
    log.debug("a365 provider disabled");
    return { app: null, shutdown: async () => {} };
  }

  const runtime: RuntimeEnv = opts.runtime ?? {
    log: console.log,
    error: console.error,
    exit: (code: number): never => {
      throw new Error(`exit ${code}`);
    },
  };

  const port = a365Cfg.webhook?.port ?? 3978;

  log.info(`starting a365 provider (port ${port})`);

  const creds = resolveA365Credentials(a365Cfg);
  if (!creds) {
    log.error("A365 credentials not configured - set appId/appPassword/tenantId in config or A365_APP_ID/A365_APP_PASSWORD/A365_TENANT_ID env vars");
    return { app: null, shutdown: async () => {} };
  }

  process.env["connections__serviceConnection__settings__clientId"] = creds.appId;
  process.env["connections__serviceConnection__settings__clientSecret"] = creds.appPassword;
  process.env["connections__serviceConnection__settings__tenantId"] = creds.tenantId;
  process.env["connectionsMap__0__connection"] = "serviceConnection";
  process.env["connectionsMap__0__serviceUrl"] = "*";

  process.env.MicrosoftAppId = creds.appId;
  process.env.MicrosoftAppPassword = creds.appPassword;
  process.env.MicrosoftAppTenantId = creds.tenantId;
  process.env.MicrosoftAppType = "SingleTenant";
  process.env.PORT = String(port);

  const { AgentApplication, MemoryStorage, TurnContext, TurnState } = await import(
    "@microsoft/agents-hosting"
  );
  const { ActivityTypes } = await import("@microsoft/agents-activity");

  type ApplicationTurnState = typeof TurnState;

  // The single `agentic` AgenticUserAuthorization handler serves both A365
  // observability token exchange and per-MCP-server OBO. Scopes union both
  // surfaces so authorization.exchangeToken("agentic", { scopes }) succeeds
  // for either at turn time.
  const storage = new MemoryStorage();
  const agentApp = new AgentApplication<ApplicationTurnState>({
    storage,
    authorization: {
      agentic: {
        type: "AgenticUserAuthorization",
        scopes: getAuthScopes(),
      },
    },
  });

  agentApp.onConversationUpdate("membersAdded", async (context: typeof TurnContext) => {
    log.debug("members added event");
    const welcomeMessage = a365Cfg?.welcomeMessage;
    if (welcomeMessage !== undefined && welcomeMessage !== "") {
      await context.sendActivity(welcomeMessage);
    }
  });

  agentApp.onActivity(
    ActivityTypes.Message,
    async (context: typeof TurnContext, _state: ApplicationTurnState) => {
      const activity = context.activity;
      const text = activity.text?.trim();

      if (!text) {
        log.debug("skipping empty message");
        return;
      }

      const metadata = extractMessageMetadata(activity);
      log.info("received message", {
        from: metadata.userName || metadata.userId,
        isGroup: metadata.isGroup,
        textLength: text.length,
      });

      const obsAgentId =
        (activity.recipient as { agenticAppId?: string } | undefined)?.agenticAppId ??
        a365Cfg?.graph?.aaInstanceId ??
        "";
      const obsTenantId = activity.recipient?.tenantId ?? creds.tenantId ?? "";
      try {
        await preloadObservabilityToken(
          obsAgentId,
          obsTenantId,
          context,
          agentApp.authorization,
          getObservabilityScopes(),
        );
      } catch (err) {
        log.warn(
          `preloadObservabilityToken failed: agentId=${obsAgentId} tenantId=${obsTenantId} error=${err instanceof Error ? err.message : String(err)}`,
        );
        console.warn("[a365] preloadObservabilityToken failed:", err);
      }

      // Lazy MCP discovery — first turn populates the module-level cache that
      // channel.ts's sync `agentTools` factory reads on every subsequent run.
      try {
        await primeMcpToolsForTurn(a365Cfg, context, agentApp.authorization, "agentic");
      } catch (err) {
        log.warn(
          `primeMcpToolsForTurn failed: ${err instanceof Error ? err.message : String(err)}`,
        );
      }

      const invokeScope = InvokeAgentScope.start(
        { content: text },
        {},
        {
          agentId: a365Cfg?.graph?.aaInstanceId || a365Cfg?.agentIdentity || creds.appId,
          agentEmail: a365Cfg?.agentIdentity,
          tenantId: creds.tenantId,
        },
      );

      try {
        await invokeScope.withActiveSpanAsync(async () => {
          invokeScope.recordInputMessages(text);

      try {
        const convRef = activity.getConversationReference();
        log.info("Saving conversation reference", {
          conversationId: convRef.conversation?.id,
          serviceUrl: convRef.serviceUrl,
          agentRole: convRef.agent?.role,
        });
        await saveConversationReference(convRef, metadata.userAadId);
        log.info("Conversation reference saved successfully");
      } catch (err) {
        log.error(`Failed to save conversation reference: ${String(err)}`)
      }

      const allowFrom = a365Cfg?.allowFrom;
      if (allowFrom && allowFrom.length > 0 && !allowFrom.includes("*")) {
        const userAllowed =
          allowFrom.includes(metadata.userId) ||
          allowFrom.includes(metadata.userEmail || "") ||
          allowFrom.includes(metadata.userAadId || "");
        if (!userAllowed) {
          log.debug("user not in allowlist", { userId: metadata.userId });
          return;
        }
      }

      const isOwner =
        (a365Cfg?.owner &&
          metadata.userEmail?.toLowerCase() === a365Cfg.owner.toLowerCase()) ||
        (a365Cfg?.ownerAadId && metadata.userAadId === a365Cfg.ownerAadId);
      const userRole = isOwner ? "Owner" : "Requester";

      // Per-turn MCP context for tool execute closures (read via ALS).
      await runWithMcpToolContext(
        {
          turnContext: context,
          authorization: agentApp.authorization,
          authHandlerName: "agentic",
        },
        async () => {
          const senderId = metadata.userAadId || metadata.userId;
          const conversationId = metadata.conversationId;
          const isDirectMessage = !metadata.isGroup;

          const route = core.channel.routing.resolveAgentRoute({
            cfg,
            channel: "a365",
            peer: {
              kind: isDirectMessage ? "dm" : "group",
              id: isDirectMessage ? senderId : conversationId,
            },
          });

          const a365From = isDirectMessage
            ? `a365:${senderId}`
            : `a365:group:${conversationId}`;
          const a365To = isDirectMessage ? `user:${senderId}` : `conversation:${conversationId}`;

          const ctxPayload = core.channel.reply.finalizeInboundContext({
            Body: text,
            RawBody: text,
            CommandBody: text,
            From: a365From,
            To: a365To,
            SessionKey: route.sessionKey,
            AccountId: route.accountId,
            ChatType: isDirectMessage ? "direct" : "group",
            ConversationLabel: metadata.userName || senderId,
            SenderName: metadata.userName || senderId,
            SenderId: senderId,
            Provider: "a365" as const,
            Surface: "a365" as const,
            MessageSid: metadata.activityId,
            Timestamp: Date.now(),
            WasMentioned: true,
            CommandAuthorized: isOwner,
            OriginatingChannel: "a365" as const,
            OriginatingTo: conversationId,
          });

          let replyCount = 0;
          const pendingSends: Promise<void>[] = [];
          const collectedReplies: string[] = [];

          const sendReply = async (replyText: string) => {
            try {
              log.debug("sendReply called", { length: replyText.length });
              const result = await context.sendActivity(replyText);
              replyCount++;
              collectedReplies.push(replyText);
              log.debug("reply sent successfully", { replyCount, resultId: result?.id });
            } catch (sendErr) {
              const err = sendErr as Error;
              log.error("sendActivity failed", { error: err?.message });
            }
          };

          const queuedCounts = { tool: 0, block: 0, final: 0 };
          const dispatcher = {
            sendToolResult: (payload: { text?: string }) => {
              if (payload.text) {
                queuedCounts.tool++;
                pendingSends.push(sendReply(payload.text));
              }
              return true;
            },
            sendBlockReply: (payload: { text?: string }) => {
              if (payload.text) {
                queuedCounts.block++;
                pendingSends.push(sendReply(payload.text));
              }
              return true;
            },
            sendFinalReply: (payload: { text?: string }) => {
              if (payload.text) {
                queuedCounts.final++;
                pendingSends.push(sendReply(payload.text));
              }
              return true;
            },
            waitForIdle: async () => {
              await Promise.all(pendingSends);
            },
            getQueuedCounts: () => queuedCounts,
          };

          const replyOptions = {
            onReplyStart: async () => {
              try {
                log.debug("sending typing indicator");
                await context.sendActivity({ type: "typing" });
                log.debug("typing indicator sent");
              } catch (typingErr) {
                const err = typingErr as Error;
                log.debug("typing indicator failed", {
                  error: String(err),
                  message: err?.message,
                  stack: err?.stack,
                });
              }
            },
            onTypingController: () => {},
            onTypingCleanup: () => {},
          };

          try {
            log.info("dispatching to agent", { sessionKey: route.sessionKey });

            const { queuedFinal, counts } = await core.channel.reply.dispatchReplyFromConfig({
              ctx: ctxPayload,
              cfg,
              dispatcher,
              replyOptions,
            });

            await Promise.all(pendingSends);

            log.info("dispatch complete", { queuedFinal, textCount: counts?.text ?? 0, repliesSent: replyCount });

            if (collectedReplies.length > 0) {
              try {
                const parentContext = invokeScope.getSpanContext();
                const response: OutputResponse = { messages: collectedReplies };
                const outputScope = OutputScope.start(
                  { content: text },
                  response,
                  {
                    agentId: a365Cfg?.graph?.aaInstanceId || a365Cfg?.agentIdentity || creds.appId,
                    agentEmail: a365Cfg?.agentIdentity,
                    tenantId: creds.tenantId,
                  },
                  undefined,
                  { parentContext } as A365SpanDetails,
                );
                outputScope.dispose();
              } catch (outErr) {
                log.warn(`OutputScope emit failed: ${String(outErr)}`);
              }
            }

            try {
              const storePath = core.channel.session.resolveStorePath(cfg.session?.store);
              const mainSessionKey = "agent:main:main";
              await core.channel.session.updateLastRoute({
                storePath,
                sessionKey: mainSessionKey,
                channel: "a365",
                to: conversationId,
              });
              log.info("Updated main session for cron delivery", { conversationId });
            } catch (updateErr) {
              log.error(`Failed to update main session: ${String(updateErr)}`);
            }
          } catch (err) {
            log.error("handler failed", { error: String(err) });
            runtime.error?.(`a365 handler failed: ${String(err)}`);

            try {
              await context.sendActivity(
                "I encountered an error processing your message. Please try again.",
              );
            } catch {
              // Ignore send failure
            }
          }
        },
      );
        });
      } catch (err) {
        invokeScope.recordError(err as Error);
        throw err;
      } finally {
        invokeScope.dispose();
      }
    },
  );

  const { CloudAdapter } = await import("@microsoft/agents-hosting");
  const adapter = (agentApp.adapter ?? new CloudAdapter()) as InstanceType<typeof CloudAdapter>;
  (adapter as unknown as { use: (m: unknown) => void }).use(new BaggageMiddleware());
  setAdapter(adapter);

  const blueprintClientId =
    a365Cfg?.graph?.blueprintClientAppId?.trim() ||
    process.env.BLUEPRINT_CLIENT_APP_ID?.trim() ||
    creds.appId;
  setBlueprintClientId(blueprintClientId);
  log.info("Stored adapter and blueprint client ID for proactive messaging", { blueprintClientId });

  const { startServer } = await import("@microsoft/agents-hosting-express");

  const serverPromise = startServer(agentApp);

  log.info(`a365 provider started on port ${port}`);

  const shutdown = async () => {
    log.info("shutting down a365 provider");
  };

  if (opts.abortSignal) {
    opts.abortSignal.addEventListener("abort", () => {
      void shutdown();
    });
  }

  return { app: agentApp, shutdown };
}
