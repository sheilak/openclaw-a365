import { useMicrosoftOpenTelemetry, shutdownMicrosoftOpenTelemetry, AgenticTokenCacheInstance } from "@microsoft/opentelemetry";
import { resourceFromAttributes } from "@opentelemetry/resources";
import { SimpleSpanProcessor, type SpanExporter, type ReadableSpan } from "@opentelemetry/sdk-trace-base";
import { ExportResultCode, type ExportResult } from "@opentelemetry/core";
import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";

let initialized = false;

/**
 * File-based span exporter. The openclaw gateway daemon eats stdout, so the
 * default ConsoleSpanExporter output is invisible. This writes one JSON line
 * per span to a dedicated file we can tail.
 */
class FileSpanExporter implements SpanExporter {
  private stream: fs.WriteStream;

  constructor(filePath: string) {
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    this.stream = fs.createWriteStream(filePath, { flags: "a" });
  }

  export(spans: ReadableSpan[], resultCallback: (result: ExportResult) => void): void {
    try {
      for (const s of spans) {
        const ctx = s.spanContext();
        const line = JSON.stringify({
          time: new Date().toISOString(),
          name: s.name,
          kind: s.kind,
          traceId: ctx.traceId,
          spanId: ctx.spanId,
          parentSpanId: s.parentSpanContext?.spanId,
          startTime: s.startTime,
          endTime: s.endTime,
          durationMs: (s.endTime[0] - s.startTime[0]) * 1000 + (s.endTime[1] - s.startTime[1]) / 1e6,
          status: s.status,
          attributes: s.attributes,
          events: s.events,
        }) + "\n";
        this.stream.write(line);
      }
      resultCallback({ code: ExportResultCode.SUCCESS });
    } catch (err) {
      resultCallback({ code: ExportResultCode.FAILED, error: err as Error });
    }
  }

  async shutdown(): Promise<void> {
    await new Promise<void>((resolve) => this.stream.end(() => resolve()));
  }
}

export function initObservability(): void {
  if (initialized) return;

  const a365Enabled = process.env.ENABLE_A365_OBSERVABILITY_EXPORTER !== "false";

  try {
    // Always log exporter diagnostics (token resolution + OtelWrite POST results)
    // so export failures are visible without needing an env var to debug.
    // configureA365Logger isn't exported from the package entry; the exporter
    // reads this env var at init, so set it in-process before useMicrosoftOpenTelemetry.
    if (!process.env.A365_OBSERVABILITY_LOG_LEVEL) {
      process.env.A365_OBSERVABILITY_LOG_LEVEL = "info|warn|error";
    }
    // The A365 logger forwards through @opentelemetry/api `diag`, which defaults
    // to WARN — without this, info-level export diagnostics get dropped.
    if (!process.env.OTEL_LOG_LEVEL) {
      process.env.OTEL_LOG_LEVEL = "INFO";
    }

    const spanFile = path.join(os.homedir(), ".openclaw", "logs", "a365-spans.jsonl");
    useMicrosoftOpenTelemetry({
      resource: resourceFromAttributes({
        "service.name": "openclaw-a365",
        "service.version": "2026.2.8-2",
      }),
      enableConsoleExporters: false,
      spanProcessors: [new SimpleSpanProcessor(new FileSpanExporter(spanFile))],
      azureMonitor: { enabled: false },
      a365: {
        enabled: a365Enabled,
        enableObservabilityExporter: a365Enabled,
        tokenResolver: (agentId, tenantId) =>
          AgenticTokenCacheInstance.getObservabilityToken(agentId, tenantId) ?? "",
      },
    });

    initialized = true;
    console.log(
      `[a365] OpenTelemetry initialized; a365 export ${a365Enabled ? "ENABLED" : "disabled"}; local spans → ${spanFile}`,
    );

    process.once("beforeExit", () => {
      void shutdownMicrosoftOpenTelemetry();
    });
  } catch (err) {
    console.error("[a365] OpenTelemetry init failed; continuing without:", err);
  }
}

/**
 * Preload the observability token into the AgenticTokenCache instance used by
 * the exporter's tokenResolver. Lives here (not monitor.ts) because
 * `require-in-the-middle` (loaded by `@microsoft/opentelemetry`) can split the
 * module realm so dynamically-imported modules see a different
 * `AgenticTokenCacheInstance` singleton than this one.
 */
export async function preloadObservabilityToken(
  agentId: string,
  tenantId: string,
  turnContext: unknown,
  authorization: unknown,
  scopes: string[],
): Promise<void> {
  await AgenticTokenCacheInstance.refreshObservabilityToken(
    agentId,
    tenantId,
    turnContext as Parameters<typeof AgenticTokenCacheInstance.refreshObservabilityToken>[2],
    authorization as Parameters<typeof AgenticTokenCacheInstance.refreshObservabilityToken>[3],
    scopes,
  );
}
