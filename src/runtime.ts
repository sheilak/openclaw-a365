import type { PluginRuntime } from "openclaw/plugin-sdk";

/**
 * Module-level singleton for the A365 plugin runtime, pinned on globalThis.
 *
 * The pin matters because `@microsoft/opentelemetry` installs
 * `require-in-the-middle`, which can cause this module to be evaluated more
 * than once across realms. `setA365Runtime` is called from `index.ts` at
 * eager registration; `getA365Runtime` is called from modules loaded later
 * via dynamic import (monitor.ts, token.ts, graph-tools.ts). Without the
 * shared symbol, those modules see a fresh `null` runtime and throw.
 */
const RUNTIME_KEY = Symbol.for("openclaw-a365.runtime");
type RuntimeHolder = { value: PluginRuntime | null };
const holder = ((): RuntimeHolder => {
  const g = globalThis as unknown as Record<symbol, RuntimeHolder | undefined>;
  return (g[RUNTIME_KEY] ??= { value: null });
})();

/**
 * Set the A365 plugin runtime. Called once during plugin registration.
 */
export function setA365Runtime(next: PluginRuntime): void {
  holder.value = next;
}

/**
 * Get the A365 plugin runtime.
 * @throws Error if runtime has not been initialized via setA365Runtime
 */
export function getA365Runtime(): PluginRuntime {
  if (!holder.value) {
    throw new Error("A365 runtime not initialized - ensure plugin is registered before using runtime");
  }
  return holder.value;
}

/**
 * Reset the runtime singleton. For testing purposes only.
 * @internal
 */
export function resetA365Runtime(): void {
  holder.value = null;
}
