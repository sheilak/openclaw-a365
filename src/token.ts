import type { A365Config } from "./types.js";

export type A365Credentials = {
  appId: string;
  appPassword: string;
  tenantId: string;
};

/**
 * Resolve A365 Bot Framework credentials from config or environment variables.
 */
export function resolveA365Credentials(cfg?: A365Config): A365Credentials | undefined {
  const appId = cfg?.appId?.trim() || process.env.A365_APP_ID?.trim();
  const appPassword = cfg?.appPassword?.trim() || process.env.A365_APP_PASSWORD?.trim();
  const tenantId = cfg?.tenantId?.trim() || process.env.A365_TENANT_ID?.trim();

  if (!appId || !appPassword || !tenantId) {
    return undefined;
  }

  return { appId, appPassword, tenantId };
}
