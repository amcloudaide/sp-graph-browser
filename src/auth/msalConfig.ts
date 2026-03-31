import { Configuration, LogLevel } from "@azure/msal-browser";

const DEFAULT_CLIENT_ID = "dcdb302b-f293-489f-9227-c8922f8e819d";
const DEFAULT_TENANT_ID = "fb677867-2d1e-40d2-a687-cb0979be2d90";

export function getMsalConfig(customClientId?: string | null, customTenantId?: string | null): Configuration {
  const tenantId = customTenantId || DEFAULT_TENANT_ID;
  return {
    auth: {
      clientId: customClientId || DEFAULT_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${tenantId}/`,
      redirectUri: window.location.href.split("?")[0].split("#")[0],
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false,
    },
    system: {
      loggerOptions: {
        logLevel: LogLevel.Warning,
      },
    },
  };
}

export const graphScopesBase = ["User.Read", "Sites.Read.All", "TermStore.Read.All"];
export const graphScopesFiles = ["Files.Read.All"];

export function getGraphScopes(enableFilesAccess = false): string[] {
  return enableFilesAccess
    ? [...graphScopesBase, ...graphScopesFiles]
    : graphScopesBase;
}

/** Default scopes for initial login (without files) */
export const graphScopes = graphScopesBase;

export function getSharePointScopes(tenantName: string): string[] {
  return [`https://${tenantName}.sharepoint.com/AllSites.Read`];
}
