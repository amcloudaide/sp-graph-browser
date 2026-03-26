import { Configuration, LogLevel } from "@azure/msal-browser";

const DEFAULT_CLIENT_ID = "YOUR_MULTI_TENANT_CLIENT_ID"; // TODO: replace after app registration

export function getMsalConfig(customClientId?: string | null): Configuration {
  return {
    auth: {
      clientId: customClientId || DEFAULT_CLIENT_ID,
      authority: "https://login.microsoftonline.com/common/",
      redirectUri: window.location.origin,
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

export const graphScopes = ["User.Read", "Sites.Read.All", "TermStore.Read.All"];

export function getSharePointScopes(tenantName: string): string[] {
  return [`https://${tenantName}.sharepoint.com/AllSites.Read`];
}
