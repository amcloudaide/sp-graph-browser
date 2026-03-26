import { createContext, useContext, useState, useCallback, useMemo, useEffect } from "react";
import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  AccountInfo,
} from "@azure/msal-browser";
import { MsalProvider, useMsal, useIsAuthenticated } from "@azure/msal-react";
import { getMsalConfig, graphScopes, getSharePointScopes } from "./msalConfig";

interface AuthContextValue {
  isAuthenticated: boolean;
  account: AccountInfo | null;
  tenantName: string | null;
  login: () => Promise<void>;
  logout: () => void;
  getGraphToken: () => Promise<string>;
  getSharePointToken: () => Promise<string>;
}

const AuthContext = createContext<AuthContextValue | null>(null);

export function useAuth(): AuthContextValue {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error("useAuth must be used within AuthProvider");
  return ctx;
}

function AuthContextProvider({ children }: { children: React.ReactNode }) {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const account = accounts[0] || null;

  const [spTenantName, setSpTenantName] = useState<string | null>(null);

  // After login, discover the SharePoint tenant name from the root site URL
  useEffect(() => {
    if (!isAuthenticated || !account || spTenantName) return;
    (async () => {
      try {
        const token = await instance.acquireTokenSilent({ scopes: graphScopes, account });
        const resp = await fetch("https://graph.microsoft.com/v1.0/sites/root", {
          headers: { Authorization: `Bearer ${token.accessToken}` },
        });
        const site = await resp.json();
        const hostname: string = site.siteCollection?.hostname ?? "";
        const tenant = hostname.split(".")[0];
        if (tenant) setSpTenantName(tenant);
      } catch {
        // Non-blocking — SP REST features will be unavailable
      }
    })();
  }, [isAuthenticated, account, spTenantName, instance]);

  const login = useCallback(async () => {
    await instance.loginPopup({ scopes: graphScopes });
  }, [instance]);

  const logout = useCallback(() => {
    instance.logoutPopup();
  }, [instance]);

  const getGraphToken = useCallback(async (): Promise<string> => {
    if (!account) throw new Error("Not authenticated");
    try {
      const response = await instance.acquireTokenSilent({ scopes: graphScopes, account });
      return response.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const response = await instance.acquireTokenPopup({ scopes: graphScopes, account });
        return response.accessToken;
      }
      throw error;
    }
  }, [instance, account]);

  const getSharePointToken = useCallback(async (): Promise<string> => {
    if (!account || !spTenantName) throw new Error("Not authenticated or tenant unknown");
    try {
      const response = await instance.acquireTokenSilent({
        scopes: getSharePointScopes(spTenantName),
        account,
      });
      return response.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const response = await instance.acquireTokenPopup({
          scopes: getSharePointScopes(spTenantName),
          account,
        });
        return response.accessToken;
      }
      throw error;
    }
  }, [instance, account, spTenantName]);

  const value: AuthContextValue = {
    isAuthenticated,
    account,
    tenantName: spTenantName,
    login,
    logout,
    getGraphToken,
    getSharePointToken,
  };

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

interface AuthProviderProps {
  customClientId?: string | null;
  children: React.ReactNode;
}

export function AuthProvider({ customClientId, children }: AuthProviderProps) {
  const msalInstance = useMemo(
    () => new PublicClientApplication(getMsalConfig(customClientId)),
    [customClientId]
  );

  return (
    <MsalProvider instance={msalInstance}>
      <AuthContextProvider>{children}</AuthContextProvider>
    </MsalProvider>
  );
}
