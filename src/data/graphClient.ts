import { Client } from "@microsoft/microsoft-graph-client";
import {
  AuthCodeMSALBrowserAuthenticationProvider,
} from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { PublicClientApplication, InteractionType, AccountInfo } from "@azure/msal-browser";
import { graphScopes } from "../auth/msalConfig";

export class GraphClient {
  private client: Client;
  private msalInstance: PublicClientApplication;
  private account: AccountInfo;
  private proxyUrl: string | null = null;

  constructor(msalInstance: PublicClientApplication, account: AccountInfo) {
    this.msalInstance = msalInstance;
    this.account = account;
    const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(msalInstance, {
      account,
      interactionType: InteractionType.Popup,
      scopes: graphScopes,
    });
    this.client = Client.initWithMiddleware({ authProvider });
  }

  setProxyUrl(url: string | null) {
    this.proxyUrl = url;
  }

  private async getAccessToken(): Promise<string> {
    const response = await this.msalInstance.acquireTokenSilent({
      scopes: graphScopes,
      account: this.account,
    });
    return response.accessToken;
  }

  /** Call a Graph endpoint via the proxy (app-only auth). Returns the raw result. */
  async callViaProxy<T>(path: string, apiVersion = "beta"): Promise<T> {
    if (!this.proxyUrl) throw new Error("Proxy URL not configured");
    const token = await this.getAccessToken();
    const response = await fetch(`${this.proxyUrl}/api/graphProxy`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ path, apiVersion }),
    });
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(`Proxy ${response.status}: ${(err as Record<string, string>).error ?? response.statusText}`);
    }
    const result = await response.json() as { data: T; isCollection: boolean };
    return result.data;
  }

  async getAll<T>(endpoint: string): Promise<T[]> {
    try {
      const results: T[] = [];
      let url: string | null = endpoint;
      while (url) {
        const response = await this.client.api(url).get();
        if (response.value) {
          results.push(...response.value);
        }
        url = response["@odata.nextLink"] ?? null;
      }
      return results;
    } catch (e) {
      // On 403, try via proxy (app-only auth has broader access)
      if (this.proxyUrl && String(e).includes("Access denied")) {
        console.log(`[SP Graph Browser] Delegated 403 on ${endpoint}, trying proxy...`);
        return this.callViaProxy<T[]>(endpoint, "v1.0");
      }
      throw e;
    }
  }

  async get<T>(endpoint: string): Promise<T> {
    try {
      return await this.client.api(endpoint).get();
    } catch (e) {
      if (this.proxyUrl && String(e).includes("Access denied")) {
        console.log(`[SP Graph Browser] Delegated 403 on ${endpoint}, trying proxy...`);
        return this.callViaProxy<T>(endpoint, "v1.0");
      }
      throw e;
    }
  }

  /** List all sites — uses multiple search queries to maximize coverage.
   *  Note: getAllSites() only supports application permissions, not delegated. */
  async listSites() {
    // Approach 1: Use proxy if configured (tries new graphProxy, falls back to legacy getAllSites)
    if (this.proxyUrl) {
      try {
        console.log("[SP Graph Browser] Calling proxy for getAllSites...");
        const sites = await this.callViaProxy<Record<string, unknown>[]>("/sites/getAllSites()?$top=999");
        console.log(`[SP Graph Browser] Proxy returned ${sites.length} sites`);
        return sites;
      } catch (e) {
        // Try legacy endpoint for backward compatibility
        try {
          const token = await this.getAccessToken();
          const response = await fetch(`${this.proxyUrl}/api/getAllSites`, {
            method: "POST",
            headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          });
          if (response.ok) {
            const data = await response.json();
            const sites = data.sites as Record<string, unknown>[];
            console.log(`[SP Graph Browser] Legacy proxy returned ${sites.length} sites`);
            return sites;
          }
        } catch { /* fall through */ }
        console.warn("[SP Graph Browser] Proxy failed, falling back to search:", e);
      }
    }

    // Approach 2: search=* with multiple queries
    const seen = new Map<string, Record<string, unknown>>();

    // Run multiple search queries to maximize site discovery
    const queries = [
      "/sites?search=*",
      "/sites?search=site",
      "/sites?search=team",
      "/sites?search=project",
      "/sites?search=department",
    ];

    for (const query of queries) {
      try {
        let url: string | null = query;
        while (url) {
          const response: Record<string, unknown> = await this.client.api(url).get();
          const value = response.value as Record<string, unknown>[] | undefined;
          if (value) {
            for (const site of value) {
              const id = site.id as string;
              if (id && !seen.has(id)) {
                seen.set(id, site);
              }
            }
          }
          const nextLink: string | null = (response["@odata.nextLink"] as string) ?? null;
          url = nextLink ? nextLink.replace("https://graph.microsoft.com/v1.0", "") : null;
        }
      } catch (e) {
        console.warn(`[SP Graph Browser] search query failed:`, query, e);
      }
    }

    const results = Array.from(seen.values());
    console.log(`[SP Graph Browser] Site search returned ${results.length} unique sites (${queries.length} queries)`);
    return results;
  }

  /** Get tenant/organization info */
  async getOrganization() {
    const orgs = await this.getAll<Record<string, unknown>>("/organization");
    return orgs[0] ?? {};
  }

  /** Get the root site */
  async getRootSite() {
    return this.get<Record<string, unknown>>("/sites/root");
  }

  async getSite(siteId: string) {
    return this.get<Record<string, unknown>>(`/sites/${siteId}`);
  }

  async listSubsites(siteId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/sites`);
  }

  async listLists(siteId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/lists`);
  }

  async getList(siteId: string, listId: string) {
    return this.get<Record<string, unknown>>(`/sites/${siteId}/lists/${listId}`);
  }

  async listColumns(siteId: string, listId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/lists/${listId}/columns`);
  }

  async listSiteContentTypes(siteId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/contentTypes`);
  }

  async listListContentTypes(siteId: string, listId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/lists/${listId}/contentTypes`);
  }

  async listViews(siteId: string, listId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/lists/${listId}/views`);
  }

  async listSitePermissions(siteId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/permissions`);
  }

  async listSharingLinks(siteId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/drive/items/root/permissions`);
  }

  async listSiteColumns(siteId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/columns`);
  }

  /** List drives (document libraries) for a site */
  async listDrives(siteId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/drives`);
  }

  /** List children of a drive root or folder */
  async listDriveChildren(siteId: string, driveId: string, itemId?: string) {
    const path = itemId
      ? `/sites/${siteId}/drives/${driveId}/items/${itemId}/children`
      : `/sites/${siteId}/drives/${driveId}/root/children`;
    return this.getAll<Record<string, unknown>>(path);
  }

  /** Get content type columns */
  async listContentTypeColumns(siteId: string, contentTypeId: string) {
    return this.getAll<Record<string, unknown>>(`/sites/${siteId}/contentTypes/${contentTypeId}/columns`);
  }

  async listTermStoreGroups(siteId: string) {
    const response = await this.client.api(`/sites/${siteId}/termStore/groups`).version("beta").get();
    return response.value ?? [];
  }

  async listTermSets(siteId: string, groupId: string) {
    const response = await this.client
      .api(`/sites/${siteId}/termStore/groups/${groupId}/sets`)
      .version("beta")
      .get();
    return response.value ?? [];
  }

  async listTerms(siteId: string, setId: string) {
    const response = await this.client
      .api(`/sites/${siteId}/termStore/sets/${setId}/terms`)
      .version("beta")
      .get();
    return response.value ?? [];
  }
}
