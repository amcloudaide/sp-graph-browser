import { Client } from "@microsoft/microsoft-graph-client";
import {
  AuthCodeMSALBrowserAuthenticationProvider,
} from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { PublicClientApplication, InteractionType, AccountInfo } from "@azure/msal-browser";
import { graphScopes } from "../auth/msalConfig";

export class GraphClient {
  private client: Client;

  constructor(msalInstance: PublicClientApplication, account: AccountInfo) {
    const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(msalInstance, {
      account,
      interactionType: InteractionType.Popup,
      scopes: graphScopes,
    });
    this.client = Client.initWithMiddleware({ authProvider });
  }

  async getAll<T>(endpoint: string): Promise<T[]> {
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
  }

  async get<T>(endpoint: string): Promise<T> {
    return this.client.api(endpoint).get();
  }

  /** List all sites using getAllSites (returns all sites including Teams) */
  async listSites() {
    try {
      // getAllSites returns all sites in the tenant (requires Sites.Read.All)
      return await this.getAll<Record<string, unknown>>("/sites/getAllSites");
    } catch {
      // Fallback to search if getAllSites is not available
      return this.getAll<Record<string, unknown>>("/sites?search=*");
    }
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
