import { PublicClientApplication, InteractionRequiredAuthError, AccountInfo } from "@azure/msal-browser";
import { getSharePointScopes } from "../auth/msalConfig";

export class SpRestClient {
  private msalInstance: PublicClientApplication;
  private account: AccountInfo;
  private tenantName: string;

  constructor(msalInstance: PublicClientApplication, account: AccountInfo, tenantName: string) {
    this.msalInstance = msalInstance;
    this.account = account;
    this.tenantName = tenantName;
  }

  private async getToken(): Promise<string> {
    try {
      const response = await this.msalInstance.acquireTokenSilent({
        scopes: getSharePointScopes(this.tenantName),
        account: this.account,
      });
      return response.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const response = await this.msalInstance.acquireTokenPopup({
          scopes: getSharePointScopes(this.tenantName),
          account: this.account,
        });
        return response.accessToken;
      }
      throw error;
    }
  }

  private async get<T>(siteUrl: string, apiPath: string): Promise<T> {
    const token = await this.getToken();
    const url = `${siteUrl}/_api/${apiPath}`;
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
      },
    });
    if (!response.ok) {
      throw new Error(`SP REST ${response.status}: ${response.statusText}`);
    }
    const json = await response.json();
    return json.value ?? json;
  }

  /** List all site collections via SharePoint Admin tenant API */
  async listAllSites(): Promise<Record<string, unknown>[]> {
    const token = await this.getToken();
    const adminUrl = `https://${this.tenantName}-admin.sharepoint.com`;
    const results: Record<string, unknown>[] = [];
    let startIndex = 0;
    const batchSize = 500;

    // eslint-disable-next-line no-constant-condition
    while (true) {
      const apiPath = `_api/SPO.Tenant/GetSitePropertiesFromSharePointByFilters?$skiptoken=SDFIndex%3D${startIndex}`;
      const url = `${adminUrl}/${apiPath}`;
      console.log(`[SP Graph Browser] SP Admin listing sites from index ${startIndex}...`);

      const response = await fetch(url, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
        },
        body: JSON.stringify({
          filter: null,
          startIndex: startIndex.toString(),
          includeDetail: true,
        }),
      });

      if (!response.ok) {
        console.warn(`[SP Graph Browser] SP Admin API ${response.status}: ${response.statusText}`);
        break;
      }

      const json = await response.json();
      const sites = json._Child_Items_ ?? json.value ?? [];
      if (sites.length === 0) break;

      results.push(...sites);
      startIndex += batchSize;

      // If we got fewer than batch size, we've reached the end
      if (sites.length < batchSize) break;
    }

    console.log(`[SP Graph Browser] SP Admin returned ${results.length} sites`);
    return results;
  }

  async listRoleAssignments(siteUrl: string) {
    return this.get<unknown[]>(
      siteUrl,
      "web/roleassignments?$expand=Member,RoleDefinitionBindings"
    );
  }

  async listRecycleBin(siteUrl: string) {
    return this.get<unknown[]>(siteUrl, "web/recyclebin?$top=200");
  }
}
