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
