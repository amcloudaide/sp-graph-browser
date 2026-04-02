import type { BlobData, DscComponent } from "../types";
import type { CacheStore } from "./cacheStore";

/** Resolve M365DSC PowerShell URL expressions to actual URLs */
export function resolveDscUrl(raw: string, tenantName: string): string {
  // Replace $($OrganizationName.Split('.')[0]) with tenant name
  return raw.replace(
    /\$\(\$OrganizationName\.Split\([^)]+\)\[0\]\)/g,
    tenantName
  );
}

/** Extract the base URL and SAS token from a full SAS URL */
function parseSasUrl(sasUrl: string): { baseUrl: string; sasToken: string } {
  const idx = sasUrl.indexOf("?");
  if (idx === -1) return { baseUrl: sasUrl, sasToken: "" };
  return { baseUrl: sasUrl.substring(0, idx), sasToken: sasUrl.substring(idx) };
}

/** Fetch a JSON file from blob storage */
async function fetchBlobJson<T>(baseUrl: string, path: string, sasToken: string): Promise<T> {
  const url = `${baseUrl}/${path}${sasToken}`;
  console.log(`[SP Graph Browser] Fetching blob: ${path}`);
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Blob fetch failed for ${path}: ${response.status} ${response.statusText}`);
  }
  return response.json() as Promise<T>;
}

/** Fetch and decode the M365DSC Sharepoint-Report.json (UTF-16 LE encoded) */
async function fetchDscReport(baseUrl: string, sasToken: string): Promise<DscComponent[]> {
  const url = `${baseUrl}/sharepoint/latest/Sharepoint-Report.json${sasToken}`;
  console.log("[SP Graph Browser] Fetching M365DSC report (UTF-16)...");
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`M365DSC report fetch failed: ${response.status}`);
  }
  const buffer = await response.arrayBuffer();
  // Decode UTF-16 LE (with BOM)
  const decoder = new TextDecoder("utf-16le");
  const text = decoder.decode(buffer);
  // Strip BOM if present
  const clean = text.charCodeAt(0) === 0xFEFF ? text.slice(1) : text;
  return JSON.parse(clean) as DscComponent[];
}

/** Extract tenant name from sites-inventory data */
function extractTenantName(sites: Record<string, unknown>[]): string {
  for (const site of sites) {
    const url = site.Url as string;
    if (url) {
      const match = url.match(/https:\/\/([^.]+)\.sharepoint\.com/);
      if (match) return match[1];
    }
  }
  return "unknown";
}

export class BlobClient {
  private cache: CacheStore;

  constructor(cache: CacheStore) {
    this.cache = cache;
  }

  /** Load all data from blob storage. Returns parsed and indexed data. */
  async loadAll(sasUrl: string, onProgress?: (msg: string) => void): Promise<BlobData> {
    // Check cache first
    const cached = await this.cache.get("blob:all");
    if (cached) {
      const data = cached.data as BlobData;
      console.log(`[SP Graph Browser] Using cached blob data (loaded ${new Date(data.loadedAt).toISOString()})`);
      return data;
    }

    const { baseUrl, sasToken } = parseSasUrl(sasUrl);

    onProgress?.("Loading analytics data...");

    // Fetch analytics files in parallel
    const [sitesInventory, permissionsReport, oversharingAnalysis, oversharingSummary, externalUsersReport, sharingReport] =
      await Promise.all([
        fetchBlobJson<Record<string, unknown>[]>(baseUrl, "sharepoint/analytics/sites-inventory.json", sasToken).catch(() => []),
        fetchBlobJson<Record<string, unknown>[]>(baseUrl, "sharepoint/analytics/permissions-report.json", sasToken).catch(() => []),
        fetchBlobJson<Record<string, unknown>[]>(baseUrl, "sharepoint/analytics/oversharing-analysis.json", sasToken).catch(() => []),
        fetchBlobJson<Record<string, unknown>>(baseUrl, "sharepoint/analytics/oversharing-summary.json", sasToken).catch(() => ({})),
        fetchBlobJson<Record<string, unknown>[]>(baseUrl, "sharepoint/analytics/external-users-report.json", sasToken).catch(() => []),
        fetchBlobJson<Record<string, unknown>[]>(baseUrl, "sharepoint/analytics/sharing-report.json", sasToken).catch(() => []),
      ]);

    const tenantName = extractTenantName(sitesInventory);

    onProgress?.("Loading M365DSC configuration snapshot...");

    // Fetch M365DSC report (large file, ~50MB)
    let dscComponents: DscComponent[] = [];
    try {
      dscComponents = await fetchDscReport(baseUrl, sasToken);
      // Resolve PowerShell URL expressions
      for (const comp of dscComponents) {
        if (comp.Url && typeof comp.Url === "string") {
          comp.Url = resolveDscUrl(comp.Url, tenantName);
        }
      }
      console.log(`[SP Graph Browser] M365DSC: ${dscComponents.length} components loaded`);
    } catch (e) {
      console.warn("[SP Graph Browser] M365DSC report failed:", e);
    }

    onProgress?.("Indexing data...");

    const data: BlobData = {
      dscComponents,
      sitesInventory,
      permissionsReport,
      oversharingAnalysis,
      oversharingSummary,
      externalUsersReport,
      sharingReport,
      loadedAt: Date.now(),
      tenantName,
    };

    // Cache the parsed data
    await this.cache.set("blob:all", data, "analyticsRoot");

    console.log(`[SP Graph Browser] Blob data loaded: ${sitesInventory.length} sites, ${dscComponents.length} DSC components`);
    return data;
  }

  /** Clear cached blob data (for refresh) */
  async clearCache(): Promise<void> {
    await this.cache.invalidateByPrefix("blob:");
  }
}
