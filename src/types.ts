export type NodeType =
  | "tenant"
  | "site"
  | "subsites"
  | "lists"
  | "list"
  | "columns"
  | "contentTypes"
  | "contentType"
  | "views"
  | "permissions"
  | "sharingLinks"
  | "siteColumns"
  | "siteContentTypes"
  | "recycleBin"
  | "termStore"
  | "termGroup"
  | "termSet"
  | "term"
  | "hubSites"
  | "drives"
  | "driveItem"
  // Analytics mode node types
  | "analyticsRoot"
  | "analyticsTenantConfig"
  | "analyticsTenantConfigItem"
  | "analyticsAllSites"
  | "analyticsSite"
  | "analyticsSiteConfig"
  | "analyticsSitePermissions"
  | "analyticsSiteAudit"
  | "analyticsSiteOversharing"
  | "analyticsByOwner"
  | "analyticsOwnerGroup"
  | "analyticsByRisk"
  | "analyticsRiskLevel"
  | "analyticsRiskType"
  | "analyticsExternalUsers";

export type AppMode = "live" | "analytics";

export interface TreeNodeData {
  id: string;
  parentId: string | null;
  label: string;
  nodeType: NodeType;
  resourceId: string;
  siteId?: string;
  listId?: string;
  hasChildren: boolean;
  isLoaded: boolean;
  isLoading: boolean;
  isStale: boolean;
}

export interface CacheEntry<T = unknown> {
  key: string;
  data: T;
  fetchedAt: number;
  nodeType: NodeType;
}

export type ViewMode = "properties" | "json" | "table";

export interface AppSettings {
  cacheTtlMinutes: number;
  customClientId: string | null;
  customTenantId: string | null;
  proxyUrl: string | null;
  blobSasUrl: string | null;
  theme: "light" | "dark" | "system";
  defaultViewMode: ViewMode;
  /** Request Files.Read.All for sharing links and drive permissions */
  enableFilesAccess: boolean;
}

export const DEFAULT_SETTINGS: AppSettings = {
  cacheTtlMinutes: 30,
  customClientId: null,
  customTenantId: null,
  proxyUrl: null,
  blobSasUrl: null,
  theme: "system",
  defaultViewMode: "properties",
  enableFilesAccess: false,
};

/** Parsed M365DSC component from Sharepoint-Report.json */
export interface DscComponent {
  ResourceName: string;
  ResourceInstanceName: string;
  Url?: string;
  [key: string]: unknown;
}

/** Parsed blob data — all files loaded and indexed */
export interface BlobData {
  dscComponents: DscComponent[];
  sitesInventory: Record<string, unknown>[];
  permissionsReport: Record<string, unknown>[];
  oversharingAnalysis: Record<string, unknown>[];
  oversharingSummary: Record<string, unknown>;
  externalUsersReport: Record<string, unknown>[];
  sharingReport: Record<string, unknown>[];
  loadedAt: number;
  tenantName: string;
}
