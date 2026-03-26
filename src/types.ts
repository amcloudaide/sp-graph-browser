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
  | "hubSites";

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
  theme: "light" | "dark" | "system";
  defaultViewMode: ViewMode;
}

export const DEFAULT_SETTINGS: AppSettings = {
  cacheTtlMinutes: 30,
  customClientId: null,
  customTenantId: null,
  theme: "system",
  defaultViewMode: "properties",
};
