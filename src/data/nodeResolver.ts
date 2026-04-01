// sp-graph-browser/src/data/nodeResolver.ts
import type { NodeType, TreeNodeData } from "../types";
import type { GraphClient } from "./graphClient";
import type { SpRestClient } from "./spRestClient";
import type { CacheStore } from "./cacheStore";

export interface FetchContext {
  graph: GraphClient;
  spRest: SpRestClient | null;
  cache: CacheStore;
  ttlMinutes: number;
  enableFilesAccess: boolean;
}

export interface NodeDefinition {
  /** Fetch data for this node's properties panel */
  fetchDetails: (node: TreeNodeData, ctx: FetchContext) => Promise<unknown>;
  /** Fetch child nodes for the tree */
  fetchChildren: (node: TreeNodeData, ctx: FetchContext) => Promise<TreeNodeData[]>;
  /** Cache key for this node's data */
  cacheKey: (node: TreeNodeData) => string;
}

function makeId(parts: string[]): string {
  return parts.join(":");
}

const definitions: Record<NodeType, NodeDefinition> = {
  tenant: {
    cacheKey: () => "tenant",
    fetchDetails: async (_node, ctx) => {
      const [org, rootSite] = await Promise.all([
        ctx.graph.getOrganization(),
        ctx.graph.getRootSite(),
      ]);
      const allSites = await ctx.graph.listSites();
      const spSites = allSites.filter((s: Record<string, unknown>) => {
        const url = (s.webUrl as string) ?? "";
        return !url.includes("-my.sharepoint.com/personal/");
      });
      return {
        displayName: (org as Record<string, unknown>).displayName ?? "Unknown",
        tenantId: (org as Record<string, unknown>).id,
        verifiedDomains: (org as Record<string, unknown>).verifiedDomains,
        rootSiteUrl: (rootSite as Record<string, unknown>).webUrl,
        rootSiteId: (rootSite as Record<string, unknown>).id,
        totalSites: allSites.length,
        sharePointSites: spSites.length,
        oneDriveSites: allSites.length - spSites.length,
        createdDateTime: (org as Record<string, unknown>).createdDateTime,
        country: (org as Record<string, unknown>).country,
        preferredLanguage: (org as Record<string, unknown>).preferredLanguage,
      };
    },
    fetchChildren: async (_node, ctx) => {
      const allSites = await ctx.graph.listSites();
      // Filter out OneDrive personal sites (hosted on *-my.sharepoint.com/personal/)
      const sites = allSites.filter((site: Record<string, unknown>) => {
        const url = (site.webUrl as string) ?? "";
        return !url.includes("-my.sharepoint.com/personal/");
      });
      console.log(`[SP Graph Browser] Showing ${sites.length} SP sites (filtered ${allSites.length - sites.length} OneDrive sites)`);
      return sites.map((site: Record<string, unknown>) => ({
        id: makeId(["site", site.id as string]),
        parentId: "tenant",
        label: (site.displayName as string) || (site.webUrl as string),
        nodeType: "site" as NodeType,
        resourceId: site.id as string,
        siteId: site.id as string,
        hasChildren: true,
        isLoaded: false,
        isLoading: false,
        isStale: false,
      }));
    },
  },

  site: {
    cacheKey: (node) => `site:${node.resourceId}`,
    fetchDetails: async (node, ctx) => ctx.graph.getSite(node.siteId!),
    fetchChildren: async (node, _ctx) => {
      const siteId = node.siteId!;
      return [
        {
          id: makeId(["subsites", siteId]),
          parentId: node.id,
          label: "Subsites",
          nodeType: "subsites" as NodeType,
          resourceId: siteId,
          siteId,
          hasChildren: true,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
        {
          id: makeId(["lists", siteId]),
          parentId: node.id,
          label: "Lists & Libraries",
          nodeType: "lists" as NodeType,
          resourceId: siteId,
          siteId,
          hasChildren: true,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
        {
          id: makeId(["permissions", siteId]),
          parentId: node.id,
          label: "Permissions",
          nodeType: "permissions" as NodeType,
          resourceId: siteId,
          siteId,
          hasChildren: true,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
        {
          id: makeId(["siteContentTypes", siteId]),
          parentId: node.id,
          label: "Content Types",
          nodeType: "siteContentTypes" as NodeType,
          resourceId: siteId,
          siteId,
          hasChildren: true,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
        {
          id: makeId(["siteColumns", siteId]),
          parentId: node.id,
          label: "Site Columns",
          nodeType: "siteColumns" as NodeType,
          resourceId: siteId,
          siteId,
          hasChildren: false,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
        {
          id: makeId(["drives", siteId]),
          parentId: node.id,
          label: "Drives (Doc Libraries)",
          nodeType: "drives" as NodeType,
          resourceId: siteId,
          siteId,
          hasChildren: true,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
        {
          id: makeId(["recycleBin", siteId]),
          parentId: node.id,
          label: "Recycle Bin",
          nodeType: "recycleBin" as NodeType,
          resourceId: siteId,
          siteId,
          hasChildren: false,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
        {
          id: makeId(["termStore", siteId]),
          parentId: node.id,
          label: "Term Store",
          nodeType: "termStore" as NodeType,
          resourceId: siteId,
          siteId,
          hasChildren: true,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
      ];
    },
  },

  subsites: {
    cacheKey: (node) => `subsites:${node.siteId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listSubsites(node.siteId!),
    fetchChildren: async (node, ctx) => {
      const subsites = await ctx.graph.listSubsites(node.siteId!);
      return subsites.map((site: Record<string, unknown>) => ({
        id: makeId(["site", site.id as string]),
        parentId: node.id,
        label: (site.displayName as string) || (site.webUrl as string),
        nodeType: "site" as NodeType,
        resourceId: site.id as string,
        siteId: site.id as string,
        hasChildren: true,
        isLoaded: false,
        isLoading: false,
        isStale: false,
      }));
    },
  },

  lists: {
    cacheKey: (node) => `lists:${node.siteId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listLists(node.siteId!),
    fetchChildren: async (node, ctx) => {
      const lists = await ctx.graph.listLists(node.siteId!);
      return lists.map((list: Record<string, unknown>) => ({
        id: makeId(["list", node.siteId!, list.id as string]),
        parentId: node.id,
        label: (list.displayName as string) || (list.name as string),
        nodeType: "list" as NodeType,
        resourceId: list.id as string,
        siteId: node.siteId,
        listId: list.id as string,
        hasChildren: true,
        isLoaded: false,
        isLoading: false,
        isStale: false,
      }));
    },
  },

  list: {
    cacheKey: (node) => `list:${node.siteId}:${node.listId}`,
    fetchDetails: async (node, ctx) => ctx.graph.getList(node.siteId!, node.listId!),
    fetchChildren: async (node, _ctx) => {
      const siteId = node.siteId!;
      const listId = node.listId!;
      return [
        {
          id: makeId(["columns", siteId, listId]),
          parentId: node.id,
          label: "Columns",
          nodeType: "columns" as NodeType,
          resourceId: listId,
          siteId,
          listId,
          hasChildren: false,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
        {
          id: makeId(["contentTypes", siteId, listId]),
          parentId: node.id,
          label: "Content Types",
          nodeType: "contentTypes" as NodeType,
          resourceId: listId,
          siteId,
          listId,
          hasChildren: true,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
        {
          id: makeId(["views", siteId, listId]),
          parentId: node.id,
          label: "Views",
          nodeType: "views" as NodeType,
          resourceId: listId,
          siteId,
          listId,
          hasChildren: false,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
      ];
    },
  },

  columns: {
    cacheKey: (node) => `columns:${node.siteId}:${node.listId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listColumns(node.siteId!, node.listId!),
    fetchChildren: async () => [],
  },

  contentTypes: {
    cacheKey: (node) => `contentTypes:${node.siteId}:${node.listId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listListContentTypes(node.siteId!, node.listId!),
    fetchChildren: async (node, ctx) => {
      const cts = await ctx.graph.listListContentTypes(node.siteId!, node.listId!);
      return cts.map((ct: Record<string, unknown>) => ({
        id: makeId(["contentType", node.siteId!, ct.id as string]),
        parentId: node.id,
        label: (ct.name as string) || (ct.id as string),
        nodeType: "contentType" as NodeType,
        resourceId: ct.id as string,
        siteId: node.siteId,
        hasChildren: true,
        isLoaded: false,
        isLoading: false,
        isStale: false,
      }));
    },
  },

  contentType: {
    cacheKey: (node) => `contentType:${node.siteId}:${node.resourceId}`,
    fetchDetails: async (node, ctx) => ctx.graph.get(`/sites/${node.siteId}/contentTypes/${node.resourceId}`),
    fetchChildren: async (node, ctx) => {
      const cols = await ctx.graph.listContentTypeColumns(node.siteId!, node.resourceId);
      return cols.map((col: Record<string, unknown>) => ({
        id: makeId(["ctCol", node.siteId!, node.resourceId, col.id as string]),
        parentId: node.id,
        label: (col.displayName as string) || (col.name as string) || (col.id as string),
        nodeType: "columns" as NodeType,
        resourceId: col.id as string,
        siteId: node.siteId,
        hasChildren: false,
        isLoaded: false,
        isLoading: false,
        isStale: false,
      }));
    },
  },

  views: {
    cacheKey: (node) => `views:${node.siteId}:${node.listId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listViews(node.siteId!, node.listId!),
    fetchChildren: async () => [],
  },

  permissions: {
    cacheKey: (node) => `permissions:${node.siteId}`,
    fetchDetails: async (node, ctx) => {
      const site = await ctx.graph.getSite(node.siteId!);
      const s = site as Record<string, unknown>;
      const siteUrl = s.webUrl as string;

      const result: Record<string, unknown> = {
        owner: s.owner,
        sharingCapability: s.sharingCapability,
        externalSharingEnabled: s.sharingCapability !== "Disabled",
      };

      // Try SP REST via proxy for real role assignments (groups, members, roles)
      if (siteUrl) {
        try {
          const roleAssignments = await ctx.graph.callSpRestViaProxy<unknown[]>(
            siteUrl,
            "web/roleassignments?$expand=Member,RoleDefinitionBindings"
          );
          result.roleAssignments = roleAssignments;
          result.roleAssignmentsCount = Array.isArray(roleAssignments) ? roleAssignments.length : 0;
        } catch (e) {
          result.roleAssignmentsNote = `SP REST permissions failed: ${e}. Ensure proxy has Sites.FullControl.All.`;
        }
      }

      return result;
    },
    fetchChildren: async (node, _ctx) => {
      const siteId = node.siteId!;
      return [
        {
          id: makeId(["sharingLinks", siteId]),
          parentId: node.id,
          label: "Sharing Links",
          nodeType: "sharingLinks" as NodeType,
          resourceId: siteId,
          siteId,
          hasChildren: false,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        },
      ];
    },
  },

  sharingLinks: {
    cacheKey: (node) => `sharingLinks:${node.siteId}`,
    fetchDetails: async (node, ctx) => {
      // Try proxy for sharing links (needs Files.Read.All app permission)
      try {
        return await ctx.graph.callViaProxy<unknown[]>(
          `/sites/${node.siteId}/drive/items/root/permissions`, "v1.0"
        );
      } catch {
        return {
          note: "Sharing links require the proxy with Files.Read.All (application) permission. Set the Proxy URL in Settings and ensure Files.Read.All is granted on the proxy's app registration.",
        };
      }
    },
    fetchChildren: async () => [],
  },

  siteColumns: {
    cacheKey: (node) => `siteColumns:${node.siteId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listSiteColumns(node.siteId!),
    fetchChildren: async () => [],
  },

  siteContentTypes: {
    cacheKey: (node) => `siteContentTypes:${node.siteId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listSiteContentTypes(node.siteId!),
    fetchChildren: async (node, ctx) => {
      const cts = await ctx.graph.listSiteContentTypes(node.siteId!);
      return cts.map((ct: Record<string, unknown>) => ({
        id: makeId(["contentType", node.siteId!, ct.id as string]),
        parentId: node.id,
        label: (ct.name as string) || (ct.id as string),
        nodeType: "contentType" as NodeType,
        resourceId: ct.id as string,
        siteId: node.siteId,
        hasChildren: true,
        isLoaded: false,
        isLoading: false,
        isStale: false,
      }));
    },
  },

  recycleBin: {
    cacheKey: (node) => `recycleBin:${node.siteId}`,
    fetchDetails: async (node, ctx) => {
      // Try SP REST via proxy (avoids CORS issues)
      const siteEntry = await ctx.cache.get(`site:${node.siteId}`);
      const siteUrl = (siteEntry?.data as Record<string, unknown>)?.webUrl as string;
      if (siteUrl) {
        try {
          return await ctx.graph.callSpRestViaProxy<unknown[]>(siteUrl, "web/recyclebin?$top=200");
        } catch (e) {
          console.warn("[SP Graph Browser] Recycle bin via proxy failed:", e);
        }
      }
      return { note: "Recycle bin requires the proxy. Set Proxy URL in Settings." };
    },
    fetchChildren: async () => [],
  },

  termStore: {
    cacheKey: (node) => `termStore:${node.siteId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listTermStoreGroups(node.siteId!),
    fetchChildren: async (node, ctx) => {
      const groups = await ctx.graph.listTermStoreGroups(node.siteId!);
      return groups.map((group: Record<string, unknown>) => ({
        id: makeId(["termGroup", node.siteId!, group.id as string]),
        parentId: node.id,
        label: (group.displayName as string) || (group.id as string),
        nodeType: "termGroup" as NodeType,
        resourceId: group.id as string,
        siteId: node.siteId,
        hasChildren: true,
        isLoaded: false,
        isLoading: false,
        isStale: false,
      }));
    },
  },

  termGroup: {
    cacheKey: (node) => `termGroup:${node.siteId}:${node.resourceId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listTermSets(node.siteId!, node.resourceId),
    fetchChildren: async (node, ctx) => {
      const sets = await ctx.graph.listTermSets(node.siteId!, node.resourceId);
      return sets.map((set: Record<string, unknown>) => ({
        id: makeId(["termSet", node.siteId!, set.id as string]),
        parentId: node.id,
        label: (set.localizedNames as Array<{ name: string }>)?.[0]?.name || (set.id as string),
        nodeType: "termSet" as NodeType,
        resourceId: set.id as string,
        siteId: node.siteId,
        hasChildren: true,
        isLoaded: false,
        isLoading: false,
        isStale: false,
      }));
    },
  },

  termSet: {
    cacheKey: (node) => `termSet:${node.siteId}:${node.resourceId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listTerms(node.siteId!, node.resourceId),
    fetchChildren: async (node, ctx) => {
      const terms = await ctx.graph.listTerms(node.siteId!, node.resourceId);
      return terms.map((term: Record<string, unknown>) => ({
        id: makeId(["term", node.siteId!, term.id as string]),
        parentId: node.id,
        label: (term.labels as Array<{ name: string }>)?.[0]?.name || (term.id as string),
        nodeType: "term" as NodeType,
        resourceId: term.id as string,
        siteId: node.siteId,
        hasChildren: false,
        isLoaded: false,
        isLoading: false,
        isStale: false,
      }));
    },
  },

  term: {
    cacheKey: (node) => `term:${node.siteId}:${node.resourceId}`,
    fetchDetails: async (node, ctx) =>
      ctx.graph.get(`/sites/${node.siteId}/termStore/terms/${node.resourceId}`),
    fetchChildren: async () => [],
  },

  drives: {
    cacheKey: (node) => `drives:${node.siteId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listDrives(node.siteId!),
    fetchChildren: async (node, ctx) => {
      const drives = await ctx.graph.listDrives(node.siteId!);
      return drives.map((drive: Record<string, unknown>) => ({
        id: makeId(["driveItem", node.siteId!, drive.id as string, "root"]),
        parentId: node.id,
        label: (drive.name as string) || (drive.id as string),
        nodeType: "driveItem" as NodeType,
        resourceId: "root",
        siteId: node.siteId,
        listId: drive.id as string, // reuse listId to store driveId
        hasChildren: true,
        isLoaded: false,
        isLoading: false,
        isStale: false,
      }));
    },
  },

  driveItem: {
    cacheKey: (node) => `driveItem:${node.siteId}:${node.listId}:${node.resourceId}`,
    fetchDetails: async (node, ctx) => {
      const children = await ctx.graph.listDriveChildren(node.siteId!, node.listId!, node.resourceId === "root" ? undefined : node.resourceId);
      return children;
    },
    fetchChildren: async (node, ctx) => {
      const children = await ctx.graph.listDriveChildren(node.siteId!, node.listId!, node.resourceId === "root" ? undefined : node.resourceId);
      // Only show folders as expandable children
      return children
        .filter((item: Record<string, unknown>) => item.folder)
        .map((item: Record<string, unknown>) => ({
          id: makeId(["driveItem", node.siteId!, node.listId!, item.id as string]),
          parentId: node.id,
          label: (item.name as string) || (item.id as string),
          nodeType: "driveItem" as NodeType,
          resourceId: item.id as string,
          siteId: node.siteId,
          listId: node.listId, // driveId carried forward
          hasChildren: ((item.folder as Record<string, unknown>)?.childCount as number) > 0,
          isLoaded: false,
          isLoading: false,
          isStale: false,
        }));
    },
  },

  hubSites: {
    cacheKey: () => "hubSites",
    fetchDetails: async (_node, ctx) => {
      const sites = await ctx.graph.listSites();
      return sites.filter((s: Record<string, unknown>) => s.isHubSite);
    },
    fetchChildren: async () => [],
  },
};

export function getNodeDefinition(nodeType: NodeType): NodeDefinition {
  return definitions[nodeType];
}
