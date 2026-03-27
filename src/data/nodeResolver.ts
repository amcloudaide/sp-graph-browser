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
      const sites = await ctx.graph.listSites();
      return {
        displayName: (org as Record<string, unknown>).displayName ?? "Unknown",
        tenantId: (org as Record<string, unknown>).id,
        verifiedDomains: (org as Record<string, unknown>).verifiedDomains,
        rootSiteUrl: (rootSite as Record<string, unknown>).webUrl,
        rootSiteId: (rootSite as Record<string, unknown>).id,
        totalSites: sites.length,
        createdDateTime: (org as Record<string, unknown>).createdDateTime,
        country: (org as Record<string, unknown>).country,
        preferredLanguage: (org as Record<string, unknown>).preferredLanguage,
      };
    },
    fetchChildren: async (_node, ctx) => {
      const sites = await ctx.graph.listSites();
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
          hasChildren: false,
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
          hasChildren: false,
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
          hasChildren: false,
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
    fetchChildren: async () => [],
  },

  contentType: {
    cacheKey: (node) => `contentType:${node.resourceId}`,
    fetchDetails: async (node, ctx) => ctx.graph.get(`/sites/${node.siteId}/contentTypes/${node.resourceId}`),
    fetchChildren: async () => [],
  },

  views: {
    cacheKey: (node) => `views:${node.siteId}:${node.listId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listViews(node.siteId!, node.listId!),
    fetchChildren: async () => [],
  },

  permissions: {
    cacheKey: (node) => `permissions:${node.siteId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listSitePermissions(node.siteId!),
    fetchChildren: async () => [],
  },

  sharingLinks: {
    cacheKey: (node) => `sharingLinks:${node.siteId}`,
    fetchDetails: async (node, ctx) => ctx.graph.listSharingLinks(node.siteId!),
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
    fetchChildren: async () => [],
  },

  recycleBin: {
    cacheKey: (node) => `recycleBin:${node.siteId}`,
    fetchDetails: async (node, ctx) => {
      if (!ctx.spRest) return [];
      // Need site URL — extract from site details in cache
      const siteEntry = await ctx.cache.get(`site:${node.siteId}`);
      const siteUrl = (siteEntry?.data as Record<string, unknown>)?.webUrl as string;
      if (!siteUrl) return [];
      return ctx.spRest.listRecycleBin(siteUrl);
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
