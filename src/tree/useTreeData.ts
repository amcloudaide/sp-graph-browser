import { useState, useCallback, useRef } from "react";
import type { TreeNodeData, AppSettings } from "../types";
import { getNodeDefinition, FetchContext } from "../data/nodeResolver";
import type { GraphClient } from "../data/graphClient";
import type { SpRestClient } from "../data/spRestClient";
import type { CacheStore } from "../data/cacheStore";

const ROOT_NODE: TreeNodeData = {
  id: "tenant",
  parentId: null,
  label: "Tenant",
  nodeType: "tenant",
  resourceId: "root",
  hasChildren: true,
  isLoaded: false,
  isLoading: false,
  isStale: false,
};

export function useTreeData(
  graph: GraphClient | null,
  spRest: SpRestClient | null,
  cache: CacheStore,
  settings: AppSettings
) {
  const [nodes, setNodes] = useState<TreeNodeData[]>([ROOT_NODE]);
  const [selectedNodeId, setSelectedNodeId] = useState<string | null>(null);
  const [selectedNodeData, setSelectedNodeData] = useState<unknown>(null);

  // Keep a ref to nodes so callbacks always see latest
  const nodesRef = useRef(nodes);
  nodesRef.current = nodes;

  const updateNode = useCallback((id: string, updates: Partial<TreeNodeData>) => {
    setNodes((prev) =>
      prev.map((n) => (n.id === id ? { ...n, ...updates } : n))
    );
  }, []);

  const expandNode = useCallback(
    async (nodeId: string) => {
      if (!graph) return;

      const node = nodesRef.current.find((n) => n.id === nodeId);
      if (!node || !node.hasChildren || node.isLoading) return;

      const definition = getNodeDefinition(node.nodeType);
      if (!definition) return;
      const cacheKey = definition.cacheKey(node) + ":children";

      // Check cache first
      const isCached = await cache.isFresh(cacheKey, settings.cacheTtlMinutes);
      if (isCached && node.isLoaded) return;

      updateNode(nodeId, { isLoading: true });

      try {
        const ctx: FetchContext = { graph, spRest, cache, ttlMinutes: settings.cacheTtlMinutes, enableFilesAccess: settings.enableFilesAccess ?? false };

        let children: TreeNodeData[];
        if (isCached) {
          const cached = await cache.get(cacheKey);
          children = cached!.data as TreeNodeData[];
        } else {
          children = await definition.fetchChildren(node, ctx);
          await cache.set(cacheKey, children, node.nodeType);
        }

        setNodes((prev) => {
          // Remove old children for this parent, add new ones
          const withoutOldChildren = prev.filter((n) => n.parentId !== nodeId);
          return [...withoutOldChildren, ...children].map((n) =>
            n.id === nodeId ? { ...n, isLoaded: true, isLoading: false, isStale: false } : n
          );
        });
      } catch (error) {
        console.error(`Failed to expand ${nodeId}:`, error);
        updateNode(nodeId, { isLoading: false });
      }
    },
    [graph, spRest, cache, settings.cacheTtlMinutes, updateNode]
  );

  const selectNode = useCallback(
    async (nodeId: string) => {
      if (!graph) return;

      setSelectedNodeId(nodeId);
      const node = nodesRef.current.find((n) => n.id === nodeId);
      if (!node) return;

      const definition = getNodeDefinition(node.nodeType);
      if (!definition) return;
      const cacheKey = definition.cacheKey(node);

      // Check cache
      const isCached = await cache.isFresh(cacheKey, settings.cacheTtlMinutes);
      if (isCached) {
        const cached = await cache.get(cacheKey);
        setSelectedNodeData(cached!.data);
        return;
      }

      try {
        const ctx: FetchContext = { graph, spRest, cache, ttlMinutes: settings.cacheTtlMinutes, enableFilesAccess: settings.enableFilesAccess ?? false };
        const data = await definition.fetchDetails(node, ctx);
        await cache.set(cacheKey, data, node.nodeType);
        setSelectedNodeData(data);
      } catch (error) {
        console.error(`Failed to fetch details for ${nodeId}:`, error);
        setSelectedNodeData({ error: String(error) });
      }
    },
    [graph, spRest, cache, settings.cacheTtlMinutes]
  );

  const refreshNode = useCallback(
    async (nodeId: string) => {
      const node = nodesRef.current.find((n) => n.id === nodeId);
      if (!node) return;

      const definition = getNodeDefinition(node.nodeType);
      if (!definition) return;
      await cache.invalidate(definition.cacheKey(node));
      await cache.invalidate(definition.cacheKey(node) + ":children");

      updateNode(nodeId, { isLoaded: false, isStale: true });

      if (node.hasChildren) {
        await expandNode(nodeId);
      }
      if (selectedNodeId === nodeId) {
        await selectNode(nodeId);
      }
    },
    [cache, expandNode, selectNode, selectedNodeId, updateNode]
  );

  const breadcrumb = (() => {
    if (!selectedNodeId) return [];
    const path: TreeNodeData[] = [];
    let current = nodes.find((n) => n.id === selectedNodeId);
    while (current) {
      path.unshift(current);
      current = current.parentId ? nodes.find((n) => n.id === current!.parentId) : undefined;
    }
    return path;
  })();

  return {
    nodes,
    selectedNodeId,
    selectedNodeData,
    breadcrumb,
    expandNode,
    selectNode,
    refreshNode,
  };
}
