import { useState, useMemo } from "react";
import {
  FlatTree,
  FlatTreeItem,
  useHeadlessFlatTree_unstable,
  TreeItemValue,
} from "@fluentui/react-tree";
import { Input } from "@fluentui/react-components";
import { Search20Regular, Dismiss20Regular } from "@fluentui/react-icons";
import { TreeNodeContent } from "./TreeNode";
import type { TreeNodeData } from "../types";

interface TreePanelProps {
  nodes: TreeNodeData[];
  selectedNodeId: string | null;
  onExpand: (nodeId: string) => void;
  onSelect: (nodeId: string) => void;
}

export function TreePanel({ nodes, selectedNodeId, onExpand, onSelect }: TreePanelProps) {
  const [searchText, setSearchText] = useState("");

  // When searching, filter site nodes and keep structural nodes (tenant, folders)
  const filteredNodes = useMemo(() => {
    if (!searchText) return nodes;
    const lower = searchText.toLowerCase();

    // Find all site node IDs that match the search
    const matchingSiteIds = new Set<string>();
    for (const node of nodes) {
      if (node.nodeType === "site" && node.label.toLowerCase().includes(lower)) {
        matchingSiteIds.add(node.id);
      }
    }

    // Keep: tenant root, matching sites, and all children of matching sites
    const keep = new Set<string>();
    // Always keep tenant
    keep.add("tenant");

    for (const node of nodes) {
      // Keep matching sites
      if (matchingSiteIds.has(node.id)) {
        keep.add(node.id);
      }
      // Keep children of matching sites (their parent is a matching site)
      if (node.parentId && matchingSiteIds.has(node.parentId)) {
        keep.add(node.id);
      }
      // Keep deeper descendants — walk up to see if any ancestor is a matching site
      let current = node;
      while (current.parentId) {
        if (matchingSiteIds.has(current.parentId)) {
          keep.add(node.id);
          break;
        }
        const parent = nodes.find((n) => n.id === current.parentId);
        if (!parent) break;
        current = parent;
      }
    }

    return nodes.filter((n) => keep.has(n.id));
  }, [nodes, searchText]);

  const siteCount = nodes.filter((n) => n.nodeType === "site" && n.parentId === "tenant").length;
  const filteredSiteCount = filteredNodes.filter((n) => n.nodeType === "site" && n.parentId === "tenant").length;

  const flatTreeItems = filteredNodes.map((node) => ({
    value: node.id as TreeItemValue,
    parentValue: (node.parentId as TreeItemValue) ?? undefined,
    itemType: node.hasChildren ? ("branch" as const) : ("leaf" as const),
  }));

  const flatTree = useHeadlessFlatTree_unstable(flatTreeItems, {
    onOpenChange: (_event, data) => {
      const nodeId = data.value as string;
      if (data.open) {
        // Defer expand to next microtask to avoid React error #300
        // (Fluent UI's tree is still processing the event when expand triggers state updates)
        setTimeout(() => onExpand(nodeId), 0);
      }
    },
  });

  const nodeMap = new Map(filteredNodes.map((n) => [n.id, n]));

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      <div style={{ padding: "8px 8px 4px", borderBottom: "1px solid var(--colorNeutralStroke2)" }}>
        <Input
          size="small"
          placeholder="Search sites..."
          value={searchText}
          onChange={(_, d) => setSearchText(d.value)}
          contentBefore={<Search20Regular />}
          contentAfter={
            searchText ? (
              <Dismiss20Regular
                style={{ cursor: "pointer" }}
                onClick={() => setSearchText("")}
              />
            ) : undefined
          }
          style={{ width: "100%" }}
        />
        {searchText && (
          <div style={{ fontSize: 11, color: "var(--colorNeutralForeground3)", padding: "2px 4px" }}>
            {filteredSiteCount} of {siteCount} sites
          </div>
        )}
      </div>
      <div style={{ flex: 1, overflow: "auto" }}>
        <FlatTree {...flatTree.getTreeProps()} aria-label="SharePoint structure">
          {Array.from(flatTree.items(), (item) => {
            const {
              value,
              "aria-level": ariaLevel,
              "aria-setsize": ariaSetsize,
              "aria-posinset": ariaPosinset,
              itemType,
              parentValue,
            } = item.getTreeItemProps();
            const node = nodeMap.get(item.value as string);
            if (!node) return null;
            return (
              <FlatTreeItem
                key={node.id}
                value={value}
                aria-level={ariaLevel as number}
                aria-setsize={ariaSetsize as number}
                aria-posinset={ariaPosinset as number}
                itemType={itemType}
                parentValue={parentValue}
                onClick={() => setTimeout(() => onSelect(node.id), 0)}
                style={{
                  backgroundColor: selectedNodeId === node.id ? "var(--colorNeutralBackground1Selected)" : undefined,
                }}
              >
                <TreeNodeContent
                  label={node.label}
                  nodeType={node.nodeType}
                  isLoading={node.isLoading}
                  isStale={node.isStale}
                />
              </FlatTreeItem>
            );
          })}
        </FlatTree>
      </div>
    </div>
  );
}
