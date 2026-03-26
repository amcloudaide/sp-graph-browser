import {
  FlatTree,
  FlatTreeItem,
  useHeadlessFlatTree_unstable,
  TreeItemValue,
} from "@fluentui/react-tree";
import { TreeNodeContent } from "./TreeNode";
import type { TreeNodeData } from "../types";

interface TreePanelProps {
  nodes: TreeNodeData[];
  selectedNodeId: string | null;
  onExpand: (nodeId: string) => void;
  onSelect: (nodeId: string) => void;
}

export function TreePanel({ nodes, selectedNodeId, onExpand, onSelect }: TreePanelProps) {
  const flatTreeItems = nodes.map((node) => ({
    value: node.id as TreeItemValue,
    parentValue: (node.parentId as TreeItemValue) ?? undefined,
    itemType: node.hasChildren ? ("branch" as const) : ("leaf" as const),
  }));

  const flatTree = useHeadlessFlatTree_unstable(flatTreeItems, {
    onOpenChange: (_event, data) => {
      const nodeId = data.value as string;
      if (data.open) {
        onExpand(nodeId);
      }
    },
  });

  const nodeMap = new Map(nodes.map((n) => [n.id, n]));

  return (
    <div style={{ height: "100%", overflow: "auto" }}>
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
              onClick={() => onSelect(node.id)}
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
  );
}
