import { TreeItemLayout } from "@fluentui/react-tree";
import { Spinner } from "@fluentui/react-components";
import {
  Globe20Regular,
  Folder20Regular,
  List20Regular,
  TableSimple20Regular,
  Shield20Regular,
  Link20Regular,
  Tag20Regular,
  Delete20Regular,
  Building20Regular,
} from "@fluentui/react-icons";
import type { NodeType } from "../types";
import type { ReactElement } from "react";

const iconMap: Record<NodeType, ReactElement> = {
  tenant: <Globe20Regular />,
  site: <Building20Regular />,
  subsites: <Folder20Regular />,
  lists: <Folder20Regular />,
  list: <List20Regular />,
  columns: <TableSimple20Regular />,
  contentTypes: <Tag20Regular />,
  contentType: <Tag20Regular />,
  views: <List20Regular />,
  permissions: <Shield20Regular />,
  sharingLinks: <Link20Regular />,
  siteColumns: <TableSimple20Regular />,
  siteContentTypes: <Tag20Regular />,
  recycleBin: <Delete20Regular />,
  termStore: <Tag20Regular />,
  termGroup: <Folder20Regular />,
  termSet: <Folder20Regular />,
  term: <Tag20Regular />,
  hubSites: <Building20Regular />,
};

interface TreeNodeProps {
  label: string;
  nodeType: NodeType;
  isLoading: boolean;
  isStale: boolean;
}

export function TreeNodeContent({ label, nodeType, isLoading, isStale }: TreeNodeProps) {
  return (
    <TreeItemLayout
      iconBefore={isLoading ? <Spinner size="tiny" /> : iconMap[nodeType]}
    >
      <span style={{ opacity: isStale ? 0.6 : 1 }}>{label}</span>
    </TreeItemLayout>
  );
}
