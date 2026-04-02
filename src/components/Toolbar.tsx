import {
  Toolbar as FluentToolbar,
  ToolbarButton,
  TabList,
  Tab,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbButton,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
} from "@fluentui/react-components";
import {
  ArrowClockwise20Regular,
  ArrowDownload20Regular,
  Database20Regular,
  Globe20Regular,
} from "@fluentui/react-icons";
import type { TreeNodeData, ViewMode, AppMode } from "../types";

interface ToolbarProps {
  breadcrumb: TreeNodeData[];
  viewMode: ViewMode;
  appMode: AppMode;
  onViewModeChange: (mode: ViewMode) => void;
  onAppModeChange: (mode: AppMode) => void;
  onRefresh: () => void;
  onExport: (format: "json" | "csv" | "html") => void;
  onBreadcrumbClick: (nodeId: string) => void;
  blobDataAge?: string | null;
}

export function Toolbar({
  breadcrumb,
  viewMode,
  appMode,
  onViewModeChange,
  onAppModeChange,
  onRefresh,
  onExport,
  onBreadcrumbClick,
  blobDataAge,
}: ToolbarProps) {
  return (
    <div style={{ borderBottom: "1px solid var(--colorNeutralStroke1)", padding: "4px 8px" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 4, marginRight: 8 }}>
          <ToolbarButton
            icon={<Globe20Regular />}
            appearance={appMode === "live" ? "primary" : "subtle"}
            onClick={() => onAppModeChange("live")}
          >
            Live
          </ToolbarButton>
          <ToolbarButton
            icon={<Database20Regular />}
            appearance={appMode === "analytics" ? "primary" : "subtle"}
            onClick={() => onAppModeChange("analytics")}
          >
            Analytics
            {blobDataAge && <span style={{ fontSize: 10, marginLeft: 4, opacity: 0.7 }}>{blobDataAge}</span>}
          </ToolbarButton>
        </div>

        <Breadcrumb size="small">
          {breadcrumb.map((node, i) => (
            <BreadcrumbItem key={node.id}>
              <BreadcrumbButton
                onClick={() => onBreadcrumbClick(node.id)}
                current={i === breadcrumb.length - 1}
              >
                {node.label}
              </BreadcrumbButton>
            </BreadcrumbItem>
          ))}
        </Breadcrumb>

        <FluentToolbar>
          <ToolbarButton
            icon={<ArrowClockwise20Regular />}
            appearance="subtle"
            onClick={onRefresh}
          >
            Refresh
          </ToolbarButton>

          <Menu>
            <MenuTrigger>
              <ToolbarButton icon={<ArrowDownload20Regular />} appearance="subtle">
                Export
              </ToolbarButton>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem onClick={() => onExport("json")}>JSON</MenuItem>
                <MenuItem onClick={() => onExport("csv")}>CSV</MenuItem>
                <MenuItem onClick={() => onExport("html")}>HTML Report</MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        </FluentToolbar>
      </div>

      <TabList
        size="small"
        selectedValue={viewMode}
        onTabSelect={(_, data) => onViewModeChange(data.value as ViewMode)}
      >
        <Tab value="properties">Properties</Tab>
        <Tab value="json">Raw JSON</Tab>
        <Tab value="table">Table</Tab>
      </TabList>
    </div>
  );
}
