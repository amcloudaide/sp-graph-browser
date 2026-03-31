import { useMemo, useState, useCallback, useEffect } from "react";
import { useMsal } from "@azure/msal-react";
import type { PublicClientApplication } from "@azure/msal-browser";
import { Button, Tooltip } from "@fluentui/react-components";
import { SignOut20Regular, Person20Regular } from "@fluentui/react-icons";
import { useAuth } from "./auth/AuthProvider";
import { SpRestClient } from "./data/spRestClient";
import { LandingPage } from "./components/LandingPage";
import { TreePanel } from "./tree/TreePanel";
import { Toolbar } from "./components/Toolbar";
import { SettingsDialog } from "./components/SettingsDialog";
import { useTreeData } from "./tree/useTreeData";
import { GraphClient } from "./data/graphClient";
import { CacheStore } from "./data/cacheStore";
import { PropertiesView } from "./views/PropertiesView";
import { JsonView } from "./views/JsonView";
import { TableView } from "./views/TableView";
import { downloadJson } from "./export/jsonExporter";
import { downloadCsv } from "./export/csvExporter";
import { downloadHtmlReport } from "./export/htmlReportBuilder";
import { DEFAULT_SETTINGS } from "./types";
import type { ViewMode, AppSettings } from "./types";

function loadSettings(): AppSettings {
  try {
    const stored = localStorage.getItem("sp-graph-browser-settings");
    if (stored) return { ...DEFAULT_SETTINGS, ...JSON.parse(stored) };
  } catch { /* ignore */ }
  return DEFAULT_SETTINGS;
}

const cacheStore = new CacheStore();

export default function App() {
  const { isAuthenticated, account, tenantName, logout } = useAuth();
  const { instance } = useMsal();
  const [settings, setSettings] = useState<AppSettings>(loadSettings);
  const [viewMode, setViewMode] = useState<ViewMode>(settings.defaultViewMode);

  const handleSettingsChange = useCallback((newSettings: AppSettings) => {
    setSettings(newSettings);
  }, []);

  const graphClient = useMemo(() => {
    if (!isAuthenticated || !account) return null;
    return new GraphClient(instance as PublicClientApplication, account);
  }, [isAuthenticated, account, instance]);

  useEffect(() => {
    if (graphClient) {
      graphClient.setProxyUrl(settings.proxyUrl);
    }
  }, [graphClient, settings.proxyUrl]);

  const spRestClient = useMemo(() => {
    if (!isAuthenticated || !account || !tenantName) return null;
    return new SpRestClient(instance as PublicClientApplication, account, tenantName);
  }, [isAuthenticated, account, tenantName, instance]);

  const {
    nodes,
    selectedNodeId,
    selectedNodeData,
    breadcrumb,
    expandNode,
    selectNode,
    refreshNode,
  } = useTreeData(graphClient, spRestClient, cacheStore, settings);

  if (!isAuthenticated) {
    return <LandingPage />;
  }

  const selectedNode = nodes.find((n) => n.id === selectedNodeId);

  const handleExport = (format: "json" | "csv" | "html") => {
    if (!selectedNodeData || !selectedNode) return;
    const name = selectedNode.label;
    switch (format) {
      case "json":
        downloadJson(selectedNodeData, name);
        break;
      case "csv":
        if (Array.isArray(selectedNodeData)) {
          downloadCsv(selectedNodeData as Record<string, unknown>[], name);
        }
        break;
      case "html":
        downloadHtmlReport(selectedNodeData, name, breadcrumb.map((n) => n.label));
        break;
    }
  };

  // Navigate from table row click — find matching child node in the tree
  const handleTableNavigate = useCallback((item: Record<string, unknown>) => {
    if (!selectedNode) return;
    const itemId = (item.id as string) ?? "";
    if (!itemId) return;

    // Look for a child node whose resourceId or siteId matches the clicked item's id
    const childNode = nodes.find((n) =>
      n.parentId === selectedNode.id && (n.resourceId === itemId || n.id.includes(itemId))
    );
    if (childNode) {
      selectNode(childNode.id);
      expandNode(childNode.id);
    } else {
      // Item might be a site at tenant level
      const siteNode = nodes.find((n) => n.nodeType === "site" && n.resourceId === itemId);
      if (siteNode) {
        selectNode(siteNode.id);
        expandNode(siteNode.id);
      }
    }
  }, [selectedNode, nodes, selectNode, expandNode]);

  // Determine if current data is navigable (array of items with ids that have matching tree nodes)
  const isTableNavigable = useMemo(() => {
    if (!selectedNodeData || !Array.isArray(selectedNodeData) || !selectedNode) return false;
    // Check if any item has an id that matches a child node
    const firstItem = selectedNodeData[0] as Record<string, unknown> | undefined;
    if (!firstItem?.id) return false;
    // Navigable node types: tenant (sites), lists folder, subsites, content types, drives
    const navigableTypes = ["tenant", "lists", "subsites", "siteContentTypes", "contentTypes", "drives", "driveItem"];
    return navigableTypes.includes(selectedNode.nodeType);
  }, [selectedNodeData, selectedNode]);

  const renderView = () => {
    if (!selectedNodeData) {
      return <p style={{ color: "var(--colorNeutralForeground3)" }}>Select a node to view its properties</p>;
    }
    switch (viewMode) {
      case "properties": return <PropertiesView data={selectedNodeData} />;
      case "json": return <JsonView data={selectedNodeData} />;
      case "table": return <TableView data={selectedNodeData} onNavigate={isTableNavigable ? handleTableNavigate : undefined} />;
    }
  };

  return (
    <div style={{ display: "flex", height: "100vh" }}>
      <div style={{ width: 300, borderRight: "1px solid var(--colorNeutralStroke1)", overflow: "hidden" }}>
        <TreePanel nodes={nodes} selectedNodeId={selectedNodeId} onExpand={expandNode} onSelect={selectNode} />
      </div>
      <div style={{ flex: 1, display: "flex", flexDirection: "column" }}>
        <div style={{ display: "flex", alignItems: "center" }}>
          <div style={{ flex: 1 }}>
            <Toolbar
              breadcrumb={breadcrumb}
              viewMode={viewMode}
              onViewModeChange={setViewMode}
              onRefresh={() => selectedNodeId && refreshNode(selectedNodeId)}
              onExport={handleExport}
              onBreadcrumbClick={selectNode}
            />
          </div>
          <div style={{ padding: "4px 8px", display: "flex", alignItems: "center", gap: 4 }}>
            <SettingsDialog settings={settings} onSave={handleSettingsChange} />
            <Tooltip content={account?.username ?? ""} relationship="label">
              <Button icon={<Person20Regular />} appearance="subtle" size="small" />
            </Tooltip>
            <Tooltip content="Sign out and clear session" relationship="label">
              <Button
                icon={<SignOut20Regular />}
                appearance="subtle"
                size="small"
                onClick={() => {
                  // Clear all cached auth state
                  localStorage.clear();
                  cacheStore.clear();
                  logout();
                }}
              />
            </Tooltip>
          </div>
        </div>
        <div style={{ flex: 1, padding: 16, overflow: "auto" }}>
          {renderView()}
        </div>
      </div>
    </div>
  );
}
