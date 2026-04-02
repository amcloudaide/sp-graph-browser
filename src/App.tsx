import { useMemo, useState, useCallback, useEffect } from "react";
import { useMsal } from "@azure/msal-react";
import type { PublicClientApplication } from "@azure/msal-browser";
import { Button, Tooltip } from "@fluentui/react-components";
import { SignOut20Regular, Person20Regular, PanelLeft20Regular, PanelLeftContract20Regular } from "@fluentui/react-icons";
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
import { BlobClient } from "./data/blobClient";
import { buildAnalyticsTree } from "./data/analyticsTreeBuilder";
import { downloadJson } from "./export/jsonExporter";
import { downloadCsv } from "./export/csvExporter";
import { downloadHtmlReport } from "./export/htmlReportBuilder";
import { DEFAULT_SETTINGS } from "./types";
import type { ViewMode, AppSettings, AppMode, BlobData } from "./types";

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
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [appMode, setAppMode] = useState<AppMode>("live");
  const [blobData, setBlobData] = useState<BlobData | null>(null);
  const [blobLoading, setBlobLoading] = useState(false);
  const [blobError, setBlobError] = useState<string | null>(null);

  const blobClient = useMemo(() => new BlobClient(cacheStore), []);

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
    nodes: liveNodes,
    selectedNodeId: liveSelectedNodeId,
    selectedNodeData: liveSelectedNodeData,
    breadcrumb: liveBreadcrumb,
    expandNode: liveExpandNode,
    selectNode: liveSelectNode,
    refreshNode,
  } = useTreeData(graphClient, spRestClient, cacheStore, settings);

  // Analytics mode — local selection state (no Graph calls)
  const [analyticsSelectedNodeId, setAnalyticsSelectedNodeId] = useState<string | null>(null);

  // Load blob data when analytics mode is active
  useEffect(() => {
    if (appMode !== "analytics" || !settings.blobSasUrl) return;
    if (blobData) return; // Already loaded

    setBlobLoading(true);
    setBlobError(null);
    blobClient.loadAll(settings.blobSasUrl, (msg) => console.log(`[Blob] ${msg}`))
      .then((data) => {
        setBlobData(data);
        setBlobLoading(false);
      })
      .catch((err) => {
        console.error("[SP Graph Browser] Blob load failed:", err);
        setBlobError(String(err));
        setBlobLoading(false);
      });
  }, [appMode, settings.blobSasUrl, blobData, blobClient]);

  // Build analytics tree from blob data
  const analyticsTree = useMemo(() => {
    if (!blobData) return null;
    return buildAnalyticsTree(blobData);
  }, [blobData]);

  const isAnalytics = appMode === "analytics";

  // Active nodes, selection, and data depend on mode
  const activeNodes = isAnalytics && analyticsTree ? analyticsTree.nodes : liveNodes;
  const selectedNodeId = isAnalytics ? analyticsSelectedNodeId : liveSelectedNodeId;
  const activeSelectedData = isAnalytics && analyticsTree && analyticsSelectedNodeId
    ? analyticsTree.nodeData.get(analyticsSelectedNodeId) ?? null
    : liveSelectedNodeData;

  // Breadcrumb: in analytics mode, walk parent chain from selected node
  const breadcrumb = useMemo(() => {
    if (!selectedNodeId) return [];
    if (!isAnalytics) return liveBreadcrumb;
    const path: typeof activeNodes = [];
    let current = activeNodes.find((n) => n.id === selectedNodeId);
    while (current) {
      path.unshift(current);
      current = current.parentId ? activeNodes.find((n) => n.id === current!.parentId) : undefined;
    }
    return path;
  }, [selectedNodeId, isAnalytics, liveBreadcrumb, activeNodes]);

  // Mode-aware select and expand
  const selectNode = useCallback((nodeId: string) => {
    if (isAnalytics) {
      setTimeout(() => setAnalyticsSelectedNodeId(nodeId), 0);
    } else {
      liveSelectNode(nodeId);
    }
  }, [isAnalytics, liveSelectNode]);

  const expandNode = useCallback((nodeId: string) => {
    if (!isAnalytics) {
      liveExpandNode(nodeId);
    }
    // Analytics tree is pre-built — no expand needed, tree already has all nodes
  }, [isAnalytics, liveExpandNode]);

  // Blob data age for the toolbar badge
  const blobDataAge = blobData
    ? `${Math.round((Date.now() - blobData.loadedAt) / 3600000)}h ago`
    : null;

  const selectedNode = activeNodes.find((n) => n.id === selectedNodeId) ?? null;

  const handleAppModeChange = useCallback((mode: AppMode) => {
    setAppMode(mode);
    if (mode === "analytics" && !settings.blobSasUrl) {
      setBlobError("Configure Blob SAS URL in Settings to use Analytics mode.");
    }
  }, [settings.blobSasUrl]);

  // Navigate from table row click — find matching child node in the tree
  const handleTableNavigate = useCallback((item: Record<string, unknown>) => {
    const itemId = (item.id as string) ?? "";
    if (!itemId) return;

    setTimeout(() => {
      const childNode = activeNodes.find((n) =>
        n.parentId === selectedNodeId && (n.resourceId === itemId || n.id.includes(itemId))
      );
      if (childNode) {
        selectNode(childNode.id);
        expandNode(childNode.id);
      } else {
        const siteNode = activeNodes.find((n) =>
          (n.nodeType === "site" || n.nodeType === "analyticsSite") && n.resourceId === itemId
        );
        if (siteNode) {
          selectNode(siteNode.id);
          expandNode(siteNode.id);
        }
      }
    }, 0);
  }, [selectedNodeId, activeNodes, selectNode, expandNode]);

  // Determine if current data is navigable (array of items with ids that have matching tree nodes)
  const isTableNavigable = useMemo(() => {
    if (!activeSelectedData || !Array.isArray(activeSelectedData) || !selectedNode) return false;
    const firstItem = activeSelectedData[0] as Record<string, unknown> | undefined;
    if (!firstItem?.id) return false;
    const navigableTypes = [
      "tenant", "lists", "subsites", "siteContentTypes", "contentTypes", "drives", "driveItem",
      "analyticsAllSites", "analyticsOwnerGroup", "analyticsRiskLevel", "analyticsRiskType",
    ];
    return navigableTypes.includes(selectedNode.nodeType);
  }, [activeSelectedData, selectedNode]);

  const handleExport = useCallback((format: "json" | "csv" | "html") => {
    if (!activeSelectedData || !selectedNode) return;
    const name = selectedNode.label;
    switch (format) {
      case "json":
        downloadJson(activeSelectedData, name);
        break;
      case "csv":
        if (Array.isArray(activeSelectedData)) {
          downloadCsv(activeSelectedData as Record<string, unknown>[], name);
        }
        break;
      case "html":
        downloadHtmlReport(activeSelectedData, name, breadcrumb.map((n) => n.label));
        break;
    }
  }, [activeSelectedData, selectedNode, breadcrumb]);

  if (!isAuthenticated) {
    return <LandingPage />;
  }

  const renderView = () => {
    if (appMode === "analytics" && blobLoading) {
      return <p style={{ color: "var(--colorNeutralForeground3)" }}>Loading analytics data...</p>;
    }
    if (appMode === "analytics" && blobError) {
      return <p style={{ color: "var(--colorPaletteRedForeground1)" }}>{blobError}</p>;
    }
    const viewData = activeSelectedData;
    if (!viewData) {
      return <p style={{ color: "var(--colorNeutralForeground3)" }}>Select a node to view its properties</p>;
    }
    switch (viewMode) {
      case "properties": return <PropertiesView data={viewData} />;
      case "json": return <JsonView data={viewData} />;
      case "table": return <TableView data={viewData} onNavigate={isTableNavigable ? handleTableNavigate : undefined} />;
    }
  };

  return (
    <div style={{ display: "flex", height: "100vh", overflow: "hidden" }}>
      {/* Sidebar */}
      <div style={{
        width: sidebarOpen ? 300 : 0,
        minWidth: sidebarOpen ? 300 : 0,
        flexShrink: 0,
        borderRight: sidebarOpen ? "1px solid var(--colorNeutralStroke1)" : "none",
        overflow: "hidden",
        transition: "width 0.2s ease, min-width 0.2s ease",
        display: "flex",
        flexDirection: "column",
      }}>
        {sidebarOpen && (
          <TreePanel nodes={activeNodes} selectedNodeId={selectedNodeId} onExpand={expandNode} onSelect={selectNode} />
        )}
      </div>

      {/* Main content */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0 }}>
        <div style={{ display: "flex", alignItems: "center", borderBottom: "1px solid var(--colorNeutralStroke2)" }}>
          <Tooltip content={sidebarOpen ? "Hide tree panel" : "Show tree panel"} relationship="label">
            <Button
              icon={sidebarOpen ? <PanelLeftContract20Regular /> : <PanelLeft20Regular />}
              appearance="subtle"
              size="small"
              onClick={() => setSidebarOpen(!sidebarOpen)}
              style={{ margin: "4px 4px 4px 8px" }}
            />
          </Tooltip>
          <div style={{ flex: 1 }}>
            <Toolbar
              breadcrumb={breadcrumb}
              viewMode={viewMode}
              appMode={appMode}
              onViewModeChange={setViewMode}
              onAppModeChange={handleAppModeChange}
              onRefresh={() => {
                if (appMode === "analytics") {
                  blobClient.clearCache();
                  setBlobData(null);
                } else if (selectedNodeId) {
                  refreshNode(selectedNodeId);
                }
              }}
              onExport={handleExport}
              onBreadcrumbClick={selectNode}
              blobDataAge={blobDataAge}
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
