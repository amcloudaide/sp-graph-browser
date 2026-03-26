import { useMemo, useState, useCallback } from "react";
import { useMsal } from "@azure/msal-react";
import type { PublicClientApplication } from "@azure/msal-browser";
import { useAuth } from "./auth/AuthProvider";
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
  const { isAuthenticated, account } = useAuth();
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

  const {
    nodes,
    selectedNodeId,
    selectedNodeData,
    breadcrumb,
    expandNode,
    selectNode,
    refreshNode,
  } = useTreeData(graphClient, null, cacheStore, settings);

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

  const renderView = () => {
    if (!selectedNodeData) {
      return <p style={{ color: "var(--colorNeutralForeground3)" }}>Select a node to view its properties</p>;
    }
    switch (viewMode) {
      case "properties": return <PropertiesView data={selectedNodeData} />;
      case "json": return <JsonView data={selectedNodeData} />;
      case "table": return <TableView data={selectedNodeData} />;
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
          <div style={{ padding: "4px 8px" }}>
            <SettingsDialog settings={settings} onSave={handleSettingsChange} />
          </div>
        </div>
        <div style={{ flex: 1, padding: 16, overflow: "auto" }}>
          {renderView()}
        </div>
      </div>
    </div>
  );
}
