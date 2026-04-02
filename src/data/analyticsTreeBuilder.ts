import type { TreeNodeData, NodeType, BlobData, DscComponent } from "../types";

function makeId(parts: string[]): string {
  return "a:" + parts.join(":");
}

function makeNode(
  id: string,
  parentId: string | null,
  label: string,
  nodeType: NodeType,
  hasChildren: boolean,
  resourceId = "",
): TreeNodeData {
  return { id, parentId, label, nodeType, resourceId, hasChildren, isLoaded: true, isLoading: false, isStale: false };
}

/** Build the full analytics tree from blob data. All nodes are pre-built (no lazy loading). */
export function buildAnalyticsTree(data: BlobData): {
  nodes: TreeNodeData[];
  /** Map from node ID to its detail data (shown in Properties/JSON/Table views) */
  nodeData: Map<string, unknown>;
} {
  const nodes: TreeNodeData[] = [];
  const nodeData = new Map<string, unknown>();

  // Index DSC components by ResourceName and URL
  const dscByType = new Map<string, DscComponent[]>();
  for (const comp of data.dscComponents) {
    const list = dscByType.get(comp.ResourceName) ?? [];
    list.push(comp);
    dscByType.set(comp.ResourceName, list);
  }

  const dscSites = dscByType.get("SPOSite") ?? [];
  const dscGroups = dscByType.get("SPOSiteGroup") ?? [];

  // Index analytics data by site URL
  const permsByUrl = new Map<string, Record<string, unknown>[]>();
  for (const p of data.permissionsReport) {
    const url = p.SiteUrl as string;
    if (url) {
      const list = permsByUrl.get(url) ?? [];
      list.push(p);
      permsByUrl.set(url, list);
    }
  }

  const oversharingByUrl = new Map<string, Record<string, unknown>[]>();
  for (const o of data.oversharingAnalysis) {
    const url = o.SiteUrl as string;
    if (url) {
      const list = oversharingByUrl.get(url) ?? [];
      list.push(o);
      oversharingByUrl.set(url, list);
    }
  }

  const dscGroupsByUrl = new Map<string, DscComponent[]>();
  for (const g of dscGroups) {
    const url = g.Url as string;
    if (url) {
      const list = dscGroupsByUrl.get(url) ?? [];
      list.push(g);
      dscGroupsByUrl.set(url, list);
    }
  }

  // ── Tenant Configuration ──
  const tenantConfigId = makeId(["tenantConfig"]);
  nodes.push(makeNode(tenantConfigId, null, "Tenant Configuration", "analyticsTenantConfig", true));

  const tenantTypes = [
    "SPOTenantSettings", "SPOAccessControlSettings", "SPOSharingSettings",
    "SPOBrowserIdleSignout", "SPOHomeSite", "SPOOrgAssetsLibrary",
    "SPOTenantCdnEnabled", "SPOTenantCdnPolicy", "SPOStorageEntity",
    "SPOUserProfileProperty",
  ];
  for (const typeName of tenantTypes) {
    const items = dscByType.get(typeName) ?? [];
    if (items.length === 0) continue;
    const itemId = makeId(["tenantConfig", typeName]);
    nodes.push(makeNode(itemId, tenantConfigId, `${typeName} (${items.length})`, "analyticsTenantConfigItem", false));
    nodeData.set(itemId, items.length === 1 ? items[0] : items);
  }
  nodeData.set(tenantConfigId, {
    totalDscComponents: data.dscComponents.length,
    componentTypes: Object.fromEntries([...dscByType.entries()].map(([k, v]) => [k, v.length])),
    loadedAt: new Date(data.loadedAt).toISOString(),
  });

  // ── All Sites ──
  // Prefer sites-inventory (has owner email), fall back to DSC SPOSite
  const allSites = data.sitesInventory.length > 0 ? data.sitesInventory : dscSites;
  const allSitesId = makeId(["allSites"]);
  nodes.push(makeNode(allSitesId, null, `All Sites (${allSites.length})`, "analyticsAllSites", true));
  nodeData.set(allSitesId, allSites);

  for (const site of allSites) {
    const siteUrl = (site.Url ?? site.url ?? "") as string;
    const siteTitle = (site.Title ?? site.title ?? site.ResourceInstanceName ?? siteUrl) as string;
    const siteId = makeId(["site", siteUrl]);

    nodes.push(makeNode(siteId, allSitesId, siteTitle, "analyticsSite", true, siteUrl));
    nodeData.set(siteId, site);

    // Site Config (from DSC)
    const dscSite = dscSites.find((s) => s.Url === siteUrl);
    if (dscSite) {
      const configId = makeId(["site", siteUrl, "config"]);
      nodes.push(makeNode(configId, siteId, "Site Configuration (DSC)", "analyticsSiteConfig", false));
      nodeData.set(configId, dscSite);
    }

    // Permissions (from DSC SPOSiteGroup)
    const siteGroups = dscGroupsByUrl.get(siteUrl) ?? [];
    if (siteGroups.length > 0) {
      const permsId = makeId(["site", siteUrl, "permissions"]);
      nodes.push(makeNode(permsId, siteId, `Permissions (${siteGroups.length} groups)`, "analyticsSitePermissions", false));
      nodeData.set(permsId, siteGroups.map((g) => ({
        group: g.Identity,
        roles: g.PermissionLevels,
        ...g,
      })));
    }

    // Also add analytics permissions if available (has member count)
    const analyticsPerms = permsByUrl.get(siteUrl);
    if (analyticsPerms && analyticsPerms.length > 0 && siteGroups.length === 0) {
      const permsId = makeId(["site", siteUrl, "permissions"]);
      nodes.push(makeNode(permsId, siteId, `Permissions (${analyticsPerms.length} groups)`, "analyticsSitePermissions", false));
      nodeData.set(permsId, analyticsPerms);
    }

    // Audit Settings (from DSC)
    const auditSettings = (dscByType.get("SPOSiteAuditSettings") ?? []).find((a) => a.Url === siteUrl);
    if (auditSettings) {
      const auditId = makeId(["site", siteUrl, "audit"]);
      nodes.push(makeNode(auditId, siteId, "Audit Settings", "analyticsSiteAudit", false));
      nodeData.set(auditId, auditSettings);
    }

    // Oversharing Flags (from analytics)
    const oversharing = oversharingByUrl.get(siteUrl);
    if (oversharing && oversharing.length > 0) {
      const overId = makeId(["site", siteUrl, "oversharing"]);
      nodes.push(makeNode(overId, siteId, `Oversharing (${oversharing.length} flags)`, "analyticsSiteOversharing", false));
      nodeData.set(overId, oversharing);
    }
  }

  // ── By Owner ──
  const byOwnerId = makeId(["byOwner"]);
  const ownerMap = new Map<string, Record<string, unknown>[]>();
  for (const site of allSites) {
    const owner = ((site.Owner ?? site.owner ?? "") as string) || "(no owner)";
    const list = ownerMap.get(owner) ?? [];
    list.push(site);
    ownerMap.set(owner, list);
  }
  const sortedOwners = [...ownerMap.entries()].sort((a, b) => b[1].length - a[1].length);
  nodes.push(makeNode(byOwnerId, null, `By Owner (${ownerMap.size})`, "analyticsByOwner", true));
  nodeData.set(byOwnerId, sortedOwners.map(([owner, sites]) => ({ owner, siteCount: sites.length })));

  for (const [owner, sites] of sortedOwners) {
    const ownerGroupId = makeId(["owner", owner]);
    nodes.push(makeNode(ownerGroupId, byOwnerId, `${owner} (${sites.length})`, "analyticsOwnerGroup", true));
    nodeData.set(ownerGroupId, sites);

    for (const site of sites) {
      const siteUrl = (site.Url ?? site.url ?? "") as string;
      const siteTitle = (site.Title ?? site.title ?? siteUrl) as string;
      const refId = makeId(["owner", owner, siteUrl]);
      nodes.push(makeNode(refId, ownerGroupId, siteTitle, "analyticsSite", false, siteUrl));
      nodeData.set(refId, site);
    }
  }

  // ── By Risk Level ──
  const byRiskId = makeId(["byRisk"]);
  nodes.push(makeNode(byRiskId, null, "By Risk Level", "analyticsByRisk", true));
  nodeData.set(byRiskId, data.oversharingSummary);

  const riskLevels = ["Critical", "High", "Medium", "Low"];
  for (const level of riskLevels) {
    const risksAtLevel = data.oversharingAnalysis.filter((o) => o.RiskLevel === level);
    const levelId = makeId(["risk", level]);
    nodes.push(makeNode(levelId, byRiskId, `${level} (${risksAtLevel.length})`, "analyticsRiskLevel", risksAtLevel.length > 0));
    nodeData.set(levelId, risksAtLevel);

    // Group by risk type within level
    const typeMap = new Map<string, Record<string, unknown>[]>();
    for (const r of risksAtLevel) {
      const t = (r.RiskType as string) ?? "Unknown";
      const list = typeMap.get(t) ?? [];
      list.push(r);
      typeMap.set(t, list);
    }
    for (const [riskType, items] of typeMap) {
      const typeId = makeId(["risk", level, riskType]);
      nodes.push(makeNode(typeId, levelId, `${riskType} (${items.length})`, "analyticsRiskType", false));
      nodeData.set(typeId, items);
    }
  }

  // ── External Users ──
  const extUsersId = makeId(["externalUsers"]);
  nodes.push(makeNode(extUsersId, null, `External Users (${data.externalUsersReport.length})`, "analyticsExternalUsers", false));
  nodeData.set(extUsersId, data.externalUsersReport.length > 0 ? data.externalUsersReport : { note: "No external users data available." });

  console.log(`[SP Graph Browser] Analytics tree: ${nodes.length} nodes`);
  return { nodes, nodeData };
}
