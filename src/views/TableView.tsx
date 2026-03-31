import { useState, useMemo, useCallback, useRef } from "react";
import { Input } from "@fluentui/react-components";
import { ArrowUp20Regular, ArrowDown20Regular, ChevronRight20Regular } from "@fluentui/react-icons";

interface TableViewProps {
  data: unknown;
  /** Called when user clicks a row to navigate deeper (e.g., click a list to open it) */
  onNavigate?: (item: Record<string, unknown>) => void;
}

/** Format a cell value — handles nested objects, arrays, nulls */
function formatCell(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "number") return String(value);
  if (typeof value === "string") return value;
  if (Array.isArray(value)) {
    if (value.length === 0) return "[]";
    return value.map((item) => {
      if (typeof item === "string") return item;
      if (typeof item === "object" && item !== null) {
        const obj = item as Record<string, unknown>;
        return String(obj.displayName ?? obj.name ?? obj.value ?? obj.email ?? obj.id ?? JSON.stringify(obj));
      }
      return String(item);
    }).join(", ");
  }
  if (typeof value === "object") {
    const obj = value as Record<string, unknown>;
    const label = obj.displayName ?? obj.name ?? obj.email ?? obj.value ?? obj.id;
    if (label !== undefined) return String(label);
    const entries = Object.entries(obj).filter(([k]) => !k.startsWith("@odata"));
    if (entries.length <= 3) {
      return entries.map(([k, v]) => `${k}: ${formatCell(v)}`).join(", ");
    }
    return JSON.stringify(obj);
  }
  return String(value);
}

export function TableView({ data, onNavigate }: TableViewProps) {
  const [sortKey, setSortKey] = useState<string | null>(null);
  const [sortAsc, setSortAsc] = useState(true);
  const [filter, setFilter] = useState("");
  const [columnWidths, setColumnWidths] = useState<Record<string, number>>({});
  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);

  const items = Array.isArray(data) ? data : typeof data === "object" && data ? [data] : [];
  if (items.length === 0) return <p>No tabular data available</p>;

  const columns = useMemo(() => {
    const keys = new Set<string>();
    items.forEach((item) => {
      if (typeof item === "object" && item) {
        Object.keys(item as Record<string, unknown>).forEach((k) => {
          if (!k.startsWith("@odata") && !k.startsWith("_")) keys.add(k);
        });
      }
    });
    return Array.from(keys);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [JSON.stringify(items)]);

  const filtered = useMemo(() => {
    if (!filter) return items;
    const lower = filter.toLowerCase();
    return items.filter((item) =>
      Object.values(item as Record<string, unknown>).some((v) =>
        formatCell(v).toLowerCase().includes(lower)
      )
    );
  }, [items, filter]);

  const sorted = useMemo(() => {
    if (!sortKey) return filtered;
    return [...filtered].sort((a, b) => {
      const va = formatCell((a as Record<string, unknown>)[sortKey]);
      const vb = formatCell((b as Record<string, unknown>)[sortKey]);
      return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
    });
  }, [filtered, sortKey, sortAsc]);

  const handleSort = (key: string) => {
    if (sortKey === key) {
      setSortAsc(!sortAsc);
    } else {
      setSortKey(key);
      setSortAsc(true);
    }
  };

  // Column resize handlers
  const onResizeStart = useCallback((col: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startWidth = columnWidths[col] || 150;
    resizingRef.current = { col, startX: e.clientX, startWidth };

    const onMouseMove = (ev: MouseEvent) => {
      if (!resizingRef.current) return;
      const delta = ev.clientX - resizingRef.current.startX;
      const newWidth = Math.max(60, resizingRef.current.startWidth + delta);
      setColumnWidths((prev) => ({ ...prev, [resizingRef.current!.col]: newWidth }));
    };

    const onMouseUp = () => {
      resizingRef.current = null;
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
    };

    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
  }, [columnWidths]);

  const isNavigable = !!onNavigate;

  return (
    <div>
      <Input
        placeholder="Filter..."
        value={filter}
        onChange={(_, d) => setFilter(d.value)}
        style={{ marginBottom: 8, width: 300 }}
      />
      <div style={{ overflow: "auto", maxHeight: "calc(100vh - 160px)" }}>
        <table style={{
          borderCollapse: "collapse",
          fontSize: 12,
          fontFamily: "monospace",
        }}>
          <thead>
            <tr style={{ position: "sticky", top: 0, background: "var(--colorNeutralBackground3)", zIndex: 1 }}>
              {isNavigable && (
                <th style={{ width: 28, padding: "6px 4px", borderBottom: "1px solid var(--colorNeutralStroke1)" }} />
              )}
              {columns.map((col) => (
                <th
                  key={col}
                  onClick={() => handleSort(col)}
                  style={{
                    cursor: "pointer",
                    userSelect: "none",
                    padding: "6px 8px",
                    textAlign: "left",
                    borderBottom: "1px solid var(--colorNeutralStroke1)",
                    whiteSpace: "nowrap",
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    width: columnWidths[col] || 150,
                    minWidth: 60,
                    position: "relative",
                  }}
                  title={col}
                >
                  {col}{" "}
                  {sortKey === col && (sortAsc ? <ArrowUp20Regular /> : <ArrowDown20Regular />)}
                  {/* Resize handle */}
                  <div
                    onMouseDown={(e) => onResizeStart(col, e)}
                    style={{
                      position: "absolute",
                      right: 0,
                      top: 0,
                      bottom: 0,
                      width: 5,
                      cursor: "col-resize",
                      background: "transparent",
                    }}
                    onMouseOver={(e) => { (e.target as HTMLElement).style.background = "var(--colorBrandBackground)"; }}
                    onMouseOut={(e) => { (e.target as HTMLElement).style.background = "transparent"; }}
                  />
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sorted.map((item, i) => (
              <tr
                key={i}
                onClick={isNavigable ? () => onNavigate(item as Record<string, unknown>) : undefined}
                style={{
                  borderBottom: "1px solid var(--colorNeutralStroke2)",
                  cursor: isNavigable ? "pointer" : undefined,
                }}
                onMouseOver={isNavigable ? (e) => { (e.currentTarget as HTMLElement).style.background = "var(--colorNeutralBackground1Hover)"; } : undefined}
                onMouseOut={isNavigable ? (e) => { (e.currentTarget as HTMLElement).style.background = ""; } : undefined}
              >
                {isNavigable && (
                  <td style={{ padding: "4px", color: "var(--colorNeutralForeground3)" }}>
                    <ChevronRight20Regular />
                  </td>
                )}
                {columns.map((col) => {
                  const formatted = formatCell((item as Record<string, unknown>)[col]);
                  return (
                    <td
                      key={col}
                      title={formatted}
                      style={{
                        padding: "4px 8px",
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                        width: columnWidths[col] || 150,
                        maxWidth: columnWidths[col] || 150,
                      }}
                    >
                      {formatted}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p style={{ color: "var(--colorNeutralForeground3)", fontSize: 12, marginTop: 8 }}>
        {sorted.length} of {items.length} items
        {isNavigable && " — click a row to navigate"}
      </p>
    </div>
  );
}
