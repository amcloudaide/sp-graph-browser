import { useState, useMemo } from "react";
import {
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Input,
} from "@fluentui/react-components";
import { ArrowUp20Regular, ArrowDown20Regular } from "@fluentui/react-icons";

interface TableViewProps {
  data: unknown;
}

export function TableView({ data }: TableViewProps) {
  const [sortKey, setSortKey] = useState<string | null>(null);
  const [sortAsc, setSortAsc] = useState(true);
  const [filter, setFilter] = useState("");

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
        String(v ?? "").toLowerCase().includes(lower)
      )
    );
  }, [items, filter]);

  const sorted = useMemo(() => {
    if (!sortKey) return filtered;
    return [...filtered].sort((a, b) => {
      const va = String((a as Record<string, unknown>)[sortKey] ?? "");
      const vb = String((b as Record<string, unknown>)[sortKey] ?? "");
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

  return (
    <div>
      <Input
        placeholder="Filter..."
        value={filter}
        onChange={(_, d) => setFilter(d.value)}
        style={{ marginBottom: 8, width: 300 }}
      />
      <div style={{ overflow: "auto", maxHeight: "calc(100vh - 160px)" }}>
        <Table size="small" aria-label="Data table">
          <TableHeader>
            <TableRow>
              {columns.map((col) => (
                <TableHeaderCell
                  key={col}
                  onClick={() => handleSort(col)}
                  style={{ cursor: "pointer", userSelect: "none" }}
                >
                  {col}{" "}
                  {sortKey === col && (sortAsc ? <ArrowUp20Regular /> : <ArrowDown20Regular />)}
                </TableHeaderCell>
              ))}
            </TableRow>
          </TableHeader>
          <TableBody>
            {sorted.map((item, i) => (
              <TableRow key={i}>
                {columns.map((col) => (
                  <TableCell key={col} style={{ fontFamily: "monospace", fontSize: 12 }}>
                    {String((item as Record<string, unknown>)[col] ?? "")}
                  </TableCell>
                ))}
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>
      <p style={{ color: "var(--colorNeutralForeground3)", fontSize: 12, marginTop: 8 }}>
        {sorted.length} of {items.length} items
      </p>
    </div>
  );
}
