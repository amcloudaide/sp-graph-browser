import type { ReactNode } from "react";
import {
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Badge,
  Link,
} from "@fluentui/react-components";

interface PropertiesViewProps {
  data: unknown;
}

/** Flatten nested objects into dot-notation key-value pairs */
function flattenObject(obj: Record<string, unknown>, prefix = ""): [string, unknown][] {
  const entries: [string, unknown][] = [];
  for (const [key, value] of Object.entries(obj)) {
    const fullKey = prefix ? `${prefix}.${key}` : key;
    if (fullKey.startsWith("@odata") || fullKey.startsWith("_")) continue;
    if (value !== null && typeof value === "object" && !Array.isArray(value)) {
      entries.push(...flattenObject(value as Record<string, unknown>, fullKey));
    } else {
      entries.push([fullKey, value]);
    }
  }
  return entries;
}

function formatValue(key: string, value: unknown): ReactNode {
  if (value === null || value === undefined) return <span style={{ color: "#666" }}>—</span>;
  if (typeof value === "boolean") {
    return <Badge appearance="filled" color={value ? "success" : "danger"}>{String(value)}</Badge>;
  }
  if (typeof value === "string") {
    if (value.startsWith("http://") || value.startsWith("https://")) {
      return <Link href={value} target="_blank">{value}</Link>;
    }
    if (key.toLowerCase().includes("sharing") && value.includes("External")) {
      return <Badge appearance="filled" color="warning">{value}</Badge>;
    }
  }
  if (Array.isArray(value)) {
    if (value.length === 0) return <span style={{ color: "#666" }}>[]</span>;
    // For arrays of simple objects, show a compact summary
    return (
      <span style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
        {value.map((item, i) => {
          if (typeof item === "object" && item !== null) {
            // Show the most useful field (name, displayName, or first string value)
            const obj = item as Record<string, unknown>;
            const label = obj.displayName ?? obj.name ?? obj.value ?? obj.id ?? JSON.stringify(obj);
            return <span key={i}>{i > 0 ? ", " : ""}{String(label)}</span>;
          }
          return <span key={i}>{i > 0 ? ", " : ""}{String(item)}</span>;
        })}
      </span>
    );
  }
  if (typeof value === "object") {
    // Should not happen after flattening, but handle gracefully
    const obj = value as Record<string, unknown>;
    const label = obj.displayName ?? obj.name ?? obj.email ?? obj.id;
    if (label) return String(label);
    return JSON.stringify(value);
  }
  return String(value);
}

export function PropertiesView({ data }: PropertiesViewProps) {
  if (!data || typeof data !== "object") return <p>No data</p>;

  // If data is an array, show count and suggest table view
  if (Array.isArray(data)) {
    return <p>{data.length} items — switch to Table view for details</p>;
  }

  const entries = flattenObject(data as Record<string, unknown>);

  return (
    <Table size="small" aria-label="Properties">
      <TableHeader>
        <TableRow>
          <TableHeaderCell style={{ width: "30%" }}>Property</TableHeaderCell>
          <TableHeaderCell>Value</TableHeaderCell>
        </TableRow>
      </TableHeader>
      <TableBody>
        {entries.map(([key, value]) => (
          <TableRow key={key}>
            <TableCell style={{ color: "var(--colorNeutralForeground3)", fontFamily: "monospace", fontSize: 12 }}>
              {key}
            </TableCell>
            <TableCell style={{ fontFamily: "monospace", fontSize: 12 }}>
              {formatValue(key, value)}
            </TableCell>
          </TableRow>
        ))}
      </TableBody>
    </Table>
  );
}
